import io
import os
import re
import subprocess
import tempfile
from typing import Optional
from collections.abc import Callable
import shutil

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

from talkingdb.helpers import file_store
from talkingdb.logger.console import logger
from talkingdb.models.document.document import DocumentModel

from ..killable_subprocess import ReadCancelled, run_killable

# Toggle without a redeploy if soffice is missing/misbehaving in an env.
PAGINATE_DOCX_ENABLED = os.getenv(
    "CE_DOCX_PAGINATE", "1") not in ("0", "false", "False")

CONVERT_TIMEOUT_SECONDS = int(
    os.getenv("CE_DOCX_PAGINATE_TIMEOUT_SECONDS", "1800"))

_ANCHOR_LEN = 40
_MIN_ANCHOR_LEN = 8

if PAGINATE_DOCX_ENABLED and shutil.which("soffice") is None:
    raise EnvironmentError(
        "CE_DOCX_PAGINATE is enabled but soffice is not on PATH. "
        "Either install libreoffice-writer or set CE_DOCX_PAGINATE=0."
    )


class PaginationError(Exception):
    pass


def paginate_docx(
    docx_bytes: bytes,
    model: DocumentModel,
    cancel_check: Optional[Callable[[], bool]] = None,
    channel: Optional[str] = None,
    file_hash: Optional[str] = None,
) -> None:
    """Best-effort DOCX pagination: render to PDF, extract page text, and set elem.page.

    Pagination failures are logged and leave page=None. ReadCancelled propagates
    so cancellation is not treated as a pagination failure.

    When `channel` and `file_hash` are given (the MinIO coordinates of the
    docx module-ttt already uploaded at request time), this also bakes
    matching page breaks into the docx and overwrites that SAME object with
    the result. This is a pure overwrite of an already-safely-stored file -
    every step here is best-effort and failures are swallowed (logged only):
    on any failure the original, already-uploaded docx is simply left as-is.
    """
    try:
        pdf_bytes = _render_to_pdf(
            docx_bytes, CONVERT_TIMEOUT_SECONDS, cancel_check=cancel_check
        )
        page_texts = _extract_page_texts(pdf_bytes)
        _assign_pages(model, page_texts)
    except ReadCancelled:
        raise
    except Exception as exc:
        logger.warning(f"docx pagination skipped: {exc}")
        return None

    if channel and file_hash:
        _bake_and_store_page_breaks(docx_bytes, page_texts, channel, file_hash)


def _render_to_pdf(
    docx_bytes: bytes,
    timeout_seconds: int,
    cancel_check: Optional[Callable[[], bool]] = None,
) -> bytes:
    with tempfile.TemporaryDirectory(prefix="tdb-docx-paginate-") as tmp_dir:
        docx_path = os.path.join(tmp_dir, "input.docx")
        with open(docx_path, "wb") as fh:
            fh.write(docx_bytes)
        # handle is now closed — soffice can read the file cleanly

        profile_uri = f"file://{tmp_dir}/lo_profile"

        try:
            returncode, stdout, stderr = run_killable(
                [
                    "soffice",
                    "--headless",
                    "--norestore",
                    f"-env:UserInstallation={profile_uri}",
                    "--convert-to", "pdf",
                    "--outdir", tmp_dir,
                    docx_path,
                ],
                timeout_seconds=timeout_seconds,
                cancel_check=cancel_check,
            )
        except subprocess.TimeoutExpired:
            raise PaginationError(
                f"docx->pdf render exceeded {timeout_seconds}s"
            )
        except FileNotFoundError:
            raise PaginationError("soffice binary not found on PATH")

        if returncode != 0:
            detail = (stderr or stdout or "").strip()
            raise PaginationError(
                f"soffice exited with code {returncode}: {detail}"
            )

        # Step 3: read the pdf soffice wrote — separate handle, read mode
        pdf_path = os.path.join(tmp_dir, "input.pdf")
        if not os.path.exists(pdf_path):
            raise PaginationError("soffice did not produce a pdf output")

        with open(pdf_path, "rb") as fh:
            return fh.read()
    # tmp_dir and all its contents are deleted here after the bytes are returned


def _extract_page_texts(pdf_bytes: bytes) -> list[str]:
    import fitz  # PyMuPDF — transitive dep via pdf2docx

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        return [page.get_text() for page in doc]
    finally:
        doc.close()


def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def _assign_pages(model: DocumentModel, page_texts: list[str]) -> None:
    """Walk elements in document order, advancing a page cursor forward-only
    by anchor-matching each element's text against the rendered page text.

    Forward-only + cursor-after-match avoids false matches on repeated
    phrases (e.g. running headers). Best effort: if an anchor can't be
    found ahead, the element keeps the current page rather than failing.
    """
    norm_pages = [_normalize(t) for t in page_texts]
    if not norm_pages:
        return

    page_idx = 0
    cursor = 0  # consumed offset within norm_pages[page_idx]

    for elem in model.iter_elements():
        text = elem.to_text() if hasattr(elem, "to_text") else ""
        norm = _normalize(text)

        if not norm:
            elem.page = page_idx + 1
            continue

        anchor = norm[:_ANCHOR_LEN] if len(norm) >= _MIN_ANCHOR_LEN else norm

        matched = False
        search_idx = page_idx
        while search_idx < len(norm_pages):
            haystack = norm_pages[search_idx]
            start = cursor if search_idx == page_idx else 0
            pos = haystack.find(anchor, start)
            if pos != -1:
                page_idx = search_idx
                cursor = pos + len(anchor)
                matched = True
                break
            search_idx += 1

        elem.page = page_idx + 1

        if not matched:
            logger.debug(
                f"docx pagination: no anchor match for element "
                f"{getattr(elem, 'id', None)}; kept page {page_idx + 1}"
            )


def _iter_body_blocks(doc: "Document"):
    """Yield (text, oxml_element) for every top-level body block - paragraph
    OR table - in document order. Surfacing tables (unlike doc.paragraphs)
    is what lets page breaks land correctly when a *table* starts a page.
    """
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            p = Paragraph(child, doc)
            yield p.text, child
        elif isinstance(child, CT_Tbl):
            t = Table(child, doc)
            cell_text = " ".join(
                cell.text for row in t.rows for cell in row.cells if cell.text
            )
            yield cell_text, child


def _page_break_targets(doc: "Document", page_texts: list[str]) -> list[tuple[object, int]]:
    """Anchor-match every top-level body block (paragraph or table) against
    rendered PDF pages. Returns (elem, gap) pairs: `elem` is the oxml
    element a break should be inserted before, and `gap` is how many page
    boundaries were crossed to get there (normally 1; >1 when anchor
    matching jumps over pages with no extractable text.
    """
    norm_pages = [_normalize(t) for t in page_texts]
    if not norm_pages:
        return []

    page_idx = 0
    cursor = 0
    prev_page = 0
    seen_first = False
    targets: list[tuple[object, int]] = []

    for text, elem in _iter_body_blocks(doc):
        norm = _normalize(text)
        if not norm:
            continue  # no anchor to match on; leave cursor where it is

        anchor = norm[:_ANCHOR_LEN] if len(norm) >= _MIN_ANCHOR_LEN else norm

        search_idx = page_idx
        while search_idx < len(norm_pages):
            haystack = norm_pages[search_idx]
            start = cursor if search_idx == page_idx else 0
            pos = haystack.find(anchor, start)
            if pos != -1:
                page_idx = search_idx
                cursor = pos + len(anchor)
                break
            search_idx += 1
        # no match anywhere ahead -> best effort, keep current page/cursor

        if seen_first and page_idx > prev_page:
            targets.append((elem, page_idx - prev_page))
        prev_page = page_idx
        seen_first = True

    return targets


def _has_break_immediately_before(elem) -> bool:
    if isinstance(elem, CT_P):
        pPr = elem.find(qn("w:pPr"))
        if pPr is not None and pPr.find(qn("w:pageBreakBefore")) is not None:
            return True
        for br in elem.iter(qn("w:br")):
            if br.get(qn("w:type")) == "page":
                return True

    prev = elem.getprevious()
    if prev is not None:
        for br in prev.iter(qn("w:br")):
            if br.get(qn("w:type")) == "page":
                return True

    return False


_BAKED_DOCVAR_NAME = "CE_DOCX_PAGE_BREAKS_BAKED"
_BAKED_DOCVAR_VALUE = "1"


def _is_page_breaks_baked(doc: "Document") -> bool:
    """Return True when this DOCX has already had page breaks baked into it.

    The marker lives in w:settings/w:docVars, which python-docx preserves when
    loading/saving the document. This makes baking idempotent across retries
    without introducing a separate custom XML part.
    """
    settings = doc.settings.element
    doc_vars = settings.find(qn("w:docVars"))
    if doc_vars is None:
        return False

    for doc_var in doc_vars.findall(qn("w:docVar")):
        if (
            doc_var.get(qn("w:name")) == _BAKED_DOCVAR_NAME
            and doc_var.get(qn("w:val")) == _BAKED_DOCVAR_VALUE
        ):
            return True
    return False


def _mark_page_breaks_baked(doc: "Document") -> None:
    """Stamp the document so a later retry will not bake it again."""
    settings = doc.settings.element
    doc_vars = settings.find(qn("w:docVars"))
    if doc_vars is None:
        doc_vars = OxmlElement("w:docVars")
        settings.append(doc_vars)

    for doc_var in doc_vars.findall(qn("w:docVar")):
        if doc_var.get(qn("w:name")) == _BAKED_DOCVAR_NAME:
            doc_var.set(qn("w:val"), _BAKED_DOCVAR_VALUE)
            return

    doc_var = OxmlElement("w:docVar")
    doc_var.set(qn("w:name"), _BAKED_DOCVAR_NAME)
    doc_var.set(qn("w:val"), _BAKED_DOCVAR_VALUE)
    doc_vars.append(doc_var)


def _append_page_break_to_previous_paragraph(ct_element) -> bool:
    """Append a page break to the nearest preceding top-level paragraph.

    This avoids creating a new paragraph mark, which can itself occupy space
    and, for table targets, can turn stacked breaks into a genuinely blank page.
    Returns True when a preceding paragraph was found and updated.
    """
    prev = ct_element.getprevious()
    while prev is not None:
        if isinstance(prev, CT_P):
            run = OxmlElement("w:r")
            br = OxmlElement("w:br")
            br.set(qn("w:type"), "page")
            run.append(br)
            prev.append(run)
            return True
        prev = prev.getprevious()

    return False


def _set_page_break_before(ct_p) -> None:
    """Mark a paragraph with w:pageBreakBefore - a formatting flag, not
    content. Idempotent (checked before calling) and, unlike inserting a
    new paragraph, it can't push anything onto an extra page or invalidate
    page numbers already computed.
    """
    pPr = ct_p.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        ct_p.insert(0, pPr)
    if pPr.find(qn("w:pageBreakBefore")) is None:
        pPr.append(OxmlElement("w:pageBreakBefore"))


def _insert_page_break_paragraph_before(ct_element) -> None:
    """Insert a new paragraph containing a hard page break directly before
    `ct_element` (a CT_P or CT_Tbl), via raw oxml insertion.
    """
    new_p = OxmlElement("w:p")
    run = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run.append(br)
    new_p.append(run)
    ct_element.addprevious(new_p)


def _add_page_breaks(docx_bytes: bytes, page_texts: list[str]) -> bytes:
    doc = Document(io.BytesIO(docx_bytes))

    # Baking must be idempotent. A retry may receive the object that was already
    # overwritten by a previous successful bake.
    if _is_page_breaks_baked(doc):
        return docx_bytes

    targets = _page_break_targets(doc, page_texts)

    for elem, gap in targets:
        already_broken = _has_break_immediately_before(elem)
        needed = gap - 1 if already_broken else gap
        if needed <= 0:
            continue  # LibreOffice's own break already accounts for this boundary

        if isinstance(elem, CT_P) and needed == 1:
            _set_page_break_before(elem)
        else:
            # Tables cannot carry w:pageBreakBefore. Prefer putting the break
            # at the end of the existing preceding paragraph so we do not add
            # a new paragraph mark that can consume space on a page.
            attached = False
            for _ in range(needed):
                if not _append_page_break_to_previous_paragraph(elem):
                    _insert_page_break_paragraph_before(elem)
                else:
                    attached = True

            if attached:
                continue

    # Stamp only the successfully constructed document. If save fails, the
    # original MinIO object remains untouched and an unbaked document is not
    # marked as baked.
    _mark_page_breaks_baked(doc)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def _bake_and_store_page_breaks(
    docx_bytes: bytes,
    page_texts: list[str],
    channel: str,
    file_hash: str,
) -> None:
    """Insert explicit page breaks at LibreOffice's page boundaries and
    overwrite the already-uploaded docx object in MinIO with the result, at
    the SAME (channel, file_hash) key the original upload used.
    """
    if not file_store.is_configured():
        return

    try:
        # _add_page_breaks performs the authoritative idempotency check. Keeping
        # the check there is important because this function may receive bytes
        # that were downloaded from MinIO after a previous successful bake.
        paginated_bytes = _add_page_breaks(docx_bytes, page_texts)
        with tempfile.NamedTemporaryFile(suffix=".docx") as tmp:
            tmp.write(paginated_bytes)
            tmp.flush()
            file_store.overwrite_file(channel, file_hash, tmp.name)
    except Exception as exc:
        logger.warning(f"docx page-break baking/storage skipped: {exc}")
