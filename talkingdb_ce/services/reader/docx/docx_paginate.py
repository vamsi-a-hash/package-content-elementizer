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
) -> Optional[str]:
    """Best-effort DOCX pagination: render to PDF, extract page text, and set elem.page.

    Pagination failures leave page=None. ReadCancelled propagates so
    cancellation is not treated as a pagination failure.

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
        return f"pagination failed: {exc}"

    if channel and file_hash:
        return _bake_and_store_page_breaks(docx_bytes, page_texts, channel, file_hash)

    return None


def _render_to_pdf(
    docx_bytes: bytes,
    timeout_seconds: int,
    cancel_check: Optional[Callable[[], bool]] = None,
) -> bytes:
    with tempfile.TemporaryDirectory(prefix="tdb-docx-paginate-") as tmp_dir:
        docx_path = os.path.join(tmp_dir, "input.docx")
        try:
            with open(docx_path, "wb") as fh:
                fh.write(docx_bytes)
            # handle is now closed — soffice can read the file cleanly
        except Exception as exc:
            raise PaginationError(f"failed to write docx to temp file: {exc}") from exc

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


def _append_break_before(ct_element) -> None:
    prev = ct_element.getprevious()
    if prev is not None and isinstance(prev, CT_P):
        run = OxmlElement("w:r")
        br = OxmlElement("w:br")
        br.set(qn("w:type"), "page")
        run.append(br)
        prev.append(run)
        return
    if prev is None:
        return  # first block in the body - nothing to break before
    new_p = OxmlElement("w:p")
    run = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run.append(br)
    new_p.append(run)
    ct_element.addprevious(new_p)


def _add_page_breaks(docx_bytes: bytes, page_texts: list[str]) -> bytes:
    doc = Document(io.BytesIO(docx_bytes))
    targets = _page_break_targets(doc, page_texts)

    for elem, gap in targets:
        already_broken = _has_break_immediately_before(elem)
        needed = gap - 1 if already_broken else gap
        if needed <= 0:
            continue

        if isinstance(elem, CT_P):
            _set_page_break_before(elem, needed)
        else:
            for _ in range(needed):
                _append_break_before(elem)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

def _set_page_break_before(elem: CT_P, needed: int) -> None:
    pPr = elem.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        elem.insert(0, pPr)
    if pPr.find(qn("w:pageBreakBefore")) is None:
        pPr.append(OxmlElement("w:pageBreakBefore"))
        needed -= 1
    for _ in range(needed):
        run = OxmlElement("w:r")
        br = OxmlElement("w:br")
        br.set(qn("w:type"), "page")
        run.append(br)
        elem.insert(0, run)


def _bake_and_store_page_breaks(
    docx_bytes: bytes, page_texts: list[str], channel: str, file_hash: str,
) -> Optional[str]:
    if not file_store.is_configured():
        return None
    try:
        paginated_bytes = _add_page_breaks(docx_bytes, page_texts)
        with tempfile.NamedTemporaryFile(suffix=".docx") as tmp:
            tmp.write(paginated_bytes)
            tmp.flush()
            return file_store.overwrite_file(channel, file_hash, tmp.name)
    except Exception as exc:
        logger.warning(f"docx page-break baking/storage skipped: {exc}")
        return None
