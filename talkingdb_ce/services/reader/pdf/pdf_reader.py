import io
import os
import json
import shutil
import subprocess
import sys
import tempfile
from collections import Counter
from typing import Callable, List, Optional, Tuple

import fitz

from talkingdb.logger.console import logger
from talkingdb.models.document.document import DocumentModel
from talkingdb.models.failure.failure import DocumentFailure
from talkingdb.models.failure.reason import FailureReason
from talkingdb.models.document.elements.primitive.paragraph import (
    ParagraphModel,
    ParagraphStyleModel,
)

from ..docx.docx_reader import DocxReader
from ..killable_subprocess import run_killable


CONVERT_TIMEOUT_SECONDS = int(
    os.getenv("CE_PDF_CONVERT_TIMEOUT_SECONDS", "3600"))

_CONVERT_MAX_MEMORY_MB = int(
    os.getenv("CE_PDF_CONVERT_MAX_MEMORY_MB", "0"))
CONVERT_MAX_MEMORY_BYTES = (
    _CONVERT_MAX_MEMORY_MB * 1024 * 1024 if _CONVERT_MAX_MEMORY_MB > 0 else None
)

MIN_EXTRACTABLE_TEXT_CHARS = int(os.getenv("CE_PDF_MIN_TEXT_CHARS", "16"))

MAX_HEADING_LEVEL = 6

HEADING_MAX_CHARS = 200
BOLD_HEADING_MAX_CHARS = 120

_CONVERT_MODULE = "talkingdb_ce.services.reader.pdf._convert_main"


def _round_size(size: Optional[float]) -> Optional[int]:
    """Round a font size to the nearest point for tier clustering."""
    return round(size) if size else None


class PdfReader:
    """Reads a PDF by converting it to DOCX and reusing :class:`DocxReader`."""

    def __init__(self) -> None:
        self.docx_reader = DocxReader()

    # --------------------------------------------------------------- public API
    def read_document(
        self,
        io_buffer,
        file_name,
        cancel_check: Optional[Callable[[], bool]] = None,
        checkpoint_dir: Optional[str] = None,
    ) -> DocumentModel:
        """Read a PDF by first converting it to DOCX. Optionally provide a persistent
          checkpoint directory to resume failed conversions across retries; otherwise,
          a temporary directory is created and cleaned up automatically.
        """
        self._reject_if_password_protected(io_buffer, file_name)

        docx_bytes, page_numbers = self._to_docx_bytes(
            io_buffer, cancel_check=cancel_check, checkpoint_dir=checkpoint_dir
        )

        model, _ = self.docx_reader.read_document(
            io.BytesIO(docx_bytes), file_name, paginate=False)

        self._reject_if_textless(model, file_name)
        self._apply_page_numbers(model, page_numbers)
        self._remap_headings(model)
        model.build_hierarchy()
        return model

    # --------------------------------------------------------------- preflight
    def _reject_if_password_protected(self, io_buffer, file_name) -> None:
        io_buffer.seek(0)
        pdf_data = io_buffer.read()
        io_buffer.seek(0)

        try:
            document = fitz.open(stream=pdf_data, filetype="pdf")
        except Exception as exc:
            logger.warning(f"pdf preflight could not open '{file_name}': {exc}")
            return

        try:
            needs_password = bool(document.needs_pass)
        finally:
            document.close()

        if needs_password:
            raise DocumentFailure(
                FailureReason.PASSWORD_PROTECTED,
                detail=f"PDF '{file_name}' requires a password to open",
            )

    # ----------------------------------------------------------------- convert
    def _to_docx_bytes(
        self,
        io_buffer,
        cancel_check: Optional[Callable[[], bool]] = None,
        checkpoint_dir: Optional[str] = None,
    ) -> Tuple[bytes, List[int]]:
        io_buffer.seek(0)
        pdf_data = io_buffer.read()

        pdf_path: Optional[str] = None
        docx_path: Optional[str] = None
        pages_path: Optional[str] = None
        owns_checkpoint_dir = checkpoint_dir is None
        if owns_checkpoint_dir:
            checkpoint_dir = tempfile.mkdtemp(prefix="tdb-pdf-ckpt-")
        try:
            with tempfile.NamedTemporaryFile(
                prefix="tdb-pdf-", suffix=".pdf", delete=False
            ) as tmp_pdf:
                tmp_pdf.write(pdf_data)
                pdf_path = tmp_pdf.name

            with tempfile.NamedTemporaryFile(
                prefix="tdb-pdf-", suffix=".docx", delete=False
            ) as tmp_docx:
                docx_path = tmp_docx.name

            with tempfile.NamedTemporaryFile(
                prefix="tdb-pdf-", suffix=".pages.json", delete=False
            ) as tmp_pages:
                pages_path = tmp_pages.name

            page_numbers = self._convert(
                pdf_path, docx_path, pages_path, checkpoint_dir,
                cancel_check=cancel_check,
            )

            with open(docx_path, "rb") as fh:
                return fh.read(), page_numbers
        finally:
            for path in (pdf_path, docx_path, pages_path):
                if path and os.path.exists(path):
                    try:
                        os.remove(path)
                    except OSError:
                        logger.warning(f"failed to remove temp file: {path}")
            if owns_checkpoint_dir:
                shutil.rmtree(checkpoint_dir, ignore_errors=True)

    def _convert(
        self,
        pdf_path: str,
        docx_path: str,
        pages_path: str,
        checkpoint_dir: str,
        cancel_check: Optional[Callable[[], bool]] = None,
    ) -> List[int]:
        """Run pdf2docx in a killable child process with a wall-clock cap.

        Process pages in batches and checkpoint each batch. If interrupted,
        a later attempt resumes from the last completed batch instead of page 0.

        Returns the ordered list of 1-based PDF page numbers, one per docx
        section pdf2docx produced (it starts a new section per page).
        """
        try:
            returncode, stdout, stderr = run_killable(
                [sys.executable, "-m", _CONVERT_MODULE,
                    pdf_path, docx_path, pages_path, checkpoint_dir],
                timeout_seconds=CONVERT_TIMEOUT_SECONDS,
                cancel_check=cancel_check,
                max_memory_bytes=CONVERT_MAX_MEMORY_BYTES,
            )
        except subprocess.TimeoutExpired:
            raise ValueError(
                f"PDF conversion exceeded {CONVERT_TIMEOUT_SECONDS}s and was aborted"
            )

        if returncode != 0:
            detail = (stderr or stdout or "").strip()
            detail = detail or f"converter exited with code {returncode}"
            raise ValueError(f"PDF could not be converted ({detail})")

        try:
            with open(pages_path, "r") as fh:
                page_numbers = json.load(fh)
        except (OSError, ValueError) as exc:
            logger.warning(
                f"could not read page map, page numbers will be unset: {exc}")
            return []

        return page_numbers if isinstance(page_numbers, list) else []

    def _apply_page_numbers(self, model: DocumentModel, page_numbers: List[int]) -> None:
        """Assign page numbers using the section->page mapping from pdf2docx.

        pdf2docx emits one docx section per source PDF page, and DocxReader
        already splits the document into one LayoutModel per section in order,
        so model.layouts[i] corresponds to page_numbers[i].
        """
        if not page_numbers:
            return

        if len(model.layouts) != len(page_numbers):
            logger.warning(
                f"Layout/page-number mismatch: {len(model.layouts)} layouts "
                f"but {len(page_numbers)} page numbers. "
                f"This can happen with multi-column PDFs. "
                f"Layouts beyond index {len(page_numbers) - 1} will have page=None."
            )

        for layout_idx, layout in enumerate(model.layouts):
            if layout_idx >= len(page_numbers):
                break
            page_no = page_numbers[layout_idx]
            for elem in layout.elements:
                elem.page = page_no

    # ------------------------------------------------------------- text guard
    def _reject_if_textless(self, model: DocumentModel, file_name) -> None:
        total = 0
        for elem in model.iter_elements():
            text = elem.to_text() if hasattr(elem, "to_text") else ""
            if text:
                total += len(text.strip())
            if total >= MIN_EXTRACTABLE_TEXT_CHARS:
                return
        raise ValueError(
            f"No extractable text found in PDF '{file_name}' "
            f"(possibly a scanned/image-only document; OCR is not supported)"
        )

    # --------------------------------------------------------- heading remap
    def _remap_headings(self, model: DocumentModel) -> None:
        paragraphs = [
            elem
            for elem in model.iter_elements()
            if isinstance(elem, ParagraphModel)
        ]

        for para in paragraphs:
            if para.style is None:
                para.style = ParagraphStyleModel(name="Normal")

        sized: List[Tuple[ParagraphModel, str, Optional[int], bool]] = []
        size_weight: Counter = Counter()
        for para in paragraphs:
            text = para.to_text().strip()
            if not text:
                continue
            size = self._para_size(para)
            sized.append((para, text, size, self._para_bold(para)))
            if size is not None:
                size_weight[size] += len(text)

        if not size_weight:
            return  # no font metrics to reason about; leave as a flat document

        body_size = size_weight.most_common(1)[0][0]
        tiers = sorted({s for s in size_weight if s > body_size}, reverse=True)
        size_level = {s: min(i + 1, MAX_HEADING_LEVEL)
                      for i, s in enumerate(tiers)}
        bold_level = min(len(tiers) + 1, MAX_HEADING_LEVEL)

        for para, text, size, bold in sized:
            level: Optional[int] = None
            if size in size_level and len(text) <= HEADING_MAX_CHARS:
                level = size_level[size]
            elif (
                bold
                and len(text) <= BOLD_HEADING_MAX_CHARS
                and (size is None or size <= body_size)
            ):
                level = bold_level
            if level:
                para.style.name = f"Heading {level}"

    @staticmethod
    def _para_size(para: ParagraphModel) -> Optional[int]:
        sizes = [
            r.attributes.font_size
            for r in para.runs
            if r.text and r.text.strip() and r.attributes and r.attributes.font_size
        ]
        if sizes:
            return _round_size(max(sizes))
        if para.style and para.style.font_size:
            return _round_size(para.style.font_size)
        return None

    @staticmethod
    def _para_bold(para: ParagraphModel) -> bool:
        runs = [r for r in para.runs if r.text and r.text.strip()]
        if not runs:
            return False
        return all(r.attributes and r.attributes.bold for r in runs)
