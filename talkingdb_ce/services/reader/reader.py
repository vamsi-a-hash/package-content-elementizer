from typing import Callable, Optional

from .docx.docx_reader import DocxReader
from .pdf.pdf_reader import PdfReader
from talkingdb.models.document.document import DocumentModel
from talkingdb.models.failure.failure import DocumentFailure
from talkingdb.models.failure.reason import FailureReason


class ReaderFactory:
    """
    Factory class to pick the right reader based on file type.
    """

    @staticmethod
    def get_reader(file_type: str):
        readers = {
            "docx": DocxReader(),
            "pdf": PdfReader(),
        }
        return readers.get(file_type.lower())


def parse_document(
    io_buffer,
    file_type: str,
    file_name: str,
    cancel_check: Optional[Callable[[], bool]] = None,
    checkpoint_dir: Optional[str] = None,
    *,
    bake_page_breaks: bool = True,
    baked_docx_path: Optional[str] = None,
    channel: Optional[str] = None,
    side_effects: Optional[dict] = None,
) -> DocumentModel:
    reader = ReaderFactory.get_reader(file_type)
    if not reader:
        raise DocumentFailure(
            FailureReason.UNSUPPORTED_FILE_TYPE,
            detail=f"unsupported file type: {file_type or '(none)'}",
        )
    if checkpoint_dir and (file_type or "").lower() == "pdf":
        return reader.read_document(
            io_buffer, file_name,
            cancel_check=cancel_check, checkpoint_dir=checkpoint_dir,
        )
    if (file_type or "").lower() == "docx":
        document = reader.read_document(
            io_buffer, file_name, cancel_check=cancel_check,
            bake_page_breaks=bake_page_breaks,
            baked_docx_path=baked_docx_path,
            channel=channel,
        )
        if side_effects is not None:
            stored_hash = getattr(reader, "stored_file_hash", None)
            if stored_hash:
                side_effects["stored_file_hash"] = stored_hash
        return document
    return reader.read_document(io_buffer, file_name, cancel_check=cancel_check)