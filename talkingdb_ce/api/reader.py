import os
import io
import re
import hashlib
from typing import Callable, Optional
from fastapi import APIRouter, UploadFile, File, Form
from talkingdb.models.metadata.metadata import Metadata, DEFAULT_METADATA
from talkingdb.helpers.metadata import update_metadata
from talkingdb.helpers.event import create_event, update_event
from talkingdb.helpers.request import set_request_context, clear_request_context
from talkingdb.logger.context import set_log_context, clear_log_context
from talkingdb.helpers.dataclass import to_json
from talkingdb.models.event.status import EventStatus
from talkingdb.models.event.type import EventType
from talkingdb.logger.track import track
from talkingdb_ce.services.reader.reader import parse_document
from talkingdb_ce.model.api.reader import RequestModel

router = APIRouter(prefix="/content", tags=["Elementizer"])

CHANNEL_PATTERN = re.compile(r"^(localhost|ttt-rc[1-9][0-9]*)$")

async def run_parse(
    document_file: UploadFile,
    metadata: Optional[str],
    cancel_check: Optional[Callable[[], bool]] = None,
    checkpoint_dir: Optional[str] = None,
):
    """Parse an uploaded document, supporting HTTP and direct in-process calls
    with cancel_check for the latter.
    """
    request = RequestModel(
        document_file=document_file,
        metadata=metadata
    )

    meta = Metadata.from_json(metadata)
    meta = Metadata.ensure_metadata(meta)
    event = create_event(meta, EventType.CONTENT_ELEMENT)
    meta = update_metadata(meta, event)
    _meta = meta.model_dump(mode="json")

    set_request_context(**_meta)
    set_log_context(service="talkingdb_ce",
                    function="parse_file",
                    **_meta)

    try:
        file_bytes = await request.document_file.read()
        file_name = request.document_file.filename
        io_buffer = io.BytesIO(file_bytes)
        document_path = file_name

        event_data = {
            "document_path": document_path,
            "file_name": file_name,
            "type": _meta.get("type", "unknown")
        }
        event = update_event(event, EventStatus.ONGOING, event_data)

        _, ext = os.path.splitext(file_name)
        file_type = ext.lstrip(".").lower()

        channel = getattr(meta, "channel", None)
        if not channel or not CHANNEL_PATTERN.match(channel):
            channel = None
        file_hash = hashlib.sha256(file_bytes).hexdigest()

        @track()
        def _parse_document():
            return parse_document(
                io_buffer, file_type, file_name,
                cancel_check=cancel_check, checkpoint_dir=checkpoint_dir,
                channel=channel, file_hash=file_hash,
            )

        document, new_file_hash = _parse_document()
        file_index = document.build_index()

        update_event(event, EventStatus.COMPLETED)

        @track(log_response=True, response_key=file_name)
        def _parsed_document():
            return {
                "document": to_json(document),
                "file_index": file_index.model_dump(mode="json")
            }

        return _parsed_document()
    finally:
        clear_request_context()
        clear_log_context()


@router.post("/parse")
async def parse_file(
        document_file: UploadFile = File(None),
        metadata: Optional[str] = Form(DEFAULT_METADATA)
):
    return await run_parse(document_file, metadata)