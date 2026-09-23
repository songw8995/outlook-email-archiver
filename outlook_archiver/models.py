from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import datetime
from typing import Any


@dataclass(slots=True)
class FolderInfo:
    store_id: str
    store_name: str
    store_path: str
    folder_id: str
    folder_name: str
    folder_path: str
    segments: list[str]
    parent_id: str | None
    kind: str = "custom"
    selectable: bool = True
    warning: str = ""
    segment_ids: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class Candidate:
    store_id: str
    store_name: str
    store_path: str
    folder_id: str
    folder_path: str
    folder_segments: list[str]
    folder_kind: str
    entry_id: str
    effective_at: datetime | None
    effective_basis: str
    item_class: int
    last_modified: datetime | None = None
    size: int = 0
    segment_ids: list[str] = field(default_factory=list)
    source_store_key: str = ""
    archive_store_id: str = ""
    subject: str = ""


@dataclass(slots=True)
class AttachmentRecord:
    index: int
    original_name: str
    actual_name: str
    relative_path: str
    content_id: str = ""
    content_location: str = ""
    mime_type: str = ""
    hidden: bool = False
    inline: bool = False
    size: int = 0
    sha256: str = ""
    status: str = "saved"
    error: str = ""
    attachment_type: int = 1

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class MessageSnapshot:
    source_key: str
    content_signature: str
    store_id: str
    store_name: str
    store_path: str
    folder_id: str
    folder_path: str
    folder_segments: list[str]
    folder_kind: str
    entry_id: str
    internet_message_id: str
    conversation_id: str
    subject: str
    sender_name: str
    sender_email: str
    to: list[str]
    cc: list[str]
    bcc: list[str]
    sent_at: datetime | None
    received_at: datetime | None
    created_at: datetime | None
    modified_at: datetime | None
    effective_at: datetime | None
    effective_basis: str
    direction: str
    html_body: str
    text_body: str
    segment_ids: list[str] = field(default_factory=list)
    attachments: list[AttachmentRecord] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
