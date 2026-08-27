"""Classify Outlook attachments as user-visible files vs hidden/inline parts.

Outlook's Attachments collection includes signature images, cid: HTML inlines,
OLE embeddings, and hidden MAPI parts. This module filters that collection
down to the files the Outlook UI shows on the paperclip.
"""
from __future__ import annotations

from dataclasses import dataclass

# OlAttachmentType
OL_BY_VALUE = 1
OL_BY_REFERENCE = 4
OL_EMBEDDEDITEM = 5
OL_OLE = 6

PR_ATTACH_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACH_CONTENT_ID_ANSI = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
PR_ATTACH_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x37140003"
ATT_MHTML_REF = 4

KIND_FILE = "file"
KIND_EMBEDDED_ITEM = "embedded_item"
KIND_INLINE = "inline"
KIND_HIDDEN = "hidden"
KIND_OLE = "ole"

_VISIBLE_KINDS = {KIND_FILE, KIND_EMBEDDED_ITEM}


@dataclass(frozen=True)
class ClassifiedAttachment:
    com_index: int
    filename: str
    kind: str
    is_visible: bool
    storage_size: int
    content_id: str


def _pa_get(att, prop):
    try:
        return att.PropertyAccessor.GetProperty(prop)
    except Exception:
        return None


def _filename(att) -> str:
    try:
        name = att.FileName or ""
    except Exception:
        name = ""
    if not name:
        try:
            name = att.DisplayName or ""
        except Exception:
            name = ""
    return name or "attachment"


def _storage_size(att) -> int:
    try:
        return int(att.Size or 0)
    except Exception:
        return 0


def _attach_type(att) -> int:
    try:
        return int(att.Type)
    except Exception:
        return OL_BY_VALUE


def _content_id(att) -> str:
    raw = _pa_get(att, PR_ATTACH_CONTENT_ID)
    if raw is None:
        raw = _pa_get(att, PR_ATTACH_CONTENT_ID_ANSI)
    if not raw:
        return ""
    return str(raw).strip().strip("<>").strip()


def _cid_in_html(content_id: str, html_body: str) -> bool:
    if not content_id or not html_body:
        return False
    return f"cid:{content_id.lower()}" in html_body.lower()


def classify_attachment(att, com_index: int, html_body: str = "") -> ClassifiedAttachment:
    """Classify a single Outlook Attachment relative to optional HTML."""
    filename = _filename(att)
    storage_size = _storage_size(att)
    attach_type = _attach_type(att)
    content_id = _content_id(att)

    if attach_type == OL_OLE:
        kind = KIND_OLE
    elif _pa_get(att, PR_ATTACH_HIDDEN):
        kind = KIND_HIDDEN
    else:
        flags = _pa_get(att, PR_ATTACH_FLAGS)
        try:
            mhtml_ref = bool(int(flags) & ATT_MHTML_REF)
        except (TypeError, ValueError):
            mhtml_ref = False
        if mhtml_ref or _cid_in_html(content_id, html_body):
            kind = KIND_INLINE
        elif attach_type == OL_EMBEDDEDITEM:
            kind = KIND_EMBEDDED_ITEM
        else:
            kind = KIND_FILE

    return ClassifiedAttachment(
        com_index=com_index,
        filename=filename,
        kind=kind,
        is_visible=kind in _VISIBLE_KINDS,
        storage_size=storage_size,
        content_id=content_id,
    )


def classify_item_attachments(
    item,
    *,
    html_body: str | None = None,
    match_html_cid: bool = False,
) -> list[ClassifiedAttachment]:
    """Classify every attachment on a MailItem or AppointmentItem."""
    if html_body is None and match_html_cid:
        try:
            html_body = item.HTMLBody or ""
        except Exception:
            html_body = ""
    if html_body is None:
        html_body = ""

    try:
        collection = item.Attachments
        count = collection.Count
    except Exception:
        return []

    results: list[ClassifiedAttachment] = []
    for com_index in range(1, count + 1):
        try:
            att = collection.Item(com_index)
        except Exception:
            continue
        results.append(classify_attachment(att, com_index, html_body))
    return results


def attachment_counts(
    item,
    *,
    match_html_cid: bool = False,
) -> tuple[int, int]:
    """Return (visible_count, raw_count) for an Outlook item."""
    rows = classify_item_attachments(item, match_html_cid=match_html_cid)
    visible = sum(1 for row in rows if row.is_visible)
    return visible, len(rows)


def serialize_attachments(
    rows: list[ClassifiedAttachment],
    *,
    include_hidden: bool = False,
) -> list[dict]:
    """JSON-ready attachment dicts. Default omits non-visible parts.

    `index` matches `visible_index` so save_attachment(attachment_index=)
    stays aligned with list_attachments output.
    """
    visible_n = 0
    payload: list[dict] = []
    for row in rows:
        visible_index = None
        if row.is_visible:
            visible_n += 1
            visible_index = visible_n
        if not include_hidden and not row.is_visible:
            continue
        payload.append({
            "index": visible_index if visible_index is not None else row.com_index,
            "visible_index": visible_index,
            "com_index": row.com_index,
            "filename": row.filename,
            "kind": row.kind,
            "is_visible": row.is_visible,
            "content_id": row.content_id or None,
            "storage_size": row.storage_size,
        })
    return payload


def resolve_attachment_com_index(
    rows: list[ClassifiedAttachment],
    attachment_index: int,
    *,
    use_com_index: bool = False,
) -> int:
    """Map a save_attachment index to Attachments.Item (1-based).

    By default `attachment_index` is the visible paperclip index.
    With use_com_index=True it is the raw COM slot.
    """
    if attachment_index < 1:
        raise ValueError("attachment_index must be >= 1")

    if use_com_index:
        com_indexes = {row.com_index for row in rows}
        if attachment_index not in com_indexes:
            raise ValueError(
                f"No attachment at COM index {attachment_index} "
                f"(raw count: {len(rows)})"
            )
        return attachment_index

    visible = [row for row in rows if row.is_visible]
    if not visible:
        raise ValueError(
            f"No user-visible file attachments "
            f"({len(rows)} hidden/inline/OLE part(s) skipped)"
        )
    if attachment_index > len(visible):
        raise ValueError(
            f"Only {len(visible)} visible attachment(s), "
            f"requested index {attachment_index}"
        )
    return visible[attachment_index - 1].com_index
