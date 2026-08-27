"""Helpers for extracting and formatting Outlook item data."""
import re

from outlook_desktop_mcp.tools._folder_constants import (
    BUSY_STATUS_NAMES,
    MEETING_STATUS_NAMES,
    RESPONSE_NAMES,
    TASK_STATUS_NAMES,
    IMPORTANCE_NAMES,
)
from outlook_desktop_mcp.utils.attachments import (
    attachment_counts,
    classify_item_attachments,
    serialize_attachments,
)
from outlook_desktop_mcp.utils.tasks import outlook_date_or_none


def truncate(text: str, max_length: int = 2000) -> str:
    if len(text) <= max_length:
        return text
    return text[:max_length] + "\n... [truncated]"


def strip_html(html: str) -> str:
    text = re.sub(r"<[^>]+>", "", html)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def get_item_categories(item) -> list[str]:
    """Return an Outlook item's categories as a safely parsed list."""
    try:
        raw = getattr(item, "Categories", "") or ""
    except Exception:
        return []

    try:
        return [value.strip() for value in raw.split(",") if value.strip()]
    except Exception:
        return []


def merge_categories(
    current: list[str],
    requested: list[str],
    mode: str,
) -> list[str]:
    """Apply a category update mode without losing unrelated categories."""
    if mode == "replace":
        return list(requested)

    if mode == "add":
        updated = list(current)
        seen = {value.casefold() for value in updated}
        for value in requested:
            if value.casefold() not in seen:
                updated.append(value)
                seen.add(value.casefold())
        return updated

    if mode == "remove":
        remove_set = {value.casefold() for value in requested}
        return [
            value
            for value in current
            if value.casefold() not in remove_set
        ]

    raise ValueError("mode must be one of: replace, add, remove")


def format_email_summary(item) -> dict:
    """Extract key fields from an Outlook MailItem into a dict.

    attachment_count is the paperclip file count (signature/inline/OLE
    parts excluded). raw_attachment_count is the unfiltered COM collection.
    """
    visible_count, raw_count = attachment_counts(item, match_html_cid=False)
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "sender": getattr(item, "SenderEmailAddress", "unknown"),
        "sender_name": getattr(item, "SenderName", "unknown"),
        "received_time": str(item.ReceivedTime),
        "unread": bool(item.UnRead),
        "has_attachments": visible_count > 0,
        "attachment_count": visible_count,
        "raw_attachment_count": raw_count,
        "categories": get_item_categories(item),
    }


def format_email_full(item, body_max_length: int = 5000) -> dict:
    """Extract full email details including body and visible attachments."""
    result = format_email_summary(item)
    result["to"] = item.To or ""
    result["cc"] = item.CC or ""
    result["body"] = truncate(item.Body or "", body_max_length)
    rows = classify_item_attachments(item, match_html_cid=True)
    visible = [row for row in rows if row.is_visible]
    result["has_attachments"] = len(visible) > 0
    result["attachment_count"] = len(visible)
    result["raw_attachment_count"] = len(rows)
    result["attachments"] = serialize_attachments(rows)
    return result


# --- Calendar formatting ---


def format_event_summary(item) -> dict:
    """Extract key fields from an Outlook AppointmentItem."""
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "start": str(item.Start),
        "end": str(item.End),
        "duration": item.Duration,
        "location": item.Location or "",
        "organizer": item.Organizer or "",
        "is_recurring": bool(item.IsRecurring),
        "all_day": bool(item.AllDayEvent),
        "busy_status": BUSY_STATUS_NAMES.get(item.BusyStatus, "unknown"),
        "meeting_status": MEETING_STATUS_NAMES.get(item.MeetingStatus, "unknown"),
        "required_attendees": item.RequiredAttendees or "",
        "optional_attendees": item.OptionalAttendees or "",
        "categories": get_item_categories(item),
    }


def format_event_full(item, body_max_length: int = 5000) -> dict:
    """Full event details including body."""
    result = format_event_summary(item)
    result["body"] = truncate(item.Body or "", body_max_length)
    result["reminder_set"] = bool(item.ReminderSet)
    result["reminder_minutes"] = (
        item.ReminderMinutesBeforeStart if item.ReminderSet else None
    )
    result["response_status"] = RESPONSE_NAMES.get(item.ResponseStatus, "unknown")
    return result


# --- Task formatting ---


def format_task_summary(item) -> dict:
    """Extract key fields from an Outlook TaskItem into a dict."""
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "status": TASK_STATUS_NAMES.get(item.Status, "unknown"),
        "percent_complete": item.PercentComplete,
        "due_date": outlook_date_or_none(item.DueDate),
        "start_date": outlook_date_or_none(item.StartDate),
        "importance": IMPORTANCE_NAMES.get(item.Importance, "normal"),
        "complete": bool(item.Complete),
        "categories": get_item_categories(item),
        "owner": item.Owner or "",
    }


def format_task_full(item, body_max_length: int = 5000) -> dict:
    """Full task details including body and reminder time."""
    result = format_task_summary(item)
    result["body"] = truncate(item.Body or "", body_max_length)
    result["reminder_set"] = bool(item.ReminderSet)
    result["reminder_time"] = (
        outlook_date_or_none(item.ReminderTime) if item.ReminderSet else None
    )
    result["date_completed"] = (
        outlook_date_or_none(item.DateCompleted) if item.Complete else None
    )
    return result
