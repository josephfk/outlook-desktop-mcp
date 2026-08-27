"""Mail follow-up flags (To-Do bar), independent of Outlook TaskItems."""
from __future__ import annotations

from outlook_desktop_mcp.tools._folder_constants import (
    FLAG_STATUS_NAMES,
    OL_FLAG_COMPLETE,
    OL_FLAG_MARKED,
    OL_MARK_COMPLETE,
    OL_MARK_NO_DATE,
    OL_NO_FLAG,
)
from outlook_desktop_mcp.utils.tasks import (
    outlook_date_or_none,
    parse_iso_datetime,
)

FLAG_STATUS_ALIASES = {
    "flag": "flagged",
    "follow_up": "flagged",
    "followup": "flagged",
    "mark": "flagged",
    "todo": "flagged",
    "done": "complete",
    "completed": "complete",
    "none": "clear",
    "unflag": "clear",
    "remove": "clear",
    "no_flag": "clear",
    "noflag": "clear",
}

ALLOWED_FLAG_STATUSES = ("flagged", "complete", "clear")


def normalize_flag_status(status: str) -> str:
    key = status.strip().lower().replace("-", "_").replace(" ", "_")
    key = FLAG_STATUS_ALIASES.get(key, key)
    if key not in ALLOWED_FLAG_STATUSES:
        allowed = ", ".join(ALLOWED_FLAG_STATUSES)
        raise ValueError(f"status must be one of: {allowed}")
    return key


def flag_info(item) -> dict:
    """JSON-ready flag fields for a MailItem (or similar)."""
    try:
        code = int(item.FlagStatus)
    except Exception:
        code = OL_NO_FLAG
    try:
        marked = bool(item.IsMarkedAsTask)
    except Exception:
        marked = code == OL_FLAG_MARKED
    due = None
    try:
        due = outlook_date_or_none(item.TaskDueDate)
    except Exception:
        due = None
    if due is None:
        try:
            due = outlook_date_or_none(item.FlagDueBy)
        except Exception:
            due = None
    try:
        request = (item.FlagRequest or "").strip() or None
    except Exception:
        request = None
    return {
        "flag_status": FLAG_STATUS_NAMES.get(code, "unknown"),
        "flagged_as_task": marked,
        "flag_due": due,
        "flag_request": request,
    }


def apply_mail_flag(
    item,
    *,
    status: str,
    due_date: str = "",
    reminder: bool = False,
    flag_request: str = "",
) -> None:
    """Set, complete, or clear a follow-up flag on a MailItem-like object.

    Prefers MarkAsTask / ClearTaskFlag (To-Do bar). Falls back to FlagStatus
    on older object models.
    """
    action = normalize_flag_status(status)

    if action == "clear":
        _clear_flag(item)
        return

    if action == "complete":
        _complete_flag(item)
        return

    _mark_flagged(item)
    request = (flag_request or "").strip() or "Follow up"
    try:
        item.FlagRequest = request
    except Exception:
        pass

    if due_date.strip():
        due = parse_iso_datetime(due_date)
        for attr in ("TaskDueDate", "FlagDueBy"):
            try:
                setattr(item, attr, due)
            except Exception:
                pass
        if reminder:
            try:
                item.ReminderSet = True
                item.ReminderTime = due
            except Exception:
                pass
    elif reminder:
        try:
            item.ReminderSet = True
        except Exception:
            pass


def _clear_flag(item) -> None:
    try:
        item.ClearTaskFlag()
        return
    except Exception:
        pass
    item.FlagStatus = OL_NO_FLAG


def _complete_flag(item) -> None:
    try:
        item.MarkAsTask(OL_MARK_COMPLETE)
        return
    except Exception:
        pass
    item.FlagStatus = OL_FLAG_COMPLETE


def _mark_flagged(item) -> None:
    try:
        item.MarkAsTask(OL_MARK_NO_DATE)
        return
    except Exception:
        pass
    item.FlagStatus = OL_FLAG_MARKED
