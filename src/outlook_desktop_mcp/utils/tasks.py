"""Task field mapping, Outlook null dates, and list filters."""
from __future__ import annotations

from datetime import datetime, timedelta
import re

from outlook_desktop_mcp.tools._folder_constants import (
    IMPORTANCE_NAMES,
    OL_TASK_COMPLETE,
    OL_TASK_IN_PROGRESS,
    OL_TASK_NOT_STARTED,
    TASK_STATUS_NAMES,
)

OUTLOOK_NULL_YEAR = 4500
OUTLOOK_NULL_DATE = datetime(4501, 1, 1)

STATUS_ALIASES = {
    "completed": "complete",
    "done": "complete",
    "todo": "not_started",
    "not started": "not_started",
    "in progress": "in_progress",
    "in-progress": "in_progress",
}

STATUS_NAME_TO_VALUE = {name: code for code, name in TASK_STATUS_NAMES.items()}
IMPORTANCE_NAME_TO_VALUE = {name: code for code, name in IMPORTANCE_NAMES.items()}

_YEAR_TOKEN = re.compile(r"\b(450[01])\b")


def outlook_date_or_none(value) -> str | None:
    """Return an ISO date string, or None for Outlook's unset sentinel (year 4500/4501).

    Compares the datetime year when available so non-US string forms of
    01/01/4501 are not treated as real due dates.
    """
    if value is None:
        return None

    year = getattr(value, "year", None)
    if isinstance(year, int):
        if year >= OUTLOOK_NULL_YEAR:
            return None
        if hasattr(value, "isoformat"):
            try:
                return value.isoformat()
            except Exception:
                pass
        return str(value)

    text = str(value).strip()
    if not text:
        return None
    try:
        dt = datetime.fromisoformat(text.replace("Z", "+00:00"))
        if dt.year >= OUTLOOK_NULL_YEAR:
            return None
        return dt.isoformat()
    except ValueError:
        pass
    if _YEAR_TOKEN.search(text):
        return None
    return text


def parse_iso_datetime(value: str) -> datetime:
    return datetime.fromisoformat(value.strip().replace("Z", "+00:00"))


def normalize_status_name(status: str) -> str:
    key = status.strip().lower().replace("_", " ")
    key = STATUS_ALIASES.get(key, key.replace(" ", "_"))
    if key not in STATUS_NAME_TO_VALUE:
        allowed = ", ".join(STATUS_NAME_TO_VALUE)
        raise ValueError(f"status must be one of: {allowed}")
    return key


def parse_task_status(status: str) -> int:
    return STATUS_NAME_TO_VALUE[normalize_status_name(status)]


def normalize_importance_name(importance: str) -> str:
    key = importance.strip().lower()
    if key not in IMPORTANCE_NAME_TO_VALUE:
        allowed = ", ".join(IMPORTANCE_NAME_TO_VALUE)
        raise ValueError(f"importance must be one of: {allowed}")
    return key


def parse_importance(importance: str) -> int:
    return IMPORTANCE_NAME_TO_VALUE[normalize_importance_name(importance)]


def compute_reminder_time(
    *,
    reminder_time: str = "",
    reminder_minutes: int = 0,
    due_date: str = "",
) -> datetime | None:
    """TaskItem uses ReminderTime (absolute), not ReminderMinutesBeforeStart."""
    if reminder_time.strip():
        return parse_iso_datetime(reminder_time)
    if reminder_minutes > 0:
        if not due_date.strip():
            raise ValueError(
                "reminder_minutes requires due_date (or pass reminder_time)"
            )
        return parse_iso_datetime(due_date) - timedelta(minutes=reminder_minutes)
    return None


def apply_task_fields(
    item,
    *,
    subject: str | None = None,
    body: str | None = None,
    due_date: str | None = None,
    start_date: str | None = None,
    status: str | None = None,
    percent_complete: int | None = None,
    importance: str | None = None,
    reminder_time: str | None = None,
    reminder_minutes: int | None = None,
    clear_reminder: bool = False,
) -> None:
    """Mutate a TaskItem-like object. Empty/None fields are left unchanged.

    due_date/start_date of 'none' or 'clear' wipe the Outlook date.
    """
    if subject is not None:
        item.Subject = subject
    if body is not None:
        item.Body = body
    if due_date is not None:
        item.DueDate = _assign_outlook_date(due_date)
    if start_date is not None:
        item.StartDate = _assign_outlook_date(start_date)
    if importance is not None:
        item.Importance = parse_importance(importance)

    if status is not None:
        code = parse_task_status(status)
        item.Status = code
        if code == OL_TASK_COMPLETE:
            item.Complete = True
            if percent_complete is None:
                item.PercentComplete = 100
        else:
            item.Complete = False
            if percent_complete is None and code == OL_TASK_NOT_STARTED:
                item.PercentComplete = 0

    if percent_complete is not None:
        if percent_complete < 0 or percent_complete > 100:
            raise ValueError("percent_complete must be between 0 and 100")
        item.PercentComplete = percent_complete
        if percent_complete >= 100:
            item.Status = OL_TASK_COMPLETE
            item.Complete = True
        elif getattr(item, "Complete", False) and percent_complete < 100:
            item.Complete = False
            if status is None:
                item.Status = OL_TASK_IN_PROGRESS

    if clear_reminder:
        item.ReminderSet = False
    else:
        computed = compute_reminder_time(
            reminder_time=reminder_time or "",
            reminder_minutes=reminder_minutes or 0,
            due_date=due_date or "",
        )
        if computed is not None:
            item.ReminderSet = True
            item.ReminderTime = computed


def _assign_outlook_date(value: str):
    key = value.strip().lower()
    if key in {"none", "clear", ""}:
        return OUTLOOK_NULL_DATE
    return parse_iso_datetime(value)


def task_matches_filters(
    summary: dict,
    *,
    status: str = "",
    importance: str = "",
    category: str = "",
    due_before: str = "",
) -> bool:
    """True if a format_task_summary dict matches optional list_tasks filters."""
    if status:
        if summary.get("status") != normalize_status_name(status):
            return False
    if importance:
        if summary.get("importance") != normalize_importance_name(importance):
            return False
    if category:
        needle = category.casefold()
        cats = [str(c).casefold() for c in (summary.get("categories") or [])]
        if needle not in cats:
            return False
    if due_before:
        due = summary.get("due_date")
        if not due:
            return False
        due_dt = _coerce_datetime(due)
        before = parse_iso_datetime(due_before)
        if due_dt is None or _as_naive(due_dt) > _as_naive(before):
            return False
    return True


def _as_naive(value: datetime) -> datetime:
    if value.tzinfo is not None:
        return value.replace(tzinfo=None)
    return value


def _coerce_datetime(value) -> datetime | None:
    if isinstance(value, datetime):
        if value.year >= OUTLOOK_NULL_YEAR:
            return None
        return value
    parsed = outlook_date_or_none(value)
    if not parsed:
        return None
    try:
        return parse_iso_datetime(parsed)
    except ValueError:
        return None


def is_outlook_task(item) -> bool:
    """True if item is an Outlook TaskItem (olTask / Class 48)."""
    try:
        return int(item.Class) == 48
    except Exception:
        return False
