"""Locale-safe Outlook Restrict date formatting.

Jet filters (`[Start] >= '...'`) follow the Windows user short-date format.
Hardcoding mm/dd/yyyy silently returns the wrong window on en-AU (and other
day-first) locales. DASL `@SQL=` filters are locale-independent when dates
use ISO `yyyy-mm-dd HH:MM`.

Calendar `IncludeRecurrences` still requires Jet, so those filters go through
the user-locale formatter. Mail list/search uses DASL ISO dates.
"""
from __future__ import annotations

from ctypes import Structure, byref, create_unicode_buffer, windll, wintypes
from datetime import datetime

LOCALE_USER_DEFAULT = 0x0400
DATE_SHORTDATE = 0x00000001
TIME_NOSECONDS = 0x00000002


class SYSTEMTIME(Structure):
    _fields_ = [
        ("wYear", wintypes.WORD),
        ("wMonth", wintypes.WORD),
        ("wDayOfWeek", wintypes.WORD),
        ("wDay", wintypes.WORD),
        ("wHour", wintypes.WORD),
        ("wMinute", wintypes.WORD),
        ("wSecond", wintypes.WORD),
        ("wMilliseconds", wintypes.WORD),
    ]


def parse_filter_datetime(value: str) -> datetime:
    """Parse an ISO 8601 date/time. Timezone-aware values become local naive."""
    text = value.strip().replace("Z", "+00:00")
    parsed = datetime.fromisoformat(text)
    if parsed.tzinfo is not None:
        return parsed.astimezone().replace(tzinfo=None)
    return parsed


def dasl_datetime(dt: datetime) -> str:
    """Locale-independent DASL date-time (`yyyy-mm-dd HH:MM`)."""
    return dt.strftime("%Y-%m-%d %H:%M")


def jet_datetime(dt: datetime) -> str:
    """Jet Restrict date-time in the current Windows user locale.

    Uses kernel32 GetDateFormatW/GetTimeFormatW with a SYSTEMTIME so the
    calendar day is not shifted through UTC the way pywin32 Time objects are.
    """
    st = SYSTEMTIME(
        dt.year, dt.month, 0, dt.day,
        dt.hour, dt.minute, dt.second, 0,
    )
    date_buf = create_unicode_buffer(80)
    time_buf = create_unicode_buffer(80)
    kernel32 = windll.kernel32
    if not kernel32.GetDateFormatW(
        LOCALE_USER_DEFAULT, DATE_SHORTDATE, byref(st), None, date_buf, 80,
    ):
        raise OSError("GetDateFormatW failed")
    if not kernel32.GetTimeFormatW(
        LOCALE_USER_DEFAULT, TIME_NOSECONDS, byref(st), None, time_buf, 80,
    ):
        raise OSError("GetTimeFormatW failed")
    return f"{date_buf.value} {time_buf.value}"


def dasl_received_clauses(
    *,
    unread_only: bool = False,
    start: datetime | None = None,
    end: datetime | None = None,
) -> list[str]:
    """DASL clauses for received-time and unread filters."""
    parts: list[str] = []
    if unread_only:
        parts.append('"urn:schemas:httpmail:read" = 0')
    if start is not None:
        parts.append(
            f"\"urn:schemas:httpmail:datereceived\" >= '{dasl_datetime(start)}'"
        )
    if end is not None:
        parts.append(
            f"\"urn:schemas:httpmail:datereceived\" <= '{dasl_datetime(end)}'"
        )
    return parts


def dasl_filter(clauses: list[str]) -> str | None:
    """Join DASL clauses with `@SQL=`. Empty input returns None (no Restrict)."""
    if not clauses:
        return None
    return "@SQL=" + " AND ".join(clauses)


def jet_start_range(start: datetime, end: datetime) -> str:
    """Jet `[Start]` range for calendar Restrict + IncludeRecurrences."""
    return (
        f"[Start] >= '{jet_datetime(start)}' "
        f"AND [Start] <= '{jet_datetime(end)}'"
    )
