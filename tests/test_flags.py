"""Unit tests for mail follow-up flags."""
import unittest
from datetime import datetime

from outlook_desktop_mcp.tools._folder_constants import (
    OL_FLAG_COMPLETE,
    OL_FLAG_MARKED,
    OL_MARK_COMPLETE,
    OL_MARK_NO_DATE,
    OL_NO_FLAG,
)
from outlook_desktop_mcp.utils.flags import (
    apply_mail_flag,
    flag_info,
    normalize_flag_status,
)
from outlook_desktop_mcp.utils.tasks import OUTLOOK_NULL_DATE


class _FlagMail:
    def __init__(self, **kwargs):
        self.FlagStatus = kwargs.get("FlagStatus", OL_NO_FLAG)
        self.FlagRequest = kwargs.get("FlagRequest", "")
        self.FlagDueBy = kwargs.get("FlagDueBy", OUTLOOK_NULL_DATE)
        self.TaskDueDate = kwargs.get("TaskDueDate", OUTLOOK_NULL_DATE)
        self.IsMarkedAsTask = kwargs.get("IsMarkedAsTask", False)
        self.ReminderSet = False
        self.ReminderTime = OUTLOOK_NULL_DATE
        self.mark_interval = None
        self.cleared = False
        self.use_mark = kwargs.get("use_mark", True)

    def MarkAsTask(self, interval):
        if not self.use_mark:
            raise RuntimeError("MarkAsTask unavailable")
        self.mark_interval = interval
        if interval == OL_MARK_COMPLETE:
            self.FlagStatus = OL_FLAG_COMPLETE
            self.IsMarkedAsTask = True
        else:
            self.FlagStatus = OL_FLAG_MARKED
            self.IsMarkedAsTask = True

    def ClearTaskFlag(self):
        if not self.use_mark:
            raise RuntimeError("ClearTaskFlag unavailable")
        self.cleared = True
        self.FlagStatus = OL_NO_FLAG
        self.IsMarkedAsTask = False
        self.FlagRequest = ""
        self.TaskDueDate = OUTLOOK_NULL_DATE


class NormalizeFlagTests(unittest.TestCase):
    def test_aliases(self):
        self.assertEqual(normalize_flag_status("follow-up"), "flagged")
        self.assertEqual(normalize_flag_status("done"), "complete")
        self.assertEqual(normalize_flag_status("unflag"), "clear")

    def test_unknown_rejected(self):
        with self.assertRaises(ValueError):
            normalize_flag_status("snooze")


class FlagInfoTests(unittest.TestCase):
    def test_unflagged(self):
        info = flag_info(_FlagMail())
        self.assertEqual(info["flag_status"], "none")
        self.assertFalse(info["flagged_as_task"])
        self.assertIsNone(info["flag_due"])

    def test_flagged_with_due(self):
        item = _FlagMail(
            FlagStatus=OL_FLAG_MARKED,
            IsMarkedAsTask=True,
            FlagRequest="Follow up",
            TaskDueDate=datetime(2026, 3, 12, 17, 0),
        )
        info = flag_info(item)
        self.assertEqual(info["flag_status"], "flagged")
        self.assertTrue(info["flagged_as_task"])
        self.assertEqual(info["flag_request"], "Follow up")
        self.assertTrue(info["flag_due"].startswith("2026-03-12"))


class ApplyFlagTests(unittest.TestCase):
    def test_flag_uses_mark_as_task(self):
        item = _FlagMail()
        apply_mail_flag(item, status="follow_up", due_date="2026-03-12 17:00")
        self.assertEqual(item.mark_interval, OL_MARK_NO_DATE)
        self.assertEqual(item.FlagStatus, OL_FLAG_MARKED)
        self.assertEqual(item.FlagRequest, "Follow up")
        self.assertEqual(item.TaskDueDate, datetime(2026, 3, 12, 17, 0))
        self.assertEqual(item.FlagDueBy, datetime(2026, 3, 12, 17, 0))

    def test_complete_and_clear(self):
        item = _FlagMail()
        apply_mail_flag(item, status="flagged")
        apply_mail_flag(item, status="complete")
        self.assertEqual(item.mark_interval, OL_MARK_COMPLETE)
        self.assertEqual(item.FlagStatus, OL_FLAG_COMPLETE)

        apply_mail_flag(item, status="clear")
        self.assertTrue(item.cleared)
        self.assertEqual(item.FlagStatus, OL_NO_FLAG)

    def test_fallback_without_mark_as_task(self):
        item = _FlagMail(use_mark=False)
        apply_mail_flag(item, status="flagged")
        self.assertEqual(item.FlagStatus, OL_FLAG_MARKED)
        apply_mail_flag(item, status="complete")
        self.assertEqual(item.FlagStatus, OL_FLAG_COMPLETE)
        apply_mail_flag(item, status="clear")
        self.assertEqual(item.FlagStatus, OL_NO_FLAG)

    def test_reminder_with_due_date(self):
        item = _FlagMail()
        apply_mail_flag(
            item, status="flagged", due_date="2026-03-12 09:00", reminder=True,
        )
        self.assertTrue(item.ReminderSet)
        self.assertEqual(item.ReminderTime, datetime(2026, 3, 12, 9, 0))


if __name__ == "__main__":
    unittest.main()
