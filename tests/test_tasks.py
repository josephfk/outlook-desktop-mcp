"""Unit tests for task dates, status mapping, filters, and field updates."""
import unittest
from datetime import datetime

from outlook_desktop_mcp.tools._folder_constants import (
    OL_TASK_COMPLETE,
    OL_TASK_IN_PROGRESS,
    OL_TASK_NOT_STARTED,
)
from outlook_desktop_mcp.utils.formatting import format_task_full, format_task_summary
from outlook_desktop_mcp.utils.tasks import (
    OUTLOOK_NULL_DATE,
    apply_task_fields,
    compute_reminder_time,
    is_outlook_task,
    outlook_date_or_none,
    parse_importance,
    parse_task_status,
    task_matches_filters,
)


class _Task:
    def __init__(self, **kwargs):
        self.EntryID = kwargs.get("EntryID", "eid")
        self.Subject = kwargs.get("Subject", "Task")
        self.Status = kwargs.get("Status", OL_TASK_NOT_STARTED)
        self.PercentComplete = kwargs.get("PercentComplete", 0)
        self.DueDate = kwargs.get("DueDate", OUTLOOK_NULL_DATE)
        self.StartDate = kwargs.get("StartDate", OUTLOOK_NULL_DATE)
        self.Importance = kwargs.get("Importance", 1)
        self.Complete = kwargs.get("Complete", False)
        self.Categories = kwargs.get("Categories", "")
        self.Owner = kwargs.get("Owner", "me")
        self.Body = kwargs.get("Body", "")
        self.ReminderSet = kwargs.get("ReminderSet", False)
        self.ReminderTime = kwargs.get("ReminderTime", OUTLOOK_NULL_DATE)
        self.DateCompleted = kwargs.get("DateCompleted", OUTLOOK_NULL_DATE)
        self.Class = kwargs.get("Class", 48)


class OutlookDateTests(unittest.TestCase):
    def test_datetime_sentinel_year_is_none(self):
        self.assertIsNone(outlook_date_or_none(datetime(4501, 1, 1)))
        self.assertIsNone(outlook_date_or_none(datetime(4500, 12, 31)))

    def test_real_datetime_is_isoformat(self):
        self.assertEqual(
            outlook_date_or_none(datetime(2026, 3, 1, 9, 0)),
            "2026-03-01T09:00:00",
        )

    def test_us_and_eu_sentinel_strings_are_none(self):
        self.assertIsNone(outlook_date_or_none("01/01/4501"))
        self.assertIsNone(outlook_date_or_none("1/1/4501"))
        self.assertIsNone(outlook_date_or_none("01.01.4501"))
        self.assertIsNone(outlook_date_or_none("4501-01-01"))

    def test_iso_string_passthrough(self):
        self.assertEqual(
            outlook_date_or_none("2026-03-01T00:00:00"),
            "2026-03-01T00:00:00",
        )

    def test_none_and_blank(self):
        self.assertIsNone(outlook_date_or_none(None))
        self.assertIsNone(outlook_date_or_none("  "))


class FormatTaskTests(unittest.TestCase):
    def test_summary_hides_sentinel_due_date(self):
        summary = format_task_summary(_Task())
        self.assertIsNone(summary["due_date"])
        self.assertIsNone(summary["start_date"])
        self.assertEqual(summary["status"], "not_started")

    def test_summary_keeps_real_due_date(self):
        summary = format_task_summary(
            _Task(DueDate=datetime(2026, 3, 15)),
        )
        self.assertTrue(summary["due_date"].startswith("2026-03-15"))

    def test_full_includes_reminder_time_not_appointment_minutes(self):
        item = _Task(
            ReminderSet=True,
            ReminderTime=datetime(2026, 3, 14, 17, 0),
            Complete=True,
            DateCompleted=datetime(2026, 3, 10),
            Status=OL_TASK_COMPLETE,
            PercentComplete=100,
        )
        full = format_task_full(item)
        self.assertEqual(full["reminder_time"], "2026-03-14T17:00:00")
        self.assertNotIn("reminder_minutes", full)
        self.assertTrue(full["date_completed"].startswith("2026-03-10"))


class StatusAndImportanceTests(unittest.TestCase):
    def test_status_aliases(self):
        self.assertEqual(parse_task_status("completed"), OL_TASK_COMPLETE)
        self.assertEqual(parse_task_status("in progress"), OL_TASK_IN_PROGRESS)
        self.assertEqual(parse_task_status("NOT_STARTED"), OL_TASK_NOT_STARTED)

    def test_unknown_status_rejected(self):
        with self.assertRaisesRegex(ValueError, "status must be one of"):
            parse_task_status("maybe")

    def test_importance(self):
        self.assertEqual(parse_importance("HIGH"), 2)
        with self.assertRaises(ValueError):
            parse_importance("urgent")


class ReminderTests(unittest.TestCase):
    def test_explicit_reminder_time_wins(self):
        dt = compute_reminder_time(
            reminder_time="2026-03-01T08:00:00",
            reminder_minutes=30,
            due_date="2026-03-01T09:00:00",
        )
        self.assertEqual(dt, datetime(2026, 3, 1, 8, 0))

    def test_minutes_before_due_date(self):
        dt = compute_reminder_time(
            reminder_minutes=30,
            due_date="2026-03-01T09:00:00",
        )
        self.assertEqual(dt, datetime(2026, 3, 1, 8, 30))

    def test_minutes_without_due_date_rejected(self):
        with self.assertRaisesRegex(ValueError, "requires due_date"):
            compute_reminder_time(reminder_minutes=15)


class ApplyFieldsTests(unittest.TestCase):
    def test_update_status_complete_sets_percent(self):
        item = _Task()
        apply_task_fields(item, status="complete")
        self.assertEqual(item.Status, OL_TASK_COMPLETE)
        self.assertTrue(item.Complete)
        self.assertEqual(item.PercentComplete, 100)

    def test_uncomplete_resets_complete_flag(self):
        item = _Task(
            Status=OL_TASK_COMPLETE,
            Complete=True,
            PercentComplete=100,
        )
        apply_task_fields(item, status="not_started")
        self.assertEqual(item.Status, OL_TASK_NOT_STARTED)
        self.assertFalse(item.Complete)
        self.assertEqual(item.PercentComplete, 0)

    def test_percent_100_marks_complete(self):
        item = _Task()
        apply_task_fields(item, percent_complete=100)
        self.assertTrue(item.Complete)
        self.assertEqual(item.Status, OL_TASK_COMPLETE)

    def test_clear_due_date(self):
        item = _Task(DueDate=datetime(2026, 3, 1))
        apply_task_fields(item, due_date="none")
        self.assertEqual(item.DueDate, OUTLOOK_NULL_DATE)

    def test_set_due_and_reminder_from_minutes(self):
        item = _Task()
        apply_task_fields(
            item,
            due_date="2026-04-01T17:00:00",
            reminder_minutes=60,
        )
        self.assertEqual(item.DueDate, datetime(2026, 4, 1, 17, 0))
        self.assertTrue(item.ReminderSet)
        self.assertEqual(item.ReminderTime, datetime(2026, 4, 1, 16, 0))

    def test_clear_reminder(self):
        item = _Task(ReminderSet=True, ReminderTime=datetime(2026, 3, 1))
        apply_task_fields(item, clear_reminder=True)
        self.assertFalse(item.ReminderSet)

    def test_does_not_use_appointment_reminder_property(self):
        item = _Task()
        apply_task_fields(
            item,
            due_date="2026-04-01T17:00:00",
            reminder_minutes=15,
        )
        self.assertFalse(hasattr(item, "ReminderMinutesBeforeStart"))


class FilterTests(unittest.TestCase):
    def test_status_and_importance_and_category(self):
        summary = {
            "status": "in_progress",
            "importance": "high",
            "categories": ["Work", "Follow-up"],
            "due_date": "2026-03-10T00:00:00",
        }
        self.assertTrue(
            task_matches_filters(
                summary,
                status="in progress",
                importance="high",
                category="work",
                due_before="2026-03-15",
            )
        )
        self.assertFalse(task_matches_filters(summary, status="complete"))
        self.assertFalse(task_matches_filters(summary, category="Home"))
        self.assertFalse(
            task_matches_filters(summary, due_before="2026-03-01")
        )

    def test_undated_task_fails_due_before(self):
        self.assertFalse(
            task_matches_filters({"due_date": None}, due_before="2026-03-01")
        )


class ClassFilterTests(unittest.TestCase):
    def test_task_class_48(self):
        self.assertTrue(is_outlook_task(_Task(Class=48)))
        self.assertFalse(is_outlook_task(_Task(Class=43)))

    def test_missing_class_is_not_a_task(self):
        self.assertFalse(is_outlook_task(object()))


if __name__ == "__main__":
    unittest.main()
