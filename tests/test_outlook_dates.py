"""Unit tests for locale-safe Outlook Restrict date formatting."""
import unittest
from datetime import datetime, timezone

from outlook_desktop_mcp.utils.outlook_dates import (
    dasl_datetime,
    dasl_filter,
    dasl_received_clauses,
    jet_datetime,
    jet_start_range,
    parse_filter_datetime,
)


class DaslDatetimeTests(unittest.TestCase):
    def test_iso_not_us_month_first(self):
        dt = datetime(2026, 3, 10, 9, 0)
        self.assertEqual(dasl_datetime(dt), "2026-03-10 09:00")
        self.assertNotEqual(dasl_datetime(dt), dt.strftime("%m/%d/%Y %H:%M"))

    def test_october_is_not_ambiguous_with_march(self):
        march = dasl_datetime(datetime(2026, 3, 10, 14, 30))
        october = dasl_datetime(datetime(2026, 10, 3, 14, 30))
        self.assertEqual(march, "2026-03-10 14:30")
        self.assertEqual(october, "2026-10-03 14:30")
        self.assertNotEqual(march, october)


class DaslFilterTests(unittest.TestCase):
    def test_received_clauses_use_iso_and_unread_flag(self):
        clauses = dasl_received_clauses(
            unread_only=True,
            start=datetime(2026, 3, 10, 9, 0),
            end=datetime(2026, 3, 11, 17, 0),
        )
        sql = dasl_filter(clauses)
        self.assertTrue(sql.startswith("@SQL="))
        self.assertIn("2026-03-10 09:00", sql)
        self.assertIn("2026-03-11 17:00", sql)
        self.assertNotIn("03/10/2026", sql)
        self.assertNotIn("10/03/2026", sql)
        self.assertIn('"urn:schemas:httpmail:read" = 0', sql)

    def test_empty_clauses_mean_no_restrict(self):
        self.assertIsNone(dasl_filter([]))
        self.assertEqual(dasl_received_clauses(), [])


class JetDatetimeTests(unittest.TestCase):
    def test_march_and_october_are_distinct(self):
        march = jet_datetime(datetime(2026, 3, 10, 14, 30))
        october = jet_datetime(datetime(2026, 10, 3, 14, 30))
        self.assertNotEqual(march, october)
        self.assertIn("2026", march)
        self.assertIn("2026", october)

    def test_does_not_shift_the_calendar_day_through_utc(self):
        dt = datetime(2026, 3, 10, 9, 0)
        formatted = jet_datetime(dt)
        self.assertTrue(
            "10" in formatted,
            f"expected day 10 in locale date, got {formatted!r}",
        )

    def test_start_range_uses_jet_not_iso(self):
        start = datetime(2026, 3, 10, 9, 0)
        end = datetime(2026, 3, 11, 17, 0)
        restrict = jet_start_range(start, end)
        self.assertTrue(restrict.startswith("[Start] >="))
        self.assertNotIn("2026-03-10 09:00", restrict)


class ParseFilterDatetimeTests(unittest.TestCase):
    def test_date_only_and_space_separated(self):
        self.assertEqual(
            parse_filter_datetime("2026-03-10"),
            datetime(2026, 3, 10),
        )
        self.assertEqual(
            parse_filter_datetime("2026-03-10 09:00"),
            datetime(2026, 3, 10, 9, 0),
        )

    def test_aware_values_become_naive_local(self):
        parsed = parse_filter_datetime("2026-03-10T09:00:00Z")
        self.assertIsNone(parsed.tzinfo)
        expected = datetime(2026, 3, 10, 9, 0, tzinfo=timezone.utc).astimezone().replace(
            tzinfo=None,
        )
        self.assertEqual(parsed, expected)


if __name__ == "__main__":
    unittest.main()
