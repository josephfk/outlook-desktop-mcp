"""Unit tests for category parsing and update behavior."""
import unittest

from outlook_desktop_mcp.utils.formatting import (
    get_item_categories,
    merge_categories,
)


class _Item:
    def __init__(self, categories):
        self.Categories = categories


class _InaccessibleCategoriesItem:
    @property
    def Categories(self):
        raise RuntimeError("COM property unavailable")


class CategoryTests(unittest.TestCase):
    def test_parser_returns_trimmed_nonempty_categories(self):
        item = _Item(" Action List,Completed, , Client ")

        self.assertEqual(
            get_item_categories(item),
            ["Action List", "Completed", "Client"],
        )

    def test_parser_safely_handles_blank_missing_and_inaccessible_values(self):
        self.assertEqual(get_item_categories(_Item(None)), [])
        self.assertEqual(get_item_categories(_Item("  ")), [])
        self.assertEqual(get_item_categories(object()), [])
        self.assertEqual(get_item_categories(_InaccessibleCategoriesItem()), [])

    def test_replace_uses_only_requested_categories(self):
        self.assertEqual(
            merge_categories(["Existing"], ["New", "Another"], "replace"),
            ["New", "Another"],
        )

    def test_add_preserves_existing_and_avoids_case_insensitive_duplicates(self):
        self.assertEqual(
            merge_categories(
                ["Action List", "Client"],
                ["action list", "Waiting on Client"],
                "add",
            ),
            ["Action List", "Client", "Waiting on Client"],
        )

    def test_remove_matches_case_insensitively(self):
        self.assertEqual(
            merge_categories(
                ["Action List", "Client", "Completed"],
                ["action list", "COMPLETED"],
                "remove",
            ),
            ["Client"],
        )

    def test_unknown_mode_is_rejected(self):
        with self.assertRaises(ValueError):
            merge_categories([], [], "invalid")


if __name__ == "__main__":
    unittest.main()
