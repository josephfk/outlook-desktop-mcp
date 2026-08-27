"""Unit tests for folder listing fields (no Items.Count unless requested)."""
import unittest

from outlook_desktop_mcp.utils.folders import folder_fields


class _Items:
    def __init__(self):
        self.count_reads = 0

    @property
    def Count(self):
        self.count_reads += 1
        return 42


class _Folder:
    def __init__(self, name="Inbox", unread=3):
        self.Name = name
        self.UnReadItemCount = unread
        self.Items = _Items()


class FolderFieldsTests(unittest.TestCase):
    def test_default_omits_item_count_and_does_not_touch_items(self):
        folder = _Folder()
        payload = folder_fields(folder, full_path="Inbox", include_counts=False)
        self.assertEqual(payload["name"], "Inbox")
        self.assertEqual(payload["full_path"], "Inbox")
        self.assertEqual(payload["unread_count"], 3)
        self.assertNotIn("item_count", payload)
        self.assertEqual(folder.Items.count_reads, 0)

    def test_include_counts_reads_items_count(self):
        folder = _Folder(unread=0)
        payload = folder_fields(folder, full_path="Inbox/Receipts", include_counts=True)
        self.assertEqual(payload["item_count"], 42)
        self.assertEqual(payload["unread_count"], 0)
        self.assertEqual(folder.Items.count_reads, 1)


if __name__ == "__main__":
    unittest.main()
