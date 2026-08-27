"""Unit tests for compose helpers: attachments, draft, Display."""
import os
import tempfile
import unittest

from outlook_desktop_mcp.utils.mail import (
    attach_local_files,
    finish_mail_item,
    parse_attachment_paths,
)


class _Attachments:
    def __init__(self):
        self.added = []

    def Add(self, path):
        self.added.append(path)


class _Mail:
    def __init__(self, subject="Hello"):
        self.Subject = subject
        self.EntryID = "draft-eid"
        self.Attachments = _Attachments()
        self.sent = False
        self.saved = False
        self.displayed = None

    def Send(self):
        self.sent = True

    def Save(self):
        self.saved = True

    def Display(self, modal):
        self.displayed = modal


class ParseAttachmentPathsTests(unittest.TestCase):
    def test_semicolon_split_and_quote_strip(self):
        self.assertEqual(parse_attachment_paths(""), [])
        self.assertEqual(
            parse_attachment_paths(r'C:\a.pdf; "D:\b notes.docx"'),
            [r"C:\a.pdf", r"D:\b notes.docx"],
        )


class AttachLocalFilesTests(unittest.TestCase):
    def test_missing_path_is_rejected(self):
        mail = _Mail()
        with self.assertRaisesRegex(ValueError, "Attachment not found"):
            attach_local_files(mail, r"C:\definitely-missing-outlook-mcp.pdf")

    def test_existing_file_is_attached(self):
        mail = _Mail()
        with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as handle:
            handle.write(b"hi")
            path = handle.name
        try:
            attached = attach_local_files(mail, path)
            self.assertEqual(len(attached), 1)
            self.assertEqual(mail.Attachments.added[0], os.path.realpath(path))
        finally:
            os.unlink(path)


class FinishMailItemTests(unittest.TestCase):
    def test_send_does_not_save_or_display(self):
        mail = _Mail("Report")
        result = finish_mail_item(mail, send=True, display=True)
        self.assertTrue(mail.sent)
        self.assertFalse(mail.saved)
        self.assertIsNone(mail.displayed)
        self.assertEqual(result["status"], "sent")
        self.assertEqual(result["subject"], "Report")

    def test_draft_saves_and_returns_entry_id(self):
        mail = _Mail("Draft")
        result = finish_mail_item(mail, send=False, display=False)
        self.assertFalse(mail.sent)
        self.assertTrue(mail.saved)
        self.assertIsNone(mail.displayed)
        self.assertEqual(result["status"], "draft")
        self.assertEqual(result["entry_id"], "draft-eid")
        self.assertFalse(result["displayed"])

    def test_display_is_modeless(self):
        mail = _Mail()
        result = finish_mail_item(mail, send=False, display=True)
        self.assertTrue(mail.saved)
        self.assertIs(mail.displayed, False)
        self.assertTrue(result["displayed"])


if __name__ == "__main__":
    unittest.main()
