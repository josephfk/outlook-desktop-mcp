"""Unit tests for visible vs hidden/inline Outlook attachment classification."""
import unittest

from outlook_desktop_mcp.utils.attachments import (
    ATT_MHTML_REF,
    KIND_EMBEDDED_ITEM,
    KIND_FILE,
    KIND_HIDDEN,
    KIND_INLINE,
    KIND_OLE,
    OL_BY_VALUE,
    OL_EMBEDDEDITEM,
    OL_OLE,
    PR_ATTACH_CONTENT_ID,
    PR_ATTACH_FLAGS,
    PR_ATTACH_HIDDEN,
    ClassifiedAttachment,
    attachment_counts,
    classify_attachment,
    classify_item_attachments,
    resolve_attachment_com_index,
    serialize_attachments,
)
from outlook_desktop_mcp.utils.formatting import format_email_full, format_email_summary


class _FakePA:
    def __init__(self, props):
        self._props = props

    def GetProperty(self, key):
        if key not in self._props:
            raise RuntimeError("property not found")
        return self._props[key]


class _FakeAttachment:
    def __init__(
        self,
        filename="report.pdf",
        attach_type=OL_BY_VALUE,
        size=1000,
        props=None,
        display_name="",
    ):
        self.FileName = filename
        self.Type = attach_type
        self.Size = size
        self.DisplayName = display_name
        self.PropertyAccessor = _FakePA(props or {})


class _FakeAttachments:
    def __init__(self, items):
        self._items = items

    @property
    def Count(self):
        return len(self._items)

    def Item(self, index):
        return self._items[index - 1]


class _FakeMail:
    def __init__(self, attachments, html=""):
        self.EntryID = "eid-1"
        self.Subject = "Hello"
        self.SenderEmailAddress = "a@example.com"
        self.SenderName = "A"
        self.ReceivedTime = "2026-03-10"
        self.UnRead = False
        self.Categories = ""
        self.HTMLBody = html
        self.Body = ""
        self.Attachments = _FakeAttachments(attachments)


def _file(**kwargs):
    return _FakeAttachment(**kwargs)


class ClassifyAttachmentTests(unittest.TestCase):
    def test_regular_file_is_visible(self):
        row = classify_attachment(_file(filename="report.pdf"), 1)
        self.assertEqual(row.kind, KIND_FILE)
        self.assertTrue(row.is_visible)
        self.assertEqual(row.com_index, 1)
        self.assertEqual(row.filename, "report.pdf")
        self.assertEqual(row.storage_size, 1000)

    def test_embedded_outlook_item_is_visible(self):
        row = classify_attachment(
            _file(filename="Invite.msg", attach_type=OL_EMBEDDEDITEM),
            2,
        )
        self.assertEqual(row.kind, KIND_EMBEDDED_ITEM)
        self.assertTrue(row.is_visible)

    def test_ole_embedding_is_not_visible(self):
        row = classify_attachment(
            _file(filename="Chart", attach_type=OL_OLE, size=80000),
            1,
        )
        self.assertEqual(row.kind, KIND_OLE)
        self.assertFalse(row.is_visible)

    def test_hidden_flag_is_not_visible(self):
        row = classify_attachment(
            _file(props={PR_ATTACH_HIDDEN: True}),
            1,
        )
        self.assertEqual(row.kind, KIND_HIDDEN)
        self.assertFalse(row.is_visible)

    def test_hidden_flag_accepts_integer_true(self):
        row = classify_attachment(
            _file(props={PR_ATTACH_HIDDEN: 1}),
            1,
        )
        self.assertEqual(row.kind, KIND_HIDDEN)
        self.assertFalse(row.is_visible)

    def test_mhtml_ref_flag_is_inline(self):
        row = classify_attachment(
            _file(
                filename="image001.png",
                props={PR_ATTACH_FLAGS: ATT_MHTML_REF},
            ),
            1,
        )
        self.assertEqual(row.kind, KIND_INLINE)
        self.assertFalse(row.is_visible)

    def test_content_id_referenced_in_html_is_inline(self):
        row = classify_attachment(
            _file(
                filename="logo.png",
                props={PR_ATTACH_CONTENT_ID: "logo@example"},
            ),
            1,
            html_body='<img src="cid:logo@example">',
        )
        self.assertEqual(row.kind, KIND_INLINE)
        self.assertFalse(row.is_visible)

    def test_content_id_not_in_html_stays_a_file(self):
        row = classify_attachment(
            _file(
                filename="photo.png",
                props={PR_ATTACH_CONTENT_ID: "photo@example"},
            ),
            1,
            html_body="<p>no image tag</p>",
        )
        self.assertEqual(row.kind, KIND_FILE)
        self.assertTrue(row.is_visible)

    def test_strips_angle_brackets_on_content_id(self):
        row = classify_attachment(
            _file(props={PR_ATTACH_CONTENT_ID: "<logo@cid>"}),
            1,
            html_body='<img src="CID:LOGO@CID">',
        )
        self.assertEqual(row.kind, KIND_INLINE)
        self.assertEqual(row.content_id, "logo@cid")

    def test_missing_property_accessor_still_classifies_files(self):
        att = _file()
        att.PropertyAccessor = None
        row = classify_attachment(att, 1)
        self.assertEqual(row.kind, KIND_FILE)
        self.assertTrue(row.is_visible)


class ItemAttachmentTests(unittest.TestCase):
    def test_summary_counts_visible_not_raw(self):
        item = _FakeMail([
            _file(filename="image001.png", props={PR_ATTACH_FLAGS: ATT_MHTML_REF}),
            _file(filename="image002.png", props={PR_ATTACH_HIDDEN: True}),
            _file(filename="invoice.pdf"),
        ])
        visible, raw = attachment_counts(item)
        self.assertEqual(visible, 1)
        self.assertEqual(raw, 3)

        summary = format_email_summary(item)
        self.assertTrue(summary["has_attachments"])
        self.assertEqual(summary["attachment_count"], 1)
        self.assertEqual(summary["raw_attachment_count"], 3)

    def test_signature_only_mail_has_no_visible_attachments(self):
        item = _FakeMail([
            _file(filename="image001.png", props={PR_ATTACH_FLAGS: ATT_MHTML_REF}),
        ])
        summary = format_email_summary(item)
        self.assertFalse(summary["has_attachments"])
        self.assertEqual(summary["attachment_count"], 0)
        self.assertEqual(summary["raw_attachment_count"], 1)

    def test_list_payload_reindexes_visible_files(self):
        item = _FakeMail([
            _file(filename="image001.png", props={PR_ATTACH_FLAGS: ATT_MHTML_REF}),
            _file(filename="notes.docx", size=4096),
            _file(filename="Chart", attach_type=OL_OLE),
        ])
        rows = classify_item_attachments(item)
        payload = serialize_attachments(rows)
        self.assertEqual(len(payload), 1)
        self.assertEqual(payload[0]["filename"], "notes.docx")
        self.assertEqual(payload[0]["index"], 1)
        self.assertEqual(payload[0]["visible_index"], 1)
        self.assertEqual(payload[0]["com_index"], 2)
        self.assertEqual(payload[0]["kind"], KIND_FILE)
        self.assertEqual(payload[0]["storage_size"], 4096)

    def test_include_hidden_keeps_all_parts(self):
        item = _FakeMail([
            _file(filename="image001.png", props={PR_ATTACH_FLAGS: ATT_MHTML_REF}),
            _file(filename="notes.docx"),
        ])
        payload = serialize_attachments(
            classify_item_attachments(item),
            include_hidden=True,
        )
        self.assertEqual(len(payload), 2)
        self.assertFalse(payload[0]["is_visible"])
        self.assertEqual(payload[0]["com_index"], 1)
        self.assertIsNone(payload[0]["visible_index"])
        self.assertEqual(payload[1]["visible_index"], 1)
        self.assertEqual(payload[1]["com_index"], 2)

    def test_html_cid_match_requires_opt_in(self):
        html = '<img src="cid:logo@x">'
        att = _file(filename="logo.png", props={PR_ATTACH_CONTENT_ID: "logo@x"})
        item = _FakeMail([att], html=html)

        without_html = classify_item_attachments(item, match_html_cid=False)
        self.assertTrue(without_html[0].is_visible)

        with_html = classify_item_attachments(item, match_html_cid=True)
        self.assertEqual(with_html[0].kind, KIND_INLINE)
        self.assertFalse(with_html[0].is_visible)

    def test_full_email_lists_visible_attachments_and_matches_html_cid(self):
        html = '<img src="cid:logo@x">'
        item = _FakeMail(
            [
                _file(filename="logo.png", props={PR_ATTACH_CONTENT_ID: "logo@x"}),
                _file(filename="spec.pdf"),
            ],
            html=html,
        )
        item.To = "b@example.com"
        item.CC = ""
        full = format_email_full(item)
        self.assertEqual(full["attachment_count"], 1)
        self.assertEqual(full["raw_attachment_count"], 2)
        self.assertEqual(len(full["attachments"]), 1)
        self.assertEqual(full["attachments"][0]["filename"], "spec.pdf")
        self.assertEqual(full["attachments"][0]["com_index"], 2)


class ResolveIndexTests(unittest.TestCase):
    def test_visible_index_skips_inline_parts(self):
        rows = [
            ClassifiedAttachment(1, "image001.png", KIND_INLINE, False, 10, "x"),
            ClassifiedAttachment(2, "file.pdf", KIND_FILE, True, 20, ""),
        ]
        self.assertEqual(resolve_attachment_com_index(rows, 1), 2)

    def test_visible_index_out_of_range(self):
        rows = [
            ClassifiedAttachment(1, "file.pdf", KIND_FILE, True, 20, ""),
        ]
        with self.assertRaisesRegex(ValueError, "Only 1 visible"):
            resolve_attachment_com_index(rows, 2)

    def test_no_visible_attachments(self):
        rows = [
            ClassifiedAttachment(1, "image001.png", KIND_INLINE, False, 10, "x"),
        ]
        with self.assertRaisesRegex(ValueError, "No user-visible"):
            resolve_attachment_com_index(rows, 1)

    def test_com_index_passthrough(self):
        rows = [
            ClassifiedAttachment(1, "image001.png", KIND_INLINE, False, 10, "x"),
            ClassifiedAttachment(2, "file.pdf", KIND_FILE, True, 20, ""),
        ]
        self.assertEqual(
            resolve_attachment_com_index(rows, 1, use_com_index=True),
            1,
        )

    def test_com_index_missing(self):
        rows = [
            ClassifiedAttachment(1, "file.pdf", KIND_FILE, True, 20, ""),
        ]
        with self.assertRaisesRegex(ValueError, "COM index 3"):
            resolve_attachment_com_index(rows, 3, use_com_index=True)


if __name__ == "__main__":
    unittest.main()
