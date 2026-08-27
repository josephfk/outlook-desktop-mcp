"""Unit tests for SMTP sender resolution, preview, and conversation id."""
import unittest

from outlook_desktop_mcp.utils.formatting import (
    body_preview,
    format_email_summary,
    resolve_sender_smtp,
)


class _PA:
    def __init__(self, props):
        self._props = props

    def GetProperty(self, key):
        if key not in self._props:
            raise RuntimeError("missing")
        return self._props[key]


class _ExchangeUser:
    def __init__(self, smtp):
        self.PrimarySmtpAddress = smtp


class _Sender:
    def __init__(self, smtp=None, pa_props=None):
        self._smtp = smtp
        self.PropertyAccessor = _PA(pa_props or {})

    def GetExchangeUser(self):
        if self._smtp is None:
            return None
        return _ExchangeUser(self._smtp)


class _Mail:
    def __init__(self, **kwargs):
        self.EntryID = kwargs.get("EntryID", "eid")
        self.Subject = kwargs.get("Subject", "Hello")
        self.SenderEmailAddress = kwargs.get("SenderEmailAddress", "a@example.com")
        self.SenderEmailType = kwargs.get("SenderEmailType", "SMTP")
        self.SenderName = kwargs.get("SenderName", "A")
        self.Sender = kwargs.get("Sender")
        self.ReceivedTime = kwargs.get("ReceivedTime", "2026-03-10")
        self.UnRead = kwargs.get("UnRead", False)
        self.Categories = kwargs.get("Categories", "")
        self.HTMLBody = kwargs.get("HTMLBody", "")
        self.Body = kwargs.get("Body", "")
        self.ConversationID = kwargs.get("ConversationID", "conv-1")
        self.Attachments = kwargs.get("Attachments")
        self.PropertyAccessor = _PA(kwargs.get("pa_props") or {})
        if self.Attachments is None:
            class _Empty:
                Count = 0

                def Item(self, index):
                    raise IndexError
            self.Attachments = _Empty()


class BodyPreviewTests(unittest.TestCase):
    def test_collapses_whitespace_and_truncates(self):
        self.assertEqual(body_preview("  hello\n\nworld  "), "hello world")
        long = "x" * 300
        preview = body_preview(long, max_length=240)
        self.assertEqual(len(preview), 243)
        self.assertTrue(preview.endswith("..."))
        self.assertEqual(preview[:240], "x" * 240)


class SmtpResolutionTests(unittest.TestCase):
    def test_smtp_type_is_passed_through(self):
        item = _Mail(SenderEmailAddress="alice@contoso.com", SenderEmailType="SMTP")
        self.assertEqual(resolve_sender_smtp(item), "alice@contoso.com")

    def test_exchange_dn_uses_get_exchange_user(self):
        item = _Mail(
            SenderEmailAddress=(
                "/O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP"
                "/CN=RECIPIENTS/CN=alice"
            ),
            SenderEmailType="EX",
            Sender=_Sender(smtp="alice@contoso.com"),
        )
        self.assertEqual(resolve_sender_smtp(item), "alice@contoso.com")

    def test_exchange_dn_falls_back_to_mapi_smtp_property(self):
        from outlook_desktop_mcp.utils.formatting import PR_SENDER_SMTP_ADDRESS

        item = _Mail(
            SenderEmailAddress="/O=EXCHANGELABS/CN=alice",
            SenderEmailType="EX",
            Sender=_Sender(smtp=None),
            pa_props={PR_SENDER_SMTP_ADDRESS: "alice@contoso.com"},
        )
        self.assertEqual(resolve_sender_smtp(item), "alice@contoso.com")

    def test_unresolvable_dn_returns_original(self):
        item = _Mail(
            SenderEmailAddress="/O=EXCHANGELABS/CN=alice",
            SenderEmailType="EX",
            Sender=_Sender(smtp=None),
        )
        self.assertEqual(resolve_sender_smtp(item), "/O=EXCHANGELABS/CN=alice")


class SummaryTests(unittest.TestCase):
    def test_includes_smtp_preview_and_conversation_id(self):
        item = _Mail(
            SenderEmailAddress="/O=EXCHANGELABS/CN=alice",
            SenderEmailType="EX",
            Sender=_Sender(smtp="alice@contoso.com"),
            Body="Please review the attached\n\nquarterly report.",
            ConversationID="AAMkAD-conversation",
        )
        summary = format_email_summary(item)
        self.assertEqual(summary["sender"], "alice@contoso.com")
        self.assertEqual(
            summary["preview"],
            "Please review the attached quarterly report.",
        )
        self.assertEqual(summary["conversation_id"], "AAMkAD-conversation")
        self.assertEqual(summary["flag_status"], "none")
        self.assertFalse(summary["flagged_as_task"])


if __name__ == "__main__":
    unittest.main()
