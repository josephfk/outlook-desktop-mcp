"""Compose helpers for send/reply: attachments, draft, and Display."""
from __future__ import annotations

import os
from typing import Any


def parse_attachment_paths(spec: str) -> list[str]:
    """Split a semicolon-separated attachment path list."""
    if not spec or not spec.strip():
        return []
    paths: list[str] = []
    for part in spec.split(";"):
        cleaned = part.strip().strip('"').strip("'")
        if cleaned:
            paths.append(cleaned)
    return paths


def attach_local_files(mail: Any, spec: str) -> list[str]:
    """Attach local files to a MailItem. Raises ValueError if a path is missing."""
    attached: list[str] = []
    for raw in parse_attachment_paths(spec):
        path = os.path.realpath(os.path.expanduser(raw))
        if not os.path.isfile(path):
            raise ValueError(f"Attachment not found: {raw}")
        mail.Attachments.Add(path)
        attached.append(path)
    return attached


def finish_mail_item(mail: Any, *, send: bool, display: bool) -> dict:
    """Send, or save a draft and optionally open a modeless Inspector.

    Display(True) is modal and would block the COM STA thread until the user
    closes the window, so Display is always modeless.
    """
    subject = getattr(mail, "Subject", "") or ""
    if send:
        mail.Send()
        return {
            "status": "sent",
            "subject": subject,
        }

    mail.Save()
    result = {
        "status": "draft",
        "subject": subject,
        "entry_id": mail.EntryID,
        "displayed": False,
    }
    if display:
        mail.Display(False)
        result["displayed"] = True
    return result
