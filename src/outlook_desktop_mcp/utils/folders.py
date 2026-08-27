"""Folder listing payload helpers."""
from __future__ import annotations


def folder_fields(folder, *, full_path: str, include_counts: bool) -> dict:
    """Serialize a MAPIFolder without touching Items unless counts are requested.

    UnReadItemCount is a folder property (cheap). Items.Count enumerates the
    folder and can freeze Outlook in Online mode.
    """
    result = {
        "name": folder.Name,
        "full_path": full_path,
        "unread_count": int(getattr(folder, "UnReadItemCount", 0) or 0),
    }
    if include_counts:
        result["item_count"] = int(folder.Items.Count)
    return result
