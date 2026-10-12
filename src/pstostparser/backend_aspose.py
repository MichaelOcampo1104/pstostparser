"""Aspose backend: opens both .pst and .ost via PersonalStorage."""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Any, Iterator


def _get(obj: Any, *names: str, default: Any = None) -> Any:
    for n in names:
        try:
            v = getattr(obj, n)
            return v() if callable(v) else v
        except Exception:
            continue
    return default


def open_store(path: str):
    from aspose.email.storage.pst import PersonalStorage
    return PersonalStorage.from_file(path)


def walk_folders(store) -> Iterator[tuple[str, Any]]:
    """Yield (folder_path, folder_obj) depth-first."""
    root = store.root_folder

    def _recurse(folder, prefix: str):
        name = _get(folder, "display_name", "name") or "(root)"
        cur = f"{prefix}/{name}" if prefix else name
        yield cur, folder
        try:
            subs = folder.get_sub_folders()
        except Exception:
            return
        for sf in subs:
            yield from _recurse(sf, cur)

    yield from _recurse(root, "")


def folder_message_count(folder) -> int:
    # content_count is the true total; len(get_contents()) caps at 50 in eval mode.
    try:
        return int(folder.content_count)
    except Exception:
        pass
    try:
        return len(list(folder.get_contents()))
    except Exception:
        return -1


def _folder_pages(folder, page_size: int = 5000):
    """Yield MessageInfo pages; evaluation caps each call at 50, offset paging works."""
    try:
        total = int(folder.content_count)
    except Exception:
        total = -1
    seen: set[str] = set()
    start = 0
    while True:
        if 0 <= total <= start:
            break
        try:
            page = list(folder.get_contents(start, page_size))
        except TypeError:
            page = list(folder.get_contents()) if start == 0 else []
        if not page:
            break
        ids = []
        for mi in page:
            try:
                ids.append(mi.entry_id_string)
            except Exception:
                ids.append(None)
        new = [(mi, i) for mi, i in zip(page, ids) if i is None or i not in seen]
        if not new:
            break  # offset ignored by backend -> avoid infinite loop
        seen.update(i for _, i in new if i is not None)
        yield [mi for mi, _ in new]
        start += len(page)
        if total < 0 and len(page) < page_size and len(page) < 50:
            break


def _safe_str(v: Any) -> str:
    if v is None:
        return ""
    try:
        return str(v)
    except Exception:
        return ""


def _safe_dt(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, datetime):
        if v.tzinfo is None:
            v = v.replace(tzinfo=timezone.utc)
        return v.isoformat()
    return _safe_str(v)


def message_to_dict(folder_path: str, msg, include_body: bool = True) -> dict:
    sender = _get(msg, "sender_email_address", "sender_smtp_address", "sender_address", default="") or ""
    sender_smtp = _get(msg, "sender_smtp_address", default="") or ""
    sender_name = _get(msg, "sender_name", default="") or ""
    subject = _get(msg, "subject", default="") or ""
    to = _get(msg, "display_to", "display_to_string", default="") or ""
    cc = _get(msg, "display_cc", default="") or ""
    bcc = _get(msg, "display_bcc", default="") or ""
    date = _get(msg, "delivery_time", "client_submit_time", "sent_on", "sent_date", default=None)
    msg_class = _get(msg, "message_class", default="") or ""
    size = _get(msg, "message_size", default=0) or 0

    att_names: list[str] = []
    att_count = 0
    try:
        atts = msg.attachments or []
        att_count = len(atts)
        for a in atts:
            n = _get(a, "long_file_name", "file_name", "name", default="") or ""
            att_names.append(_safe_str(n))
    except Exception:
        pass

    body_text = ""
    body_html = ""
    if include_body:
        try:
            body_text = _safe_str(_get(msg, "body", "body_text", default=""))
        except Exception:
            pass
        try:
            body_html = _safe_str(_get(msg, "body_html", "html_body", default=""))
        except Exception:
            pass

    entry_id = ""
    try:
        entry_id = _safe_str(_get(msg, "entry_id_string", "entry_id", default=""))
    except Exception:
        pass

    return {
        "folder": folder_path,
        "subject": _safe_str(subject),
        "sender": _safe_str(sender),
        "sender_smtp": _safe_str(sender_smtp),
        "sender_name": _safe_str(sender_name),
        "to": _safe_str(to),
        "cc": _safe_str(cc),
        "bcc": _safe_str(bcc),
        "date": _safe_dt(date),
        "message_class": _safe_str(msg_class),
        "internet_message_id": _safe_str(_get(msg, "internet_message_id", default="")),
        "conversation_topic": _safe_str(_get(msg, "conversation_topic", default="")),
        "size": int(size) if isinstance(size, int) else 0,
        "has_attachments": att_count > 0,
        "attachment_count": att_count,
        "attachment_names": "; ".join(att_names),
        "body_text": body_text,
        "body_html": body_html,
        "entry_id": entry_id,
    }


def extract_full_message(store, message_info):
    """get_contents() returns lightweight MessageInfo; this loads the full MapiMessage."""
    try:
        return store.extract_message(message_info)
    except Exception:
        pass
    try:
        return store.extract_message(message_info.entry_id_string)
    except Exception:
        return None


def iter_messages(store, include_body: bool = True, skip_ids: set[str] | None = None,
                  stats: dict | None = None) -> Iterator[tuple[str, Any, dict]]:
    """Yield (folder_path, full_msg, msg_dict), paging past the 50/call cap.

    skip_ids: entry_ids from a previous export — matching messages are skipped
    BEFORE full extraction (fast path for --incremental). Skips are tallied in
    stats["skipped"] when stats is given.
    """
    for path, folder in walk_folders(store):
        try:
            pages = _folder_pages(folder)
        except Exception:
            continue
        for page in pages:
            for info in page:
                if skip_ids:
                    try:
                        eid = _safe_str(info.entry_id_string)
                    except Exception:
                        eid = ""
                    if eid and eid in skip_ids:
                        if stats is not None:
                            stats["skipped"] = stats.get("skipped", 0) + 1
                        continue
                try:
                    full = extract_full_message(store, info)
                    if full is None:
                        continue
                    d = message_to_dict(path, full, include_body)
                    if not d.get("entry_id"):
                        try:
                            d["entry_id"] = _safe_str(info.entry_id_string)
                        except Exception:
                            pass
                    yield path, full, d
                except Exception:
                    continue


def save_attachments(msg, dest_dir: str, prefix: str = "") -> list[str]:
    import os
    try:
        atts = msg.attachments or []
    except Exception:
        return []
    if not atts:
        return []  # don't create empty per-mail folders
    os.makedirs(dest_dir, exist_ok=True)
    saved: list[str] = []
    for i, a in enumerate(atts):
        name = _get(a, "long_file_name", "file_name", "name", default=f"attachment_{i}") or f"attachment_{i}"
        name = "".join(c if c.isalnum() or c in "._- " else "_" for c in str(name)).strip() or f"attachment_{i}"
        out = os.path.join(dest_dir, f"{prefix}{i:02d}_{name}")
        try:
            a.save(out)
            saved.append(out)
        except Exception:
            continue
    return saved


def save_eml(msg, dest_path: str) -> bool:
    try:
        from aspose.email import EmlSaveOptions, MailMessageSaveType
        msg.save(dest_path, EmlSaveOptions(MailMessageSaveType.eml))
        return True
    except Exception:
        try:
            msg.save(dest_path)
            return True
        except Exception:
            return False


RECOVERED_LABEL = "Top of Outlook data file [recovered soft-deleted]"


def count_soft_deleted(store) -> int:
    try:
        return len(list(store.find_and_extract_soft_deleted_items()))
    except Exception:
        return -1


def iter_soft_deleted(store, include_body: bool = True) -> Iterator[tuple[str, Any, dict]]:
    """Yield (label, full MapiMessage, msg_dict) for recoverable deleted messages."""
    try:
        entries = store.find_and_extract_soft_deleted_items()
    except Exception:
        return
    for entry in entries:
        try:
            full = entry.item  # already a full MapiMessage
            if full is None:
                continue
            d = message_to_dict(RECOVERED_LABEL, full, include_body)
            if not d.get("entry_id"):
                try:
                    d["entry_id"] = _safe_str(entry.item_id)
                except Exception:
                    pass
            d["recovered_from_deleted"] = True
            yield RECOVERED_LABEL, full, d
        except Exception:
            continue
