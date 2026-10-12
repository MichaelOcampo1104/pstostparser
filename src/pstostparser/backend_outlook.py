"""Outlook COM backend (Windows + Outlook required).

Reads through MAPI exactly as Outlook sees it — the fallback when a file's
low-level tables don't enumerate fully via Aspose. Read-only.

Store lookup: attaches to a running Outlook instance (or starts one) and finds
the store whose FilePath matches the given .pst/.ost path.

Live mode (`open_live_store` + `iter_messages_on_date`) skips the file lookup
and reads every mail store of the running Outlook, limited to one local day.
"""

from __future__ import annotations

import datetime
import os
from typing import Any, Iterable, Iterator

OL_MAIL_CLASS = 43  # olMail: leaves out meeting requests, NDRs and other non-mail items
OL_MAIL_FOLDER = 0  # Folder.DefaultItemType of a mail folder
SKIP_FOLDER_NAMES = frozenset({"Search Folders"})  # virtual copies of items held elsewhere


def _prop(obj: Any, name: str, default: Any = "") -> Any:
    try:
        v = getattr(obj, name)
        v = v() if callable(v) else v
        return default if v is None else v
    except Exception:
        return default


class _OutlookStore:
    def __init__(self, app, folder):
        self.app = app
        self.folder = folder

    def close(self) -> None:
        pass  # never quit the user's Outlook


class _LiveStore:
    """Every mail store of the running Outlook; no .pst/.ost path involved."""

    def __init__(self, app, namespace):
        self.app = app
        self.namespace = namespace

    def roots(self) -> list:
        folders = self.namespace.Folders
        return [folders.Item(i) for i in range(1, folders.Count + 1)]

    def close(self) -> None:
        pass  # never quit the user's Outlook


def _norm(p: str) -> str:
    return os.path.abspath(p).lower()


def _attach():
    """The running Outlook.Application; starts Outlook if it isn't open."""
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError("Outlook backend needs pywin32: uv pip install pywin32")
    try:
        try:
            return win32com.client.GetActiveObject('Outlook.Application')
        except Exception:
            return win32com.client.Dispatch('Outlook.Application')
    except Exception as e:
        raise RuntimeError(f"Could not connect to Outlook (is it installed?): {e}")


def open_live_store() -> _LiveStore:
    app = _attach()
    return _LiveStore(app, app.GetNamespace('MAPI'))


def open_store(path: str) -> _OutlookStore:
    app = _attach()
    ns = app.GetNamespace('MAPI')
    want = _norm(path)
    for i in range(1, ns.Folders.Count + 1):
        top = ns.Folders.Item(i)
        try:
            fp = _norm(top.Store.FilePath)
        except Exception:
            continue
        if fp == want or os.path.basename(fp) == os.path.basename(want):
            return _OutlookStore(app, top)
    raise RuntimeError(
        f"No Outlook store found for {path!r}. Open the file in Outlook first "
        "(File > Open & Export > Open Outlook Data File).")


def walk_folders(store: _OutlookStore, skip_names: Iterable[str] = ()) -> Iterator[tuple[str, Any]]:
    skip = set(skip_names)

    def _recurse(folder, prefix: str):
        name = _prop(folder, 'Name', '(root)')
        if str(name) in skip:
            return  # prunes the folder and everything under it
        cur = f"{prefix}/{name}" if prefix else str(name)
        yield cur, folder
        try:
            subs = list(folder.Folders)
        except Exception:
            return
        for sf in subs:
            yield from _recurse(sf, cur)

    yield from _recurse(store.folder, "")


def folder_message_count(folder) -> int:
    try:
        return int(folder.Items.Count)
    except Exception:
        return -1


def _iso(v: Any) -> str:
    if v is None or v == "":
        return ""
    try:
        import datetime
        if isinstance(v, datetime.datetime):
            return v.isoformat()
    except Exception:
        pass
    try:
        return str(v)
    except Exception:
        return ""


def _s(v: Any) -> str:
    if v is None:
        return ""
    try:
        return str(v)
    except Exception:
        return ""


def _internet_message_id(mail) -> str:
    try:
        pa = mail.PropertyAccessor
        return _s(pa.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x1035001E'))
    except Exception:
        return ""


def message_to_dict(folder_path: str, mail, include_body: bool = True) -> dict:
    sender = _prop(mail, 'SenderEmailAddress')
    sender_name = _prop(mail, 'SenderName')
    subject = _prop(mail, 'Subject')
    to = _prop(mail, 'To')
    cc = _prop(mail, 'CC')
    bcc = _prop(mail, 'BCC')
    received = _prop(mail, 'ReceivedTime', None)
    date = _prop(mail, 'SentOn', None) or received
    msg_class = _prop(mail, 'MessageClass')
    try:
        size = int(_prop(mail, 'Size', 0) or 0)
    except Exception:
        size = 0

    att_names: list[str] = []
    try:
        atts = mail.Attachments
        for a in atts:
            att_names.append(_s(_prop(a, 'FileName')))
    except Exception:
        pass

    body_text, body_html = "", ""
    if include_body:
        body_text = _s(_prop(mail, 'Body'))
        try:
            body_html = _s(_prop(mail, 'HTMLBody'))
        except Exception:
            pass

    return {
        "folder": folder_path,
        "subject": subject,
        "sender": sender,
        "sender_smtp": sender if "@" in str(sender) else "",
        "sender_name": sender_name,
        "to": to,
        "cc": cc,
        "bcc": bcc,
        "date": _iso(date),
        "received_date": _local_iso(received),
        "message_class": _s(msg_class),
        "internet_message_id": _internet_message_id(mail) if include_body else "",
        "conversation_topic": _s(_prop(mail, 'ConversationTopic')),
        "size": size,
        "has_attachments": len(att_names) > 0,
        "attachment_count": len(att_names),
        "attachment_names": "; ".join(att_names),
        "body_text": body_text,
        "body_html": body_html,
        "entry_id": _s(_prop(mail, 'EntryID')),
    }


def iter_messages(store: _OutlookStore, include_body: bool = True,
                  skip_ids: set[str] | None = None,
                  stats: dict | None = None) -> Iterator[tuple[str, Any, dict]]:
    """skip_ids/stats behave as in backend_aspose.iter_messages (fast EntryID check)."""
    for path, folder in walk_folders(store):
        try:
            items = folder.Items
            n = items.Count
        except Exception:
            continue
        for i in range(1, n + 1):
            try:
                mail = items.Item(i)
                if skip_ids:
                    eid = _s(_prop(mail, 'EntryID'))
                    if eid and eid in skip_ids:
                        if stats is not None:
                            stats["skipped"] = stats.get("skipped", 0) + 1
                        continue
                yield path, mail, message_to_dict(path, mail, include_body)
            except Exception:
                continue


def folder_selected(path: str, wanted: Iterable[str]) -> bool:
    """True when a --folder value names this folder: its name, its path under the store, or its full path."""
    parts = path.split("/")
    names = {path.lower(), parts[-1].lower(), "/".join(parts[1:]).lower()}
    return any(w.lower() in names for w in wanted)


def _local_naive(rt: datetime.datetime) -> datetime.datetime:
    """Outlook's COM times are local wall-clock values. pywin32 attaches a zone label that can be
    wrong (on a Singapore PC it says 'GMT Standard Time'), so the label is dropped, not converted."""
    return rt.replace(tzinfo=None)


def _local_iso(v: Any) -> str:
    """ISO text of an Outlook local time, without the zone label pywin32 attaches."""
    if isinstance(v, datetime.datetime):
        return _local_naive(v).isoformat()
    return _iso(v)


def _count(stats: dict, key: str) -> None:
    stats[key] = stats.get(key, 0) + 1


def _is_mail_folder(folder) -> bool:
    """Calendars, contacts, tasks and notes hold no mail, so they are not walked."""
    try:
        return int(folder.DefaultItemType) == OL_MAIL_FOLDER
    except Exception:
        return True  # type unknown: try the folder and count any errors


def _received_on_day(folder, start: datetime.datetime, end: datetime.datetime,
                     stats: dict) -> Iterator[Any]:
    """Yield the items of `folder` received in [start, end).

    A date-only Restrict filter missed mail in one folder during testing, so each folder
    is walked instead. Sorted newest first, the walk stops at the day's start.
    """
    try:
        items = folder.Items
        count = items.Count
    except Exception:
        _count(stats, "errors")
        return
    try:
        items.Sort("[ReceivedTime]", True)
        newest_first = True
    except Exception:
        newest_first = False  # the folder would not sort: read every item instead
    for i in range(1, count + 1):
        try:
            mail = items.Item(i)
            if int(mail.Class) != OL_MAIL_CLASS:
                _count(stats, "non_mail")
                continue
            rt = _local_naive(mail.ReceivedTime)
        except Exception:
            _count(stats, "errors")
            continue
        if newest_first and rt < start:
            break
        if not (start <= rt < end):
            _count(stats, "outside")
            continue
        yield mail


def iter_messages_on_date(store: _LiveStore, day: datetime.date, folders: Iterable[str] = (),
                          include_body: bool = True, skip_ids: set[str] | None = None,
                          stats: dict | None = None) -> Iterator[tuple[str, Any, dict]]:
    """Yield (folder_path, mail, record) for mail received on `day`, in local time.

    Only kept mail becomes a record, so bodies and attachments are read for kept items alone.
    """
    stats = {} if stats is None else stats
    start = datetime.datetime.combine(day, datetime.time())
    end = start + datetime.timedelta(days=1)
    wanted = list(folders)
    for root in store.roots():
        for path, folder in walk_folders(_OutlookStore(store.app, root), skip_names=SKIP_FOLDER_NAMES):
            if wanted and not folder_selected(path, wanted):
                continue
            if not _is_mail_folder(folder):
                continue
            _count(stats, "folders")
            for mail in _received_on_day(folder, start, end, stats):
                try:
                    eid = _s(_prop(mail, 'EntryID'))
                    if skip_ids and eid and eid in skip_ids:
                        _count(stats, "skipped")
                        continue
                    record = message_to_dict(path, mail, include_body)
                except Exception:
                    _count(stats, "errors")
                    continue
                yield path, mail, record


def save_attachments(mail, dest_dir: str, prefix: str = "") -> list[str]:
    # NOTE: Outlook's SaveAsFile rejects forward slashes — use native abs path.
    dest_dir = os.path.abspath(dest_dir)
    try:
        atts = mail.Attachments
        count = atts.Count
    except Exception:
        return []
    if not count:
        return []  # don't create empty per-mail folders
    os.makedirs(dest_dir, exist_ok=True)
    saved: list[str] = []
    idx = 0
    for a in atts:
        raw_name = _s(_prop(a, 'FileName')) or 'attachment'
        name = "".join(c if c.isalnum() or c in "._- " else "_" for c in raw_name).strip() or 'attachment'
        out = os.path.join(dest_dir, f"{prefix}{idx:02d}_{name}")
        idx += 1
        try:
            a.SaveAsFile(out)
            saved.append(out)
        except Exception:
            continue
    return saved


def save_eml(mail, dest_path: str) -> bool:
    """Outlook has no native EML export; save .msg instead (caller uses .msg name)."""
    try:
        base, _ = os.path.splitext(dest_path)
        msg_path = os.path.abspath(base + ".msg")
        mail.SaveAs(msg_path, 3)  # olMSG
        return True
    except Exception:
        return False
