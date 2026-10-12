"""pypff backend (libyal). Free/open-source; good for forensics.

Install: uv pip install pypff
Note: on some Windows setups pypff needs a prebuilt wheel. If the import
fails, use the aspose backend instead.
"""

from __future__ import annotations

from typing import Any, Iterator


def open_file(path: str):
    import pypff
    f = pypff.file()
    f.open(path)
    return f


def _s(v: Any) -> str:
    if v is None:
        return ""
    try:
        return str(v)
    except Exception:
        return ""


def walk_folders(pff_file) -> Iterator[tuple[str, Any]]:
    root = pff_file.get_root_folder()

    def _recurse(folder, prefix: str):
        try:
            name = folder.get_name() or "(root)"
        except Exception:
            name = "(root)"
        cur = f"{prefix}/{name}" if prefix else _s(name)
        yield cur, folder
        try:
            n = folder.get_number_of_sub_folders()
        except Exception:
            return
        for i in range(n):
            try:
                yield from _recurse(folder.get_sub_folder(i), cur)
            except Exception:
                continue

    yield from _recurse(root, "")


def message_to_dict(folder_path: str, msg) -> dict:
    def g(*meths: str) -> str:
        for m in meths:
            try:
                fn = getattr(msg, m)
                v = fn() if callable(fn) else fn
                if v:
                    return _s(v)
            except Exception:
                continue
        return ""

    subject = g("get_subject")
    sender = g("get_sender_email_address", "get_sender_name")
    try:
        n_sub = 0
        try:
            n_sub = msg.get_number_of_attachments()
        except Exception:
            pass
        att_names = []
        for i in range(n_sub or 0):
            try:
                att_names.append(_s(msg.get_attachment(i).get_name()))
            except Exception:
                continue
    except Exception:
        n_sub, att_names = 0, []

    body = g("get_plain_text_body", "get_body")
    return {
        "folder": folder_path,
        "subject": subject,
        "sender": sender,
        "sender_name": g("get_sender_name"),
        "to": g("get_display_to", "get_header_string"),
        "cc": "",
        "bcc": "",
        "date": g("get_delivery_time", "get_client_submit_time", "get_creation_time"),
        "message_class": "",
        "size": 0,
        "has_attachments": (n_sub or 0) > 0,
        "attachment_count": n_sub or 0,
        "attachment_names": "; ".join(att_names),
        "body_text": body,
        "body_html": g("get_html_body"),
        "entry_id": g("get_identifier"),
    }


def iter_messages(pff_file, include_body: bool = True) -> Iterator[tuple[str, Any, dict]]:
    for path, folder in walk_folders(pff_file):
        try:
            n = folder.get_number_of_sub_messages()
        except Exception:
            continue
        for i in range(n):
            try:
                m = folder.get_sub_message(i)
                yield path, m, message_to_dict(path, m)
            except Exception:
                continue
