"""Exporters: JSON / CSV / SQLite."""

from __future__ import annotations

import csv
import json
import os
import re
import sqlite3

FIELDS = [
    "stem", "folder", "recovered_from_deleted", "subject", "sender", "sender_smtp", "sender_name",
    "to", "cc", "bcc", "date", "received_date", "message_class", "internet_message_id",
    "size", "has_attachments", "attachment_count", "attachment_names",
    "attachment_dir", "md_path", "eml_path", "entry_id",
]
FIELDS_WITH_BODY = FIELDS + ["body_text", "body_html"]


def matches_filters(m: dict, sender_contains="", keyword="", after="", before="") -> bool:
    if sender_contains and sender_contains.lower() not in (
        (m.get("sender") or "") + " " + (m.get("sender_name") or "")
    ).lower():
        return False
    if keyword:
        hay = " ".join([m.get("subject", ""), m.get("body_text", ""), m.get("sender", "")]).lower()
        if keyword.lower() not in hay:
            return False
    # ISO-date string compare works for YYYY-MM-DD prefixes
    if after and (m.get("date") or "") < after:
        return False
    if before and (m.get("date") or "") > before:
        return False
    return True


def write_json(rows: list[dict], path: str) -> None:
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)


def write_csv(rows: list[dict], path: str, include_body: bool = False) -> None:
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    fields = FIELDS_WITH_BODY if include_body else FIELDS
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in fields})


def write_sqlite(rows: list[dict], path: str) -> None:
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    con = sqlite3.connect(path)
    try:
        con.execute("DROP TABLE IF EXISTS emails")
        con.execute(
            """CREATE TABLE emails(
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              stem TEXT, folder TEXT, recovered_from_deleted INTEGER,
              subject TEXT, sender TEXT, sender_smtp TEXT,
              sender_name TEXT, [to] TEXT, cc TEXT, bcc TEXT, date TEXT, received_date TEXT,
              message_class TEXT, internet_message_id TEXT,
              size INTEGER, has_attachments INTEGER, attachment_count INTEGER,
              attachment_names TEXT, attachment_dir TEXT, md_path TEXT, eml_path TEXT,
              entry_id TEXT, body_text TEXT, body_html TEXT)"""
        )
        for r in rows:
            con.execute(
                """INSERT INTO emails(stem,folder,recovered_from_deleted,subject,sender,sender_smtp,sender_name,
                   [to],cc,bcc,date,received_date,message_class,internet_message_id,
                   size,has_attachments,attachment_count,
                   attachment_names,attachment_dir,md_path,eml_path,entry_id,body_text,body_html)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    r.get("stem", ""), r.get("folder", ""),
                    1 if r.get("recovered_from_deleted") else 0,
                    r.get("subject", ""),
                    r.get("sender", ""), r.get("sender_smtp", ""),
                    r.get("sender_name", ""), r.get("to", ""), r.get("cc", ""),
                    r.get("bcc", ""), r.get("date", ""), r.get("received_date", ""),
                    r.get("message_class", ""),
                    r.get("internet_message_id", ""),
                    r.get("size", 0), 1 if r.get("has_attachments") else 0,
                    r.get("attachment_count", 0), r.get("attachment_names", ""),
                    r.get("attachment_dir", ""), r.get("md_path", ""), r.get("eml_path", ""),
                    r.get("entry_id", ""), r.get("body_text", ""), r.get("body_html", ""),
                ),
            )
        con.commit()
        con.execute("CREATE INDEX IF NOT EXISTS idx_emails_date ON emails(date)")
        con.execute("CREATE INDEX IF NOT EXISTS idx_emails_sender ON emails(sender)")
        con.commit()
    finally:
        con.close()


# ── Incremental support: stable identity + manifest reload ────────────────────

_STEM_IDX_RE = re.compile(r"^(\d+)_")

_INT_FIELDS = frozenset({
    "size", "has_attachments", "attachment_count", "recovered_from_deleted",
})


def message_key(m: dict) -> str:
    """Stable identity for dedup: entry_id, falling back to internet_message_id."""
    return str(m.get("entry_id") or m.get("internet_message_id") or "")


def manifest_keys(rows: list[dict]) -> set[str]:
    return {k for k in (message_key(r) for r in rows) if k}


def max_stem_index(rows: list[dict]) -> int:
    """Highest numeric prefix in existing stems (-1 when none parse)."""
    best = -1
    for r in rows:
        mt = _STEM_IDX_RE.match(str(r.get("stem") or ""))
        if mt:
            try:
                best = max(best, int(mt.group(1)))
            except ValueError:
                continue
    return best


def _normalize_row(r: dict) -> dict:
    row = {k: r.get(k, "") for k in FIELDS_WITH_BODY}
    for k in _INT_FIELDS:
        try:
            row[k] = int(row[k] or 0)
        except (TypeError, ValueError):
            row[k] = 0
    if not isinstance(row.get("stem"), str):
        row["stem"] = str(row.get("stem") or "")
    return row


def load_manifest(out_dir: str, base: str) -> list[dict]:
    """Reload a previous export manifest (JSON, else SQLite, else CSV).

    Returns [] when no manifest exists yet (fresh export).
    """
    jp = os.path.join(out_dir, f"{base}.json")
    if os.path.isfile(jp):
        try:
            with open(jp, encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                return [_normalize_row(r) for r in data if isinstance(r, dict)]
        except (OSError, ValueError):
            pass
    dp = os.path.join(out_dir, f"{base}.db")
    if os.path.isfile(dp):
        try:
            con = sqlite3.connect(dp)
            try:
                con.row_factory = sqlite3.Row
                return [_normalize_row(dict(r))
                        for r in con.execute("SELECT * FROM emails")]
            finally:
                con.close()
        except sqlite3.Error:
            pass
    cp = os.path.join(out_dir, f"{base}.csv")
    if os.path.isfile(cp):
        try:
            with open(cp, encoding="utf-8-sig", newline="") as f:
                return [_normalize_row(dict(r)) for r in csv.DictReader(f)]
        except (OSError, csv.Error):
            pass
    return []
