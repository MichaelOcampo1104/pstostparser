"""Per-email Markdown export with traceability front-matter."""

from __future__ import annotations

import html as _html
import os
import re


def slugify(text: str, max_len: int = 60) -> str:
    text = (text or "no-subject").strip().lower()
    text = re.sub(r"[^\w\s-]", "", text, flags=re.UNICODE)
    text = re.sub(r"[\s_-]+", "-", text).strip("-")
    if not text:
        text = "no-subject"
    return text[:max_len].strip("-") or "no-subject"


def mail_stem(idx: int, msg: dict) -> str:
    """Traceable, sortable, filesystem-safe stem: 000001_20240301_subject-slug."""
    date = (msg.get("date") or "")[:10].replace("-", "")
    subj = slugify(msg.get("subject", ""))
    if date:
        return f"{idx:06d}_{date}_{subj}"
    return f"{idx:06d}_{subj}"


def html_to_markdown_basic(html_text: str) -> str:
    """Small dependency-free HTML -> Markdown conversion (links, breaks, lists)."""
    if not html_text:
        return ""
    t = html_text
    t = re.sub(r"(?is)<head.*?</head>", "", t)  # drop <title> etc. (e.g. "External Disclaimer")
    t = re.sub(r"(?i)<br\s*/?>", "\n", t)
    t = re.sub(r"(?i)</p\s*>", "\n\n", t)
    t = re.sub(r"(?i)<p[^>]*>", "", t)
    t = re.sub(r"(?i)</(div|tr|h[1-6]|li|ul|ol|table)>", "\n", t)
    t = re.sub(r"(?i)<li[^>]*>", "\n- ", t)

    def _link(m: re.Match) -> str:
        href, label = m.group(1), re.sub(r"<[^>]+>", "", m.group(2)).strip()
        label = label or href
        return f"[{label}]({href})"

    t = re.sub(r"(?is)<a[^>]+href=[\"']([^\"']+)[\"'][^>]*>(.*?)</a>", _link, t)
    t = re.sub(r"<[^>]+>", "", t)
    t = _html.unescape(t)
    t = re.sub(r"[ \t]+\n", "\n", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()


def _yaml_quote(v: object) -> str:
    s = "" if v is None else str(v)
    s = s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ")
    return f'"{s}"'


def effective_body_text(msg: dict) -> str:
    """Readable body: cleaned copy if stripping ran, else text, else converted HTML."""
    if msg.get("body_clean") is not None:
        return (msg.get("body_clean") or "").strip()
    body = (msg.get("body_text") or "").strip()
    if not body and msg.get("body_html"):
        body = html_to_markdown_basic(msg.get("body_html", ""))
    return body
    s = "" if v is None else str(v)
    s = s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ")
    return f'"{s}"'


def write_markdown(
    msg: dict,
    dest_path: str,
    source_file: str = "",
    attachment_rel_paths: list[str] | None = None,
    eml_rel_path: str = "",
) -> str:
    """Write one email as Markdown with YAML front-matter. Returns dest_path."""
    os.makedirs(os.path.dirname(os.path.abspath(dest_path)), exist_ok=True)
    atts = attachment_rel_paths or []
    body = effective_body_text(msg)
    if not body:
        body = "_(no body content)_"
    stripped_labels = str(msg.get("boilerplate_labels") or "")

    lines: list[str] = ["---"]
    for k in ("folder", "subject", "sender", "sender_smtp", "sender_name",
              "to", "cc", "bcc", "date", "message_class",
              "internet_message_id", "entry_id"):
        lines.append(f"{k}: {_yaml_quote(msg.get(k, ''))}")
    lines.append(f"source_file: {_yaml_quote(source_file)}")
    if msg.get("recovered_from_deleted"):
        lines.append("recovered_from_deleted: true")
    lines.append(f"attachment_count: {msg.get('attachment_count', 0)}")
    lines.append(f"boilerplate_removed: {int(msg.get('boilerplate_removed') or 0)}")
    if stripped_labels:
        lines.append(f"boilerplate_labels: {_yaml_quote(stripped_labels)}")
    lines.append(f"md_file: {_yaml_quote(os.path.basename(dest_path))}")
    if eml_rel_path:
        lines.append(f"eml_file: {_yaml_quote(eml_rel_path)}")
    lines.append("attachments:")
    for a in atts:
        lines.append(f"  - {_yaml_quote(a)}")
    if not atts:
        lines.append("  []")
    lines.append("---")
    lines.append("")
    lines.append(f"# {msg.get('subject') or '(no subject)'}")
    lines.append("")
    lines.append(f"- **From:** {msg.get('sender_name') or ''} <{msg.get('sender') or ''}>")
    lines.append(f"- **To:** {msg.get('to') or ''}")
    if msg.get("cc"):
        lines.append(f"- **Cc:** {msg.get('cc')}")
    lines.append(f"- **Date:** {msg.get('date') or ''}")
    lines.append(f"- **Folder:** {msg.get('folder') or ''}")
    lines.append("")
    lines.append("## Body")
    lines.append("")
    lines.append(body)
    lines.append("")
    if stripped_labels:
        lines.append(f"*Boilerplate notice(s) removed during conversion: {stripped_labels}.*")
        lines.append("")
    lines.append("## Attachments")
    lines.append("")
    if atts:
        for a in atts:
            lines.append(f"- [{os.path.basename(a)}]({a})")
    else:
        lines.append("_(none)_")
    lines.append("")

    with open(dest_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return dest_path
