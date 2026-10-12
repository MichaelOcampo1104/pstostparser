"""CLI: pstostparser info/list/export."""

from __future__ import annotations

import argparse
import datetime
import os
import sys

from . import __version__
from .backends import BACKEND_CHOICES, detect_backends, guess_format, resolve_backend
from .boilerplate import load_custom_patterns, strip_boilerplate
from .exporter import (load_manifest, manifest_keys, matches_filters,
                   max_stem_index, message_key, write_csv, write_json,
                   write_sqlite,
                   is_already_exported, manifest_message_ids)
from .markdown import effective_body_text, mail_stem, write_markdown


def _add_common(p: argparse.ArgumentParser, input_optional: bool = False) -> None:
    if input_optional:
        p.add_argument("input", nargs="?", default="",
                       help="Path to .pst or .ost file (leave out with --live)")
    else:
        p.add_argument("input", help="Path to .pst or .ost file")
    p.add_argument("--backend", default="auto", choices=BACKEND_CHOICES,
                   help="Parsing engine (default: auto)")


def _add_filters(p: argparse.ArgumentParser) -> None:
    p.add_argument("--sender", default="", help="Only keep mails whose sender contains this (case-insensitive)")
    p.add_argument("--keyword", default="", help="Only keep mails with keyword in subject/body/sender")
    p.add_argument("--after", default="", help="Only keep mails with date >= this ISO string (e.g. 2024-01-01)")
    p.add_argument("--before", default="", help="Only keep mails with date <= this ISO string")
    p.add_argument("--limit", type=int, default=0, help="Max messages to export (0 = no limit)")


def live_arg_error(args) -> str | None:
    """Why these export arguments don't fit together, or None when they do."""
    live = getattr(args, "live", False)
    date = getattr(args, "date", "")
    folder = getattr(args, "folder", None)
    name = getattr(args, "name", "")
    if not live:
        if date or folder or name:
            return "--date, --folder and --name only apply with --live"
        if not args.input:
            return "give a .pst/.ost path, or use --live --date YYYY-MM-DD"
        return None
    if args.input:
        return "--live reads the running Outlook; leave out the .pst/.ost path"
    if args.backend not in ("auto", "outlook"):
        return "--live always uses the outlook backend; drop --backend"
    if name and any(c in name for c in '\\/:*?"<>|'):
        return '--name must be a plain file name, without folders or any of \\ / : * ? " < > |'
    if not date:
        return "--live needs --date YYYY-MM-DD"
    try:
        datetime.date.fromisoformat(date)
    except ValueError:
        return f"--date must be YYYY-MM-DD, got {date!r}"
    if getattr(args, "after", "") or getattr(args, "before", ""):
        return "--after/--before do not apply with --live; use --date"
    return None


def live_base_name(day: str, name: str = "") -> str:
    """File stem for a live export: --name when given, else e.g. outlook-live_2026-10-12."""
    return name or f"outlook-live_{day}"


def _messages(backend_mod, live: bool, args, store, skip_ids, stats):
    """Records to export: the file's messages, or only the chosen day's when live."""
    if live:
        return backend_mod.iter_messages_on_date(store, datetime.date.fromisoformat(args.date),
                                                 folders=args.folder or (), include_body=not args.no_body,
                                                 skip_ids=skip_ids, stats=stats)
    return backend_mod.iter_messages(store, include_body=not args.no_body,
                                     skip_ids=skip_ids, stats=stats)


def cmd_info(args) -> int:
    if not os.path.isfile(args.input):
        print(f"File not found: {args.input}", file=sys.stderr)
        return 2
    backend = resolve_backend(args.backend)
    fmt = guess_format(args.input)
    size = os.path.getsize(args.input)
    print(f"File   : {args.input}")
    print(f"Type   : {fmt} ({size:,} bytes)")
    print(f"Backend: {backend}")
    if backend == "aspose":
        from . import backend_aspose as B
        store = B.open_store(args.input)
        try:
            folders = list(B.walk_folders(store))
            print(f"Folders: {len(folders)}")
            total = 0
            for path, folder in folders:
                try:
                    c = B.folder_message_count(folder)
                except Exception:
                    c = -1
                total += max(c, 0)
                print(f"  - {path}  [{c} msgs]")
            print(f"Total messages (approx): {total}")
            try:
                n_del = B.count_soft_deleted(store)
                print(f"Soft-deleted (recoverable with --include-soft-deleted): {n_del}")
            except Exception:
                pass
        finally:
            try:
                store.dispose()
            except Exception:
                pass
    elif backend == "outlook":
        from . import backend_outlook as B
        store = B.open_store(args.input)
        try:
            folders = list(B.walk_folders(store))
            print(f"Folders: {len(folders)}")
            total = 0
            for path, folder in folders:
                try:
                    c = B.folder_message_count(folder)
                except Exception:
                    c = -1
                total += max(c, 0)
                print(f"  - {path}  [{c} msgs]")
            print(f"Total messages (approx): {total}")
        finally:
            try:
                store.close()
            except Exception:
                pass
    else:
        from . import backend_pypff as B
        f = B.open_file(args.input)
        try:
            folders = list(B.walk_folders(f))
            print(f"Folders: {len(folders)}")
            for path, _ in folders:
                print(f"  - {path}")
        finally:
            try:
                f.close()
            except Exception:
                pass
    return 0


def cmd_list(args) -> int:
    args.include_body = False
    return cmd_export(args, list_only=True)


def cmd_export(args, list_only: bool = False) -> int:
    err = live_arg_error(args)
    if err:
        print(f"Error: {err}", file=sys.stderr)
        return 2
    live = bool(getattr(args, "live", False))
    if not live and not os.path.isfile(args.input):
        print(f"File not found: {args.input}", file=sys.stderr)
        return 2
    backend = "outlook" if live else resolve_backend(args.backend)
    base = live_base_name(args.date, getattr(args, "name", "")) if live else os.path.splitext(os.path.basename(args.input))[0]
    rows: list[dict] = []
    extra_boiler: list = []
    # ── incremental: skip ids already in the output manifest ──
    existing_rows: list[dict] = []
    skip_ids: set[str] = set()
    skip_message_ids: set[tuple[str, str]] = set()
    next_idx = 0
    n_skipped_post = 0
    stats: dict = {}
    if getattr(args, "incremental", False) and not list_only:
        existing_rows = load_manifest(args.out, base)
        if existing_rows:
            skip_ids = manifest_keys(existing_rows)
            skip_message_ids = manifest_message_ids(existing_rows)
            next_idx = max_stem_index(existing_rows) + 1
            print(f"Incremental: {len(existing_rows)} message(s) already in {args.out}, "
                  f"skipping {len(skip_ids)} known id(s); new stems from {next_idx:06d}")
        else:
            print(f"Incremental: no existing manifest for {base!r} in {args.out} — full export")
    if getattr(args, "strip_boilerplate", False) and getattr(args, "boilerplate_file", ""):
        try:
            extra_boiler = load_custom_patterns(args.boilerplate_file)
            print(f"Loaded {len(extra_boiler)} custom boilerplate pattern(s) from {args.boilerplate_file}")
        except (OSError, ValueError) as e:
            print(f"Error: {e}", file=sys.stderr)
            return 1

    if backend in ("aspose", "outlook"):
        if backend == "aspose":
            from . import backend_aspose as B
        else:
            from . import backend_outlook as B
        msg_ext = ".msg" if backend == "outlook" else ".eml"  # Outlook saves MSG, not EML
        store = B.open_live_store() if live else B.open_store(args.input)
        try:
            if list_only:
                for path, folder in B.walk_folders(store):
                    print(f"{path}  [{B.folder_message_count(folder)} msgs]")
                return 0
            if args.include_soft_deleted and backend != "aspose":
                print("Note: --include-soft-deleted is only supported by the aspose backend; ignoring.",
                      file=sys.stderr)
            out_dir = args.out
            os.makedirs(out_dir, exist_ok=True)
            eml_dir = os.path.join(out_dir, "_eml") if args.save_eml else None
            md_dir = os.path.join(out_dir, "_md") if args.save_md else None
            att_root = os.path.join(out_dir, "_attachments") if args.attachments else None
            if eml_dir:
                os.makedirs(eml_dir, exist_ok=True)
            if md_dir:
                os.makedirs(md_dir, exist_ok=True)
            if att_root:
                os.makedirs(att_root, exist_ok=True)
            source_label = "Outlook (live)" if live else os.path.basename(args.input)
            n = 0

            def _already_exported(m) -> bool:
                # Safety net for ids the backend fast path can't see: the same mail from
                # another backend (internet_message_id) and soft-deleted items.
                return is_already_exported(m, skip_ids, skip_message_ids)

            def _handle(raw, m) -> None:
                nonlocal n, next_idx
                stem = mail_stem(next_idx, m)
                next_idx += 1
                m["stem"] = stem
                if args.strip_boilerplate:
                    cleaned, labels = strip_boilerplate(effective_body_text(m), extra_boiler)
                    m["body_clean"] = cleaned
                    m["boilerplate_removed"] = len(labels)
                    m["boilerplate_labels"] = "; ".join(labels)
                if args.attachments:
                    # Traceability: one folder per email -> _attachments/<stem>/
                    per_mail_dir = os.path.join(att_root, stem)
                    try:
                        saved = B.save_attachments(raw, per_mail_dir, prefix="")
                        if saved:
                            rel = [os.path.relpath(s, out_dir).replace(os.sep, "/") for s in saved]
                            m["attachment_dir"] = os.path.relpath(per_mail_dir, out_dir).replace(os.sep, "/")
                            m["attachment_paths"] = rel
                            m["extracted_attachments"] = "; ".join(rel)
                    except Exception as e:
                        m["attachment_error"] = str(e)
                if args.save_eml:
                    try:
                        ep = os.path.join(eml_dir, f"{stem}{msg_ext}")
                        if B.save_eml(raw, ep):
                            m["eml_path"] = os.path.relpath(ep, out_dir).replace(os.sep, "/")
                    except Exception:
                        pass
                if args.save_md:
                    try:
                        md_path = os.path.join(md_dir, f"{stem}.md")
                        atts_rel = m.get("attachment_paths", [])
                        eml_rel = m.get("eml_path", "")
                        write_markdown(m, md_path, source_file=source_label,
                                       attachment_rel_paths=atts_rel, eml_rel_path=eml_rel)
                        m["md_path"] = os.path.relpath(md_path, out_dir).replace(os.sep, "/")
                    except Exception as e:
                        m["md_error"] = str(e)
                rows.append(m)
                n += 1
                if not args.quiet:
                    rec = " [recovered]" if m.get("recovered_from_deleted") else ""
                    line = (f"  [{n}]{rec} {m.get('date','')[:10]} | {m.get('subject','')[:80]}"
                            + (f" | +{m.get('attachment_count',0)} att" if m.get("attachment_count") else ""))
                    try:
                        print(line, flush=True)
                    except UnicodeEncodeError:
                        print(line.encode("ascii", "replace").decode(), flush=True)

            for path, raw, m in _messages(B, live, args, store,
                                          skip_ids=skip_ids or None, stats=stats):
                if not matches_filters(m, args.sender, args.keyword, args.after, args.before):
                    continue
                if _already_exported(m):
                    n_skipped_post += 1
                    continue
                _handle(raw, m)
                if args.limit and n >= args.limit:
                    break
            if args.include_soft_deleted and backend == "aspose" and not (args.limit and n >= args.limit):
                print("  Scanning soft-deleted items...", flush=True)
                for _, raw, m in B.iter_soft_deleted(store, include_body=not args.no_body):
                    if not matches_filters(m, args.sender, args.keyword, args.after, args.before):
                        continue
                    if _already_exported(m):
                        n_skipped_post += 1
                        continue
                    _handle(raw, m)
                    if args.limit and n >= args.limit:
                        break
        finally:
            try:
                if backend == "aspose":
                    store.dispose()
                else:
                    store.close()
            except Exception:
                pass
    else:
        from . import backend_pypff as B
        f = B.open_file(args.input)
        try:
            if list_only:
                for path, _ in B.walk_folders(f):
                    print(path)
                return 0
            for _, _, m in B.iter_messages(f):
                if not matches_filters(m, args.sender, args.keyword, args.after, args.before):
                    continue
                if skip_ids:
                    k = message_key(m)
                    if k and k in skip_ids:
                        stats["skipped"] = stats.get("skipped", 0) + 1
                        continue
                m["stem"] = mail_stem(next_idx, m)
                next_idx += 1
                if args.strip_boilerplate:
                    cleaned, labels = strip_boilerplate(effective_body_text(m), extra_boiler)
                    m["body_clean"] = cleaned
                    m["boilerplate_removed"] = len(labels)
                    m["boilerplate_labels"] = "; ".join(labels)
                rows.append(m)
                if args.limit and len(rows) >= args.limit:
                    break
        finally:
            try:
                f.close()
            except Exception:
                pass

    out_dir = getattr(args, "out", ".")
    os.makedirs(out_dir, exist_ok=True)
    fmt = getattr(args, "format", "json")
    include_body = not getattr(args, "no_body", False)
    all_rows = existing_rows + rows if existing_rows else rows

    if fmt in ("json", "all"):
        write_json(all_rows, os.path.join(out_dir, f"{base}.json"))
    if fmt in ("csv", "all"):
        write_csv(all_rows, os.path.join(out_dir, f"{base}.csv"), include_body=include_body)
    if fmt in ("sqlite", "all"):
        write_sqlite(all_rows, os.path.join(out_dir, f"{base}.db"))
    if live:
        print(f"Live scan {args.date}: {stats.get('folders', 0)} folder(s); "
              f"left out {stats.get('non_mail', 0)} non-mail and {stats.get('outside', 0)} other-day item(s); "
              f"{stats.get('errors', 0)} error(s)")
    if getattr(args, "incremental", False):
        n_skip = stats.get("skipped", 0) + n_skipped_post
        print(f"Exported {len(rows)} new message(s) ({n_skip} already-exported skipped) "
              f"-> {out_dir}  (format={fmt})")
    else:
        print(f"Exported {len(rows)} messages -> {out_dir}  (format={fmt})")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="pstostparser",
                                description="Parse Outlook .pst/.ost: list folders, extract emails, export JSON/CSV/SQLite/EML.")
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    sub = p.add_subparsers(dest="cmd", required=True)

    pi = sub.add_parser("info", help="Show file info + folder tree with counts")
    _add_common(pi)
    pi.set_defaults(func=cmd_info)

    pl = sub.add_parser("list", help="List folder tree")
    _add_common(pl)
    pl.set_defaults(func=cmd_list)

    pe = sub.add_parser("export", help="Export messages to JSON/CSV/SQLite (+MD/EML/attachments)")
    _add_common(pe, input_optional=True)
    _add_filters(pe)
    pe.add_argument("-o", "--out", default="./out", help="Output directory (default: ./out)")
    pe.add_argument("--format", default="json", choices=["json", "csv", "sqlite", "all"])
    pe.add_argument("--no-body", action="store_true", help="Skip bodies (smaller/faster metadata-only export)")
    pe.add_argument("--attachments", action="store_true", help="Extract attachments to out/_attachments/<mail-stem>/ (aspose only)")
    pe.add_argument("--save-eml", action="store_true", help="Save each message as .eml to out/_eml (aspose only)")
    pe.add_argument("--include-soft-deleted", action="store_true",
                    help="Also recover soft-deleted messages (marked [recovered], aspose only)")
    pe.add_argument("--incremental", action="store_true",
                    help="Skip messages already in the output manifest (matched by entry_id) and "
                         "number only new mail continuing from the highest existing stem; "
                         "manifests are merged old+new")
    pe.add_argument("--save-md", action="store_true", help="Save each message as .md to out/_md with front-matter + body")
    pe.add_argument("--strip-boilerplate", action="store_true",
                    help="Remove boilerplate notices (external-sender warnings, confidentiality footers, ...) from Markdown bodies")
    pe.add_argument("--boilerplate-file", default="",
                    help="Text file with extra patterns (literal per line, or /regex/); requires --strip-boilerplate")
    pe.add_argument("--live", action="store_true",
                    help="Read the running Outlook instead of a file; needs --date")
    pe.add_argument("--date", default="", metavar="YYYY-MM-DD",
                    help="With --live: the local calendar day to export (matched on received time)")
    pe.add_argument("--folder", action="append", metavar="NAME",
                    help="With --live: only this folder (a name, or a path under the store such as "
                         "Inbox/Clients); repeat for more")
    pe.add_argument("--name", default="", metavar="STEM",
                    help="With --live: output file stem. Give the existing manifest's name (for example the "
                         "OST file's) with --incremental to continue that export")
    pe.add_argument("--quiet", action="store_true")
    pe.set_defaults(func=cmd_export)
    return p


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    avail = detect_backends()
    if not any(avail.values()) and getattr(args, "backend", "auto") == "auto":
        print("No backend installed yet. Install one first:", file=sys.stderr)
        print("  uv pip install 'Aspose.Email-for-Python-via-NET'   # easiest on Windows", file=sys.stderr)
        print("  uv pip install pywin32  # Outlook COM backend (needs Outlook)", file=sys.stderr)
        return 3
    try:
        return args.func(args)
    except RuntimeError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
