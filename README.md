# pstostparser

Parse Outlook `.pst` and `.ost` files on Windows: list folders, extract emails, export to JSON / CSV / SQLite, dump `.eml` + attachments.

Both formats share the same underlying store, so one code path handles both.

## Setup (Windows + uv)

You already have `uv` with Python 3.12. From this folder:

```powershell
# create env + install the tool editable
uv venv
uv pip install -e .

# install a parsing backend (pick one; aspose is easiest on Windows)
uv pip install "Aspose.Email-for-Python-via-NET"
# optional free/open-source alternative (needs a working wheel):
# uv pip install pypff
```

Check:

```powershell
uv run pstostparser --help
```

## Usage

```powershell
# folder tree + counts
uv run pstostparser info "C:\path\to\mailbox.pst"
uv run pstostparser info "C:\path\to\mailbox.ost" --backend aspose

# list folders only
uv run pstostparser list "C:\path\to\mailbox.pst"

# export metadata + bodies to JSON
uv run pstostparser export "C:\path\to\mailbox.pst" -o ./out --format json

# CSV (Excel-friendly), SQLite (large mailboxes), or everything
uv run pstostparser export mailbox.ost -o ./out --format csv --no-body
uv run pstostparser export mailbox.pst -o ./out --format sqlite
uv run pstostparser export mailbox.pst -o ./out --format all --attachments --save-eml

# Markdown per email + traceable per-email attachment folders
uv run pstostparser export mailbox.pst -o ./out --format all --save-md --attachments

# Markdown with boilerplate notices stripped (external-sender warnings,
# confidentiality footers, AV lines, ...) + your own extra patterns
uv run pstostparser export mailbox.pst -o ./out --format all --save-md --attachments --strip-boilerplate
uv run pstostparser export mailbox.pst -o ./out --save-md --strip-boilerplate --boilerplate-file ./boilerplate.txt

# Also recover soft-deleted messages (marked [recovered] in logs,
# recovered_from_deleted in manifest/Markdown)
uv run pstostparser export mailbox.pst -o ./out --format all --save-md --attachments --include-soft-deleted

# Incremental: PST/OST grew with new mail? Only export what's new.
# Already-exported messages (matched by entry_id against the output manifest)
# are skipped before extraction; new mail is numbered continuing from the
# highest existing stem and manifests are merged old+new:
uv run pstostparser export mailbox.pst -o ./out --format all --save-md --attachments --strip-boilerplate --incremental

# Outlook backend: reads through Outlook/MAPI exactly as Outlook sees it
# (needs Outlook open or installed; use when counts don't match `info`)
uv run pstostparser export mailbox.pst -o ./out --format all --save-md --attachments --backend outlook

# search / filter
uv run pstostparser export mailbox.pst -o ./out --format json `
  --sender "boss@company.com" `
  --keyword "invoice" `
  --after "2024-01-01" --before "2024-12-31T23:59:59" `
  --limit 200
```

Outputs in `./out`:

- `<mailbox>.json` / `.csv` / `.db` (SQLite table `emails` — manifest with `stem`, `md_path`, `attachment_dir`, ... )
- `_md/<stem>.md` (with `--save-md`: YAML front-matter + header + body; attachment links point at the mail's folder)
- `_attachments/<stem>/filename.pdf` (with `--attachments`: one folder per email, named `<index>_<date>_<subject-slug>`)
- `_eml/<stem>.eml` (with `--save-eml`, aspose only)

## Live Outlook: one day, no file

`--live` reads the Outlook that is already open (it starts Outlook if it isn't running) instead of a `.pst`/`.ost`, and exports only the mail received on one local calendar day:

```powershell
uv run pstostparser export --live --date 2026-10-12 -o ./out --format all --save-md --attachments

# only some folders: a folder name, or a path under the account
uv run pstostparser export --live --date 2026-10-12 -o ./out --folder Inbox --folder "Inbox/Clients" --format json
```

- Works with classic Outlook for Windows. The new Outlook app is not supported.
- Every mail store in Outlook is searched (accounts and data files), except Search Folders, which only repeat items held in other folders. Calendar, contact, task and note folders are not read.
- Only mail items are kept. Meeting requests, NDR bounces and other non-mail items are left out.
- The day is the local calendar day the mail arrived on. `received_date` holds that local time with no zone suffix. `date` is unchanged: SentOn first, as in the other Outlook exports.
- `received_date` is filled for Outlook exports, file or live. Aspose and pypff rows leave it blank.
- `--after`/`--before` do not apply with `--live`; use `--date`.
- Output files are named `outlook-live_<date>.*`, or the name given with `--name`. The export never quits Outlook.
- Older dates take longer, because each folder is read from its newest item back to that day: about 40 seconds for a day three months back on this mailbox, a few seconds for recent days.

### Continue an existing export

`--name` sets the output name, and `--incremental` continues an existing export. Give the base name of that export's `.json`, `.csv` and `.db` files. For an export made from `mailbox.ost`, that is `mailbox`. Numbering picks up after the highest existing stem, and mail the export already holds is skipped. Run this from the project folder, replacing the placeholders:

```powershell
uv run pstostparser export --live --date 2026-10-12 -o "D:\path\to\export-folder" --name "mailbox" --incremental --format all --save-md --attachments --strip-boilerplate
```

- Use the same options as the original export. `--strip-boilerplate` matches exports made with it.
- The `.json`, `.csv` and `.db` files are rewritten with the old and new rows, so keep a copy first.
- For another day, change `--date`. Running the command again adds only mail that isn't in the export yet.

## Boilerplate stripping

`--strip-boilerplate` removes machine-added notice blocks from the Markdown
bodies (the raw `body_text`/`body_html` stay untouched in the manifest, and the
cleaned copy is stored as `body_clean`). Each `.md` records
`boilerplate_removed: N` plus which patterns matched, with a one-line note at
the end of the body. Built-ins cover: external-sender warnings ("This email
comes from outside the organization..."), generic CAUTION/External headers,
confidentiality disclaimers, antivirus footers, "Sent from my iPhone",
print-environment notices, and unsubscribe lines.

`--boilerplate-file` adds your own patterns, one per line: plain text matches
case-insensitively, or wrap in slashes for regex. Blank lines and `#` comments
are ignored:

```text
# boilerplate.txt
Kindly acknowledge receipt
/Do not (click|open)[^.]{0,200}unexpected[^.]{0,200}\./
```

## Notes

- **Backends:** `aspose` (default, no Outlook needed) opens PST/OST directly;
  `outlook` (`--backend outlook`, needs `uv pip install pywin32` + Outlook)
  reads through MAPI exactly as Outlook displays it — use it if `info` counts
  don't match Outlook. With the Outlook backend, `--save-eml` saves `.msg`
  files (Outlook has no native EML export) and `--include-soft-deleted` does
  not apply. `pypff` (libyal) is a third option where a working build exists —
  note PyPI's `pypff` 1.x is an unrelated package, do not install it.
- **50-item evaluation cap:** unlicensed Aspose returns at most 50 messages
  per folder-listing call. The parser pages through automatically (`info`
  reports true `content_count` totals), so exports are complete — just slower
  on huge files. A licensed Aspose removes the cap.
- **Long runs:** for multi-GB files, launch detached so a terminal timeout
  can't kill the export, with `PYTHONUTF8=1` so console logging never crashes
  on emoji (file contents are always written UTF-8 regardless):

```powershell
# D:\run_export.bat
@echo off
set PYTHONUTF8=1
"C:\path\to\.venv\Scripts\python.exe" -m pstostparser.cli export "D:\Outlook\mail.ost" -o "D:\out" --format all --save-md --attachments --strip-boilerplate --backend aspose > "D:\export.log" 2>&1
```

```powershell
Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{CommandLine = 'cmd /c "D:\run_export.bat"'}
Get-Content D:\export.log -Tail 5   # poll progress
```

- **Backend `aspose`** (`aspose-email`, commercial, evaluation limits may apply) is the path of least resistance on Windows and opens both PST and OST via `PersonalStorage.from_file()`.
- **Backend `pypff`** (libyal bindings, free) is better for forensics but has no official Windows wheels on some Python versions — use it if `import pypff` works in your env.
- OST files may be locked while Outlook is running — close Outlook or copy the file first.
- Large mailboxes: prefer `--format sqlite --no-body` first for a fast inventory, then re-export the subset you need with bodies + attachments.
