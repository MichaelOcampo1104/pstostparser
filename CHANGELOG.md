# Changelog

Changes to pstostparser, newest first.

## Unreleased

### Added

- `export --live --date YYYY-MM-DD` reads the mail of the running Outlook for one local day, with no `.pst` or `.ost` file. `--folder` narrows it to some folders.
- `--name` (with `--live`) sets the output file name. With `--incremental` it continues an existing export: give the base name of its `.json`, `.csv` and `.db` files.
- `received_date` column (local wall-clock time) in JSON, CSV and SQLite exports.
- README: the command for continuing an export into an existing folder.

### Changed

- Incremental exports skip a message already in the manifest when its entry ID matches, or when its internet message ID and received time to the second both match. The internet message ID alone is not enough: one ID can belong to many mails, such as a recurring digest.
- `.gitignore` excludes `inputs/`, `*.pst` and `*.ost`, so mail data stays out of the repository.
