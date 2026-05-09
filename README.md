# PEC IMAP Backup

A desktop backup utility written in Python and PySide6 for IMAP and Italian PEC mailboxes.

The application saves each message as a complete RAW `.eml` file, creates an Excel `.xlsx` index, and can optionally mark backed-up IMAP messages as `\Deleted` only after the local backup has completed and the user confirms it.

## Main features

- Saves messages as byte-for-byte RAW `.eml` files, preserving body text, attachments, MIME structure, signatures, and attached messages.
- Filters messages by cutoff date: only messages older than the selected date are processed.
- Optional size filter: process only messages larger than a chosen B/KB/MB/GB threshold.
- Loads real IMAP folder names from the server and lets you select the folders to process.
- Deduplicates messages using SHA-256, so interrupted backups can be resumed safely.
- Creates `indice_backup_pec.xlsx` with message metadata, local path, SHA-256, and PEC-specific fields.
- Lets you choose which columns are shown in the visible Excel sheet.
- Keeps a hidden internal Excel sheet with the technical fields needed for deduplication and deletion tracking.
- Parses PEC metadata from `daticert.xml` / `postacert.eml` to identify the real sender/counterparty instead of generic providers such as `posta-certificata@...`.
- Lets you configure the `.eml` file-name template and a maximum file-name length.
- Sanitizes saved file and folder names for Windows/NTFS, Linux, and macOS by replacing unsafe characters such as `/`, `\`, `"`, `:`, `*`, `?`, `<`, `>`, `|` and control characters.
- Preserves a short SHA-256 hash in the saved `.eml` file name to reduce collision risk when long names are truncated.
- Provides a preview table; double-click a row to inspect the RAW source of the message.
- Supports cooperative Stop for long-running operations.
- Provides a clearable GUI log and an optional log/debug file saved to a configurable folder.
- Supports Italian and English UI translations through JSON files in `i18n/`.
- Includes application icons in `resources/`.

## Installation

On Linux/macOS:

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python imap_backup.py
```

On Windows:

```bat
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python imap_backup.py
```

## Recommended workflow

1. Enter host, port, security mode, username, and password.
2. Choose the interface language.
3. Press **Test / list folders** to load the real mailbox names from the IMAP server.
4. Tick the folders you want to process.
5. Select **Save messages before**.
6. Enable the size filter only if needed.
7. Use **Configure export/log…** to choose:
   - visible columns in `indice_backup_pec.xlsx`;
   - the `.eml` file-name template;
   - the maximum `.eml` file-name length;
   - whether to save the GUI/debug log to a file and where.
8. Press **Preview** to check the messages that will be processed.
9. Press **Run backup**.
10. If the backup completes without errors, the application may ask whether to set `\Deleted` on the saved IMAP messages.
11. `EXPUNGE` is separate and must be explicitly enabled because it permanently deletes server-side messages already marked as `\Deleted`.

## Advanced export and file-name configuration

Open **Configure export/log…** to adjust export and naming behavior.

### XLSX columns

The visible sheet named `indice` is regenerated with only the selected columns. The workbook also contains a hidden `_internal` sheet that stores all technical columns required by the program, including SHA-256, status, relative path, and deletion metadata.

This means you can keep the Excel file readable while the application still has the data it needs to resume backups and update deletion status safely.

### File-name template

The default `.eml` template is:

```text
{date} - {party} - {subject} __{hash}
```

Available fields are:

- `{date}`: message date formatted for file names.
- `{party}`: PEC counterparty/real sender when available, otherwise the From header.
- `{subject}`: PEC subject when available, otherwise the message Subject header.
- `{folder}`: IMAP folder name.
- `{uid}`: IMAP UID.
- `{direction}`: inferred direction, for example incoming or outgoing.
- `{pec_type}`: PEC type from `daticert.xml` when available.
- `{message_id}`: Message-ID header.
- `{hash}`: first 12 characters of the message SHA-256.
- `{full_hash}`: full SHA-256.

The generated name is sanitized after template expansion, so unsafe characters from mail subjects, senders, folders, or message IDs are replaced. The final `.eml` name is also truncated to the configured maximum length, including the `.eml` extension.

If the template does not end with `{hash}`, the application appends the short hash automatically to keep names unique and stable.

## Log/debug file

The GUI log is always shown in the application. In **Configure export/log…** you can also enable saving the same log to disk, including detailed IMAP debug lines when **Detailed IMAP debug** is enabled.

The log file is created lazily in the selected folder with a name like:

```text
imap_backup_YYYYMMDD_HHMMSS.log
```

Passwords are never printed in the IMAP debug output.

## Languages

Translations are stored as JSON files:

- `i18n/it.json`
- `i18n/en.json`

To correct or extend a translation, edit the value associated with an existing key. Keep keys unchanged.

## IMAP deletion notes

Backup and deletion are separate phases. The program does not mark messages as `\Deleted` while saving them. The `\Deleted` step is proposed only at the end, after backup and verification, and requires explicit confirmation.

`EXPUNGE` permanently removes from the server messages already marked as `\Deleted`. It is disabled by default.

## Debug and crash options

Startup options:

```bash
python imap_backup.py --safe-mode
python imap_backup.py --reset-settings
python imap_backup.py --qt-debug-plugins
python imap_backup.py --force-xcb
python imap_backup.py --force-wayland
```

The early crash log is written to `pec_imap_backup_crash.log` by default. You can override the path with:

```bash
PEC_IMAP_BACKUP_CRASH_LOG=/path/to/log.txt python imap_backup.py
```

## Password storage

If **save password locally** is enabled, the password is stored through QSettings using simple base64/XOR obfuscation. This is not strong encryption. Use it only on a trusted machine.
