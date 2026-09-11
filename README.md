# VBAExportTools

VBAExportTools exports Microsoft Office VBA source code into ordinary files
suitable for Git and other version-control systems.

## Status

Outlook `.OTM` exporting has been tested and verified against a real Outlook
VBA project. Excel `.xlsm` and `.xlam` exporting has also been tested and
verified. Microsoft Access VBA export is planned future work and is not
currently supported.

## Usage

```text
ExportVBA "<source file>" "<destination folder>" ["commit message" or NOCOMMIT]
```

The source file and destination folder are required. Enclose a multi-word
commit message in quotes. `NOCOMMIT` is case-insensitive and performs the
export without a Git commit. If the optional third argument is omitted, the
exporter uses its timestamped default Git commit message.

`ExportVBA.bat` rejects more than three command-line arguments. This helps
catch a forgotten pair of quotes around a multi-word commit message.

Verified Outlook example:

```text
ExportVBA "C:\Users\mm103\AppData\Roaming\Microsoft\Outlook\VbaProject.OTM" "D:\OneDrive\DufferPools\EmailMover\Modules" NOCOMMIT
```

Example with a Git commit message:

```text
ExportVBA "C:\Users\mm103\AppData\Roaming\Microsoft\Outlook\VbaProject.OTM" "D:\OneDrive\DufferPools\EmailMover\Modules" "Refactored rules"
```

## Destination-directory contract

The destination is an exporter-managed output directory. Every top-level file
in it is considered exporter-owned and may be backed up, removed, or replaced
during a successful export. Do not place unrelated top-level files there.

Existing subdirectories and their contents are left untouched. Parser-generated
output such as `VBA_P-code.txt` may be exported; output is not restricted to
specific VBA filename extensions.

## Safety behavior

The complete export is staged in a temporary directory before the destination
is modified. Existing top-level destination files are temporarily backed up
before replacement.

- Extraction, staging, or backup failures leave the destination untouched.
- A replacement failure triggers restoration of the original files.
- If restoration cannot be completed, the backup is retained and its location
  is reported.
- An empty extraction leaves the destination untouched.

## Git behavior

With a commit message, the exporter stages and commits the resulting
destination changes. With `NOCOMMIT`, it exports without committing.

Git tracks file-content changes: regenerated files with unchanged contents do
not become Git changes merely because their filesystem timestamps changed.
Files intentionally excluded by the destination repository's `.gitignore`,
such as `VBA_P-code.txt`, remain excluded normally.

## Requirements

- Python 3
- `oletools`
- Git, when automatic Git commits are desired

`pywin32` is not a requirement.
