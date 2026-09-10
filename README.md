# VBAExportTools

Utilities for exporting Microsoft Office VBA source code to files suitable
for version control with Git.

## Current Status

The existing Python exporter successfully exports VBA modules from an Outlook
`.OTM` file using `oletools.olevba.VBA_Parser`.

The project is currently being generalized so that the source VBA file and
destination folder are supplied on the command line.

Outlook `.OTM` export is the existing known-working implementation.

Excel `.xlsm` / `.xlam` support is planned and still needs to be verified.

Microsoft Access VBA export is a future goal after Outlook and Excel exporting
have been completed and thoroughly tested.

## Intended Usage

```text
ExportVBA "<source file>" "<destination folder>" [commit message|NOCOMMIT]
```

Example:

```text
ExportVBA "C:\Users\mm103\AppData\Roaming\Microsoft\Outlook\VbaProject.OTM" "D:\OneDrive\DufferPools\EmailMover\Modules" "Refactored rules"
```

## Intended Features

- Export VBA modules from Outlook `.OTM` files
- Export VBA modules from Excel `.xlsm` / `.xlam` files
- Automatically determine the source type from the file extension
- Export modules to a specified destination folder
- Produce stable output suitable for Git comparisons
- Optionally commit exported changes to Git
- Support `NOCOMMIT` to export without committing

## Future

After Outlook and Excel exporting are complete and verified, investigate
support for exporting VBA from Microsoft Access databases.
