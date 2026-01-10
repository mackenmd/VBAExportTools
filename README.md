# VBAExportTools

Python and batch scripts for exporting Outlook and Excel VBA modules
to version control (GitHub).

## Usage

```bash
ExportVBA "C:\Users\mm103\AppData\Roaming\Microsoft\Outlook\VbaProject.OTM" "D:\OneDrive\DufferPools\EmailMover\Modules" "Refactored rules"
```

### Features
- Detects Outlook .OTM or Excel .xlsm source files automatically
- Exports all modules into a Git repository folder
- Optionally performs a Git commit

---
