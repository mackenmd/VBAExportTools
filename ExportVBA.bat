@echo off
:: ====================================================================
::  ExportVBA.bat
::  Universal VBA exporter for Outlook (.OTM) and Excel (.xlsm/.xlam)
:: --------------------------------------------------------------------
::  Usage:
::     ExportVBA "<source file>" "<destination folder>" [commit message|nocommit]
:: ====================================================================

setlocal enableextensions enabledelayedexpansion

if "%~1"=="" (
    echo ❌ ERROR: Missing source VBA file path (.OTM or .XLSM).
    echo Usage: ExportVBA "C:\path\to\VbaProject.OTM" "D:\ExportFolder" [commit message|nocommit]
    pause
    exit /b 1
)

if "%~2"=="" (
    echo ❌ ERROR: Missing destination export folder.
    echo Usage: ExportVBA "C:\path\to\VbaProject.OTM" "D:\ExportFolder" [commit message|nocommit]
    pause
    exit /b 1
)

set "SRC_PATH=%~1"
set "DEST_DIR=%~2"
shift
shift
set "COMMIT_MSG=%*"

where python >nul 2>nul
if errorlevel 1 (
    echo ❌ Python not found in PATH. Please install Python 3.x or add it to PATH.
    pause
    exit /b 1
)

cd /d "%~dp0"
python "export_vba_code.py" "%SRC_PATH%" "%DEST_DIR%" %COMMIT_MSG%
pause
