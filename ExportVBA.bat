@echo off
:: ====================================================================
::  ExportVBA.bat
::  Universal VBA exporter for Outlook (.OTM) and Excel (.xlsm/.xlam)
:: --------------------------------------------------------------------
::  Usage:
::     ExportVBA "<source file>" "<destination folder>" ["commit message" or NOCOMMIT]
:: ====================================================================

setlocal enableextensions enabledelayedexpansion

if "%~1"=="" (
    echo ❌ ERROR: Missing source VBA file path. Supported extensions: .OTM, .XLSM, or .XLAM.
    echo Usage: ExportVBA "C:\path\to\VbaProject.OTM" "D:\ExportFolder" ["commit message" or NOCOMMIT]
    pause
    exit /b 1
)

if "%~2"=="" (
    echo ❌ ERROR: Missing destination export folder.
    echo Usage: ExportVBA "C:\path\to\VbaProject.OTM" "D:\ExportFolder" ["commit message" or NOCOMMIT]
    pause
    exit /b 1
)

if not "%4"=="" (
    echo ❌ ERROR: Too many arguments supplied. The optional commit message must be a single argument.
    echo You may have forgotten quotation marks around a multi-word commit message.
    echo Usage: ExportVBA "C:\path\to\VbaProject.OTM" "D:\ExportFolder" ["commit message" or NOCOMMIT]
    pause
    exit /b 1
)

set "SRC_PATH=%~1"
set "DEST_DIR=%~2"
set "COMMIT_MSG=%~3"

where python >nul 2>nul
if errorlevel 1 (
    echo ❌ Python not found in PATH. Please install Python 3.x or add it to PATH.
    pause
    exit /b 1
)

cd /d "%~dp0"
if "%COMMIT_MSG%"=="" (
    python "export_vba_code.py" "%SRC_PATH%" "%DEST_DIR%"
) else (
    python "export_vba_code.py" "%SRC_PATH%" "%DEST_DIR%" "%COMMIT_MSG%"
)
pause
