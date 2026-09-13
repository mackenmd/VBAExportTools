@echo off
:: ====================================================================
::  ExportVBA.bat
::  Universal VBA exporter for Outlook (.OTM) and Excel (.xlsm/.xlam)
:: --------------------------------------------------------------------
::  Usage:
::     ExportVBA "<source file>" "<destination folder>" "commit message|NOCOMMIT" [-init]
:: ====================================================================

setlocal enableextensions enabledelayedexpansion

if not "%~1"=="" goto arguments_supplied
if not "%~2"=="" goto arguments_supplied
if not "%1"=="" goto arguments_supplied
echo ExportVBA - Export VBA modules from Outlook and Excel files.
echo.
echo Usage:
echo   ExportVBA "source-file" "destination-folder" "commit-message|NOCOMMIT" [-init]
echo.
echo Example:
echo   ExportVBA "C:\MyWork\Golf.xlsm" "C:\MyWork\Golf-Modules" "Update email generation"
echo   ExportVBA "C:\MyWork\Golf.xlsm" "C:\MyWork\Golf-Modules" "Initial export" -init
echo.
echo source-file:
echo   Outlook or Excel file containing VBA. Supported: .OTM, .XLSM, .XLAM.
echo.
echo destination-folder:
echo   Folder where exported VBA files will be stored.
echo.
echo commit-message^|NOCOMMIT:
echo   Git commit message, or NOCOMMIT to export without committing.
echo.
echo -init:
echo   Create a new destination folder, initialize its local Git repository, and export into it.
exit /b 0

:arguments_supplied

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

if "%~3"=="" (
    echo ERROR: Missing commit message or NOCOMMIT.
    echo Usage: ExportVBA "C:\path\to\VbaProject.OTM" "D:\ExportFolder" "commit message^|NOCOMMIT" [-init]
    pause
    exit /b 1
)

if not "%5"=="" (
    echo ❌ ERROR: Too many arguments supplied. The optional commit message must be a single argument.
    echo You may have forgotten quotation marks around a multi-word commit message.
    echo Usage: ExportVBA "C:\path\to\VbaProject.OTM" "D:\ExportFolder" "commit message^|NOCOMMIT" [-init]
    pause
    exit /b 1
)

set "SRC_PATH=%~1"
set "DEST_DIR=%~2"
set "COMMIT_MSG=%~3"
set "INIT_SWITCH=%~4"

if not "%INIT_SWITCH%"=="" if /I not "%INIT_SWITCH%"=="-init" (
    echo ERROR: Unknown fourth argument: %INIT_SWITCH%
    echo The only supported fourth argument is -init.
    pause
    exit /b 1
)

where python >nul 2>nul
if errorlevel 1 (
    echo ❌ Python not found in PATH. Please install Python 3.x or add it to PATH.
    pause
    exit /b 1
)

cd /d "%~dp0"
if "%INIT_SWITCH%"=="" (
    python "export_vba_code.py" "%SRC_PATH%" "%DEST_DIR%" "%COMMIT_MSG%"
) else (
    python "export_vba_code.py" "%SRC_PATH%" "%DEST_DIR%" "%COMMIT_MSG%" "%INIT_SWITCH%"
)
pause
