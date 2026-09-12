@echo off
setlocal enableextensions

if not "%~1"=="" goto arguments_supplied
if not "%~2"=="" goto arguments_supplied
if not "%1"=="" goto arguments_supplied
echo ExportAccessObjects - Export Access modules, queries, and macros.
echo.
echo Usage:
echo   ExportAccessObjects "database-file" "output-directory"
echo.
echo Example:
echo   ExportAccessObjects "C:\MyWork\Golf.accdb" "C:\MyWork\Golf-Access-Objects"
echo.
echo database-file:
echo   Access database file to export. Supported: .ACCDB, .MDB.
echo.
echo output-directory:
echo   Folder where exported Access objects will be stored.
exit /b 0

:arguments_supplied

if "%~1"=="" (
    echo ERROR: Missing Access database path.
    echo Usage: ExportAccessObjects "database.accdb" "output-directory"
    exit /b 1
)

if "%~2"=="" (
    echo ERROR: Missing output directory.
    echo Usage: ExportAccessObjects "database.accdb" "output-directory"
    exit /b 1
)

if not "%~3"=="" (
    echo ERROR: Too many arguments supplied.
    echo Usage: ExportAccessObjects "database.accdb" "output-directory"
    exit /b 1
)

where python >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python was not found in PATH.
    exit /b 1
)

cd /d "%~dp0"
python "export_access_objects.py" "%~1" "%~2"
