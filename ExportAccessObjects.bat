@echo off
setlocal enableextensions

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
