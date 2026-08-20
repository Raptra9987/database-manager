@echo off
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>&1
if errorlevel 1 (
    echo.
    echo  Python was not found.
    echo  Install it from https://python.org
    echo  ^(tick "Add Python to PATH" during setup^).
    echo.
    pause
    exit /b 1
)

rem --- Install the dependencies on first run.
rem     tkinterdnd2 is optional: without it only drag-and-drop is unavailable.
python -c "import openpyxl, sv_ttk, PIL" >nul 2>&1
if errorlevel 1 (
    echo  Installing dependencies, this takes a moment...
    python -m pip install --quiet --disable-pip-version-check openpyxl sv-ttk pillow tkinterdnd2
)

start "" pythonw "database_manager.py"
