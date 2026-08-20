@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Database Manager - build

rem ======================================================================
rem  Builds the application into a single .exe.
rem  Result:  dist\Database Manager.exe
rem  Copy that one file to any Windows machine - no Python and no
rem  libraries need to be installed there.
rem ======================================================================

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

echo.
echo  [1/3] Checking dependencies...
python -m pip install --quiet --disable-pip-version-check pyinstaller openpyxl sv-ttk pillow tkinterdnd2
if errorlevel 1 (
    echo.
    echo  Could not install the dependencies. Check your internet connection.
    echo.
    pause
    exit /b 1
)

echo  [2/3] Preparing the icon...
python "database_manager.py" --make-icon >nul 2>&1

set "ICONARG="
if exist "icon.ico" set "ICONARG=--icon icon.ico --add-data icon.ico;."

echo  [3/3] Building ^(this takes a few minutes^)...
echo.
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "Database Manager" --collect-data sv_ttk --collect-all tkinterdnd2 %ICONARG% "database_manager.py"

if errorlevel 1 (
    echo.
    echo  The build failed.
    echo.
    pause
    exit /b 1
)

rem --- intermediate files are not needed
if exist "build" rmdir /s /q "build"
if exist "Database Manager.spec" del /q "Database Manager.spec"

echo.
echo  ============================================================
echo   Ready:   dist\Database Manager.exe
echo.
echo   Copy that single file to any other Windows machine.
echo   Settings and the error log live in:
echo   %%APPDATA%%\DatabaseManager
echo  ============================================================
echo.
pause
