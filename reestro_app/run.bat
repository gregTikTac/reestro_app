@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

echo ========================================
echo  EGRN Parser - run from sources
echo ========================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found in PATH.
    echo Install Python 3.11+ and add to PATH.
    pause
    exit /b 1
)

if not exist "venv\Scripts\python.exe" (
    echo [1/3] Creating venv...
    python -m venv venv
    if errorlevel 1 goto fail
)

echo [2/3] Installing dependencies...
call venv\Scripts\activate.bat
if errorlevel 1 goto fail

pip install -r requirements.txt -q
if errorlevel 1 goto fail

echo [3/3] Starting GUI...
echo.
echo  Folder:  %CD%
echo  Engine:  ..\reestro\
echo.

python app_entry.py
set EXITCODE=%ERRORLEVEL%
if not "%EXITCODE%"=="0" (
    echo.
    echo Exit code: %EXITCODE%
    pause
)
exit /b %EXITCODE%

:fail
echo.
echo Setup failed. See messages above.
pause
exit /b 1
