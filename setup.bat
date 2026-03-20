@echo off
REM NSE Data Fetcher v3 - Setup Script (creates Python venv and installs deps)

echo ============================================================
echo   NSE Data Fetcher v3 - Setup
echo ============================================================
echo.

cd /d "%~dp0"

REM ── Check Python ────────────────────────────────────────────
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH.
    echo Download from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during install.
    pause
    exit /b 1
)
echo Python found:
python --version
echo.

REM ── Create / reuse venv ─────────────────────────────────────
if exist "venv\" (
    echo Virtual environment already exists.
    choice /C YN /M "Recreate it from scratch? (Y/N)"
    if errorlevel 2 goto :activate
    echo Removing old venv...
    rmdir /s /q venv
)

echo Creating virtual environment...
python -m venv venv
if %errorlevel% neq 0 (
    echo ERROR: Failed to create venv. Make sure Python includes the venv module.
    pause
    exit /b 1
)
echo Virtual environment created.
echo.

:activate
call venv\Scripts\activate.bat
echo Activated venv.
echo.

REM ── Install dependencies ────────────────────────────────────
echo Upgrading pip...
python -m pip install --upgrade pip -q
echo Installing packages from requirements.txt...
pip install -r requirements.txt -q
if %errorlevel% neq 0 (
    echo ERROR: Package installation failed. Check your internet connection.
    pause
    exit /b 1
)

REM ── Ensure runtime directories exist ────────────────────────
if not exist data   mkdir data
if not exist logs   mkdir logs
if not exist archive mkdir archive

echo.
echo ============================================================
echo   Setup complete!
echo ============================================================
echo.
echo   Virtual env : .\venv\
echo   Python      : venv\Scripts\python.exe
echo   Next step   : run_manual.bat
echo.
pause
