@echo off
setlocal enabledelayedexpansion

echo ============================================================
echo   Stock Market: NSE50 Data Updater
echo ============================================================
echo.

cd /d "%~dp0"

REM ── Environment checks ──────────────────────────────────────
if not exist "venv\Scripts\python.exe" (
    if exist "venv\" (
        echo [!] Found broken virtual environment. Cleaning up...
        rmdir /s /q venv
    )
    echo [!] Virtual environment not found.
    echo [!] Running first-time setup...
    echo.
    call setup.bat
    if !errorlevel! neq 0 (
        echo.
        echo [ERROR] Setup failed. Please check your Python installation.
        pause
        exit /b 1
    )
)

REM ── Validate Python works ───────────────────────────────────
venv\Scripts\python.exe --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python environment is broken. Deleting venv and re-running setup...
    rmdir /s /q venv
    call setup.bat
    if !errorlevel! neq 0 (
        echo [ERROR] Setup failed. Please reinstall Python.
        pause
        exit /b 1
    )
)

for /f "tokens=*" %%v in ('venv\Scripts\python.exe --version 2^>^&1') do set PYVER=%%v
echo [OK] %PYVER% ready

REM ── Ensure directories exist ────────────────────────────────
if not exist data    mkdir data
if not exist logs    mkdir logs
if not exist archive mkdir archive

echo [OK] Directories verified
echo.

REM ── Run the updater ─────────────────────────────────────────
echo Fetching latest stock market data...
echo.

venv\Scripts\python.exe -m scripts.main %*

if %errorlevel% neq 0 (
    echo.
    echo ============================================================
    echo   [ERROR] Update failed ^(exit code: %errorlevel%^)
    echo   Check logs\nse_fetcher.log for details.
    echo ============================================================
) else (
    echo.
    echo ============================================================
    echo   Update complete! Excel file is up to date.
    echo ============================================================
)

echo.
pause
