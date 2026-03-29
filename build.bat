@echo off
setlocal enabledelayedexpansion

echo ============================================================
echo   NSE Data Fetcher v3 - Build Pipeline
echo ============================================================
echo.
echo   This script builds:
echo     1. PyInstaller .exe (one-folder mode)
echo     2. Inno Setup installer (if iscc.exe is available)
echo.

cd /d "%~dp0"

REM ── Check Python venv ───────────────────────────────────────
if not exist "venv\Scripts\python.exe" (
    echo [!] Virtual environment not found. Running setup.bat first...
    call setup.bat
    if !errorlevel! neq 0 (
        echo [ERROR] Setup failed.
        pause
        exit /b 1
    )
)
call venv\Scripts\activate.bat

REM ── Install build dependencies ──────────────────────────────
echo [1/4] Installing build dependencies...
pip install -r requirements-dev.txt -q
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install build dependencies.
    pause
    exit /b 1
)
echo       Done.
echo.

REM ── Clean previous build ────────────────────────────────────
echo [2/4] Cleaning previous build...
if exist "build" rmdir /s /q build
if exist "dist\NSE Data Fetcher.exe" del /f "dist\NSE Data Fetcher.exe"
if exist "dist\NSE Data Fetcher" rmdir /s /q "dist\NSE Data Fetcher"
echo       Done.
echo.

REM ── PyInstaller ─────────────────────────────────────────────
echo [3/4] Building with PyInstaller (one-file mode)...
pyinstaller nse_fetcher.spec --noconfirm
if %errorlevel% neq 0 (
    echo [ERROR] PyInstaller build failed. Check output above.
    pause
    exit /b 1
)

echo       PyInstaller build complete.
echo       Output: dist\NSE Data Fetcher.exe
echo.

REM ── Inno Setup (optional) ──────────────────────────────────
echo [4/4] Building installer...

set "ISCC="
REM Try common install locations
if exist "C:\Program Files (x86)\Inno Setup 6\iscc.exe" (
    set "ISCC=C:\Program Files (x86)\Inno Setup 6\iscc.exe"
)
if "%ISCC%"=="" if exist "C:\Program Files\Inno Setup 6\iscc.exe" (
    set "ISCC=C:\Program Files\Inno Setup 6\iscc.exe"
)
REM Per-user install location
if "%ISCC%"=="" if exist "%LOCALAPPDATA%\Programs\Inno Setup 6\iscc.exe" (
    set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\iscc.exe"
)

REM Try PATH (use delayed expansion to read errorlevel correctly)
if "%ISCC%"=="" (
    where iscc.exe >nul 2>&1
    if !errorlevel! equ 0 set "ISCC=iscc.exe"
)

if "%ISCC%"=="" (
    echo [SKIP] Inno Setup not found. Install from https://jrsoftware.org/isinfo.php
    echo        to build the installer. The standalone exe is still available in dist\.
) else (
    "%ISCC%" installer\installer.iss
    if !errorlevel! neq 0 (
        echo [ERROR] Inno Setup compilation failed.
        pause
        exit /b 1
    )
    echo       Installer created: dist\installer\NSE_Data_Fetcher_Setup.exe
)

echo.
echo ============================================================
echo   Build complete!
echo ============================================================
echo.
echo   Standalone exe : dist\NSE Data Fetcher.exe
if not "%ISCC%"=="" echo   Installer       : dist\installer\NSE_Data_Fetcher_Setup.exe
echo.
pause
