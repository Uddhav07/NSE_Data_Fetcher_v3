@echo off
REM NSE Data Fetcher v3 - Manual Run

echo ============================================================
echo   NSE Data Fetcher v3 - Manual Run
echo ============================================================
echo.

cd /d "%~dp0"

if not exist "venv\Scripts\python.exe" (
    echo ERROR: Virtual environment not found.
    echo Please run setup.bat first.
    pause
    exit /b 1
)

echo Using venv: %cd%\venv
echo.

REM Pass any extra CLI args through (e.g. --full-refresh, --no-futures)
venv\Scripts\python.exe -m scripts.main %*

if %errorlevel% neq 0 (
    echo.
    echo Script exited with error code %errorlevel%.
) else (
    echo.
    echo Done.
)
echo.
pause
