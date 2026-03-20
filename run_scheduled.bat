@echo off
REM NSE Data Fetcher v3 - Scheduled (silent, called by Task Scheduler)

cd /d "%~dp0"

if not exist "venv\Scripts\python.exe" (
    echo [%date% %time%] ERROR: venv not found >> logs\nse_error.log
    exit /b 1
)

venv\Scripts\python.exe -m scripts.main >> logs\nse_output.log 2>&1

if %errorlevel% equ 0 (
    echo [%date% %time%] Scheduled run succeeded >> logs\nse_output.log
    start "" "%~dp0data\NSE50_Data.xlsx"
) else (
    echo [%date% %time%] Scheduled run failed (exit %errorlevel%) >> logs\nse_error.log
)

exit /b %errorlevel%
