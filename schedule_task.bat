@echo off
REM NSE Data Fetcher v3 - Windows Task Scheduler Setup

echo ============================================================
echo   NSE Data Fetcher v3 - Schedule Daily Task
echo ============================================================
echo.

set "SCRIPT_DIR=%~dp0"
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

if not exist "%SCRIPT_DIR%\venv\Scripts\python.exe" (
    echo ERROR: Virtual environment not found.
    echo Please run setup.bat first.
    pause
    exit /b 1
)

set SCHEDULE_TIME=17:30

echo Project dir : %SCRIPT_DIR%
echo Python      : venv\Scripts\python.exe
echo Schedule    : Daily at %SCHEDULE_TIME%
echo.

REM Remove existing task if present
schtasks /query /tn "NSE50_DataFetcher_v3" >nul 2>&1
if %errorlevel% equ 0 (
    echo Removing previous scheduled task...
    schtasks /delete /tn "NSE50_DataFetcher_v3" /f
)

schtasks /create ^
    /tn "NSE50_DataFetcher_v3" ^
    /tr "\"%SCRIPT_DIR%\run_scheduled.bat\"" ^
    /sc daily ^
    /st %SCHEDULE_TIME% ^
    /rl highest ^
    /f

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Could not create task. Try running as Administrator.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   Task created: NSE50_DataFetcher_v3
echo   Runs daily at %SCHEDULE_TIME%
echo ============================================================
echo.
echo   Manage via: Win+R → taskschd.msc → NSE50_DataFetcher_v3
echo   Logs: logs\nse_fetcher.log
echo.
pause
