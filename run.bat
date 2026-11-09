@echo off
rem Update the Excel file (run after 15:35 IST with the workbook closed). Extra flags are passed through.
cd /d "%~dp0"
if not exist venv\Scripts\python.exe call setup.bat
venv\Scripts\python.exe -m scripts.main %*
echo.
pause
