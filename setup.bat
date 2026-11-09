@echo off
rem One-time setup for Windows: creates venv and installs packages. Double-click to run.
cd /d "%~dp0"
python --version >nul 2>&1
if errorlevel 1 (
  echo Python was not found. Install Python 3.9+ from https://www.python.org/downloads/windows/
  echo and tick "Add python.exe to PATH", then run this again.
  pause
  exit /b 1
)
python -m venv venv
venv\Scripts\python.exe -m pip install --quiet --upgrade pip
venv\Scripts\python.exe -m pip install --quiet -r requirements.txt
if errorlevel 1 (
  echo Package install failed. See the messages above.
  pause
  exit /b 1
)
venv\Scripts\python.exe -m pip install --quiet -r requirements-optional.txt
if errorlevel 1 echo Note: optional package not installed - Moon Nakshatra column will stay blank.
echo.
echo Setup complete. Double-click run.bat to update the data.
pause
