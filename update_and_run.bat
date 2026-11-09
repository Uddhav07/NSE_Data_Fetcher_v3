@echo off
rem 1) pull the latest version, 2) update the Excel data, 3) open the workbook. Extra flags are passed through.
rem The whole block is read before it runs, so a git pull that rewrites this file cannot break it.
cd /d "%~dp0"
(
  if exist .git (
    git --version >nul 2>&1
    if errorlevel 1 (
      echo Git is not installed - skipping the update step.
    ) else (
      git diff --quiet HEAD
      if errorlevel 1 (
        echo You have local changes to project files - skipping the update so nothing is overwritten.
      ) else (
        echo Checking GitHub for a newer version...
        git pull --ff-only
        if errorlevel 1 echo Could not update - running the current version.
      )
    )
  ) else (
    echo Not a git clone - skipping the update step.
  )
  if not exist venv\Scripts\python.exe call setup.bat
  venv\Scripts\python.exe -m pip install --disable-pip-version-check -q -r requirements.txt
  venv\Scripts\python.exe -m scripts.main --open %*
  echo.
  pause
  exit /b
)
