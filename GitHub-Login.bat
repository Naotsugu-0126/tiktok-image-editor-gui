@echo off
setlocal
cd /d "%~dp0"

where gh >nul 2>nul
if errorlevel 1 (
  echo gh command not found.
  echo Please run once: winget install --id GitHub.cli -e
  pause
  exit /b 1
)

echo Opening GitHub login flow...
gh auth login --hostname github.com --git-protocol https --web --skip-ssh-key
if errorlevel 1 (
  echo Login failed or canceled.
  pause
  exit /b 1
)

gh auth setup-git >nul 2>nul
echo Login completed.
gh auth status
pause
endlocal
exit /b 0
