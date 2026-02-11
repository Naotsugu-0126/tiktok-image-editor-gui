@echo off
setlocal
cd /d "%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\setup_and_start_gui.ps1"
if errorlevel 1 (
  echo.
  echo Failed to start GUI. Check error log above.
  pause
  exit /b 1
)

endlocal
exit /b 0
