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

gh auth status >nul 2>nul
if errorlevel 1 (
  echo You are not logged in to GitHub.
  echo Run GitHub-Login.bat first.
  pause
  exit /b 1
)

set /p REPO_NAME=Repository name (example: tiktok-image-editor-v2): 
if "%REPO_NAME%"=="" (
  echo Repository name is required.
  pause
  exit /b 1
)

set /p VISIBILITY=Visibility [private/public] ^(default: private^): 
if /i "%VISIBILITY%"=="public" (
  set VIS_FLAG=--public
) else (
  set VIS_FLAG=--private
)

for /f "usebackq delims=" %%U in (`gh api user --jq .login`) do set OWNER=%%U
if "%OWNER%"=="" (
  echo Could not resolve GitHub user.
  pause
  exit /b 1
)

echo Creating repository: %OWNER%/%REPO_NAME%
gh repo create "%OWNER%/%REPO_NAME%" %VIS_FLAG% --confirm
if errorlevel 1 (
  echo Repository creation failed.
  pause
  exit /b 1
)

set /p CONNECT_LOCAL=Connect this local folder to new repo? [y/N]: 
if /i "%CONNECT_LOCAL%"=="y" (
  git rev-parse --is-inside-work-tree >nul 2>nul
  if errorlevel 1 (
    echo Current folder is not a git repository. Skipping connect.
  ) else (
    git remote remove newrepo >nul 2>nul
    git remote add newrepo "https://github.com/%OWNER%/%REPO_NAME%.git"
    echo Added remote: newrepo
    echo Push command:
    echo   git push -u newrepo HEAD
  )
)

echo Done.
pause
endlocal
exit /b 0
