@echo off
title Zee AI - update
cd /d "%~dp0"
set LOG=update-log.txt
echo ===== Zee AI update %date% %time% ===== > %LOG%
echo ============================================================
echo   Zee AI - apply the latest changes
echo ============================================================
echo.

echo [1/4] Installing any new packages, running the tests and building (about 1 minute)...
call npm install --no-audit --no-fund >> %LOG% 2>&1
if errorlevel 1 (
  echo NPM INSTALL FAILED >> %LOG%
  echo [X] Installing packages failed - check your internet connection, or send update-log.txt to Claude.
  pause
  exit /b 1
)
call npm run build >> %LOG% 2>&1
if errorlevel 1 (
  echo BUILD FAILED >> %LOG%
  echo [X] Tests or build failed - nothing was published. Send update-log.txt to Claude.
  pause
  exit /b 1
)
echo build OK >> %LOG%
echo [OK] Build succeeded.
echo.

if not exist .git goto localonly
if not exist manifest.prod.xml goto localonly

echo [2/4] Uploading to GitHub (your hosted add-in)...
git add -A >> %LOG% 2>&1
git commit -m "Zee AI update %date% %time%" >> %LOG% 2>&1
git push origin main >> %LOG% 2>&1
if errorlevel 1 (
  echo PUSH FAILED >> %LOG%
  echo [X] Upload failed - see update-log.txt. A GitHub sign-in window may have been closed.
  pause
  exit /b 1
)
echo push OK >> %LOG%
echo [OK] Uploaded. GitHub now rebuilds the site - this takes about 2 minutes.
start "" "https://github.com/Zeexx99/zee-ai/actions"
echo      A browser tab shows the progress. Wait for the GREEN check.
echo.
goto cache

:localonly
echo [2/4] No GitHub setup found - the local server will serve the new build.
echo.

:cache
echo [3/4] Close ALL Excel windows now, then press any key here.
pause >NUL
echo Clearing Office's cached copy of the add-in...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" (
  rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" >> %LOG% 2>&1
)
echo cache cleared >> %LOG%
echo [OK] Cache cleared.
echo.
echo [4/4] When the GitHub check is green, open Excel - Home tab - Zee AI.
echo      Settings - Updates shows the version and build time. Tip: update-all.bat updates Excel and Word together.
echo.
pause
