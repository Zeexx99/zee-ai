@echo off
setlocal EnableExtensions EnableDelayedExpansion
title Zee AI - update Excel + Word
rem Builds, tests and publishes BOTH add-ins, waits until GitHub Pages serves the new versions,
rem then clears Office's add-in cache. Works from either add-in folder (they sit side by side on D:\).
cd /d "%~dp0"
for %%d in ("%~dp0..\ADD_IN_CLAUDE") do set "EXCEL_DIR=%%~fd"
for %%d in ("%~dp0..\Word_AddIn") do set "WORD_DIR=%%~fd"
set "WAIT=%~dp0wait-live.ps1"

echo ============================================================
echo   Zee AI - update Excel + Word add-ins
echo ============================================================
echo   Excel: %EXCEL_DIR%
echo   Word : %WORD_DIR%
echo.

set "EXCEL_STATE=skipped"
set "WORD_STATE=skipped"
set "EXCEL_SINCE="
set "WORD_SINCE="

if exist "%EXCEL_DIR%\package.json" (
  call :publish "%EXCEL_DIR%" "Zee AI for Excel" EXCEL
) else (
  echo [!] Excel add-in folder not found - skipped.
)
if exist "%WORD_DIR%\package.json" (
  call :publish "%WORD_DIR%" "Zee AI for Word" WORD
) else (
  echo [!] Word add-in folder not found - skipped.
)

echo.
echo ------------------------------------------------------------
echo   Waiting for GitHub Pages
echo ------------------------------------------------------------
if "%EXCEL_STATE%"=="pushed" powershell -NoProfile -ExecutionPolicy Bypass -File "%WAIT%" -RepoDir "%EXCEL_DIR%" -Since "%EXCEL_SINCE%" -Name "Zee AI for Excel"
if "%EXCEL_STATE%"=="pushed" if errorlevel 1 set "EXCEL_STATE=pushed, not confirmed online"
if "%EXCEL_STATE%"=="pushed" set "EXCEL_STATE=LIVE"
if "%WORD_STATE%"=="pushed" powershell -NoProfile -ExecutionPolicy Bypass -File "%WAIT%" -RepoDir "%WORD_DIR%" -Since "%WORD_SINCE%" -Name "Zee AI for Word"
if "%WORD_STATE%"=="pushed" if errorlevel 1 set "WORD_STATE=pushed, not confirmed online"
if "%WORD_STATE%"=="pushed" set "WORD_STATE=LIVE"

echo.
echo ------------------------------------------------------------
echo   Excel: %EXCEL_STATE%
echo   Word : %WORD_STATE%
echo ------------------------------------------------------------
echo.
echo Close ALL Excel and Word windows now, then press any key here
echo (this clears Office's cached copy of the add-ins).
pause >NUL
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" (
  rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" >NUL 2>&1
)
echo [OK] Cache cleared.
echo.
echo Live - reopen Excel / Word and click Home - Zee AI.
echo (Settings - Updates shows the version and build time.)
echo.
pause
exit /b 0

rem ------------------------------------------------------------------ :publish <dir> <name> <VAR>
:publish
set "DIR=%~1"
set "NAME=%~2"
set "VAR=%~3"
set "LOG=%DIR%\update-log.txt"
echo.
echo ---- %NAME% ----
pushd "%DIR%"
echo ===== %NAME% update-all %date% %time% ===== > "%LOG%"
echo [1/2] Installing packages, running the tests and building...
call npm install --no-audit --no-fund >> "%LOG%" 2>&1
if errorlevel 1 (
  echo [X] npm install failed - see %LOG%
  set "%VAR%_STATE=FAILED (npm install) - send update-log.txt to Claude"
  popd
  exit /b 1
)
call npm run build >> "%LOG%" 2>&1
if errorlevel 1 (
  echo BUILD FAILED >> "%LOG%"
  echo [X] Tests or build failed - nothing published. Send %LOG% to Claude.
  set "%VAR%_STATE=FAILED (tests/build) - send update-log.txt to Claude"
  popd
  exit /b 1
)
echo build OK >> "%LOG%"
echo [OK] Build succeeded.
if not exist .git (
  echo [!] No GitHub setup in this folder - only built locally.
  set "%VAR%_STATE=built locally (no GitHub)"
  popd
  exit /b 0
)
set "CHANGES="
for /f "delims=" %%c in ('git status --porcelain') do set "CHANGES=1"
set "AHEAD=0"
for /f %%c in ('git rev-list --count origin/main..HEAD 2^>NUL') do set "AHEAD=%%c"
if not defined CHANGES if "!AHEAD!"=="0" (
  echo [OK] No changes since the last upload - already online.
  set "%VAR%_STATE=no changes - already online"
  popd
  exit /b 0
)
echo [2/2] Uploading to GitHub...
for /f "delims=" %%t in ('powershell -NoProfile -Command "(Get-Date).ToUniversalTime().ToString('o')"') do set "%VAR%_SINCE=%%t"
git add -A >> "%LOG%" 2>&1
git commit -m "%NAME% update %date% %time%" >> "%LOG%" 2>&1
git push origin main >> "%LOG%" 2>&1
if errorlevel 1 (
  echo PUSH FAILED >> "%LOG%"
  echo [X] Upload failed - see %LOG%. A GitHub sign-in window may have been closed.
  set "%VAR%_STATE=FAILED (git push) - see update-log.txt"
  popd
  exit /b 1
)
echo push OK >> "%LOG%"
echo [OK] Uploaded.
set "%VAR%_STATE=pushed"
popd
exit /b 0
