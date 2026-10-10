@echo off
title Zee AI - setup
cd /d "%~dp0"
echo ============================================
echo   Zee AI setup
echo ============================================
where node >NUL 2>NUL
if errorlevel 1 (
  echo [X] Node.js is not installed.
  echo     Install Node.js 20 LTS or newer from https://nodejs.org and run this file again.
  echo NODE_MISSING> setup-status.txt
  pause
  exit /b 1
)
for /f "delims=" %%v in ('node -v') do echo [OK] Node %%v
echo.
echo [1/3] Installing packages (takes a few minutes)...
call npm install > setup-log.txt 2>&1
if errorlevel 1 (
  echo [X] npm install failed - see setup-log.txt
  echo NPM_INSTALL_FAILED> setup-status.txt
  pause
  exit /b 1
)
echo [OK] Packages installed
echo.
echo [2/3] Creating the localhost HTTPS certificate...
echo       If Windows asks "Do you want to install this certificate?" click YES.
call npx office-addin-dev-certs install >> setup-log.txt 2>&1
if errorlevel 1 (
  echo [X] Certificate step failed - see setup-log.txt
  echo CERTS_FAILED> setup-status.txt
  pause
  exit /b 1
)
echo [OK] Certificate trusted
echo READY> setup-status.txt
echo.
echo [3/3] Starting Zee AI and opening Excel...
echo       Keep this window open while you use the add-in.
call "%~dp0start-zee.bat"
pause
