@echo off
title Zee AI server - keep this open
cd /d "%~dp0"
rem Local development server (https://localhost:3100) serving the built files in dist\.
for /f %%c in ('curl.exe -k -s -o NUL -w "%%{http_code}" https://localhost:3100/taskpane.html') do set CODE=%%c
if "%CODE%"=="200" (
  echo Zee AI server is already running on https://localhost:3100
  timeout /t 5
  exit /b 0
)
rem Renew the localhost certificate if it expired (it lasts 30 days).
call npx office-addin-dev-certs verify >NUL 2>NUL
if errorlevel 1 (
  echo Renewing the localhost HTTPS certificate - click YES if Windows asks.
  call npx office-addin-dev-certs install
)
if not exist dist\taskpane.html call npm run build
echo Zee AI server starting on https://localhost:3100 - keep this window open (you can minimize it).
call npx vite preview
pause
