@echo off
title Zee AI - launcher
cd /d "%~dp0"
set LOG=start-log.txt
echo ===== Zee AI start %date% %time% ===== > %LOG%

echo [1/4] Stopping any previous Zee AI server on port 3100...
powershell -NoProfile -Command "Get-NetTCPConnection -LocalPort 3100 -State Listen -ErrorAction SilentlyContinue | ForEach-Object { $p = Get-CimInstance Win32_Process -Filter ('ProcessId=' + $_.OwningProcess); if ($p.CommandLine -match 'ADD_IN_CLAUDE') { Stop-Process -Id $_.OwningProcess -Force; 'stopped old Zee AI server pid ' + $_.OwningProcess } else { 'port 3100 used by another app: ' + $p.CommandLine } }" >> %LOG% 2>&1
timeout /t 2 >NUL

call npx office-addin-dev-certs verify >NUL 2>NUL
if errorlevel 1 call npx office-addin-dev-certs install
echo [2/4] Building Zee AI (about 30 seconds)...
call npm run build >> %LOG% 2>&1
if errorlevel 1 (
  echo [X] Build failed - see start-log.txt
  echo BUILD FAILED >> %LOG%
  pause
  exit /b 1
)
echo build OK >> %LOG%

echo [3/4] Starting the Zee AI server on https://localhost:3100 ...
start "Zee AI server - keep this open" cmd /k "npx vite preview"
set TRIES=0
:waitloop
set CODE=000
for /f %%c in ('curl.exe -k -s -o NUL -w "%%{http_code}" https://localhost:3100/taskpane.html') do set CODE=%%c
if "%CODE%"=="200" goto served
set /a TRIES+=1
if %TRIES% GEQ 30 goto failed
timeout /t 2 >NUL
goto waitloop
:failed
echo [X] Server did not start. See the "Zee AI server" window. >> %LOG%
echo [X] Server did not start. Check the "Zee AI server" window.
pause
exit /b 1
:served
echo server OK (HTTP %CODE%) >> %LOG%

echo [4/4] Registering the add-in and opening Excel...
call npx office-addin-dev-settings debugging manifest.xml --disable >> %LOG% 2>&1
call npx office-addin-dev-settings register manifest.xml >> %LOG% 2>&1
call npx office-addin-dev-settings sideload manifest.xml desktop >> %LOG% 2>&1
echo sideload exit %errorlevel% >> %LOG%
echo.
echo Done. In Excel: Home tab - Zee AI.
echo Keep the "Zee AI server" window open while you use the add-in.
timeout /t 10
