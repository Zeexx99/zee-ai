@echo off
title Zee AI - switch Excel to the hosted version
cd /d "%~dp0"
if not exist manifest.prod.xml goto nomanifest
set URL=
for /f "delims=" %%u in ('node scripts\prod-url.mjs') do set URL=%%u
echo Checking %URL% ...
set CODE=000
for /f %%c in ('curl.exe -s -o NUL -w "%%{http_code}" %URL%') do set CODE=%%c
if not "%CODE%"=="200" goto notready
echo [OK] Hosted add-in is online.
call npx office-addin-dev-settings unregister manifest.xml >NUL 2>NUL
call npx office-addin-dev-settings register manifest.prod.xml
echo.
echo Done. Close ALL Excel windows, open Excel again, then Home - Zee AI.
echo serve-zee.bat is no longer needed - it is only for local development now.
goto end

:nomanifest
echo [X] manifest.prod.xml not found. Run publish-github.bat first.
goto end

:notready
echo [X] The hosted add-in is not reachable yet - HTTP status %CODE%.
echo     Wait until the GitHub Actions deploy shows a green check, then run this again.

:end
echo.
pause
