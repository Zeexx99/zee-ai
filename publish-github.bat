@echo off
title Zee AI - publish to GitHub Pages
cd /d "%~dp0"
echo ============================================================
echo   Publish Zee AI to GitHub Pages - free HTTPS hosting
echo ============================================================
echo.

where git >NUL 2>NUL
if errorlevel 1 goto nogit
where node >NUL 2>NUL
if errorlevel 1 goto nonode

set GHUSER=
set /p GHUSER=Type your GitHub username and press Enter: 
if "%GHUSER%"=="" goto nouser
set REPO=zee-ai
set SITE=https://%GHUSER%.github.io/%REPO%

echo.
echo STEP 1 of 4 - create an EMPTY public repository named %REPO%
echo   A browser tab opens now. Keep "Public", do NOT add a README, then click "Create repository".
start "" "https://github.com/new?name=%REPO%&visibility=public"
echo   When the repository exists, come back here and press any key...
pause >NUL

echo.
echo STEP 2 of 4 - generating the hosted manifest for %SITE%
call node scripts\make-manifest.mjs %SITE%
if errorlevel 1 goto failed

echo.
echo STEP 3 of 4 - uploading the code. A GitHub sign-in window may appear: sign in and allow access.
if not exist .github\workflows mkdir .github\workflows
echo   Adding the GitHub Pages deploy workflow: deploy\deploy.yml to .github\workflows\deploy.yml
copy /y deploy\deploy.yml .github\workflows\deploy.yml >NUL
if not exist .git git init
git config user.email >NUL 2>NUL
if errorlevel 1 git config user.email "%GHUSER%@users.noreply.github.com"
git config user.name >NUL 2>NUL
if errorlevel 1 git config user.name "%GHUSER%"
git checkout -B main
git add -A
git commit -m "Zee AI release"
git remote remove origin >NUL 2>NUL
git remote add origin https://github.com/%GHUSER%/%REPO%.git
git push -u origin main --force
if errorlevel 1 goto pushfailed

echo.
echo STEP 4 of 4 - turn on GitHub Pages
echo   In the page that opens: under "Build and deployment", set Source = GitHub Actions.
echo   Then open the Actions tab and wait for the green check - about 2 minutes.
start "" "https://github.com/%GHUSER%/%REPO%/settings/pages"
echo.
echo When the deploy is green, double-click use-online.bat
echo Your add-in address: %SITE%/taskpane.html
goto end

:nogit
echo [X] Git is not installed.
echo     Install "Git for Windows" from https://git-scm.com/download/win - keep all default options -
echo     then close this window and run publish-github.bat again.
start "" "https://git-scm.com/download/win"
goto end

:nonode
echo [X] Node.js was not found. Run setup.bat first.
goto end

:nouser
echo [X] No username entered.
goto end

:pushfailed
echo.
echo [X] Upload failed. Common causes:
echo     - the repository name or your username is misspelled
echo     - the repository was not created yet
echo     - the GitHub sign-in window was closed
echo     Fix it and run publish-github.bat again.
goto end

:failed
echo [X] Something went wrong - see the messages above.

:end
echo.
pause
