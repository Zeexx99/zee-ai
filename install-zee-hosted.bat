@echo off
title Zee AI - install hosted add-ins (Excel + Word)
echo ============================================================
echo   Zee AI - register the ONLINE (GitHub) versions of the
echo   Excel and Word add-ins so they load every time Office starts.
echo   No local server needed. Other add-ins are not touched.
echo ============================================================
echo.
echo Close ALL Word and Excel windows, then press any key.
pause >NUL

set K=HKCU\Software\Microsoft\Office\16.0\WEF\Developer

echo [1/3] Removing old Zee AI (localhost) registrations...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
 "$k='HKCU:\Software\Microsoft\Office\16.0\WEF\Developer';" ^
 "if (-not (Test-Path $k)) { New-Item -Path $k -Force | Out-Null };" ^
 "(Get-ItemProperty $k).PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' -and (($_.Name + ' ' + [string]$_.Value) -match 'ADD_IN_CLAUDE|Word_AddIn|ZeeAI|zee-ai') } | ForEach-Object { Remove-ItemProperty -Path $k -Name $_.Name; '   removed: ' + $_.Value }"

echo [2/3] Registering the online versions...
if not exist "D:\ADD_IN_CLAUDE\manifest.prod.xml" echo [X] Missing D:\ADD_IN_CLAUDE\manifest.prod.xml - run publish-github.bat there first.
if not exist "D:\Word_AddIn\manifest.prod.xml"    echo [X] Missing D:\Word_AddIn\manifest.prod.xml - run publish-github.bat there first.
reg add "%K%" /v ZeeAI-Excel /t REG_SZ /d "D:\ADD_IN_CLAUDE\manifest.prod.xml" /f >NUL && echo    Excel: OK
reg add "%K%" /v ZeeAI-Word  /t REG_SZ /d "D:\Word_AddIn\manifest.prod.xml"    /f >NUL && echo    Word:  OK

echo [3/3] Done.
echo.
echo Open Word or Excel - Zee AI should be on the Home tab.
echo If it is not there the first time, load it once from
echo   Home - Add-ins - Developer Add-ins, then close and reopen.
echo.
echo Do NOT run start-zee.bat for daily use - it re-adds the localhost copy.
echo To uninstall:  reg delete "%K%" /v ZeeAI-Excel /f   and   /v ZeeAI-Word /f
echo.
pause
