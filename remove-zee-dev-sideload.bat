@echo off
title Zee AI - remove developer (localhost) sideloads
echo ============================================================
echo   Removes the Zee AI "Developer Add-ins" entries for Excel
echo   and Word, so you can install the hosted versions instead.
echo   Other add-ins are not touched.
echo ============================================================
echo.
echo Close ALL Word and Excel windows, then press any key.
pause >NUL

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
 "$k='HKCU:\Software\Microsoft\Office\16.0\WEF\Developer';" ^
 "if (-not (Test-Path $k)) { 'No developer add-ins are registered.'; exit 0 };" ^
 "$n=0; (Get-ItemProperty $k).PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' -and (($_.Name + ' ' + [string]$_.Value) -match 'ADD_IN_CLAUDE|Word_AddIn|ZeeAI|zee-ai') } | ForEach-Object { Remove-ItemProperty -Path $k -Name $_.Name; 'Removed: ' + $_.Value; $n++ };" ^
 "if ($n -eq 0) { 'No Zee AI developer entries found.' }"

echo.
echo Next: open Word / Excel - Insert - Get Add-ins - My Add-ins -
echo       Upload My Add-in - pick manifest.prod.xml from
echo         D:\Word_AddIn        (for Word)
echo         D:\ADD_IN_CLAUDE     (for Excel)
echo.
pause
