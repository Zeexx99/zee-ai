@echo off
title Zee AI diagnostics
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0diagnose.ps1"
echo Diagnostics written to diag2-log.txt
timeout /t 5
