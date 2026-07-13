@echo off
chcp 936 >nul
title YZScraper
cd /d "%~dp0"
echo Starting...
echo.
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "yzscraper.ps1"
echo.
pause
