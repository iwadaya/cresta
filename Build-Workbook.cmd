@echo off
rem Double-click to build Aggregate_Tracker.xlsm from Aggregate_Tracker.xlsx + vba\*.bas
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build-Workbook.ps1" %*
pause
