@echo off
title HP Bureautique Monitor - Sync to Google Drive
echo ============================================
echo   Synchronisation automatique Google Drive
echo ============================================
echo.

REM Lancer le script PowerShell en mode surveillance
powershell -ExecutionPolicy Bypass -File "%~dp0SyncToDrive.ps1" -Watch

pause
