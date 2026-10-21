@echo off
setlocal enabledelayedexpansion
title Humble Mumble - Google Sheets Laptop Fleet Sync

set "SCRIPT_DIR=%~dp0"
set "PYTHONPATH=%SCRIPT_DIR%src;%PYTHONPATH%"

echo ======================================================================
echo    HUMBLE MUMBLE - GOOGLE SHEETS LAPTOP FLEET SYNC
echo    Target: CSVS Clinic Laptop Inventory
echo ======================================================================
echo  1. [Default] Sync Directly to Google Sheets via API
echo  2. Export Master Laptop CSV (Backup)
echo ======================================================================
set /p choice="Enter your choice (1-2) [Default is 1]: "

if "%choice%"=="" set choice=1

if "%choice%"=="1" (
    echo.
    echo Syncing latest clinic laptop fleet directly to Google Sheets...
    python -m humble_mumble.ninjaone.google_sheet_fleet --sync
    echo.
    pause
    exit /b 0
)

if "%choice%"=="2" (
    echo.
    echo Exporting master laptop CSV backup...
    python -m humble_mumble.ninjaone.google_sheet_fleet --csv
    echo.
    pause
    exit /b 0
)

echo Invalid choice.
pause
