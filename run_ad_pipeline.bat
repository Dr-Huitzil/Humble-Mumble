@echo off
title Active Directory Computer Audit & Report Pipeline
color 0b

echo ============================================================
echo   ACTIVE DIRECTORY INVENTORY & AUDIT PIPELINE
echo ============================================================
echo.
echo [Step 1/2] Querying Active Directory and exporting raw data...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Export_AD_Computers_Raw.ps1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Active Directory export failed. Ensure you have network connectivity to the domain.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [Step 2/2] Generating executive Excel inventory report...
python build_ad_report.py
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Excel report generation failed.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo ============================================================
echo   PIPELINE COMPLETE!
echo ============================================================
echo Opening generated report...
for /f "delims=" %%F in ('dir /b /o:-d AD_Computer_Inventory_Report_*.xlsx') do (
    start "" "%%F"
    goto :done
)
:done
pause
