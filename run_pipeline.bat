@echo off
echo ============================================================
echo   NINJAONE INVENTORY AUTOMATION PIPELINE
echo ============================================================
set PYTHONPATH=%~dp0src;%PYTHONPATH%
python -m humble_mumble.ninjaone.build_inventory
echo.
pause
