@echo off
cd /d "%~dp0"
set PYTHONPATH=%~dp0src;%PYTHONPATH%
start pythonw "%~dp0google_tasks_gui.py"
