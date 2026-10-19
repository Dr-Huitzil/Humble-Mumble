@echo off
set PYTHONPATH=%~dp0src;%PYTHONPATH%
start pythonw "%~dp0app_gui.py"
