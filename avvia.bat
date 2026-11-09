@echo off
cd /d "%~dp0"
python tia_gui.py
if errorlevel 1 pause
