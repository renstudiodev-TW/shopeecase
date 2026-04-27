@echo off
REM Yahoo Auction Bulk Upload Converter Tool - Ren Studio
REM Run this file to launch the GUI

cd /d "%~dp0"
python app.py
if errorlevel 1 pause
