@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "upload.ps1"
pause
