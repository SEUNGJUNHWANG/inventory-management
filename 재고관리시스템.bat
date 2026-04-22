@echo off
chcp 65001 >nul
cd /d "%~dp0"
cscript //nologo "start_app.vbs"
