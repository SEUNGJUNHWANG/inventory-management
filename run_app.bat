@echo off
chcp 65001 >nul
title Inventory Management System v1.0
cd /d "%~dp0"
"C:\Users\user\AppData\Local\Programs\Python\Python311\python.exe" main.py
if errorlevel 1 (
    echo.
    echo Error occurred.
    pause
)
