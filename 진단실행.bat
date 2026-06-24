@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   재고관리 시스템 - 오류 진단 중...
echo ============================================
echo.
"C:\Users\user\AppData\Local\Programs\Python\Python311\python.exe" 진단.py
