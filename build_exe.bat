@echo off
title Build - Inventory Management System
cd /d "%~dp0"

echo ============================================
echo   Inventory Management System - EXE Build
echo ============================================
echo.

echo [1/4] Checking Python...
"C:\Users\user\AppData\Local\Programs\Python\Python311\python.exe" --version
if errorlevel 1 (
    echo ERROR: Python 3.11 not found.
    pause
    exit /b 1
)

echo [2/4] Installing dependencies...
"C:\Users\user\AppData\Local\Programs\Python\Python311\python.exe" -m pip install -r requirements.txt -q
"C:\Users\user\AppData\Local\Programs\Python\Python311\python.exe" -m pip install pyinstaller -q

echo [3/4] Building EXE...
"C:\Users\user\AppData\Local\Programs\Python\Python311\python.exe" -m PyInstaller build.spec --noconfirm --clean

echo [4/4] Checking result...
if exist "dist\재고관리시스템.exe" (
    echo.
    echo ============================================
    echo   BUILD SUCCESS!
    echo   Output: dist\재고관리시스템.exe
    echo ============================================
    echo.
    echo NEXT STEP: Upload dist\재고관리시스템.exe to GitHub Releases
) else (
    echo.
    echo ============================================
    echo   BUILD FAILED. Check error messages above.
    echo ============================================
)
echo.
pause
