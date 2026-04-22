@echo off
title Diagnostic
cd /d "%~dp0"
echo === CURRENT DIR ===
echo %cd%
echo.

echo === PYTHON CHECK ===
"C:\Users\user\AppData\Local\Programs\Python\Python311\python.exe" --version
if errorlevel 1 (
    echo Python 3.11 NOT FOUND at expected path
    echo Trying system python...
    python --version
    if errorlevel 1 (
        echo Python NOT FOUND in PATH either
    )
) else (
    echo Python 3.11 OK
)
echo.

echo === FOLDER STRUCTURE ===
if exist "core\" (
    echo [core folder EXISTS]
    dir /b core\
) else (
    echo [core folder MISSING]
)
echo.
if exist "ui\" (
    echo [ui folder EXISTS]
    dir /b ui\
) else (
    echo [ui folder MISSING]
)
echo.
if exist "utils\" (
    echo [utils folder EXISTS]
    dir /b utils\
) else (
    echo [utils folder MISSING]
)
echo.

echo === PACKAGE CHECK ===
"C:\Users\user\AppData\Local\Programs\Python\Python311\python.exe" -c "import gspread; print('gspread: OK')"
"C:\Users\user\AppData\Local\Programs\Python\Python311\python.exe" -c "import openpyxl; print('openpyxl: OK')"
"C:\Users\user\AppData\Local\Programs\Python\Python311\python.exe" -c "import google.auth; print('google-auth: OK')"
echo.

echo === RUNNING main.py ===
"C:\Users\user\AppData\Local\Programs\Python\Python311\python.exe" main.py
echo.
echo EXIT CODE: %errorlevel%
pause
