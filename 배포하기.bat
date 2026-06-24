@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   재고관리 시스템 - 새 버전 배포
echo ============================================
echo.

:: 현재 버전 읽기
for /f "tokens=2 delims=='\"'" %%a in ('findstr "APP_VERSION" core\constants.py') do set VERSION=%%a
for /f "tokens=2 delims==" %%a in ('findstr "APP_VERSION" core\constants.py') do (
    set RAW=%%a
)

:: Python으로 버전만 추출
for /f %%a in ('"C:\Users\user\AppData\Local\Programs\Python\Python311\python.exe" -c "import re,open; print(re.search(r'APP_VERSION\s*=\s*\"([^\"]+)\"', open('core/constants.py').read()).group(1))"') do set VERSION=%%a

echo 현재 버전: v%VERSION%
echo.
echo 이 버전(v%VERSION%)을 GitHub에 배포하시겠습니까?
echo (배포 후 동료 프로그램에 자동 업데이트 알림이 전송됩니다)
echo.
set /p CONFIRM=진행하려면 Y 입력: 

if /i not "%CONFIRM%"=="Y" (
    echo 취소되었습니다.
    pause
    exit /b
)

echo.
echo [1/3] 변경사항 저장 중...
git add -A
git commit -m "v%VERSION% - 업데이트 배포"
if errorlevel 1 (
    echo 커밋할 변경사항이 없거나 이미 커밋되었습니다. 계속 진행합니다.
)

echo.
echo [2/3] GitHub에 업로드 중...
git push origin main
if errorlevel 1 (
    echo.
    echo [오류] GitHub 업로드 실패. 인터넷 연결 또는 GitHub 설정을 확인하세요.
    pause
    exit /b 1
)

echo.
echo [3/3] 버전 태그 생성 및 배포 시작...
git tag v%VERSION%
git push origin v%VERSION%
if errorlevel 1 (
    echo.
    echo [경고] 태그가 이미 존재할 수 있습니다.
    echo 버전 번호를 올린 뒤 다시 시도하세요.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   배포 완료!
echo ============================================
echo.
echo GitHub Actions에서 EXE 자동 빌드 중...
echo (보통 5-10분 소요)
echo.
echo 빌드 현황 확인:
echo https://github.com/SEUNGJUNHWANG/inventory-management/actions
echo.
echo 빌드 완료 후 동료가 프로그램을 열면
echo 자동으로 업데이트 알림이 표시됩니다.
echo ============================================
echo.
pause
