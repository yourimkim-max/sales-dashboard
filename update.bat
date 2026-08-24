@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ===================================
echo   대시보드 데이터 업데이트
echo ===================================
echo.

echo [1/2] 데이터 재생성 중...
python update.py
if %errorlevel% neq 0 (
    echo.
    echo 오류가 발생했습니다. 위 메시지를 확인해주세요.
    pause
    exit /b 1
)

echo.
echo [2/2] GitHub 배포 중...
git add index.html dashboard.html
git commit -m "daily update"
git push origin main

echo.
echo ===================================
echo  완료! 30초 후 아래 링크에 반영됩니다
echo  https://yourimkim-max.github.io/sales-dashboard/
echo ===================================
echo.
pause
