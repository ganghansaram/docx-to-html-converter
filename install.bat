@echo off
chcp 65001 >nul
echo ========================================
echo   Document to HTML Converter 의존성 설치
echo ========================================
echo.

pip install --no-index --find-links=./packages -r requirements.txt

echo.
if %errorlevel%==0 (
    echo [성공] 설치가 완료되었습니다.
) else (
    echo [실패] 설치 중 오류가 발생했습니다.
)
echo.
pause
