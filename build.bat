@echo off
echo DOCX to HTML Converter - EXE 빌드
echo ================================

cd /d %~dp0

echo 빌드 시작...
pyinstaller --noconfirm --onefile --windowed ^
    --name "DocxToHtmlConverter" ^
    --add-data "config.json;." ^
    --icon "NONE" ^
    src/main.py

echo.
echo 빌드 완료!
echo 실행 파일: dist\DocxToHtmlConverter.exe
pause
