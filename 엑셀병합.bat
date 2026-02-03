@echo off
chcp 65001 > nul
echo 엑셀 시트 병합 프로그램을 시작합니다...
python "%~dp0merge_excel_gui.py"
if errorlevel 1 (
    echo.
    echo Python이 설치되어 있지 않거나 오류가 발생했습니다.
    echo Python 설치: https://www.python.org/downloads/
    pause
)
