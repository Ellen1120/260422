@echo off
chcp 65001 > nul
cd /d "%~dp0"
set PY=C:\Users\220216\AppData\Local\Programs\Python\Python314\python.exe

echo QC 시험 준비 자동화 시스템 시작 중...
echo.
echo 브라우저에서 http://localhost:8502 를 열어주세요.
echo 종료하려면 Ctrl+C 를 누르세요.
echo.

"%PY%" -m uvicorn main:app --host 0.0.0.0 --port 8502
pause
