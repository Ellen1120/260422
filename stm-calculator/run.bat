@echo off
set PY=C:\Users\220216\AppData\Local\Programs\Python\Python314\python.exe

echo STM 시험 자원 계산기 시작 중...
echo.

REM ANTHROPIC_API_KEY 확인
if "%ANTHROPIC_API_KEY%"=="" (
    echo [경고] ANTHROPIC_API_KEY 환경변수가 설정되지 않았습니다.
    echo        STM 문서 파싱 기능을 사용하려면 키를 설정하세요.
    echo        예: set ANTHROPIC_API_KEY=sk-ant-...
    echo.
)

echo 브라우저에서 http://localhost:8000 을 열어주세요.
echo 종료하려면 Ctrl+C 를 누르세요.
echo.

"%PY%" -m uvicorn main:app --host 127.0.0.1 --port 8000
pause
