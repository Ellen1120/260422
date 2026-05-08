@echo off
chcp 65001 > nul
:: 관리자 권한 확인
net session >nul 2>&1
if %errorLevel% == 0 (
    echo 관리자 권한이 확인되었습니다.
) else (
    echo 관리자 권한을 요청합니다...
    powershell -Command "Start-Process '%~dpnx0' -Verb RunAs"
    exit /b
)

echo 방화벽에 8502번 포트를 열고 있습니다 (STM Calculator 접근 허용)...
powershell -Command "New-NetFirewallRule -DisplayName 'STM Calculator (Port 8502)' -Direction Inbound -LocalPort 8502 -Protocol TCP -Action Allow -ErrorAction SilentlyContinue"
echo.
echo 방화벽 설정이 완료되었습니다. 이제 다른 사용자가 IP 주소(10.160.76.198:8502)로 항상 접속할 수 있습니다.
pause
