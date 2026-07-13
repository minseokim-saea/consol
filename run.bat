@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ================================================
echo  연결 재무보고 통합 시스템 - 서버 실행
echo  브라우저: http://localhost:5000
echo  종료: 이 창에서 Ctrl+C  (창을 닫으면 서버도 꺼집니다)
echo ================================================
echo.
if not exist "venv\Scripts\python.exe" (
    echo [오류] venv 가 없습니다. 최초 1회는 setup_and_run.bat 로 설치하세요.
    pause
    exit /b 1
)
venv\Scripts\python.exe app.py
echo.
echo [서버가 종료되었습니다]
pause
