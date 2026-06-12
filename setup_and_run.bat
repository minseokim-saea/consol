@echo off
chcp 65001 >nul
echo ================================================
echo  연결 재무보고 통합 시스템 - 설치 및 실행
echo ================================================
echo.

:: Python 설치 확인
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo.
    echo  아래 링크에서 Python 3.11 이상을 설치하세요:
    echo  https://www.python.org/downloads/
    echo.
    echo  설치 시 "Add Python to PATH" 체크박스를 반드시 선택하세요!
    pause
    exit /b 1
)

echo [OK] Python 확인 완료
python --version
echo.

:: 가상환경 생성 (최초 1회)
if not exist venv (
    echo [설치] 가상환경 생성 중...
    python -m venv venv
)

:: 가상환경 활성화
call venv\Scripts\activate.bat

:: 패키지 설치
echo [설치] 필요 패키지 설치 중...
pip install -r requirements.txt -q

echo.
echo ================================================
echo  서버 시작 중... 브라우저에서 아래 주소 접속:
echo  http://localhost:5000
echo ================================================
echo.

python app.py

pause
