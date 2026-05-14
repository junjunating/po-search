@echo off
cd /d "%~dp0"

echo ========================================
echo Python 설치 확인
echo ========================================
py --version
if errorlevel 1 (
    echo py 명령어를 찾을 수 없습니다. python 명령어로 시도합니다.
    python --version
)

echo.
echo ========================================
echo 가상환경 생성
echo ========================================
py -m venv .venv
if errorlevel 1 (
    python -m venv .venv
)

echo.
echo ========================================
echo 가상환경 활성화
echo ========================================
call .venv\Scripts\activate.bat

echo.
echo ========================================
echo pip 업그레이드
echo ========================================
python -m pip install --upgrade pip

echo.
echo ========================================
echo 필요한 라이브러리 설치
echo ========================================
pip install -r requirements.txt

echo.
echo ========================================
echo Playwright 브라우저 설치
echo ========================================
playwright install

echo.
echo ========================================
echo 설치 완료
echo ========================================
pause