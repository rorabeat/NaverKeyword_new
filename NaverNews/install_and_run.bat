@echo off
echo ========================================
echo    네이버 뉴스 크롤러 설치 및 실행
echo ========================================
echo.

echo [1/3] Playwright 브라우저 설치 중...
pip install playwright || (
    echo 오류: pip이 설치되어 있지 않습니다.
    echo Python을 설치하고 pip을 사용할 수 있는지 확인해주세요.
    pause
    exit /b 1
)

playwright install chromium || (
    echo 오류: Playwright 브라우저 설치 실패
    echo 인터넷 연결을 확인해주세요.
    pause
    exit /b 1
)

echo 브라우저 설치 완료!
echo.

echo [2/3] 필요한 Python 패키지 설치 중...
pip install PyQt5 beautifulsoup4 || (
    echo 오류: 패키지 설치 실패
    pause
    exit /b 1
)

echo 패키지 설치 완료!
echo.

echo [3/3] 네이버 뉴스 크롤러 실행 중...
echo.
echo ========================================
echo    ★★★ exe 파일만으로 실행 가능! ★★★
echo
echo    사용법:
echo    1. NaverNewsCrawler.exe 더블클릭
echo    2. 브라우저 설치 확인창에서 '예' 클릭
echo    3. 크롤링 시작!
echo ========================================
echo.

if exist "NaverNewsCrawler.exe" (
    echo exe 파일을 찾았습니다. 실행합니다...
    start "" "NaverNewsCrawler.exe"
) else (
    echo 오류: NaverNewsCrawler.exe 파일을 찾을 수 없습니다.
    pause
    exit /b 1
)

echo.
echo 크롤러가 실행되었습니다!
echo 창이 나타나지 않으면 작업관리자에서 확인해주세요.
echo.
pause