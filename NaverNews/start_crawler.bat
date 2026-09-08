@echo off
echo ========================================
echo      네이버 뉴스 크롤러 시작
echo ========================================
echo.

if exist "NaverNewsCrawler.exe" (
    echo NaverNewsCrawler.exe 실행 중...
    start "" "NaverNewsCrawler.exe"
    echo.
    echo 프로그램이 실행되었습니다!
    echo 창이 나타나지 않으면 작업관리자에서 확인해주세요.
) else (
    echo 오류: NaverNewsCrawler.exe 파일을 찾을 수 없습니다.
    echo install_and_run.bat를 먼저 실행해주세요.
    pause
)

echo.
pause