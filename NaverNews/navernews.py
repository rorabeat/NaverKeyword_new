#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import logging
import os
import re
import time
from datetime import datetime
from typing import List, Optional, Tuple

from playwright.sync_api import sync_playwright, Page, Browser
from bs4 import BeautifulSoup


def clean_text(text: str) -> str:
    """텍스트 정리: 공백 정리, 과도한 줄바꿈 줄이기"""
    if not text:
        return ""

    # HTML 엔티티 디코딩
    from html import unescape
    text = unescape(text)

    # HTML 태그 제거
    text = re.sub(r"<[^>]+>", "", text)

    # 과도한 공백 및 줄바꿈 정리
    text = re.sub(r"\n\s*\n\s*\n+", "\n\n", text)  # 3줄 이상 연속 줄바꿈을 2줄로
    text = re.sub(r"[ \t]+", " ", text)  # 연속 공백을 하나로
    text = re.sub(r"^\s+|\s+$", "", text, flags=re.MULTILINE)  # 각 줄의 앞뒤 공백 제거

    return text.strip()


def setup_browser() -> Tuple[Browser, Page]:
    """Playwright 브라우저 설정"""
    playwright = sync_playwright().start()

    # 모바일 User-Agent 설정
    user_agent = "Mozilla/5.0 (iPhone; CPU iPhone OS 14_7_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.1.2 Mobile/15E148 Safari/604.1"

    browser = playwright.chromium.launch(
        headless=True,  # 실제 운영시에는 True로 설정
        args=[
            "--no-sandbox",
            "--disable-setuid-sandbox",
            "--disable-dev-shm-usage",
            "--disable-accelerated-2d-canvas",
            "--no-first-run",
            "--no-zygote",
            "--disable-gpu"
        ]
    )

    context = browser.new_context(
        user_agent=user_agent,
        viewport={"width": 375, "height": 667},  # iPhone SE 크기
        locale="ko-KR"
    )

    page = context.new_page()
    return browser, page


def collect_news_links(page: Page, keyword: str, max_links: int, max_scrolls: int = 20) -> List[str]:
    """모바일 네이버 뉴스 검색에서 기사 링크 수집"""
    # 최신순 검색을 위한 sort=1 파라미터 추가
    search_url = f"https://m.search.naver.com/search.naver?ssc=tab.m_news.all&where=m_news&sm=mtb_jum&query={keyword}&sort=1"

    try:
        page.goto(search_url, wait_until="networkidle")

        # 모바일 뉴스 토글 버튼 클릭 (요구사항에서 언급된 부분)
        try:
            # 모바일 뉴스 전환 버튼 찾기
            mobile_toggle = page.locator("i.spnew.ico_switch").first
            if mobile_toggle.is_visible():
                mobile_toggle.click()
                page.wait_for_timeout(1000)  # 토글 후 대기
        except Exception as e:
            logging.warning(f"모바일 뉴스 토글 실패: {e}")

        collected_links = set()
        scroll_count = 0

        while len(collected_links) < max_links and scroll_count < max_scrolls:
            # 현재 페이지의 모든 링크 수집
            links = page.query_selector_all("a[href]")

            for link in links:
                href = link.get_attribute("href")
                if href and any(domain in href for domain in ["n.news.naver.com", "m.news.naver.com", "news.naver.com"]):
                    # 중복 제거
                    if href not in collected_links:
                        collected_links.add(href)
                        if len(collected_links) >= max_links:
                            break

            # 목표 개수 달성했으면 종료
            if len(collected_links) >= max_links:
                break

            # 스크롤 다운
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            page.wait_for_timeout(2000)  # 스크롤 후 로딩 대기
            scroll_count += 1

            logging.info(f"스크롤 {scroll_count}/{max_scrolls}, 수집된 링크: {len(collected_links)}")

        return list(collected_links)[:max_links]

    except Exception as e:
        logging.error(f"링크 수집 실패: {e}")
        return []


def extract_article_title(soup: BeautifulSoup) -> str:
    """기사 제목 추출 - 여러 selector 시도"""
    title_selectors = [
        "h2.media_end_head_headline",
        "h3#articleTitle",
        "h2#title_area",
        "h1.article_title",
        ".article_header h1",
        "title"
    ]

    for selector in title_selectors:
        try:
            element = soup.select_one(selector)
            if element:
                title = clean_text(element.get_text())
                if title:
                    return title
        except:
            continue

    return ""


def extract_article_date(soup: BeautifulSoup) -> str:
    """기사 작성일자 추출"""
    date_selectors = [
        # 메타 태그
        "meta[property='article:published_time']",
        "meta[property='article:modified_time']",
        "meta[name='article:published_time']",
        "meta[name='article:modified_time']",

        # 일반적인 날짜 클래스들
        ".article_info .author em.date",
        ".article_date",
        ".byline em.date",
        ".news_date",
        ".date",
        "time",

        # 네이버 뉴스 특정 클래스들
        ".media_end_head_info_datestamp_bunch .media_end_head_info_datestamp_time",
        ".media_end_head_info_datestamp_time",
        ".t11",

        # 페이지 상단 정보 영역
        ".news_headline .byline em",
        ".article_header .byline em",

        # 추가적인 선택자들
        "[data-date]",
        ".article_date span",
        ".date em",
    ]

    for selector in date_selectors:
        try:
            element = soup.select_one(selector)
            if element:
                if element.get('content'):
                    date_text = element.get('content')
                elif element.get('datetime'):
                    date_text = element.get('datetime')
                elif element.get('data-date'):
                    date_text = element.get('data-date')
                else:
                    date_text = element.get_text()

                date_text = clean_text(date_text)
                if date_text and len(date_text) > 5:  # 최소 6자 이상
                    # 날짜 형식 정규화
                    import re

                    # 다양한 날짜 패턴 찾기
                    patterns = [
                        r'\d{4}[.-]\d{1,2}[.-]\d{1,2}',  # YYYY.MM.DD or YYYY-MM-DD
                        r'\d{4}년\s*\d{1,2}월\s*\d{1,2}일',  # YYYY년 MM월 DD일
                        r'\d{2}[.-]\d{1,2}[.-]\d{1,2}',  # YY.MM.DD (2024년 기준 20YY로 변환)
                        r'\d{4}/\d{1,2}/\d{1,2}',  # YYYY/MM/DD
                    ]

                    for pattern in patterns:
                        match = re.search(pattern, date_text)
                        if match:
                            found_date = match.group()
                            # 2자리 연도를 4자리로 변환
                            if re.match(r'\d{2}[.-]', found_date):
                                found_date = '20' + found_date
                            return found_date

                    # 패턴에 맞지 않아도 길이가 적당하면 반환 (시간 정보 등)
                    if len(date_text) >= 10 and any(char.isdigit() for char in date_text):
                        return date_text.strip()
        except:
            continue

    return "날짜 정보 없음"


def extract_article_content(page: Page, url: str) -> Tuple[str, str, str]:
    """기사 페이지에서 제목, 본문, 작성일자 추출"""
    try:
        page.goto(url, wait_until="networkidle", timeout=30000)
        page.wait_for_timeout(2000)  # 페이지 로딩 대기

        html = page.content()
        soup = BeautifulSoup(html, 'html.parser')

        # 제목 추출
        title = extract_article_title(soup)

        # 본문 추출
        content_element = soup.select_one("article#dic_area")
        if content_element:
            content = clean_text(content_element.get_text())
        else:
            content = "본문 추출 실패"

        # 작성일자 추출
        date = extract_article_date(soup)
        logging.info(f"추출된 날짜: '{date}', 정규화된 날짜: '{normalize_date_for_filename(date)}'")

        return title, content, date

    except Exception as e:
        logging.error(f"기사 내용 추출 실패 ({url}): {e}")
        return "", "본문 추출 실패", "날짜 정보 없음"


def normalize_date_for_filename(date_str: str) -> str:
    """파일명에 사용할 수 있도록 날짜 문자열을 YYYYMMDD 형식으로 정규화"""
    if not date_str or date_str == "날짜 정보 없음":
        return datetime.now().strftime("%Y%m%d")

    import re

    # 다양한 날짜 패턴 매칭
    patterns = [
        r'(\d{4})[.-](\d{1,2})[.-](\d{1,2})',  # YYYY.MM.DD or YYYY-MM-DD
        r'(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일',  # YYYY년 MM월 DD일
        r'(\d{2})[.-](\d{1,2})[.-](\d{1,2})',  # YY.MM.DD (2024년 기준 20YY로 변환)
        r'(\d{4})/(\d{1,2})/(\d{1,2})',  # YYYY/MM/DD
    ]

    for pattern in patterns:
        match = re.search(pattern, date_str)
        if match:
            groups = match.groups()
            if len(groups) == 3:
                year, month, day = groups
                # 2자리 연도를 4자리로 변환
                if len(year) == 2:
                    year = '20' + year
                # 월과 일을 2자리로 포맷팅
                month = month.zfill(2)
                day = day.zfill(2)
                return f"{year}{month}{day}"

    # 패턴에 맞지 않으면 현재 날짜 사용
    return datetime.now().strftime("%Y%m%d")


def save_article_to_file(title: str, content: str, date: str, url: str, keyword: str, index: int) -> str:
    """각 기사를 별도의 텍스트 파일로 저장"""
    # 키워드 폴더 생성
    safe_keyword = re.sub(r'[^\w\s가-힣]', '', keyword)[:30].strip()  # 한글, 영문, 숫자, 공백만 허용
    if not safe_keyword:
        safe_keyword = "unknown_keyword"

    keyword_dir = os.path.join("results", safe_keyword)
    os.makedirs(keyword_dir, exist_ok=True)

    # 파일명으로 사용할 제목 정리 (특수문자 제거)
    safe_title = re.sub(r'[^\w\s가-힣]', '', title)[:50].strip()  # 한글, 영문, 숫자, 공백만 허용, 50자 제한
    if not safe_title:
        safe_title = f"기사_{index}"

    # 파일명에 기사 작성일자 포함 (기존 timestamp 대신 date 사용)
    article_date = normalize_date_for_filename(date)
    filename = f"{safe_title}_{article_date}.txt"
    filepath = os.path.join(keyword_dir, filename)

    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(f"제목: {title or f'기사 {index}'}\n")
        f.write(f"작성일자: {date}\n")
        f.write(f"URL: {url}\n")
        f.write(f"키워드: {keyword}\n")
        f.write(f"{'='*50}\n\n")
        f.write(f"{content}\n")

    return filepath


def save_summary_file(keyword: str, requested_count: int, collected_count: int,
                     search_url: str, articles: List[Tuple[str, str, str, str]]) -> str:
    """요약 정보를 txt 파일로 저장"""
    # 키워드 폴더 생성
    safe_keyword = re.sub(r'[^\w\s가-힣]', '', keyword)[:30].strip()  # 한글, 영문, 숫자, 공백만 허용
    if not safe_keyword:
        safe_keyword = "unknown_keyword"

    keyword_dir = os.path.join("results", safe_keyword)
    os.makedirs(keyword_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"summary_{safe_keyword}_{timestamp}.txt"
    filepath = os.path.join(keyword_dir, filename)

    with open(filepath, 'w', encoding='utf-8') as f:
        # 헤더 정보
        f.write(f"키워드: {keyword}\n")
        f.write(f"요청 개수: {requested_count}\n")
        f.write(f"수집 개수: {collected_count}\n")
        f.write(f"검색 URL: {search_url}\n\n")

        # 각 기사 정보 요약
        f.write("수집된 기사 목록:\n")
        f.write(f"{'='*80}\n")
        for i, (title, url, content, date) in enumerate(articles, 1):
            f.write(f"{i:2d}")
            f.write(f"제목: {title or f'기사 {i}'}\n")
            f.write(f"작성일자: {date}\n")
            f.write(f"URL: {url}\n")
            f.write(f"{'-'*50}\n")

    return filepath


def main():
    """메인 함수"""
    # 로깅이 이미 설정되어 있지 않은 경우에만 설정
    if not logging.getLogger().hasHandlers():
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s | %(levelname)s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )

    print("네이버 뉴스 크롤러")
    print("=" * 30)

    # 사용자 입력 (테스트 모드에서는 기본값 사용)
    import sys
    if len(sys.argv) > 2:
        # 명령줄 인수로 키워드와 개수 지정 가능
        keyword = sys.argv[1]
        try:
            news_count = int(sys.argv[2])
        except:
            news_count = 3  # 기본값
    elif len(sys.argv) > 1 and sys.argv[1] == "--demo":
        # 데모 모드: 자동으로 테스트 실행
        keyword = "인공지능"
        news_count = 2
        print(f"데모 모드 실행: 키워드='{keyword}', 개수={news_count}")
    else:
        # 일반 모드: 사용자 입력
        keyword = input("키워드를 입력하세요: ").strip()
        if not keyword:
            print("키워드가 비어있습니다.")
            return 1

        try:
            news_count_input = input("수집할 뉴스 개수를 입력하세요 (기본 10): ").strip()
            news_count = int(news_count_input) if news_count_input else 10
            if news_count <= 0:
                raise ValueError("개수는 1 이상이어야 합니다.")
        except ValueError as e:
            print(f"잘못된 입력: {e}")
            return 1

    # 최신순 검색을 위한 sort=1 파라미터 추가
    search_url = f"https://m.search.naver.com/search.naver?ssc=tab.m_news.all&where=m_news&sm=mtb_jum&query={keyword}&sort=1"

    # 키워드 폴더명용 안전한 이름 생성
    safe_keyword = re.sub(r'[^\w\s가-힣]', '', keyword)[:30].strip()
    if not safe_keyword:
        safe_keyword = "unknown_keyword"

    # 브라우저 설정
    browser, page = None, None
    try:
        logging.info("브라우저 시작 중...")
        browser, page = setup_browser()

        # 뉴스 링크 수집
        logging.info(f"키워드 '{keyword}'으로 뉴스 링크 수집 시작...")
        news_links = collect_news_links(page, keyword, news_count)

        if not news_links:
            logging.warning("수집된 뉴스 링크가 없습니다.")
            return 1

        logging.info(f"총 {len(news_links)}개의 뉴스 링크 수집 완료")

        # 각 기사 내용 추출 및 개별 파일 저장
        articles = []
        saved_files = []

        for i, link in enumerate(news_links, 1):
            logging.info(f"기사 {i}/{len(news_links)} 처리 중: {link}")
            title, content, date = extract_article_content(page, link)
            articles.append((title, link, content, date))

            # 각 기사를 별도 파일로 저장
            article_file = save_article_to_file(title, content, date, link, keyword, i)
            saved_files.append(article_file)
            logging.info(f"기사 파일 저장: {article_file}")

            time.sleep(1)  # 요청 간 딜레이

        # 요약 파일 저장
        summary_file = save_summary_file(keyword, news_count, len(articles), search_url, articles)
        logging.info(f"요약 파일 저장 완료: {summary_file}")

        print(f"\n크롤링 완료!")
        print(f"요청 개수: {news_count}")
        print(f"수집 개수: {len(articles)}")
        print(f"저장 폴더: results/{safe_keyword}/")
        print(f"개별 기사 파일: {len(saved_files)}개 저장됨")
        print(f"요약 파일: {os.path.basename(summary_file)}")
        print("\n저장된 기사 파일들:")
        for file_path in saved_files[:5]:  # 처음 5개만 표시
            print(f"  - {os.path.basename(file_path)}")
        if len(saved_files) > 5:
            print(f"  ... 외 {len(saved_files) - 5}개")

        return 0

    except KeyboardInterrupt:
        logging.info("사용자에 의해 중단됨")
        return 1
    except Exception as e:
        logging.error(f"오류 발생: {e}")
        return 1
    finally:
        if page:
            page.close()
        if browser:
            browser.close()


if __name__ == "__main__":
    # 테스트 모드: 명령줄 인수가 있으면 테스트 실행
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "--test":
        print("=== 네이버 뉴스 크롤러 테스트 ===")
        print("1. playwright 설치 확인")
        try:
            from playwright.sync_api import sync_playwright
            p = sync_playwright().start()
            print("   [OK] playwright 정상")
            p.stop()
        except Exception as e:
            print(f"   [ERROR] playwright 오류: {e}")

        print("2. beautifulsoup4 설치 확인")
        try:
            import bs4
            print("   [OK] beautifulsoup4 정상")
        except Exception as e:
            print(f"   [ERROR] beautifulsoup4 오류: {e}")

        print("3. 날짜 추출 함수 테스트")
        test_html = '''
        <html><head>
        <meta property="article:published_time" content="2024.01.15. 14:30">
        </head><body>
        <div class="article_info">
            <span class="author"><em class="date">2024.01.15. 14:30</em></span>
        </div>
        </body></html>
        '''
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(test_html, 'html.parser')
        test_date = extract_article_date(soup)
        print(f"   추출된 날짜: {test_date}")
        print(f"   정규화된 날짜: {normalize_date_for_filename(test_date)}")

        # 추가 테스트: 실제 네이버 뉴스 페이지 형식
        print("4. 네이버 뉴스 실제 형식 테스트")
        naver_test_html = '''
        <html><head>
        <meta property="article:published_time" content="2026-02-08T10:30:00+09:00">
        </head><body>
        <div class="media_end_head_info_datestamp">
            <div class="media_end_head_info_datestamp_bunch">
                <em class="media_end_head_info_datestamp_time">2026.02.08. 오전 10:30</em>
            </div>
        </div>
        </body></html>
        '''
        naver_soup = BeautifulSoup(naver_test_html, 'html.parser')
        naver_test_date = extract_article_date(naver_soup)
        print(f"   네이버 형식 추출된 날짜: {naver_test_date}")
        print(f"   네이버 형식 정규화된 날짜: {normalize_date_for_filename(naver_test_date)}")

        print("=== 테스트 완료 ===")
        exit(0)

    exit(main())