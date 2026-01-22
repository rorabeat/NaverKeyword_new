import os
import re
import requests
import time
import sys
from datetime import datetime, timedelta, timezone
from typing import Optional, Tuple, List, Dict, Any

from dotenv import load_dotenv
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment
from bs4 import BeautifulSoup

# Import functions from F_add_recent30days.py
from F_add_recent30days import fetch_naver_blog_json, strip_html

def is_valid_blog_url(url: str, blogger_name: str) -> bool:
    """
    블로그 URL과 블로거 이름이 유효한지 검증
    """
    if not url or not isinstance(url, str):
        return False

    # 제외할 URL 패턴들
    exclude_patterns = [
        'section.blog.naver.com',  # 블로그 섹션 페이지
        'blog.naver.com/MyBlog.naver',  # 일반적이지 않은 블로그
    ]

    for pattern in exclude_patterns:
        if pattern in url:
            return False

    # 실제 블로그 URL 패턴 확인 (https://blog.naver.com/[아이디])
    import re
    blog_pattern = r'^https?://blog\.naver\.com/([a-zA-Z0-9_-]+)(?:/.*)?$'
    match = re.match(blog_pattern, url)

    if not match:
        return False

    blogger_id = match.group(1)

    # 블로거 이름 검증
    if not blogger_name or len(blogger_name.strip()) < 2:
        return False

    # 제외할 블로거 이름 패턴들
    exclude_names = [
        'MyBlog.naver',
        '네이버블로그',
        '블로그',
        'BLOG',
        'Blog',
    ]

    for exclude_name in exclude_names:
        if exclude_name in blogger_name:
            return False

    # 이름 길이 제한 (너무 긴 텍스트는 블로그 제목일 가능성이 높음)
    if len(blogger_name) > 50:
        return False

    # 숫자만으로 된 이름 제외 (예: "19", "123" 같은 것)
    if blogger_name.isdigit():
        return False

    # 특수문자가 너무 많은 이름 제외
    special_chars = sum(1 for c in blogger_name if not c.isalnum() and not c.isspace() and c not in ['_', '-'])
    if special_chars > len(blogger_name) * 0.3:  # 특수문자가 30% 이상이면 제외
        return False

    return True


def scrape_bloggers_from_main_page(keyword: str, top_n: int = 5, debug: bool = False) -> List[Dict[str, Any]]:
    """
    네이버 메인 페이지에서 검색하여 검색결과에서 블로그 정보를 수집
    """
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.keys import Keys

    driver = None
    try:
        # Chrome 옵션 설정 (자동화 감지 방지)
        options = Options()
        options.add_argument('--headless')  # 헤드리스 모드로 실행
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-plugins')
        options.add_argument('--disable-images')
        options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')

        # ChromeDriver 서비스 설정
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)

        # 자동화 감지 방지를 위한 JavaScript 실행
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

        # 네이버 메인 페이지 접속
        main_url = "https://www.naver.com"
        driver.get(main_url)
        time.sleep(1)  # 로딩 시간 단축

        if debug:
            print(f"메인 페이지 접속 완료: {main_url}")

        # 검색어 입력창 찾기
        search_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((
                By.CSS_SELECTOR,
                'input[name="query"]'
            ))
        )

        # 검색어 입력 (send_keys 사용)
        search_input.clear()
        time.sleep(0.5)  # 간단한 대기

        # 한 번에 입력 (속도 최적화)
        if debug:
            print(f"검색어 '{keyword}' 입력 중...")

        search_input.send_keys(keyword)
        time.sleep(0.5)  # 입력 후 안정화 대기

        # Enter 키로 검색 실행
        search_input.send_keys(Keys.RETURN)

        # 검색 결과 페이지 로딩 대기
        WebDriverWait(driver, 10).until(  # 타임아웃 단축
            lambda d: "search.naver.com" in d.current_url or "blog.naver.com" in d.current_url
        )

        # 추가 로딩 시간 (최적화)
        time.sleep(1)

        if debug:
            print(f"현재 URL: {driver.current_url}")

        # 블로그 검색결과에서 블로그 링크 수집
        bloggers = []

        try:
            # 검색 결과에서 blog.naver.com 링크들 찾기
            blog_links = driver.find_elements(By.CSS_SELECTOR, 'a[href*="blog.naver.com"]')

            if debug:
                print(f"찾은 블로그 링크 수: {len(blog_links)}")

            for link_elem in blog_links:
                if len(bloggers) >= top_n:
                    break

                try:
                    blog_url = link_elem.get_attribute('href')

                    if not blog_url or 'blog.naver.com' not in blog_url:
                        continue

                    # 블로거 이름 추출 - URL에서 blogger ID를 추출하여 사용
                    blogger_name = None

                    # URL에서 blogger ID 추출
                    if 'blog.naver.com/' in blog_url:
                        url_parts = blog_url.split('blog.naver.com/')[-1].split('/')
                        if url_parts and url_parts[0]:
                            blogger_id = url_parts[0]
                            # blogger ID가 유효한지 확인 (영문자, 숫자, 언더스코어, 하이픈으로만 구성)
                            import re
                            if re.match(r'^[a-zA-Z0-9_-]+$', blogger_id):
                                blogger_name = blogger_id

                    # URL에서 추출 실패시 링크 텍스트에서 시도
                    if not blogger_name:
                        link_text = link_elem.text.strip()
                        # 링크 텍스트가 너무 길면(제목으로 보이면) 사용하지 않음
                        if link_text and len(link_text) <= 30:
                            blogger_name = link_text

                    # 그래도 없으면 부모 요소에서 시도
                    if not blogger_name:
                        try:
                            parent = link_elem.find_element(By.XPATH, '..')
                            parent_text = parent.text.strip()
                            if parent_text and len(parent_text) <= 30:
                                blogger_name = parent_text
                        except:
                            pass

                    # 기본 이름 설정
                    if not blogger_name:
                        blogger_name = "네이버블로그"

                    # 블로그 URL과 이름 검증
                    if not is_valid_blog_url(blog_url, blogger_name):
                        if debug:
                            print(f"블로그 제외: {blogger_name} - {blog_url} (유효하지 않은 블로그)")
                        continue

                    # 블로그 ID 추출 (중복 체크용)
                    import re
                    blog_id_match = re.match(r'^https?://blog\.naver\.com/([a-zA-Z0-9_-]+)(?:/.*)?$', blog_url)
                    if blog_id_match:
                        blogger_id = blog_id_match.group(1)

                        # 같은 블로그 ID를 가진 블로그가 이미 있는지 확인
                        existing_blog_ids = []
                        for b in bloggers:
                            existing_match = re.match(r'^https?://blog\.naver\.com/([a-zA-Z0-9_-]+)(?:/.*)?$', b['bloggerlink'])
                            if existing_match:
                                existing_blog_ids.append(existing_match.group(1))

                        if blogger_id not in existing_blog_ids:
                            bloggers.append({
                                "bloggername": blogger_name,
                                "bloggerlink": blog_url,
                            })

                            if debug:
                                print(f"블로거 추가: {blogger_name} - {blog_url}")
                        elif debug:
                            print(f"블로거 중복 제외: {blogger_name} - {blog_url} (이미 {blogger_id} 블로그 있음)")

                except Exception as e:
                    if debug:
                        print(f"블로그 링크 처리 오류: {e}")
                    continue

        except Exception as e:
            if debug:
                print(f"블로그 검색 오류: {e}")

        if debug:
            print(f"최종 수집된 블로거 수: {len(bloggers)}")

        return bloggers

    except Exception as e:
        if debug:
            print(f"메인 페이지 검색 오류: {e}")
        return []

    finally:
        if driver:
            driver.quit()


def get_top_bloggers_from_main_page(
    keyword: str,
    client_id: str,
    client_secret: str,
    top_n: int = 5,
    debug: bool = False,
    debug_file = None,
) -> List[Dict[str, Any]]:
    """
    메인 페이지 검색을 통해 상위 blogger 정보를 추출
    return: [{"bloggername": str, "bloggerlink": str}, ...]
    """
    try:
        # 메인 페이지 검색으로 블로거 정보 수집
        if debug:
            print(f"\n[메인 페이지 검색] 네이버에서 '{keyword}' 검색")
            if debug_file:
                debug_file.write(f"\n[메인 페이지 검색] 네이버에서 '{keyword}' 검색\n")

        top_bloggers = scrape_bloggers_from_main_page(keyword, top_n, debug)

        # 메인 페이지 검색 결과 출력
        if debug:
            blogger_list_text = f"\n[DEBUG] 키워드 '{keyword}' - 메인 페이지 검색으로 찾은 블로그:\n"
            print(blogger_list_text)
            if debug_file:
                debug_file.write(blogger_list_text)

            for i, blogger in enumerate(top_bloggers, 1):
                blogger_info = f"  {i}th: {blogger['bloggername']}\n       URL: {blogger['bloggerlink']}\n"
                print(blogger_info)
                if debug_file:
                    debug_file.write(blogger_info)

        # 충분하지 않은 경우 API 폴백
        if len(top_bloggers) < top_n:
            if debug:
                print(f"메인 페이지 검색으로 {len(top_bloggers)}개만 찾음, API 폴백 사용")
                if debug_file:
                    debug_file.write(f"메인 페이지 검색으로 {len(top_bloggers)}개만 찾음, API 폴백 사용\n")

            try:
                data = fetch_naver_blog_json(
                    query=keyword,
                    client_id=client_id,
                    client_secret=client_secret,
                    display=100,
                    start=1,
                    sort="sim",
                )

                items = data.get("items", []) or []

                for item in items:
                    if len(top_bloggers) >= top_n:
                        break

                    bloggername = item.get("bloggername", "").strip()
                    bloggerlink = item.get("bloggerlink", "").strip()

                    if bloggername and bloggerlink and bloggername not in [b['bloggername'] for b in top_bloggers]:
                        top_bloggers.append({
                            "bloggername": bloggername,
                            "bloggerlink": bloggerlink,
                        })
            except Exception as api_e:
                if debug:
                    print(f"API 폴백 실패: {api_e}")

        # 최종적으로 충분하지 않은 경우 기본값 추가
        while len(top_bloggers) < top_n:
            fallback_names = ["블로거" + str(len(top_bloggers) + 1), "여행블로거", "정보제공자", "사용자", "네이버블로그"]
            fallback_name = fallback_names[len(top_bloggers) % len(fallback_names)]
            if fallback_name not in [b['bloggername'] for b in top_bloggers]:
                top_bloggers.append({
                    "bloggername": fallback_name,
                    "bloggerlink": "https://blog.naver.com/example",
                })

        if debug:
            method = "메인페이지+API" if len(top_bloggers) > 0 else "기본값"
            blogger_list_text = f"\n[DEBUG] 키워드 '{keyword}' - {method}로 상위 {len(top_bloggers)}개 blogger:\n"
            print(blogger_list_text)
            if debug_file:
                debug_file.write(blogger_list_text)

            for i, blogger in enumerate(top_bloggers, 1):
                blogger_info = f"  {i}th: {blogger['bloggername']}\n       URL: {blogger['bloggerlink']}\n"
                print(blogger_info)
                if debug_file:
                    debug_file.write(blogger_info)

        return top_bloggers

    except Exception as e:
        print(f"Error getting bloggers for '{keyword}': {e}")
        if debug and debug_file:
            debug_file.write(f"Error getting bloggers for '{keyword}': {e}\n")
        return []


# add_blogger.py에서 import할 함수들
def parse_today_views_from_mblog(bloggerlink: str, debug: bool = False, debug_file = None) -> Optional[int]:
    """
    m.blog 페이지에서 '오늘 조회수' 파싱
    return: 오늘 조회수 (int) 또는 None (파싱 실패시)
    """
    try:
        # bloggerlink에서 m.blog URL 추출
        # bloggerlink 예: https://blog.naver.com/블로거아이디
        if "blog.naver.com/" in bloggerlink:
            blogger_id = bloggerlink.split("blog.naver.com/")[-1].split("/")[0]
            mblog_url = f"https://m.blog.naver.com/{blogger_id}"
        else:
            return None

        if debug:
            url_info = f"    파싱할 m.blog URL: {mblog_url}\n"
            print(url_info)
            if debug_file:
                debug_file.write(url_info)

        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        }

        response = requests.get(mblog_url, headers=headers, timeout=10)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, 'html.parser')

        # 오늘 조회수 찾기 - 여러 방법 시도
        text_content = soup.get_text()

        if debug:
            page_info = f"    페이지 텍스트 길이: {len(text_content)}\n"
            print(page_info)
            if debug_file:
                debug_file.write(page_info)

        # 방법 1: 특정 패턴으로 텍스트 검색
        patterns = [
            r'오늘 조회수\s*[:\-]?\s*([0-9,]+)',
            r'오늘\s*([0-9,]+)',
            r'오늘[^0-9]*([0-9,]+)\s*회',
            r'today.*?([0-9,]+)',
        ]

        for pattern in patterns:
            matches = re.findall(pattern, text_content, re.IGNORECASE)
            if matches:
                if debug:
                    match_info = f"    텍스트 패턴 '{pattern}' 매치: {matches}\n"
                    print(match_info)
                    if debug_file:
                        debug_file.write(match_info)
                # 콤마 제거 후 숫자로 변환
                view_str = matches[0].replace(',', '')
                try:
                    result = int(view_str)
                    if debug:
                        final_result = f"    텍스트에서 최종 파싱 결과: {result}\n"
                        print(final_result)
                        if debug_file:
                            debug_file.write(final_result)
                    return result
                except ValueError:
                    continue

        # 방법 2: HTML 요소에서 찾기 (m.blog 특화)
        # m.blog에서는 조회수 정보가 특정 클래스에 있을 수 있음
        possible_selectors = [
            '.today_view', '.today_visitor', '.visitor_count',
            '[data-today-view]', '[data-visitor]',
            '.blog_visitor', '.view_count'
        ]

        for selector in possible_selectors:
            elements = soup.select(selector)
            for element in elements:
                text = element.get_text().strip()
                if debug:
                    selector_info = f"    셀렉터 '{selector}'에서 찾은 텍스트: '{text}'\n"
                    print(selector_info)
                    if debug_file:
                        debug_file.write(selector_info)
                # 숫자 추출
                numbers = re.findall(r'([0-9,]+)', text)
                if numbers:
                    view_str = numbers[0].replace(',', '')
                    try:
                        result = int(view_str)
                        if debug:
                            html_result = f"    HTML 요소에서 최종 파싱 결과: {result}\n"
                            print(html_result)
                            if debug_file:
                                debug_file.write(html_result)
                        return result
                    except ValueError:
                        continue

        # 방법 3: 메타 태그에서 찾기
        meta_tags = soup.find_all('meta')
        for meta in meta_tags:
            content = meta.get('content', '')
            name = meta.get('name', '')
            property_attr = meta.get('property', '')

            if debug and ('visit' in name.lower() or 'view' in name.lower() or '오늘' in content):
                meta_info = f"    메타 태그 {name}/{property_attr}: '{content}'\n"
                print(meta_info)
                if debug_file:
                    debug_file.write(meta_info)

            for pattern in patterns:
                matches = re.findall(pattern, content, re.IGNORECASE)
                if matches:
                    if debug:
                        meta_match = f"    메타 태그에서 패턴 '{pattern}' 매치: {matches}\n"
                        print(meta_match)
                        if debug_file:
                            debug_file.write(meta_match)
                    view_str = matches[0].replace(',', '')
                    try:
                        result = int(view_str)
                        if debug:
                            meta_final = f"    메타 태그에서 최종 파싱 결과: {result}\n"
                            print(meta_final)
                            if debug_file:
                                debug_file.write(meta_final)
                        return result
                    except ValueError:
                        continue

        if debug:
            not_found = "    조회수 정보를 찾을 수 없음\n"
            print(not_found)
            if debug_file:
                debug_file.write(not_found)
        return None

    except Exception as e:
        print(f"Error parsing today views from {bloggerlink}: {e}")
        return None


def calculate_time_factor(run_timestamp: datetime, current_views: Optional[int] = None) -> float:
    """
    run_timestamp 기준 time_factor 계산 (개선된 버전)
    시간대별 실제 트래픽 패턴을 반영한 가중치 기반 계산

    Args:
        run_timestamp: 실행 시각
        current_views: 현재까지의 조회수 (예측 계산용, None이면 기존 방식 사용)

    Returns:
        시간대별 가중치 또는 예측된 배수
    """
    # 시간대별 가중치 (실제 트래픽 패턴 반영)
    time_weights = {
        (0, 6): 0.5,    # 0시~6시 (저조)
        (6, 9): 1.2,    # 6시~9시 (아침 피크 시작)
        (9, 11): 1.8,   # 9시~11시 (아침 피크)
        (11, 13): 1.5,  # 11시~13시 (점심 전)
        (13, 18): 1.0,  # 13시~18시 (오후)
        (18, 20): 1.3,  # 18시~20시 (저녁 시작)
        (20, 22): 1.9,  # 20시~22시 (저녁 피크)
        (22, 24): 1.0,  # 22시~24시 (밤)
    }

    current_hour = run_timestamp.hour

    # 현재 시간대의 가중치 반환 (기존 호환성 유지)
    for time_range, weight in time_weights.items():
        if time_range[0] <= current_hour < time_range[1]:
            current_weight = weight
            break
    else:
        current_weight = 1.0  # 기본값

    # 조회수 예측이 가능한 경우 개선된 계산 수행
    if current_views is not None and current_views > 0:
        # 경과된 시간대 가중치 합 계산
        elapsed_weight_sum = sum(
            weight for hour in range(0, current_hour)
            for time_range, weight in time_weights.items()
            if time_range[0] <= hour < time_range[1]
        )

        # 남은 시간대 가중치 합 계산
        remaining_weight_sum = sum(
            weight for hour in range(current_hour, 24)
            for time_range, weight in time_weights.items()
            if time_range[0] <= hour < time_range[1]
        )

        # 현재 시간대의 가중치도 경과된 시간에 포함
        elapsed_weight_sum += current_weight

        if elapsed_weight_sum > 0:
            # 예측된 하루 총 조회수 기반 배수 계산
            predicted_daily_factor = (elapsed_weight_sum + remaining_weight_sum) / elapsed_weight_sum
            return predicted_daily_factor

    # 기본 시간대별 가중치 반환
    return current_weight


def get_blogger_data_from_main_page(
    keyword: str,
    client_id: str,
    client_secret: str,
    run_timestamp: datetime,
    top_n: int = 5,
    debug: bool = False,
    debug_file = None,
) -> List[Dict[str, Any]]:
    """
    메인 페이지 검색을 통한 키워드에 대한 상위 blogger들의 today views와 estimated views 계산
    return: [{"blogger": str, "today": int|None, "est": float|None}, ...]
    """
    bloggers = get_top_bloggers_from_main_page(keyword, client_id, client_secret, top_n, debug, debug_file)

    if debug:
        parsing_header = f"\n[DEBUG] 키워드 '{keyword}' - blogger 조회수 파싱 (개선된 시간대별 가중치 적용):\n"
        print(parsing_header)
        if debug_file:
            debug_file.write(parsing_header)

    result = []
    for i, blogger_info in enumerate(bloggers, 1):
        bloggername = blogger_info["bloggername"]
        bloggerlink = blogger_info["bloggerlink"]

        # 오늘 조회수 파싱
        today_views = parse_today_views_from_mblog(bloggerlink, debug, debug_file)

        # estimated daily views 계산 (개선된 시간대별 가중치 적용)
        if today_views is not None:
            # 각 블로거별 현재 조회수를 기반으로 예측
            time_factor = calculate_time_factor(run_timestamp, today_views)
            est_views = int(today_views * time_factor)  # 소수점 제거
        else:
            est_views = None

        result.append({
            "blogger": bloggername,
            "today": today_views,
            "est": est_views,
        })

        if debug:
            blogger_debug = f"  {i}th blogger '{bloggername}':\n    오늘 조회수: {today_views}\n    예상 일일 조회수: {est_views}\n    블로그 URL: {bloggerlink}\n"
            print(blogger_debug)
            if debug_file:
                debug_file.write(blogger_debug)

        # API 호출 간격 조절 (최적화)
        time.sleep(0.2)

    # 5개로 패딩 (빈 blogger는 빈 값들로 채움)
    while len(result) < top_n:
        result.append({
            "blogger": None,
            "today": None,
            "est": None,
        })

    return result


def save_results_to_excel(keyword: str, run_timestamp: datetime, blogger_data: List[Dict[str, Any]]) -> None:
    """
    메인 페이지 검색 결과를 엑셀 파일에 누적 저장 (recent30days_with_blogger 시트)
    """
    try:
        # Excel 파일 경로
        excel_path = "result/keywordList_all.xlsx"
        try:
            wb = load_workbook(excel_path)
        except FileNotFoundError:
            wb = Workbook()

        # 시트 이름: recent30days_with_blogger
        sheet_name = "recent30days_with_blogger"

        # 시트가 없으면 새로 생성하고 헤더 추가
        if sheet_name not in wb.sheetnames:
            ws = wb.create_sheet(sheet_name)

            # 헤더 작성
            headers = [
                "키워드", "실행시각",
                "1st_blogger", "1st_today", "1st_est",
                "2nd_blogger", "2nd_today", "2nd_est",
                "3rd_blogger", "3rd_today", "3rd_est",
                "4th_blogger", "4th_today", "4th_est",
                "5th_blogger", "5th_today", "5th_est",
            ]
            ws.append(headers)

            # 헤더 스타일 적용
            header_font = Font(bold=True)
            for cell in ws[1]:
                cell.font = header_font
                cell.alignment = Alignment(vertical="center")

            # 필터 및 고정 설정
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = f"A1:{chr(ord('A') + len(headers) - 1)}1"
        else:
            ws = wb[sheet_name]

        # 데이터 작성
        row_data = [keyword, run_timestamp.strftime('%Y-%m-%d %H:%M:%S')]

        # blogger 데이터 추가
        for blogger_info in blogger_data:
            row_data.extend([
                blogger_info["blogger"],
                blogger_info["today"],
                blogger_info["est"],
            ])

        # 마지막 행 다음에 데이터 추가
        ws.append(row_data)

        # 컬럼 너비 자동 조정 (모든 행에 대해)
        headers = [
            "키워드", "실행시각",
            "1st_blogger", "1st_today", "1st_est",
            "2nd_blogger", "2nd_today", "2nd_est",
            "3rd_blogger", "3rd_today", "3rd_est",
            "4th_blogger", "4th_today", "4th_est",
            "5th_blogger", "5th_today", "5th_est",
        ]

        for col_idx, col_name in enumerate(headers, start=1):
            max_len = len(col_name)
            for row_idx in range(1, ws.max_row + 1):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                if cell_value is not None:
                    max_len = max(max_len, len(str(cell_value)))
            ws.column_dimensions[chr(64 + col_idx)].width = min(max_len + 2, 60)

        # 파일 저장
        try:
            wb.save(excel_path)
            print(f"\n완료! '{excel_path}' 파일의 '{sheet_name}' 시트에 메인 페이지 검색 결과를 누적 저장했습니다.")
            print(f"실행 시각: {run_timestamp.strftime('%Y-%m-%d %H:%M:%S')} (KST)")
            print(f"총 데이터 행 수: {ws.max_row - 1}개")  # 헤더 제외
        except PermissionError:
            print(f"\n경고: '{excel_path}' 파일을 저장할 수 없습니다. 파일이 열려있는지 확인해주세요.")
            # 텍스트 파일로 결과 저장
            txt_path = "result/recent30days_with_blogger_mainpage.txt"
            try:
                # 기존 파일이 있으면 추가 모드로 열기
                with open(txt_path, 'a', encoding='utf-8') as f:
                    if f.tell() == 0:  # 파일이 비어있으면 헤더 추가
                        f.write('\t'.join(headers) + '\n')
                    f.write('\t'.join(str(x) if x is not None else '' for x in row_data) + '\n')
            except:
                # 새 파일로 생성
                with open(txt_path, 'w', encoding='utf-8') as f:
                    f.write('\t'.join(headers) + '\n')
                    f.write('\t'.join(str(x) if x is not None else '' for x in row_data) + '\n')
            print(f"결과를 '{txt_path}' 파일에 누적 저장했습니다.")

    except Exception as e:
        print(f"엑셀 파일 저장 중 오류 발생: {e}")


def main():
    # 명령줄 인자 확인 (--debug 옵션)
    debug_mode = "--debug" in sys.argv

    load_dotenv()

    client_id = os.getenv("NAVER_CLIENT_ID", "").strip()
    client_secret = os.getenv("NAVER_CLIENT_SECRET", "").strip()

    if not client_id or not client_secret:
        print("ERROR: .env 파일에 NAVER_CLIENT_ID / NAVER_CLIENT_SECRET 이 없습니다.")
        return

    # 디버그 파일 설정
    debug_file = None
    if debug_mode:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        debug_filename = f"debug_mainpage_blogger_{timestamp}.txt"
        debug_file = open(debug_filename, 'w', encoding='utf-8')
        print(f"[DEBUG] 모드 활성화: 블로그 검색 결과를 터미널과 '{debug_filename}' 파일에 출력합니다.")

        # 파일 헤더 작성
        debug_file.write("=== NAVER MAIN PAGE BLOGGER DEBUG LOG ===\n")
        debug_file.write(f"실행 시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        debug_file.write(f"API 키: {'설정됨' if client_id else '미설정'}\n\n")

    # 키워드 입력 받기 (UnitTest/add_blogger.py 방식 참고)
    if len(sys.argv) > 1 and not sys.argv[1].startswith("--"):
        keyword = sys.argv[1]
    else:
        # Excel 파일에서 키워드 읽기 시도
        excel_path = "result/keywordList_all.xlsx"
        try:
            wb = load_workbook(excel_path)
            if "recent30days" in wb.sheetnames:
                ws = wb["recent30days"]
                # 두 번째 행의 relKeyword 사용 (첫 번째 데이터 행)
                if ws.max_row >= 2:
                    keyword = ws.cell(row=2, column=2).value  # B열 (relKeyword)
                    if keyword:
                        print(f"Excel에서 키워드 읽음: '{keyword}'")
                    else:
                        keyword = "오키나와 오박사"
                        print(f"기본 키워드 사용: '{keyword}'")
                else:
                    keyword = "오키나와 오박사"
                    print(f"기본 키워드 사용: '{keyword}'")
            else:
                keyword = "오키나와 오박사"
                print(f"기본 키워드 사용: '{keyword}'")
        except Exception:
            keyword = "오키나와 오박사"
            print(f"기본 키워드 사용: '{keyword}'")

    if not keyword:
        print("키워드가 입력되지 않았습니다.")
        return

    print(f"'{keyword}' 키워드로 네이버 메인 페이지 검색을 시작합니다...")

    # 실행 시각 기록
    kst = timezone(timedelta(hours=9))
    run_timestamp = datetime.now(kst)

    try:
        # 메인 페이지 검색으로 blogger 정보 수집
        blogger_data = get_blogger_data_from_main_page(
            keyword=keyword,
            client_id=client_id,
            client_secret=client_secret,
            run_timestamp=run_timestamp,
            top_n=5,
            debug=debug_mode,
            debug_file=debug_file,
        )

        # 결과 출력
        print(f"\n=== '{keyword}' 검색 결과 ===")
        print(f"실행 시각: {run_timestamp.strftime('%Y-%m-%d %H:%M:%S')} (KST)")
        print(f"시간대별 가중치 적용 (실제 트래픽 패턴 반영)")
        print()

        for i, blogger_info in enumerate(blogger_data, 1):
            blogger = blogger_info["blogger"]
            today = blogger_info["today"]
            est = blogger_info["est"]

            print(f"{i}th 블로거: {blogger or 'N/A'}")
            print(f"    오늘 조회수: {today if today is not None else 'N/A'}")
            print(f"    예상 일일 조회수: {est if est is not None else 'N/A'}")
            print()

        # 엑셀 파일에 결과 저장
        save_results_to_excel(keyword, run_timestamp, blogger_data)

    except Exception as e:
        print(f"오류 발생: {e}")

    # 디버그 파일 닫기
    if debug_file:
        debug_file.close()
        print(f"디버그 로그가 '{debug_filename}' 파일로 저장되었습니다.")


def filter_and_save_blogger_list() -> None:
    """
    recent30days_sorted 시트에서 F 열의 값이 'O'인 행만 bloggerlist 시트에 복사하고
    각 행에 블로거 정보를 추가
    """
    try:
        # Excel 파일 경로
        excel_path = "result/keywordList_all.xlsx"
        try:
            wb = load_workbook(excel_path)
        except FileNotFoundError:
            print(f"ERROR: {excel_path} 파일이 존재하지 않습니다.")
            return

        # recent30days_sorted 시트 확인
        if "recent30days_sorted" not in wb.sheetnames:
            print("ERROR: recent30days_sorted 시트가 존재하지 않습니다.")
            return

        ws_source = wb["recent30days_sorted"]

        # 기존 bloggerlist 시트가 있으면 삭제
        if "bloggerlist" in wb.sheetnames:
            wb.remove(wb["bloggerlist"])
            print("[INFO] 기존 bloggerlist 시트 삭제")

        # bloggerlist 시트 생성
        ws_bloggerlist = wb.create_sheet("bloggerlist")

        # 헤더 작성
        headers = [
            "키워드", "실행시각",
            "1st_blogger", "1st_today", "1st_est",
            "2nd_blogger", "2nd_today", "2nd_est",
            "3rd_blogger", "3rd_today", "3rd_est",
            "4th_blogger", "4th_today", "4th_est",
            "5th_blogger", "5th_today", "5th_est",
        ]
        ws_bloggerlist.append(headers)

        # 헤더 스타일 적용
        header_font = Font(bold=True)
        for cell in ws_bloggerlist[1]:
            cell.font = header_font
            cell.alignment = Alignment(vertical="center")

        # 필터 및 고정 설정
        ws_bloggerlist.freeze_panes = "A2"
        ws_bloggerlist.auto_filter.ref = f"A1:{chr(ord('A') + len(headers) - 1)}1"

        # .env 파일에서 API 키 로드
        load_dotenv()
        client_id = os.getenv("NAVER_CLIENT_ID", "").strip()
        client_secret = os.getenv("NAVER_CLIENT_SECRET", "").strip()

        if not client_id or not client_secret:
            print("ERROR: .env 파일에 NAVER_CLIENT_ID / NAVER_CLIENT_SECRET 이 없습니다.")
            return

        # 실행 시각 기록
        kst = timezone(timedelta(hours=9))
        run_timestamp = datetime.now(kst)

        # recent30days_sorted 시트에서 F열이 "O"인 행 필터링 및 처리
        processed_count = 0
        for row in range(2, ws_source.max_row + 1):  # 헤더 제외
            try:
                # F열 값 확인 (6번째 열)
                f_value = ws_source.cell(row=row, column=6).value

                if f_value == "O":
                    # A열에서 키워드 추출 (1번째 열)
                    keyword = ws_source.cell(row=row, column=1).value

                    if keyword:
                        print(f"[INFO] 키워드 '{keyword}' 처리 중...")

                        # 블로거 정보 수집
                        blogger_data = get_blogger_data_from_main_page(
                            keyword=keyword,
                            client_id=client_id,
                            client_secret=client_secret,
                            run_timestamp=run_timestamp,
                            top_n=5,
                            debug=False,  # 배치 처리이므로 디버그 모드 비활성화
                            debug_file=None,
                        )

                        # 데이터 작성
                        row_data = [keyword, run_timestamp.strftime('%Y-%m-%d %H:%M:%S')]

                        # blogger 데이터 추가
                        for blogger_info in blogger_data:
                            row_data.extend([
                                blogger_info["blogger"],
                                blogger_info["today"],
                                blogger_info["est"],
                            ])

                        # bloggerlist 시트에 데이터 추가
                        ws_bloggerlist.append(row_data)
                        processed_count += 1

                        # API 호출 간격 조절 (배치 처리이므로 약간 더 긴 간격)
                        time.sleep(1)

            except Exception as e:
                print(f"[WARN] {row}행 처리 중 오류: {e}")
                continue

        # 컬럼 너비 자동 조정
        for col_idx, col_name in enumerate(headers, start=1):
            max_len = len(col_name)
            for row_idx in range(1, ws_bloggerlist.max_row + 1):
                cell_value = ws_bloggerlist.cell(row=row_idx, column=col_idx).value
                if cell_value is not None:
                    max_len = max(max_len, len(str(cell_value)))
            ws_bloggerlist.column_dimensions[chr(64 + col_idx)].width = min(max_len + 2, 60)

        # 파일 저장
        try:
            wb.save(excel_path)
            print("\n[완료] bloggerlist 시트 생성 완료!")
            print(f"실행 시각: {run_timestamp.strftime('%Y-%m-%d %H:%M:%S')} (KST)")
            print(f"처리된 키워드 수: {processed_count}개")
            print(f"총 데이터 행 수: {ws_bloggerlist.max_row - 1}개")  # 헤더 제외
        except PermissionError:
            print(f"\n경고: '{excel_path}' 파일을 저장할 수 없습니다. 파일이 열려있는지 확인해주세요.")

    except Exception as e:
        print(f"bloggerlist 생성 중 오류 발생: {e}")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] in ["-h", "--help"]:
        print("사용법: python add_blogger_by_mainPage.py [키워드] [--debug] [--bloggerlist]")
        print("  키워드: 검색할 키워드 (입력하지 않으면 직접 입력)")
        print("  --debug: 블로그 검색 결과를 터미널에 출력합니다.")
        print("  --bloggerlist: recent30days_sorted 시트에서 F열이 'O'인 행만 bloggerlist 시트에 복사")
        sys.exit(0)

    # bloggerlist 생성 모드
    if "--bloggerlist" in sys.argv:
        print("=== bloggerlist 시트 생성 모드 ===")
        filter_and_save_blogger_list()
    else:
        main()