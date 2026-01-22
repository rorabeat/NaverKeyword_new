#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
네이버 블로그 포스트 크롤링 스크립트

이 스크립트는 네이버 블로그 포스트 URL을 입력받아 제목과 본문을 추출하여
UTF-8 텍스트 파일로 저장합니다.

또는 키워드와 크롤링할 갯수를 입력받아 자동으로 블로그를 검색하고 크롤링합니다.

필요 패키지 설치:
    pip install requests beautifulsoup4
"""

import os
import re
import sys
from urllib.parse import urljoin, urlparse
from datetime import datetime
from typing import Optional, Tuple, List, Dict, Any
import requests
from bs4 import BeautifulSoup

# G_add_blogger_by_mainPage.py에서 필요한 함수들 import
try:
    import sys
    import os
    # 상위 디렉토리 경로 추가
    parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if parent_dir not in sys.path:
        sys.path.insert(0, parent_dir)

    from G_add_blogger_by_mainPage import get_top_bloggers_from_main_page
    BLOGGER_SEARCH_AVAILABLE = True
except ImportError as e:
    BLOGGER_SEARCH_AVAILABLE = False
    print(f"경고: G_add_blogger_by_mainPage.py를 찾을 수 없습니다. 키워드 검색 기능이 비활성화됩니다.")
    print(f"Import 오류: {e}")
    print("현재 경로:", os.getcwd())
    print("스크립트 경로:", os.path.abspath(__file__))


class NaverBlogCrawler:
    """네이버 블로그 크롤러 클래스"""

    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'ko-KR,ko;q=0.8,en-US;q=0.5,en;q=0.3',
            'Accept-Encoding': 'gzip, deflate, br',
            'DNT': '1',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
        })

    def parse_blog_url(self, post_url: str) -> Tuple[str, str]:
        """
        블로그 URL에서 blogId와 logNo를 추출
        예: https://blog.naver.com/jjpapa1107/224149729792
        """
        # URL 파싱
        parsed = urlparse(post_url)

        # 경로에서 blogId/logNo 추출
        path_parts = parsed.path.strip('/').split('/')
        if len(path_parts) >= 2:
            blog_id = path_parts[0]
            log_no = path_parts[1]
            return blog_id, log_no

        raise ValueError(f"올바른 네이버 블로그 URL 형식이 아닙니다: {post_url}")

    def get_inner_frame_url(self, post_url: str) -> str:
        """
        메인 페이지에서 iframe#mainFrame의 src를 찾아 실제 컨텐츠 URL 생성
        """
        try:
            response = self.session.get(post_url, timeout=10)
            response.raise_for_status()

            soup = BeautifulSoup(response.content, 'html.parser')

            # iframe#mainFrame 찾기
            iframe = soup.find('iframe', id='mainFrame')
            if not iframe:
                raise ValueError("iframe#mainFrame을 찾을 수 없습니다")

            src = iframe.get('src')
            if not src:
                raise ValueError("iframe#mainFrame에 src 속성이 없습니다")

            # 상대 URL이면 절대 URL로 변환
            inner_url = urljoin(post_url, src)
            return inner_url

        except Exception as e:
            raise Exception(f"iframe URL 추출 실패: {e}")

    def extract_title(self, soup: BeautifulSoup) -> str:
        """
        제목 추출 (여러 방법 시도)
        """
        # 방법 1: 구조화된 제목 선택자
        title_selectors = [
            '.se-component.se-documentTitle .se-title-text span',
            '.se-component.se-documentTitle span',
            '.se-title-text span',
            'h3.se_textarea',
        ]

        for selector in title_selectors:
            title_elem = soup.select_one(selector)
            if title_elem:
                title = title_elem.get_text(strip=True)
                if title:
                    return title

        # 방법 2: og:title 메타 태그
        og_title = soup.find('meta', property='og:title')
        if og_title and og_title.get('content'):
            return og_title['content'].strip()

        # 방법 3: 일반 title 태그
        title_tag = soup.find('title')
        if title_tag:
            title = title_tag.get_text(strip=True)
            # 네이버 블로그 특유의 접미사 제거
            title = re.sub(r'\s*[:|]\s*네이버 블로그$', '', title)
            return title

        return "제목 없음"

    def extract_blogger_info(self, soup: BeautifulSoup) -> Tuple[str, str]:
        """
        블로거명과 작성날짜 추출
        return: (blogger_name, publish_date)
        """
        blogger_name = "알 수 없음"
        publish_date = "알 수 없음"

        # 블로거명 추출
        blogger_selectors = [
            '.nick',
            '.blogger_name',
            '.pcol1',
            '.author',
            '.writer',
            '.se_author',
            '.blog_author'
        ]

        for selector in blogger_selectors:
            blogger_elem = soup.select_one(selector)
            if blogger_elem:
                name = blogger_elem.get_text(strip=True)
                if name and len(name) > 0:
                    blogger_name = name
                    break

        # 작성날짜 추출
        date_selectors = [
            '.se_publishDate',
            '.date',
            '.se_date',
            '.publish_date',
            '.post_date',
            '.blog_date',
            'time',
            '.date-fil3'
        ]

        for selector in date_selectors:
            date_elem = soup.select_one(selector)
            if date_elem:
                # datetime 속성 확인
                if date_elem.get('datetime'):
                    publish_date = date_elem['datetime']
                    break
                # 텍스트 내용 확인
                date_text = date_elem.get_text(strip=True)
                if date_text and len(date_text) > 0:
                    publish_date = date_text
                    break

        return blogger_name, publish_date

    def extract_content(self, soup: BeautifulSoup) -> str:
        """
        본문 내용 추출
        """
        content_parts = []

        # 메인 컨테이너 찾기
        main_container = soup.select_one('.se-main-container')
        if not main_container:
            # 대안 선택자들
            main_container = soup.select_one('.post_ct') or soup.select_one('#postViewArea') or soup
            if not main_container:
                return "본문 내용을 찾을 수 없습니다"

        # 텍스트 요소들 수집
        text_selectors = [
            'p.se-text-paragraph',
            'li.se-text-list-item',
            '.se-component.se-quotation',
            '.se-text-paragraph',
            '.se-text',
            'p',
        ]

        for selector in text_selectors:
            elements = main_container.select(selector)
            for elem in elements:
                text = elem.get_text(strip=True)
                if text and len(text) > 1:  # 너무 짧은 텍스트 제외
                    content_parts.append(text)

        # 중복 제거 및 정리
        seen = set()
        unique_parts = []
        for part in content_parts:
            if part not in seen:
                seen.add(part)
                unique_parts.append(part)

        return '\n\n'.join(unique_parts) if unique_parts else "본문 내용이 없습니다"

    def _is_advertisement(self, text: str) -> bool:
        """
        광고성 텍스트인지 판별
        """
        ad_keywords = [
            '광고', '협찬', '후원', '스폰서', 'PR', '프로모션',
            '공감', '댓글', '공유', '좋아요', '팔로우',
            '이 글은', '이 포스트는', '블로그 운영정책',
            '네이버 블로그', '블로그 마켓'
        ]

        text_lower = text.lower()
        return any(keyword in text_lower for keyword in ad_keywords)

    def _sanitize_filename(self, filename: str) -> str:
        """
        파일명으로 사용할 수 없는 문자들을 제거하거나 안전한 문자로 변환
        """
        import re

        # 파일명으로 사용할 수 없는 문자들 제거 또는 변환
        # \ / : * ? " < > | 와 같은 문자들
        filename = re.sub(r'[\/:*?"<>|]', '', filename)

        # 연속된 공백을 하나의 공백으로 변환
        filename = re.sub(r'\s+', ' ', filename)

        # 앞뒤 공백 제거
        filename = filename.strip()

        # 빈 문자열인 경우 기본값 설정
        if not filename:
            filename = "제목없음"

        # 파일명이 너무 긴 경우 자르기 (Windows 파일명 길이 제한 고려)
        if len(filename) > 100:
            filename = filename[:100].strip()

        return filename

    def crawl_blog_post(self, post_url: str) -> Tuple[str, str, str, str, str]:
        """
        블로그 포스트 크롤링
        return: (title, content, filename, blogger_name, publish_date)
        """
        try:
            # URL에서 blogId와 logNo 추출
            blog_id, log_no = self.parse_blog_url(post_url)

            # iframe을 통한 실제 컨텐츠 URL 가져오기
            inner_url = self.get_inner_frame_url(post_url)
            print(f"실제 컨텐츠 URL: {inner_url}")

            # 실제 컨텐츠 페이지 요청
            response = self.session.get(inner_url, timeout=15)
            response.raise_for_status()

            # HTML 파싱
            soup = BeautifulSoup(response.content, 'html.parser', from_encoding='utf-8')

            # 제목 추출
            title = self.extract_title(soup)
            print(f"추출된 제목: {title}")

            # 블로거 정보 추출
            blogger_name, publish_date = self.extract_blogger_info(soup)
            print(f"블로거: {blogger_name}, 작성일: {publish_date}")

            # 본문 추출
            content = self.extract_content(soup)
            print(f"본문 길이: {len(content)}자")

            # 파일명 생성 (제목_날짜시간 형식)
            now = datetime.now().strftime('%Y%m%d_%H%M%S')
            # 제목에서 파일명으로 사용할 수 없는 문자들 제거
            safe_title = self._sanitize_filename(title)
            filename = f"{safe_title}_{now}.txt"

            return title, content, filename, blogger_name, publish_date

        except Exception as e:
            raise Exception(f"블로그 크롤링 실패: {e}")

    def save_to_file(self, title: str, content: str, filename: str, url: str = "", blogger_name: str = "", publish_date: str = "", output_dir: str = None) -> str:
        """
        추출된 내용을 파일로 저장
        """
        # 기본 출력 디렉토리 설정 (스크립트 위치 기준)
        if output_dir is None:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            output_dir = os.path.join(script_dir, "result")

        # 출력 디렉토리 생성
        os.makedirs(output_dir, exist_ok=True)

        # 전체 파일 경로
        filepath = os.path.join(output_dir, filename)

        # 내용 포맷팅
        full_content = f"블로그 URL: {url}\n"
        if blogger_name and blogger_name != "알 수 없음":
            full_content += f"블로거: {blogger_name}\n"
        if publish_date and publish_date != "알 수 없음":
            full_content += f"작성일: {publish_date}\n"
        full_content += f"\n{title}\n\n{content}"

        # UTF-8로 저장
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(full_content)

        print(f"파일 저장 완료: {filepath}")
        return filepath


def get_blog_posts_by_keyword(keyword: str, count: int) -> List[str]:
    """
    키워드로 블로그 검색하여 포스트 URL들 반환
    """
    try:
        # 환경변수 로드
        import os
        from dotenv import load_dotenv
        load_dotenv()

        client_id = os.getenv("NAVER_CLIENT_ID", "").strip()
        client_secret = os.getenv("NAVER_CLIENT_SECRET", "").strip()

        if not client_id or not client_secret:
            raise RuntimeError("NAVER_CLIENT_ID와 NAVER_CLIENT_SECRET 환경변수가 필요합니다.")

        # 네이버 검색 API로 블로그 글 검색
        search_url = "https://openapi.naver.com/v1/search/blog.json"
        headers = {
            "X-Naver-Client-Id": client_id,
            "X-Naver-Client-Secret": client_secret
        }

        params = {
            "query": keyword,
            "display": min(count, 100),  # 최대 10개로 제한 (네이버 API 제한)
            "start": 1,
            "sort": "sim"  # 정확도순 정렬
        }

        response = requests.get(search_url, headers=headers, params=params)
        response.raise_for_status()

        search_results = response.json()

        post_urls = []
        if "items" in search_results:
            for item in search_results["items"]:
                post_url = item.get("link", "")
                if post_url and "blog.naver.com" in post_url:
                    # 네이버 블로그 URL만 필터링
                    post_urls.append(post_url)

        return post_urls[:count]  # 요청한 갯수만큼 반환

    except Exception as e:
        raise RuntimeError(f"블로그 검색 실패: {e}")


def crawl_multiple_posts(post_urls: List[str], keyword: str = None) -> None:
    """
    여러 개의 블로그 포스트를 크롤링
    keyword: 검색 키워드 (지정시 키워드 폴더에 저장)
    """
    crawler = NaverBlogCrawler()
    success_count = 0
    fail_count = 0

    # 출력 디렉토리 설정 (키워드가 있으면 키워드 폴더 사용)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if keyword:
        # 키워드를 안전한 폴더명으로 변환
        safe_keyword = crawler._sanitize_filename(keyword)
        output_dir = os.path.join(script_dir, "result", safe_keyword)
    else:
        output_dir = os.path.join(script_dir, "result")

    print(f"\n총 {len(post_urls)}개의 블로그 포스트를 크롤링합니다...")
    print(f"저장 폴더: {output_dir}")
    print("-" * 60)

    for i, post_url in enumerate(post_urls, 1):
        try:
            print(f"\n[{i}/{len(post_urls)}] 크롤링 시작: {post_url}")

            # 블로그 크롤링
            title, content, filename, blogger_name, publish_date = crawler.crawl_blog_post(post_url)

            # 파일 저장
            saved_path = crawler.save_to_file(title, content, filename, post_url, blogger_name, publish_date, output_dir)

            print(f"✓ 성공: {saved_path}")
            success_count += 1

        except Exception as e:
            print(f"✗ 실패: {e}")
            fail_count += 1

    print("\n" + "="*60)
    print("크롤링 완료!")
    print(f"성공: {success_count}개")
    print(f"실패: {fail_count}개")
    if success_count > 0:
        print(f"저장 위치: 블로그크롤링/result/ 폴더")
    print("="*60)


def get_user_input() -> Tuple[str, int]:
    """
    사용자 입력 받기 (키워드와 크롤링할 갯수)
    """
    print("네이버 블로그 크롤링 도구")
    print("="*50)

    # 키워드 입력
    print("\n📝 검색할 키워드를 입력하세요")
    print("예: 오키나와 여행, 도쿄 맛집, 후쿠오카 관광지")
    while True:
        try:
            keyword = input("\n키워드 입력: ").strip()
            if keyword:
                print(f"✅ 선택된 키워드: '{keyword}'")
                break
            print("❌ 키워드를 입력해주세요.")
        except KeyboardInterrupt:
            print("\n\n👋 프로그램을 종료합니다.")
            sys.exit(0)

    # 갯수 입력
    print(f"\n🔢 '{keyword}' 키워드로 검색할 블로그 갯수를 입력하세요")
    print("권장: 3-5개 (너무 많으면 시간이 오래 걸릴 수 있습니다)")
    while True:
        try:
            count_input = input("\n갯수 입력 (1-100): ").strip()
            count = int(count_input)
            if 1 <= count <= 100:
                print(f"✅ 크롤링할 블로그 수: {count}개")
                break
            else:
                print("❌ 1에서 100 사이의 숫자를 입력해주세요.")
        except ValueError:
            print("❌ 올바른 숫자를 입력해주세요.")
        except KeyboardInterrupt:
            print("\n\n👋 프로그램을 종료합니다.")
            sys.exit(0)

    print(f"\n🚀 '{keyword}' 키워드로 {count}개의 블로그를 검색하고 크롤링을 시작합니다!")
    print("-" * 50)

    return keyword, count




def main():
    """메인 함수"""
    try:
        # 키워드 검색 및 크롤링
        keyword, count = get_user_input()

        # 블로그 검색
        post_urls = get_blog_posts_by_keyword(keyword, count)

        if not post_urls:
            print(f"\n❌ '{keyword}' 키워드로 검색된 블로그가 없습니다.")
            print("다른 키워드를 시도해보세요.")
            return

        print(f"\n🔍 검색된 블로그 목록:")
        for i, url in enumerate(post_urls, 1):
            print(f"  {i}. {url}")

        # 크롤링 시작
        crawl_multiple_posts(post_urls, keyword)

    except KeyboardInterrupt:
        print("\n\n👋 프로그램을 종료합니다.")
    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")


if __name__ == "__main__":
    main()