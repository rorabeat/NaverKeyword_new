#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import logging
import os
import re
import time
from datetime import datetime
from html import unescape
from typing import List, Optional

import requests
from bs4 import BeautifulSoup


def strip_html(text: str) -> str:
    if not text:
        return ""
    text = unescape(text)
    text = re.sub(r"<[^>]+>", "", text)  # <b> 포함 제거
    return text.strip()


def get_news_links(query: str, date: str) -> List[str]:
    """모바일 네이버 검색에서 뉴스 링크 추출"""
    url = f"https://m.search.naver.com/search.naver?ssc=tab.m_news.all&query={query}&sm=mtb_opt&sort=0&photo=0&field=0&pd=-1&ds={date}&de={date}&docid=&related=0&mynews=0&office_type=0&office_section_code=0&news_office_checked=&nso=&is_sug_officeid=0&office_category=0&service_area=1"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    
    soup = BeautifulSoup(response.text, 'html.parser')
    links = []
    for a in soup.select('a.news_tit'):
        href = a.get('href')
        if href and 'news.naver.com' in href:
            links.append(href)
    return links


def get_news_content(url: str) -> Optional[str]:
    """뉴스 페이지에서 본문 추출"""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, 'html.parser')
        article = soup.find('article', id='dic_area')
        if article:
            return strip_html(str(article))
        return None
    except Exception as e:
        logging.error(f"본문 추출 실패: {url} - {e}")
        return None


def save_to_txt(content: str, filename: str) -> None:
    """본문을 txt 파일로 저장"""
    os.makedirs('result', exist_ok=True)
    with open(f'result/{filename}.txt', 'w', encoding='utf-8') as f:
        f.write(content)


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(message)s")

    # 검색어 입력
    query = input("검색어를 입력하세요: ").strip()
    if not query:
        print("검색어가 비었습니다.")
        return 2

    # 최대 뉴스 수 입력
    max_news_input = input("가져올 뉴스 갯수를 입력하세요 (기본 10): ").strip()
    if not max_news_input:
        max_news = 10
    else:
        try:
            max_news = int(max_news_input)
        except ValueError:
            print("숫자로 입력하세요.")
            return 2

    # 오늘 날짜
    today = datetime.now().strftime("%Y.%m.%d")

    # 뉴스 링크 가져오기
    links = get_news_links(query, today)
    if not links:
        print("뉴스를 찾을 수 없습니다.")
        return 1

    links = links[:max_news]  # 최대 수 제한

    for i, link in enumerate(links, 1):
        logging.info(f"크롤링 중: {link}")
        content = get_news_content(link)
        if content:
            filename = f"news_{i}_{query.replace(' ', '_')}"
            save_to_txt(content, filename)
            logging.info(f"저장 완료: result/{filename}.txt")
        else:
            logging.warning(f"본문 추출 실패: {link}")
        time.sleep(1)  # 딜레이

    return 0


if __name__ == "__main__":
    raise SystemExit(main())




def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(message)s")

    # 검색어 입력
    queries_input = input("검색어를 입력하세요 (여러 개는 쉼표로 구분): ").strip()
    if not queries_input:
        print("검색어가 비었습니다.")
        return 2
    queries = [q.strip() for q in queries_input.split(",") if q.strip()]

    # 최대 기사 수 입력
    max_items_input = input("가져올 기사 갯수를 입력하세요 (키워드 1개당, 최대 1000, 기본 100): ").strip()
    if not max_items_input:
        max_items = 100
    else:
        try:
            max_items = int(max_items_input)
            if not (1 <= max_items <= NAVER_MAX_START):
                print(f"기사 갯수는 1~{NAVER_MAX_START} 범위여야 합니다.")
                return 2
        except ValueError:
            print("숫자로 입력하세요.")
            return 2

    try:
        creds = load_creds_from_dotenv()
    except Exception as e:
        print(str(e))
        return 2

    all_rows: List[Dict[str, Any]] = []

    for q in queries:
        logging.info("수집 시작 | query='%s' | max=%d", q, max_items)
        try:
            rows = fetch_news(
                creds=creds,
                query=q,
                max_items=max_items,
                page_size=100,
                sort="date",
                sleep_sec=0.2,
            )
            logging.info("수집 완료 | query='%s' | got=%d", q, len(rows))
            all_rows.extend(rows)
        except Exception as e:
            logging.error("수집 실패 | query='%s' | %s", q, e)
            continue

    # 저장
    out_path = "result/naver_news_out.xlsx"
    write_xlsx(out_path, all_rows)

    logging.info("저장 완료 | out=%s | rows=%d", out_path, len(all_rows))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
