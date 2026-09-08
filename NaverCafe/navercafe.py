#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
네이버 카페 게시판 크롤러

기능:
1) 게시판 목록 페이지를 순회하며 게시글 URL 수집
2) 게시글 상세에서 제목/날짜/본문/댓글 추출
3) 사용자가 입력한 날짜의 게시글만 필터링
4) 결과를 JSON 파일로 저장

사용 예시:
    python navercafe.py --target-date 2026-02-20
    python navercafe.py --target-date 2026.02.20 --start-page 1 --end-page 20
"""

from __future__ import annotations

import argparse
import json
import re
import time
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Set
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup


BASE_URL = "https://cafe.naver.com"
DEFAULT_CAFE_ID = "23068408"
DEFAULT_MENU_ID = "407"
DEFAULT_SLEEP = 0.3


@dataclass
class Comment:
    author: str
    date: str
    content: str


@dataclass
class Article:
    article_id: str
    title: str
    date: str
    url: str
    content: str
    comments: List[Comment]


def normalize_spaces(text: str) -> str:
    if not text:
        return ""
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def parse_target_date(date_text: str) -> str:
    """
    입력 날짜를 YYYY-MM-DD 형식으로 정규화.
    지원: YYYY-MM-DD, YYYY.MM.DD, YYYY/MM/DD
    """
    date_text = date_text.strip()
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(date_text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    raise ValueError("날짜 형식이 올바르지 않습니다. 예: 2026-02-20 또는 2026.02.20")


def extract_article_date(raw_date_text: str) -> str:
    """
    예: '2026.02.20. 16:49' -> '2026-02-20'
    """
    if not raw_date_text:
        return ""
    match = re.search(r"(\d{4})[./-](\d{1,2})[./-](\d{1,2})", raw_date_text)
    if not match:
        return ""
    y, m, d = match.groups()
    return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"


class NaverCafeCrawler:
    def __init__(
        self,
        cafe_id: str,
        menu_id: str,
        start_page: int = 1,
        end_page: int = 1,
        sleep_sec: float = DEFAULT_SLEEP,
        user_agent: Optional[str] = None,
    ) -> None:
        self.cafe_id = str(cafe_id)
        self.menu_id = str(menu_id)
        self.start_page = start_page
        self.end_page = end_page
        self.sleep_sec = sleep_sec
        self.session = requests.Session()
        self.session.headers.update(
            {
                "User-Agent": user_agent
                or "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
                "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
            }
        )

    def list_page_url(self, page: int) -> str:
        return f"{BASE_URL}/f-e/cafes/{self.cafe_id}/menus/{self.menu_id}?page={page}"

    def article_url(self, article_id: str) -> str:
        return (
            f"{BASE_URL}/f-e/cafes/{self.cafe_id}/articles/{article_id}"
            f"?menuid={self.menu_id}&referrerAllArticles=false&page=1"
        )

    def fetch_html(self, url: str, timeout: int = 20) -> str:
        response = self.session.get(url, timeout=timeout)
        response.raise_for_status()
        return response.text

    def extract_article_ids_from_list(self, html: str) -> List[str]:
        """
        목록 페이지에서 article_id를 중복 없이 추출.
        """
        ids: Set[str] = set()
        soup = BeautifulSoup(html, "lxml")

        # 1차: a[href] 기반 추출
        for a_tag in soup.select("a[href]"):
            href = a_tag.get("href", "")
            if not href:
                continue
            full = urljoin(BASE_URL, href)
            m = re.search(r"/cafes/\d+/articles/(\d+)", full)
            if m:
                ids.add(m.group(1))

        # 2차: 원문 전체에서 regex 추출 (동적 링크 대응)
        for m in re.finditer(r"/cafes/\d+/articles/(\d+)", html):
            ids.add(m.group(1))

        return sorted(ids, key=int, reverse=True)

    def parse_article(self, html: str, article_url: str, article_id: str) -> Article:
        soup = BeautifulSoup(html, "lxml")

        # 제목
        title_tag = soup.select_one(".ArticleTitle .title_text") or soup.select_one("h3.title_text")
        title = normalize_spaces(title_tag.get_text(" ", strip=True) if title_tag else "")

        # 날짜
        date_tag = soup.select_one(".WriterInfo .article_info .date") or soup.select_one(".article_info .date")
        raw_date = normalize_spaces(date_tag.get_text(" ", strip=True) if date_tag else "")
        normalized_date = extract_article_date(raw_date)

        # 본문
        content_lines: List[str] = []
        for p in soup.select(".ArticleContentBox .se-main-container .se-text-paragraph"):
            text = normalize_spaces(p.get_text(" ", strip=True))
            if text:
                content_lines.append(text)

        # fallback: 일반 본문 선택자
        if not content_lines:
            for p in soup.select(".ArticleContentBox .article_viewer p"):
                text = normalize_spaces(p.get_text(" ", strip=True))
                if text:
                    content_lines.append(text)

        content = "\n".join(content_lines).strip()

        # 댓글
        comments: List[Comment] = []
        for li in soup.select(".CommentBox .comment_list li.CommentItem"):
            author_tag = li.select_one(".comment_nickname")
            date_tag = li.select_one(".comment_info_date")
            text_tag = li.select_one(".text_comment")

            author = normalize_spaces(author_tag.get_text(" ", strip=True) if author_tag else "")
            c_date = normalize_spaces(date_tag.get_text(" ", strip=True) if date_tag else "")
            c_text = normalize_spaces(text_tag.get_text(" ", strip=True) if text_tag else "")

            if author or c_text:
                comments.append(Comment(author=author, date=c_date, content=c_text))

        return Article(
            article_id=article_id,
            title=title,
            date=normalized_date,
            url=article_url,
            content=content,
            comments=comments,
        )

    def crawl(self, target_date: str) -> List[Article]:
        target_date = parse_target_date(target_date)
        seen_ids: Set[str] = set()
        matched_articles: List[Article] = []

        print(f"[INFO] 대상 날짜: {target_date}")
        print(f"[INFO] 페이지 범위: {self.start_page} ~ {self.end_page}")

        for page in range(self.start_page, self.end_page + 1):
            list_url = self.list_page_url(page)
            print(f"[INFO] 목록 페이지 수집: {list_url}")

            try:
                list_html = self.fetch_html(list_url)
            except Exception as exc:
                print(f"[WARN] 목록 페이지 수집 실패(page={page}): {exc}")
                continue

            article_ids = self.extract_article_ids_from_list(list_html)
            if not article_ids:
                print(f"[INFO] page={page} 에서 게시글 ID를 찾지 못했습니다.")
                continue

            print(f"[INFO] page={page} 게시글 후보 수: {len(article_ids)}")

            for article_id in article_ids:
                if article_id in seen_ids:
                    continue
                seen_ids.add(article_id)

                a_url = self.article_url(article_id)
                try:
                    article_html = self.fetch_html(a_url)
                    article = self.parse_article(article_html, a_url, article_id)
                except Exception as exc:
                    print(f"[WARN] 게시글 수집 실패(article_id={article_id}): {exc}")
                    continue

                # 날짜 정보가 없거나 대상 날짜와 다르면 스킵
                if article.date != target_date:
                    continue

                matched_articles.append(article)
                print(
                    f"[MATCH] id={article.article_id} / title={article.title[:40]} / "
                    f"comments={len(article.comments)}"
                )

                time.sleep(self.sleep_sec)

        return matched_articles


def save_articles_json(articles: List[Article], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    payload = []
    for a in articles:
        item = asdict(a)
        payload.append(item)

    with output_path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="네이버 카페 날짜별 게시글 크롤러")
    parser.add_argument("--target-date", required=False, help="수집 대상 날짜 (예: 2026-02-20)")
    parser.add_argument("--cafe-id", default=DEFAULT_CAFE_ID, help=f"카페 ID (기본값: {DEFAULT_CAFE_ID})")
    parser.add_argument("--menu-id", default=DEFAULT_MENU_ID, help=f"메뉴 ID (기본값: {DEFAULT_MENU_ID})")
    parser.add_argument("--start-page", type=int, default=1, help="시작 페이지")
    parser.add_argument("--end-page", type=int, default=10, help="종료 페이지")
    parser.add_argument("--sleep-sec", type=float, default=DEFAULT_SLEEP, help="요청 간 대기 시간(초)")
    parser.add_argument(
        "--output",
        default="NaverCafe/result/navercafe_result.json",
        help="출력 JSON 파일 경로",
    )
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    target_date = args.target_date
    if not target_date:
        target_date = input("수집할 날짜를 입력하세요 (예: 2026-02-20): ").strip()

    try:
        target_date = parse_target_date(target_date)
    except ValueError as exc:
        print(f"[ERROR] {exc}")
        return

    crawler = NaverCafeCrawler(
        cafe_id=args.cafe_id,
        menu_id=args.menu_id,
        start_page=args.start_page,
        end_page=args.end_page,
        sleep_sec=args.sleep_sec,
    )
    articles = crawler.crawl(target_date=target_date)

    output_path = Path(args.output)
    save_articles_json(articles, output_path)

    print("-" * 70)
    print(f"[DONE] 수집 완료: {len(articles)}건")
    print(f"[DONE] 저장 파일: {output_path.resolve()}")


if __name__ == "__main__":
    main()
