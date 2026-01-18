# -*- coding: utf-8 -*-
"""Naver Keyword Gold Finder - Integrated Excel Pipeline

이 스크립트는 아래 신호들을 한 번에 수집해서 '황금 키워드' 후보를 엑셀로 뽑아줍니다.

- 자동완성(검색창 드롭다운) 연관검색어: Selenium(undetected_chromedriver)
- 우측/하단 연관검색어: requests + BeautifulSoup
- 광고 키워드도구(/keywordstool) 검색량(PC/모바일): SearchAd API
- 최근 30일 블로그 발행 수(100에서 캡 -> 100+): Naver Blog Search API
- 검색결과 총량(약 N건): requests 파싱(옵션)
- 점수표 스코어링

환경변수(.env 권장)
  # 네이버 블로그 검색 API
  NAVER_CLIENT_ID=...
  NAVER_CLIENT_SECRET=...

  # 네이버 검색광고(SearchAd) API
  NAVER_SEARCH_ACCESS_LICENSE_KEY=...
  NAVER_SEARCH_SECRET_KEY=...
  NAVER_SEARCH_CUSTOMER_ID=...

필수 패키지
  pip install python-dotenv requests beautifulsoup4 lxml openpyxl
  pip install selenium undetected-chromedriver

주의
- 네이버 페이지 구조가 변경되면 CSS selector가 바뀔 수 있습니다.
- 자동완성은 Selenium 기반이라 키워드가 너무 많으면 시간이 오래 걸립니다.
  (기본값으로 candidate 평가 상한을 둬서 폭주를 막습니다.)
"""

from __future__ import annotations

import os
import re
import time
import hmac
import base64
import hashlib
import logging
import atexit
import gc
import random
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import quote_plus

import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

# Selenium (자동완성)
try:
    import undetected_chromedriver as uc
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    SELENIUM_AVAILABLE = True
except Exception:
    SELENIUM_AVAILABLE = False


# -----------------------------
# 설정(필요시 여기만 튜닝)
# -----------------------------
AUTOCOMPLETE_TIMEOUT_SEC = 5
AUTOCOMPLETE_FAST_SEC = 0.8
AUTOCOMPLETE_SLOW_SEC = 3.0

BLOG_LAST30_CAP = 100
BLOG_LOOKBACK_DAYS = 30

# 너무 많은 candidate를 전부 평가하면 비용/시간이 폭주할 수 있어 기본 제한을 둠
MAX_CANDIDATES_TO_EVALUATE_PER_SEED = 120

# 우측 연관/검색결과 파싱용 User-Agent
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36"
)

NAVER_SEARCH_URL = "https://search.naver.com/search.naver"
NAVER_BLOG_API_URL = "https://openapi.naver.com/v1/search/blog.json"

SEARCHAD_BASE_URL = "https://api.searchad.naver.com"
SEARCHAD_ENDPOINT = "/keywordstool"


# -----------------------------
# 로깅
# -----------------------------

def setup_logger(log_path: str) -> logging.Logger:
    """콘솔 + 파일 로깅을 함께 설정."""
    logger = logging.getLogger("naver_keyword_pipeline")
    logger.setLevel(logging.DEBUG)

    # 중복 핸들러 방지
    if logger.handlers:
        return logger

    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")

    sh = logging.StreamHandler()
    sh.setLevel(logging.INFO)
    sh.setFormatter(fmt)

    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)

    logger.addHandler(sh)
    logger.addHandler(fh)
    return logger


# -----------------------------
# 유틸
# -----------------------------

def kst_now() -> datetime:
    return datetime.now(timezone(timedelta(hours=9)))


def safe_filename(text: str, max_len: int = 60) -> str:
    text = re.sub(r"[^0-9a-zA-Z가-힣]+", "_", text).strip("_")
    return (text[:max_len] if text else "output")


def strip_html(text: str) -> str:
    if not text:
        return ""
    return re.sub(r"<[^>]+>", "", text)


def dedupe_keep_order(items: Iterable[str]) -> List[str]:
    seen = set()
    out = []
    for x in items:
        x = (x or "").strip()
        if not x:
            continue
        if x in seen:
            continue
        seen.add(x)
        out.append(x)
    return out


# -----------------------------
# 1) 자동완성(검색창 드롭다운) - Selenium
# -----------------------------

@dataclass
class AutoCompleteResult:
    keyword: str
    suggestions: List[str]
    first_suggestion_delay: Optional[float]  # None이면 timeout


class AutoCompleteSession:
    """undetected_chromedriver 를 매 키워드마다 새로 띄우면
    Windows에서 __del__ 시점에 WinError 6이 반복될 수 있어, 드라이버 1개를 재사용."""

    def __init__(self, headless: bool = True, logger: Optional[logging.Logger] = None):
        if not SELENIUM_AVAILABLE:
            raise RuntimeError("Selenium/undetected_chromedriver 가 설치되지 않았습니다.")

        self.logger = logger
        options = uc.ChromeOptions()
        if headless:
            options.add_argument("--headless=new")

        # 안정화 옵션
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-blink-features=AutomationControlled")

        self.driver = uc.Chrome(options=options)
        self.driver.set_page_load_timeout(20)
        try:
            self.driver.get("https://www.naver.com")
        except Exception as e:
            if self.logger:
                self.logger.warning(f"[AUTO] 네이버 메인 로드 실패(재시도 진행): {e}")
            try:
                self.driver.get("https://www.naver.com")
            except Exception:
                pass

        # 프로세스 종료 시 안전하게 닫기
        atexit.register(self.close)

    def close(self):
        try:
            if getattr(self, "driver", None):
                try:
                    self.driver.quit()
                except OSError:
                    if self.logger:
                        self.logger.debug("[AUTO] driver.quit() OSError(WinError 6 등) 무시")
                    pass
                except Exception:
                    if self.logger:
                        self.logger.debug("[AUTO] driver.quit() 예외 무시", exc_info=True)
                    pass
        finally:
            self.driver = None
            gc.collect()

    def fetch(self, keyword: str) -> AutoCompleteResult:
        driver = self.driver
        if driver is None:
            raise RuntimeError("AutoCompleteSession driver is closed")

        suggestions: List[str] = []
        delay: Optional[float] = None

        try:
            if self.logger:
                self.logger.debug(f"[AUTO] 자동완성 시작: keyword='{keyword}'")

            wait = WebDriverWait(driver, 10)
            search_input = wait.until(EC.element_to_be_clickable((By.ID, "query")))

            search_input.click()
            search_input.clear()

            # 사람처럼 글자별로 랜덤 딜레이를 두면서 입력 (UnitTest 코드 참고)
            if self.logger:
                self.logger.debug(f"[AUTO] 입력 시작: '{keyword}'")
            for char in keyword:
                search_input.send_keys(char)
                typing_speed = random.uniform(0.1, 0.3)
                time.sleep(typing_speed)

            # 입력 완료 후 자동완성 목록이 갱신될 수 있도록 잠시 대기
            time.sleep(1.0)
            if self.logger:
                self.logger.debug(f"[AUTO] 입력 완료, 자동완성 대기 중...")

            t0 = time.perf_counter()
            try:
                if self.logger:
                    self.logger.debug(f"[AUTO] 자동완성 요소 대기 시작 (최대 {AUTOCOMPLETE_TIMEOUT_SEC}초)")
                WebDriverWait(driver, AUTOCOMPLETE_TIMEOUT_SEC).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "li._item[data-keyword]"))
                )
                delay = time.perf_counter() - t0
                if self.logger:
                    self.logger.debug(f"[AUTO] 자동완성 요소 발견: delay={delay:.3f}초")
            except Exception as e:
                delay = None
                if self.logger:
                    self.logger.debug(f"[AUTO] 자동완성 요소 대기 타임아웃 또는 실패: {repr(e)}")

            items = driver.find_elements(By.CSS_SELECTOR, "li._item[data-keyword]")
            for it in items:
                v = it.get_attribute("data-keyword")
                if v:
                    suggestions.append(v)

            suggestions = dedupe_keep_order(suggestions)
            if self.logger:
                self.logger.debug(f"[AUTO] 자동완성 수집 완료: keyword='{keyword}', suggestions={len(suggestions)}개, delay={delay}")
                if suggestions:
                    self.logger.debug(f"[AUTO] 추출된 키워드: {suggestions[:5]}{'...' if len(suggestions) > 5 else ''}")

            return AutoCompleteResult(keyword=keyword, suggestions=suggestions, first_suggestion_delay=delay)

        except Exception as e:
            if self.logger:
                self.logger.exception(f"[AUTO] 자동완성 수집 실패: keyword='{keyword}'")
            return AutoCompleteResult(keyword=keyword, suggestions=[], first_suggestion_delay=None)


def get_autocomplete_suggestions(keyword: str, headless: bool = True, logger: Optional[logging.Logger] = None) -> AutoCompleteResult:
    """호환을 위해 남겨둔 래퍼(세션을 안 쓰는 경우)."""
    session = AutoCompleteSession(headless=headless, logger=logger)
    try:
        return session.fetch(keyword)
    finally:
        session.close()


def classify_autocomplete_speed(delay: Optional[float]) -> str:
    if delay is None:
        return "none"
    if delay <= AUTOCOMPLETE_FAST_SEC:
        return "fast"
    if delay <= AUTOCOMPLETE_SLOW_SEC:
        return "slow"
    return "none"


# -----------------------------
# 2) 우측/하단 연관검색어 - requests
# -----------------------------

@dataclass
class RightRelatedResult:
    keyword: str
    related: List[str]


def fetch_naver_search_html(query: str) -> str:
    headers = {
        "User-Agent": UA,
        "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Referer": "https://search.naver.com/",
    }

    r = requests.get(
        NAVER_SEARCH_URL,
        params={"query": query},
        headers=headers,
        timeout=15,
    )
    r.raise_for_status()
    return r.text


def extract_right_related_from_html(html: str) -> List[str]:
    """사용자가 제공한 구조 기반.
    가장 우선: .api_subject_bx._related_box .lst_related_srch .tit
    대체: div.related_srch div.tit
    """
    soup = BeautifulSoup(html, "lxml")

    nodes = soup.select(".api_subject_bx._related_box .lst_related_srch .tit")
    related = [n.get_text(strip=True) for n in nodes if n.get_text(strip=True)]

    if not related:
        nodes2 = soup.select("div.related_srch div.tit")
        related = [n.get_text(strip=True) for n in nodes2 if n.get_text(strip=True)]

    if not related:
        nodes3 = soup.select("div.related_srch a.keyword")
        related = [n.get_text(strip=True) for n in nodes3 if n.get_text(strip=True)]

    return dedupe_keep_order(related)


def get_right_related(keyword: str) -> RightRelatedResult:
    html = fetch_naver_search_html(keyword)
    related = extract_right_related_from_html(html)
    return RightRelatedResult(keyword=keyword, related=related)


# -----------------------------
# 5) 검색결과 총량(약 N건) 파싱 - 옵션
# -----------------------------

@dataclass
class SearchTotalResult:
    keyword: str
    total_results: Optional[int]


def parse_total_results_from_html(html: str) -> Optional[int]:
    """네이버 통합검색 페이지에서 '약 12,345건' 숫자를 파싱.

    구조가 자주 바뀌므로 정규식 기반으로 여러 패턴을 시도한다.
    """
    # 가장 흔한 패턴: '약 12,345건'
    m = re.search(r"약\s*([0-9,]+)\s*건", html)
    if m:
        return int(m.group(1).replace(",", ""))

    # 대체 패턴: '검색결과 12,345건'
    m = re.search(r"검색결과\s*([0-9,]+)\s*건", html)
    if m:
        return int(m.group(1).replace(",", ""))

    return None


def get_search_total(keyword: str) -> SearchTotalResult:
    html = fetch_naver_search_html(keyword)
    total = parse_total_results_from_html(html)
    return SearchTotalResult(keyword=keyword, total_results=total)


# -----------------------------
# 3) 네이버 검색광고 키워드도구(/keywordstool)
# -----------------------------

@dataclass
class SearchAdMetrics:
    keyword: str
    exists_in_tool: bool
    monthly_pc: Optional[int]
    monthly_mobile: Optional[int]
    monthly_total: Optional[int]
    comp_idx: Optional[str]
    ad_depth: Optional[str]


def make_signature(secret_key: str, timestamp: str, method: str, uri: str) -> str:
    message = f"{timestamp}.{method}.{uri}"
    digest = hmac.new(secret_key.encode("utf-8"), message.encode("utf-8"), hashlib.sha256).digest()
    return base64.b64encode(digest).decode("utf-8")


def build_searchad_headers(api_key: str, secret_key: str, customer_id: str, method: str, uri: str) -> Dict[str, str]:
    timestamp = str(int(time.time() * 1000))
    signature = make_signature(secret_key, timestamp, method, uri)

    return {
        "X-Timestamp": timestamp,
        "X-API-KEY": api_key,
        "X-Customer": customer_id,
        "X-Signature": signature,
        "Content-Type": "application/json; charset=UTF-8",
    }


def call_keywordstool(hint_keywords: str, api_key: str, secret_key: str, customer_id: str, show_detail: int = 1) -> Dict[str, Any]:
    method = "GET"
    headers = build_searchad_headers(api_key, secret_key, customer_id, method, SEARCHAD_ENDPOINT)
    params = {"hintKeywords": hint_keywords, "showDetail": str(show_detail)}

    url = SEARCHAD_BASE_URL + SEARCHAD_ENDPOINT
    r = requests.get(url, headers=headers, params=params, timeout=15)
    r.raise_for_status()
    return r.json()


def _as_int_or_none(v: Any) -> Optional[int]:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return int(v)
    s = str(v).strip()
    if not s:
        return None
    if s.lower() in {"< 10", "<10", "< 10회", "<10회"}:
        return 9
    s = s.replace(",", "")
    if not s.isdigit():
        return None
    return int(s)


def get_searchad_metrics_for_keyword(
    keyword: str,
    api_key: str,
    secret_key: str,
    customer_id: str,
) -> SearchAdMetrics:
    # hintKeywords는 최대 5개지만 우리는 1개씩 호출(정확성/단순성). 필요하면 배치 최적화 가능.
    data = call_keywordstool(keyword, api_key, secret_key, customer_id, show_detail=1)

    keyword_list = data.get("keywordList")
    if keyword_list is None and isinstance(data, list):
        keyword_list = data
    if keyword_list is None:
        keyword_list = []

    # exact match 우선
    exact = None
    for row in keyword_list:
        if str(row.get("relKeyword", "")).strip() == keyword:
            exact = row
            break

    if exact is None:
        # exact가 없으면 첫 행이 hint와 가장 가까운 경우가 많지만, '없음'으로 처리
        return SearchAdMetrics(
            keyword=keyword,
            exists_in_tool=False,
            monthly_pc=None,
            monthly_mobile=None,
            monthly_total=None,
            comp_idx=None,
            ad_depth=None,
        )

    pc = _as_int_or_none(exact.get("monthlyPcQcCnt"))
    mo = _as_int_or_none(exact.get("monthlyMobileQcCnt"))
    total = (pc or 0) + (mo or 0) if (pc is not None or mo is not None) else None

    return SearchAdMetrics(
        keyword=keyword,
        exists_in_tool=True,
        monthly_pc=pc,
        monthly_mobile=mo,
        monthly_total=total,
        comp_idx=str(exact.get("compIdx")) if exact.get("compIdx") is not None else None,
        ad_depth=str(exact.get("plAvgDepth")) if exact.get("plAvgDepth") is not None else None,
    )


# -----------------------------
# 4) 네이버 블로그 검색 API: 최근 30일 발행 수(100+ 캡)
# -----------------------------

@dataclass
class BlogLast30:
    keyword: str
    last30_count: int  # 0~100
    over_100: bool
    cutoff_yyyymmdd: int
    api_total: Optional[int]


def fetch_blog_api(
    query: str,
    client_id: str,
    client_secret: str,
    display: int = 100,
    start: int = 1,
    sort: str = "date",
) -> Dict[str, Any]:
    headers = {
        "X-Naver-Client-Id": client_id,
        "X-Naver-Client-Secret": client_secret,
    }
    params = {"query": query, "display": display, "start": start, "sort": sort}

    r = requests.get(NAVER_BLOG_API_URL, headers=headers, params=params, timeout=10)
    r.raise_for_status()
    return r.json()


def count_blog_last30_capped(
    keyword: str,
    client_id: str,
    client_secret: str,
    limit: int = BLOG_LAST30_CAP,
    days: int = BLOG_LOOKBACK_DAYS,
) -> BlogLast30:
    cutoff = (kst_now() - timedelta(days=days)).date()
    cutoff_yyyymmdd = int(cutoff.strftime("%Y%m%d"))

    count = 0
    start = 1
    display = 100

    # 첫 페이지
    first = fetch_blog_api(keyword, client_id, client_secret, display=display, start=start, sort="date")
    api_total = first.get("total")

    def consume(items: List[Dict[str, Any]]) -> Tuple[bool, bool]:
        nonlocal count
        for it in items:
            postdate = it.get("postdate")
            if not postdate or len(str(postdate)) != 8:
                continue
            if int(postdate) < cutoff_yyyymmdd:
                return True, False
            count += 1
            if count >= limit:
                return True, True
        return False, False

    items = first.get("items", []) or []
    stop_by_date, reached_limit = consume(items)
    if reached_limit:
        return BlogLast30(keyword=keyword, last30_count=limit, over_100=True, cutoff_yyyymmdd=cutoff_yyyymmdd, api_total=api_total)
    if stop_by_date:
        return BlogLast30(keyword=keyword, last30_count=count, over_100=False, cutoff_yyyymmdd=cutoff_yyyymmdd, api_total=api_total)

    # 다음 페이지
    start = 101
    while True:
        data = fetch_blog_api(keyword, client_id, client_secret, display=display, start=start, sort="date")
        items = data.get("items", []) or []
        if not items:
            break
        stop_by_date, reached_limit = consume(items)
        if reached_limit:
            return BlogLast30(keyword=keyword, last30_count=limit, over_100=True, cutoff_yyyymmdd=cutoff_yyyymmdd, api_total=api_total)
        if stop_by_date:
            break
        start += display
        if start > 1000:
            break

    return BlogLast30(keyword=keyword, last30_count=count, over_100=False, cutoff_yyyymmdd=cutoff_yyyymmdd, api_total=api_total)


# -----------------------------
# 점수표
# -----------------------------

def score_keyword(
    autocomplete_speed: str,
    right_related_count: int,
    blog_last30_count: int,
    blog_over_100: bool,
    total_results: Optional[int],
) -> int:
    score = 0

    # 자동완성 빠름/느림
    if autocomplete_speed == "fast":
        score += 3
    elif autocomplete_speed == "slow":
        score += 1

    # 우측 연관 3개 이상
    if right_related_count >= 3:
        score += 2

    # 최근 30일 발행: 0(+1), 1~3(+3), 그 외 0
    if blog_over_100:
        pass
    elif blog_last30_count == 0:
        score += 1
    elif 1 <= blog_last30_count <= 3:
        score += 3

    # 검색결과 5,000+
    if total_results is not None and total_results >= 5000:
        score += 2

    return score


# -----------------------------
# Excel
# -----------------------------

def write_excel(
    out_path: str,
    rows: List[Dict[str, Any]],
    expansions: List[Dict[str, Any]],
) -> None:
    wb = Workbook()

    # Sheet 1: keywords
    ws = wb.active
    ws.title = "keywords"

    headers = [
        "seed",
        "keyword",
        "source",  # ac / right / both
        "score",
        "ac_delay_sec",
        "ac_speed",
        "right_related_count",
        "blog_last30_display",
        "blog_last30_count",
        "blog_over_100",
        "search_total_results",
        "adtool_exists",
        "monthly_pc",
        "monthly_mobile",
        "monthly_total",
        "comp_idx",
        "ad_depth",
    ]
    ws.append(headers)

    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = Alignment(vertical="center")

    for r in rows:
        ws.append([r.get(h, "") for h in headers])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{chr(64 + len(headers))}{len(rows) + 1}"

    # column widths (rough)
    for col_idx, col_name in enumerate(headers, start=1):
        max_len = len(col_name)
        for row_idx in range(2, len(rows) + 2):
            v = ws.cell(row=row_idx, column=col_idx).value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))
        ws.column_dimensions[chr(64 + col_idx)].width = min(max_len + 2, 60)

    # Sheet 2: expansions
    ws2 = wb.create_sheet("expansions")
    headers2 = ["seed", "type", "related_keyword"]
    ws2.append(headers2)
    for cell in ws2[1]:
        cell.font = header_font
        cell.alignment = Alignment(vertical="center")

    for r in expansions:
        ws2.append([r.get("seed", ""), r.get("type", ""), r.get("related_keyword", "")])

    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = f"A1:C{len(expansions) + 1}"
    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["B"].width = 12
    ws2.column_dimensions["C"].width = 50

    wb.save(out_path)


# -----------------------------
# 파이프라인
# -----------------------------


def read_seeds_from_user() -> List[str]:
    print("시드 키워드를 입력하세요. 여러 개면 콤마(,)로 구분하거나 줄바꿈으로 입력")
    print("입력 끝내려면 빈 줄에서 Enter\n")
    seeds: List[str] = []
    while True:
        line = input("seed > ").strip()
        if not line:
            break
        parts = [p.strip() for p in line.split(",") if p.strip()]
        seeds.extend(parts)
    return dedupe_keep_order(seeds)


def main() -> None:
    load_dotenv()

    ts = kst_now().strftime("%Y%m%d_%H%M%S")
    log_path = f"naver_keyword_pipeline_{ts}.log"
    logger = setup_logger(log_path)
    logger.info("=== naver_keyword_pipeline 시작 ===")

    # Blog API
    naver_client_id = (os.getenv("NAVER_CLIENT_ID") or "").strip()
    naver_client_secret = (os.getenv("NAVER_CLIENT_SECRET") or "").strip()

    # SearchAd
    ad_api_key = (os.getenv("NAVER_SEARCH_ACCESS_LICENSE_KEY") or "").strip()
    ad_secret_key = (os.getenv("NAVER_SEARCH_SECRET_KEY") or "").strip()
    ad_customer_id = (os.getenv("NAVER_SEARCH_CUSTOMER_ID") or "").strip()

    if not naver_client_id or not naver_client_secret:
        raise RuntimeError("ENV 누락: NAVER_CLIENT_ID / NAVER_CLIENT_SECRET")

    if not ad_api_key or not ad_secret_key or not ad_customer_id:
        raise RuntimeError("ENV 누락: NAVER_SEARCH_ACCESS_LICENSE_KEY / NAVER_SEARCH_SECRET_KEY / NAVER_SEARCH_CUSTOMER_ID")

    seeds = read_seeds_from_user()
    if not seeds:
        logger.info("시드 키워드가 없습니다. 종료")
        return

    if not SELENIUM_AVAILABLE:
        logger.warning("selenium/undetected_chromedriver 미설치 -> 자동완성(점수 일부) 없이 진행됩니다.")

    out_xlsx = f"naver_keyword_pipeline_{ts}.xlsx"

    all_rows: List[Dict[str, Any]] = []
    expansion_rows: List[Dict[str, Any]] = []

    logger.info(f"실행시각(KST): {kst_now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info(f"seed 개수: {len(seeds)}")

    # Selenium 드라이버는 1개만 띄워 재사용(WinError 6 방지)
    ac_session: Optional[AutoCompleteSession] = None
    if SELENIUM_AVAILABLE:
        try:
            # UnitTest처럼 headless=False로 설정 (GUI 모드에서 자동완성이 더 안정적)
            ac_session = AutoCompleteSession(headless=False, logger=logger)
            logger.info("자동완성: Chrome 세션 생성 완료(GUI 모드, 재사용 모드)")
        except Exception as e:
            ac_session = None
            logger.warning(f"자동완성: Chrome 세션 생성 실패 -> 자동완성 없이 진행: {repr(e)}")

    try:
        for seed in seeds:
            logger.info("=" * 80)
            logger.info(f"[SEED] {seed}")

            # 1) 자동완성 후보(확장용)
            ac_suggestions: List[str] = []
            seed_ac_delay: Optional[float] = None
            if ac_session is not None:
                ac_res = ac_session.fetch(seed)
                ac_suggestions = ac_res.suggestions
                seed_ac_delay = ac_res.first_suggestion_delay
                logger.debug(f"[SEED][AUTO] seed='{seed}' suggestions={len(ac_suggestions)} delay={seed_ac_delay}")

            # 2) 우측 연관 후보(확장용)
            right_related: List[str] = []
            try:
                rr = get_right_related(seed)
                right_related = rr.related
                logger.debug(f"[SEED][RIGHT] seed='{seed}' related={len(right_related)}")
            except Exception as e:
                logger.warning(f"우측 연관 수집 실패(seed='{seed}'): {repr(e)}")

            # expansions sheet 기록
            for k in ac_suggestions:
                expansion_rows.append({"seed": seed, "type": "ac", "related_keyword": k})
            for k in right_related:
                expansion_rows.append({"seed": seed, "type": "right", "related_keyword": k})

            # 후보 결합
            candidate_set = dedupe_keep_order(ac_suggestions + right_related)
            if not candidate_set:
                logger.info("확장 후보가 없습니다. 다음 seed로")
                continue

            # 폭주 방지
            candidates = candidate_set[:MAX_CANDIDATES_TO_EVALUATE_PER_SEED]
            if len(candidate_set) > len(candidates):
                logger.info(f"후보 {len(candidate_set)}개 중 상위 {len(candidates)}개만 평가합니다(폭주 방지).")

            # 각 후보 평가
            for kw in candidates:
                src = "both" if (kw in ac_suggestions and kw in right_related) else ("ac" if kw in ac_suggestions else "right")

                # (A) 자동완성 속도(후보별 측정). 실패하면 seed delay로 폴백
                ac_delay = None
                ac_speed = "none"
                if ac_session is not None:
                    ac_kw = ac_session.fetch(kw)
                    ac_delay = ac_kw.first_suggestion_delay
                    ac_speed = classify_autocomplete_speed(ac_delay)
                else:
                    ac_delay = seed_ac_delay
                    ac_speed = classify_autocomplete_speed(seed_ac_delay)

                # (B) 우측 연관 개수(후보별)
                rr_count = 0
                try:
                    rr_kw = get_right_related(kw)
                    rr_count = len(rr_kw.related)
                except Exception as e:
                    rr_count = 0
                    logger.debug(f"[RIGHT] 실패 keyword='{kw}': {repr(e)}")

                # (C) 블로그 최근 30일 발행(100+ 캡)
                blog_count = 0
                blog_over = False
                try:
                    blog = count_blog_last30_capped(kw, naver_client_id, naver_client_secret)
                    blog_count = blog.last30_count
                    blog_over = blog.over_100
                except Exception as e:
                    blog_count = 0
                    blog_over = False
                    logger.debug(f"[BLOG] 실패 keyword='{kw}': {repr(e)}")

                blog_display = "100+" if blog_over else str(blog_count)

                # (D) 검색결과 총량(옵션)
                total_results = None
                try:
                    total_results = get_search_total(kw).total_results
                except Exception as e:
                    total_results = None
                    logger.debug(f"[TOTAL] 실패 keyword='{kw}': {repr(e)}")

                # (E) 광고 키워드도구
                try:
                    ad = get_searchad_metrics_for_keyword(kw, ad_api_key, ad_secret_key, ad_customer_id)
                except Exception as e:
                    logger.debug(f"[SEARCHAD] 실패 keyword='{kw}': {repr(e)}")
                    ad = SearchAdMetrics(keyword=kw, exists_in_tool=False, monthly_pc=None, monthly_mobile=None, monthly_total=None, comp_idx=None, ad_depth=None)

                # 점수
                score = score_keyword(
                    autocomplete_speed=ac_speed,
                    right_related_count=rr_count,
                    blog_last30_count=blog_count,
                    blog_over_100=blog_over,
                    total_results=total_results,
                )

                all_rows.append({
                    "seed": seed,
                    "keyword": kw,
                    "source": src,
                    "score": score,
                    "ac_delay_sec": (round(ac_delay, 3) if isinstance(ac_delay, (int, float)) else ""),
                    "ac_speed": ac_speed,
                    "right_related_count": rr_count,
                    "blog_last30_display": blog_display,
                    "blog_last30_count": blog_count,
                    "blog_over_100": blog_over,
                    "search_total_results": total_results if total_results is not None else "",
                    "adtool_exists": ad.exists_in_tool,
                    "monthly_pc": ad.monthly_pc if ad.monthly_pc is not None else "",
                    "monthly_mobile": ad.monthly_mobile if ad.monthly_mobile is not None else "",
                    "monthly_total": ad.monthly_total if ad.monthly_total is not None else "",
                    "comp_idx": ad.comp_idx if ad.comp_idx is not None else "",
                    "ad_depth": ad.ad_depth if ad.ad_depth is not None else "",
                })

            logger.info(f"seed '{seed}' 평가 완료: {len(candidates)}개")

    finally:
        # 드라이버 종료는 여기서 1회만
        if ac_session is not None:
            logger.info("자동완성: Chrome 세션 종료")
            ac_session.close()
            ac_session = None

    # score desc 정렬
    all_rows.sort(key=lambda r: (r.get("score", 0), r.get("monthly_total", 0) if isinstance(r.get("monthly_total"), int) else 0), reverse=True)

    write_excel(out_xlsx, all_rows, expansion_rows)
    logger.info("완료!")
    logger.info(f"엑셀 저장: {out_xlsx}")
    logger.info(f"로그 저장: {log_path}")


if __name__ == "__main__":
    main()
