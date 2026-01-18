import requests
import urllib.parse
import os
import json
from datetime import datetime, timedelta
import re

# ==============================
# 🔐 네이버 블로그 검색 API 키 (기존 값 그대로)
# ==============================
NAVER_SEARCH_ACCESS_LICENSE_KEY = "01000000007b181c36c60b62b8d6f2ce4ed3d7849dd050ce65d33cd1ee5b24f9849d835723"
NAVER_SEARCH_SECRET_KEY = "AQAAAAB7GBw2xgtiuNbyzk7T14SdHCiiceB/IQ9vq9vxETJpLA=="
NAVER_SEARCH_CUSTOMER_ID = "4155640"
NAVER_CLIENT_ID = "gWN7F4VTG9IfOP49Ft87"
NAVER_CLIENT_SECRET = "EDSWZKIjxr"


def safe_int(val, default=0) -> int:
    """네이버 키워드툴 값이 '< 10' 같은 문자열로 오는 경우가 있어 안전 변환."""
    if val is None:
        return default

    if isinstance(val, (int, float)):
        try:
            return int(val)
        except Exception:
            return default

    s = str(val).strip()
    if not s:
        return default

    if s.replace(" ", "") in ("<10", "< 10".replace(" ", "")):
        return 0

    s = s.replace(",", "")
    if s.isdigit():
        return int(s)

    m = re.search(r"\d+", s)
    return int(m.group()) if m else default


# ==============================
# 📁 로그/캐시 폴더
# ==============================
os.makedirs("result/logs", exist_ok=True)
os.makedirs("result/cache", exist_ok=True)

log_filename = datetime.now().strftime("result/logs/%Y%m%d_%H%M%S_log.txt")
cache_path = "result/cache/blog_cache.json"


def log_write(msg: str):
    msg = str(msg)
    print(msg)
    try:
        with open(log_filename, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        # 로그 실패해도 메인 로직이 죽지 않게
        pass


def load_cache() -> dict:
    if not os.path.exists(cache_path):
        return {}
    try:
        with open(cache_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_cache(cache: dict):
    tmp = cache_path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)
    os.replace(tmp, cache_path)


BLOG_CACHE = load_cache()


def _cache_key_recent(keyword: str, days: int, end_yyyymmdd: str) -> str:
    # 기간이 바뀌면(날짜가 바뀌면) 발행량도 달라지므로, end 날짜를 키에 포함
    return f"recent_blog_posts::{keyword}::days={days}::end={end_yyyymmdd}"


def get_recent_blog_posts_count_with_cache(keyword: str, days: int = 30,
                                           display: int = 100, max_start: int = 1000) -> int:
    """
    ✅ 최근 N일(기본 30일) 블로그 '발행량' 추정

    네이버 블로그 검색 API는 startDate/endDate 파라미터가 없어서,
    sort=date(최신순)으로 페이지를 넘기면서 items[].postdate(YYYYMMDD)를 보고
    cutoff(오늘- N일) 이상인 글만 카운트합니다.

    제한:
    - start 최대 1000이라서, 최신 1000건까지만 확인 가능 (display=100이면 최대 10페이지)
    """
    end_date = datetime.now().date()
    end_yyyymmdd = end_date.strftime("%Y%m%d")
    cutoff = (end_date - timedelta(days=days)).strftime("%Y%m%d")

    ck = _cache_key_recent(keyword, days, end_yyyymmdd)
    if ck in BLOG_CACHE:
        try:
            return int(BLOG_CACHE[ck])
        except Exception:
            pass

    enc = urllib.parse.quote(keyword)
    headers = {
        "X-Naver-Client-Id": NAVER_CLIENT_ID,
        "X-Naver-Client-Secret": NAVER_CLIENT_SECRET,
    }

    total_count = 0
    start = 1

    while start <= max_start:
        url = f"https://openapi.naver.com/v1/search/blog.json?query={enc}&display={display}&start={start}&sort=date"
        try:
            res = requests.get(url, headers=headers, timeout=10)
            if res.status_code != 200:
                log_write(f"⚠️ [blog] HTTP {res.status_code} keyword={keyword} start={start} body={res.text}")
                break

            data = res.json()
            items = data.get("items", []) or []

            if not items:
                break

            # 최신순이므로, 이 페이지의 가장 오래된 글(postdate)을 보고 cutoff 지나면 중단 가능
            oldest_in_page = None

            for it in items:
                postdate = str(it.get("postdate", "")).strip()
                if not postdate or len(postdate) != 8:
                    continue

                if oldest_in_page is None or postdate < oldest_in_page:
                    oldest_in_page = postdate

                if postdate >= cutoff:
                    total_count += 1

            # 페이지에서 가장 오래된 글이 cutoff보다 과거면, 다음 페이지는 더 과거라서 중단
            if oldest_in_page is not None and oldest_in_page < cutoff:
                break

            start += display

        except Exception as e:
            log_write(f"❌ [blog] exception keyword={keyword} start={start} err={e}")
            break

    BLOG_CACHE[ck] = int(total_count)
    save_cache(BLOG_CACHE)
    return int(total_count)


def get_keyword_stats_from_keywordtool(keywordtool_data, progress_callback=None,
                                       min_total_search: int = 1000,
                                       recent_days: int = 30):
    """
    keywordtool_data: related_keywords.get_related_keywords()가 반환한 dict 리스트
    - keywordstool 재호출 ❌
    - 블로그 API는 검색량 컷 통과한 키워드만 호출 ✅

    결과 컬럼(표/엑셀용):
    - 키워드, PC, Mobile, 검색수합계, 월간 블로그 발행(=최근 N일), 포화도

    참고: '월간'이라는 이름은 UI 호환용이고, 실제로는 recent_days 기간 발행량입니다.
    """
    results = []
    total = len(keywordtool_data)

    log_write(f"\n📊 분석 대상 키워드 수: {total}")
    log_write(f"🗓️ 블로그 발행량 기준: 최근 {recent_days}일 (sort=date + postdate 필터)\n")

    processed = 0
    for idx, kw_data in enumerate(keywordtool_data, start=1):
        keyword = kw_data.get("relKeyword") or kw_data.get("keyword") or ""
        if not keyword:
            continue

        monthly_pc = safe_int(kw_data.get("monthlyPcQcCnt", 0))
        monthly_mobile = safe_int(kw_data.get("monthlyMobileQcQcCnt", kw_data.get("monthlyMobileQcCnt", 0)))
        total_search = monthly_pc + monthly_mobile

        # ✅ 검색량 컷
        if total_search < min_total_search:
            if progress_callback:
                progress_callback(idx, total)
            continue

        recent_blog_posts = get_recent_blog_posts_count_with_cache(keyword, days=recent_days)
        saturation = round((recent_blog_posts / total_search) * 100, 1) if total_search else 0.0

        results.append({
            "키워드": keyword,
            "PC": monthly_pc,
            "Mobile": monthly_mobile,
            "검색수합계": total_search,
            "월간 블로그 발행": recent_blog_posts,
            "포화도": saturation,

            # 기존 호환 컬럼(필요하면 유지)
            "월간총검색량": total_search,
            "문서수": recent_blog_posts,
            "경쟁률": saturation,
        })

        processed += 1
        log_write(
            f"📊 {keyword} | PC:{monthly_pc:,} | M:{monthly_mobile:,} | 검색:{total_search:,} | "
            f"최근{recent_days}일 발행:{recent_blog_posts:,} | 포화도:{saturation}%"
        )

        if progress_callback:
            progress_callback(idx, total)

    log_write(f"\n✅ 전체 {processed}개(컷 통과 기준) 분석 완료.")
    log_write(f"📝 로그 파일: {os.path.abspath(log_filename)}\n")
    return results
