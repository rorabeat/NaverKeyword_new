import os
import time
import hmac
import hashlib
import base64
import re
from datetime import datetime, timezone, timedelta

import requests
from dotenv import load_dotenv

BASE_URL = "https://api.searchad.naver.com"
ENDPOINT = "/keywordstool"  # RelKwdStat list

def make_signature(secret_key: str, timestamp: str, method: str, uri: str) -> str:
    """
    signature = Base64( HMAC-SHA256(secret_key, f"{timestamp}.{method}.{uri}") )
    """
    message = f"{timestamp}.{method}.{uri}"
    digest = hmac.new(secret_key.encode("utf-8"), message.encode("utf-8"), hashlib.sha256).digest()
    return base64.b64encode(digest).decode("utf-8")

def build_headers(api_key: str, secret_key: str, customer_id: str, method: str, uri: str) -> dict:
    timestamp = str(int(time.time() * 1000))
    signature = make_signature(secret_key, timestamp, method, uri)

    return {
        "X-Timestamp": timestamp,
        "X-API-KEY": api_key,
        "X-Customer": customer_id,
        "X-Signature": signature,
        "Content-Type": "application/json; charset=UTF-8",
    }

def call_keywordstool(hint_keywords: str, show_detail: int = 1) -> dict:
    load_dotenv()
    api_key = os.getenv("NAVER_SEARCH_ACCESS_LICENSE_KEY", "").strip()
    secret_key = os.getenv("NAVER_SEARCH_SECRET_KEY", "").strip()
    customer_id = os.getenv("NAVER_SEARCH_CUSTOMER_ID", "").strip()

    if not api_key or not secret_key or not customer_id:
        raise RuntimeError("ENV 누락: .env에 NAVER_SEARCH_ACCESS_LICENSE_KEY / SECRET_KEY / CUSTOMER_ID 를 설정하세요.")

    method = "GET"
    headers = build_headers(api_key, secret_key, customer_id, method, ENDPOINT)

    # ⚠️ hintKeywords는 콤마 구분 최대 5개
    params = {
        "hintKeywords": hint_keywords,
        "showDetail": str(show_detail),
    }

    url = BASE_URL + ENDPOINT
    r = requests.get(url, headers=headers, params=params, timeout=15)
    r.raise_for_status()
    return r.json()

def safe_filename(text: str, max_len: int = 40) -> str:
    text = re.sub(r"[^0-9a-zA-Z가-힣]+", "_", text).strip("_")
    return text[:max_len] if text else "keywords"

def format_result_as_txt(seed: str, data: dict) -> str:
    # 응답은 보통 {"keywordList":[...]} 형태로 오는 경우가 많음(계정/버전에 따라 다를 수 있어 방어적으로 처리)
    keyword_list = data.get("keywordList")
    if keyword_list is None and isinstance(data, list):
        keyword_list = data
    if keyword_list is None:
        keyword_list = []

    lines = []
    kst = timezone(timedelta(hours=9))
    lines.append(f"실행시각(KST): {datetime.now(kst).strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"입력 키워드(hintKeywords): {seed}")
    lines.append(f"연관 키워드 개수: {len(keyword_list)}")
    lines.append("=" * 90)
    lines.append("relKeyword\tmonthlyPcQcCnt\tmonthlyMobileQcCnt\tcompIdx\tmonthlyAvePcClkCnt\tmonthlyAveMobileClkCnt\tmonthlyAvePcCtr\tmonthlyAveMobileCtr\tplAvgDepth")

    for row in keyword_list:
        lines.append(
            f"{row.get('relKeyword','')}\t"
            f"{row.get('monthlyPcQcCnt','')}\t"
            f"{row.get('monthlyMobileQcCnt','')}\t"
            f"{row.get('compIdx','')}\t"
            f"{row.get('monthlyAvePcClkCnt','')}\t"
            f"{row.get('monthlyAveMobileClkCnt','')}\t"
            f"{row.get('monthlyAvePcCtr','')}\t"
            f"{row.get('monthlyAveMobileCtr','')}\t"
            f"{row.get('plAvgDepth','')}"
        )

    lines.append("")
    return "\n".join(lines)

def main():
    # 연관 키워드 입력하면 출려해준다. 
    print("연관 키워드 + 최근 30일 검색량(PC/모바일) 조회 프로그램 (/keywordstool)")
    print("키워드를 입력하세요. 여러 개면 콤마(,)로 구분 (최대 5개).")
    seed = input("hintKeywords > ").strip()

    if not seed:
        print("입력이 비었습니다. 종료.")
        return

    # (실무 팁) 띄어쓰기 포함 키워드에서 파라미터 오류가 난다는 이슈가 보고된 적이 있어
    # 필요하면 공백을 제거/치환하는 전처리를 할 수 있어. :contentReference[oaicite:2]{index=2}
    # seed = seed.replace(" ", "")

    try:
        data = call_keywordstool(seed, show_detail=1)
        report = format_result_as_txt(seed, data)

        kst = timezone(timedelta(hours=9))
        ts = datetime.now(kst).strftime("%Y%m%d_%H%M%S")
        out_path = f"related_keywords_{safe_filename(seed)}_{ts}.txt"

        with open(out_path, "w", encoding="utf-8") as f:
            f.write(report)

        print(report)
        print(f"저장 완료: {out_path}")

    except requests.HTTPError as e:
        print("HTTPError:", e)
        try:
            print("응답 본문(앞 1000자):", e.response.text[:1000])
        except Exception:
            pass
    except Exception as e:
        print("에러:", repr(e))

if __name__ == "__main__":
    main()
