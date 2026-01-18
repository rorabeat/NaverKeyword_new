import requests
import hmac
import hashlib
import base64
import time
import urllib.parse

# ==============================
# 🔐 네이버 검색광고 API 키 (기존 값 그대로 붙여 넣으세요)
# ==============================
NAVER_SEARCH_ACCESS_LICENSE_KEY = "01000000007b181c36c60b62b8d6f2ce4ed3d7849dd050ce65d33cd1ee5b24f9849d835723"
NAVER_SEARCH_SECRET_KEY = "AQAAAAB7GBw2xgtiuNbyzk7T14SdHCiiceB/IQ9vq9vxETJpLA=="
NAVER_SEARCH_CUSTOMER_ID = "4155640"
NAVER_CLIENT_ID = "gWN7F4VTG9IfOP49Ft87"
NAVER_CLIENT_SECRET = "EDSWZKIjxr"


def generate_signature(secret_key: str, method: str, uri: str, timestamp: str) -> str:
    """
    네이버 검색광고 API signature
    message: {timestamp}.{method}.{uri}
    """
    message = f"{timestamp}.{method}.{uri}".encode("utf-8")
    secret = secret_key.encode("utf-8")
    sign = hmac.new(secret, message, hashlib.sha256).digest()
    return base64.b64encode(sign).decode("utf-8")


def normalize_hint_keywords(keyword: str) -> str:
    """
    keywordstool의 hintKeywords는 공백이 포함되면 400(BAD_REQUEST) Invalid Parameter가 발생할 수 있음.
    따라서 공백(whitespace)을 제거하여 seed로 사용.
    예) "츄라우미 수족관" -> "츄라우미수족관"
    """
    return "".join(keyword.split())


def get_related_keywords(keyword: str, limit: int = 20):
    """
    ✅ 검색광고 keywordstool API를 1회 호출하여
    연관 키워드 + 검색량 포함(dict 리스트)을 반환합니다.

    반환 예:
      [
        {
          "relKeyword": "...",
          "monthlyPcQcCnt": ...,
          "monthlyMobileQcCnt": ...,
          ...
        },
        ...
      ]
    """
    base_url = "https://api.searchad.naver.com"
    uri = "/keywordstool"
    method = "GET"

    seed = normalize_hint_keywords(keyword)
    timestamp = str(int(time.time() * 1000))

    headers = {
        "X-API-KEY": NAVER_SEARCH_ACCESS_LICENSE_KEY,
        "X-Customer": NAVER_SEARCH_CUSTOMER_ID,
        "X-Timestamp": timestamp,
        "X-Signature": generate_signature(NAVER_SEARCH_SECRET_KEY, method, uri, timestamp),
    }

    # requests가 인코딩을 처리하지만, 공백 제거는 필수
    params = {
        "hintKeywords": seed,
        "showDetail": 1,
    }

    res = requests.get(base_url + uri, headers=headers, params=params, timeout=10)

    # 400이면 body에 Invalid Parameter 메시지가 들어오는 경우가 많아 원문을 남김
    if res.status_code != 200:
        raise RuntimeError(
            f"[keywordstool] HTTP {res.status_code} | url={res.url} | body={res.text}"
        )

    data = res.json()
    keyword_list = data.get("keywordList", [])

    # limit 적용(상위 N개만)
    return keyword_list[:limit]
