import re
from datetime import datetime, timezone, timedelta
from urllib.parse import quote_plus

import requests
from bs4 import BeautifulSoup


def safe_filename(text: str, max_len: int = 40) -> str:
    text = re.sub(r"[^0-9a-zA-Z가-힣]+", "_", text).strip("_")
    return (text[:max_len] if text else "keyword")


def fetch_naver_search_html(query: str) -> str:
    """
    네이버 통합검색 페이지 HTML 가져오기
    """
    url = f"https://search.naver.com/search.naver?query={quote_plus(query)}"
    headers = {
        # 브라우저처럼 보이게(차단 확률 감소)
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Referer": "https://search.naver.com/",
    }

    r = requests.get(url, headers=headers, timeout=15)
    r.raise_for_status()
    return r.text


def extract_related_keywords_from_html(html: str) -> list[str]:
    """
    네가 준 구조:
    <div class="related_srch">
      ...
      <div class="tit">오키나와 렌트카 비용</div>
    를 기준으로 텍스트를 뽑는다.
    """
    soup = BeautifulSoup(html, "lxml")

    # 1) 가장 정확: related_srch 영역 안의 tit
    nodes = soup.select("div.related_srch div.tit")
    keywords = [n.get_text(strip=True) for n in nodes if n.get_text(strip=True)]

    # 2) 방어 로직: 혹시 클래스가 바뀌면, related_srch 내 keyword 링크 텍스트도 시도
    if not keywords:
        nodes2 = soup.select("div.related_srch a.keyword")
        keywords = [n.get_text(strip=True) for n in nodes2 if n.get_text(strip=True)]

    # 중복 제거(순서 유지)
    deduped = []
    seen = set()
    for k in keywords:
        if k not in seen:
            seen.add(k)
            deduped.append(k)

    return deduped


def save_to_txt(query: str, related: list[str]) -> str:
    kst = timezone(timedelta(hours=9))
    ts = datetime.now(kst).strftime("%Y%m%d_%H%M%S")
    out_path = f"naver_related_{safe_filename(query)}_{ts}.txt"

    lines = []
    lines.append(f"실행시각(KST): {datetime.now(kst).strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"입력 키워드: {query}")
    lines.append(f"연관 검색어 개수: {len(related)}")
    lines.append("=" * 60)

    for i, kw in enumerate(related, start=1):
        lines.append(f"{i:02d}. {kw}")

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    return out_path


def main():
    query = input("키워드 입력 > ").strip()
    if not query:
        print("키워드가 비었습니다. 종료.")
        return

    try:
        html = fetch_naver_search_html(query)
        related = extract_related_keywords_from_html(html)

        if not related:
            print("연관 검색어를 찾지 못했습니다.")
            print("가능한 원인: 페이지 구조 변경 / 봇 차단 / 연관검색어 영역이 없는 쿼리")
            return

        out_path = save_to_txt(query, related)

        print(f"\n연관 검색어 {len(related)}개 추출 완료!")
        for k in related[:10]:
            print("-", k)
        if len(related) > 10:
            print(f"... (나머지 {len(related)-10}개는 txt 파일에 저장됨)")

        print(f"\n저장 완료: {out_path}")

    except requests.HTTPError as e:
        print("HTTPError:", e)
    except Exception as e:
        print("에러:", repr(e))


if __name__ == "__main__":
    main()
