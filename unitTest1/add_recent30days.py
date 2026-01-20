import os
import re
import requests
from datetime import datetime, timedelta, timezone
from typing import Optional, Tuple, List, Dict, Any

from dotenv import load_dotenv
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment

NAVER_BLOG_URL = "https://openapi.naver.com/v1/search/blog.json"


def fetch_naver_blog_json(
    query: str,
    client_id: str,
    client_secret: str,
    display: int = 100,
    start: int = 1,
    sort: str = "sim",
) -> Dict[str, Any]:
    headers = {
        "X-Naver-Client-Id": client_id,
        "X-Naver-Client-Secret": client_secret,
    }
    params = {"query": query, "display": display, "start": start, "sort": sort}

    r = requests.get(NAVER_BLOG_URL, headers=headers, params=params, timeout=10)
    r.raise_for_status()
    return r.json()


def strip_html(text: str) -> str:
    if not text:
        return ""
    # title에 <b> 같은 태그가 포함될 수 있어 간단 제거
    return re.sub(r"<[^>]+>", "", text)


def recent_30d_count_capped(
    query: str,
    client_id: str,
    client_secret: str,
    first_data: Optional[Dict[str, Any]] = None,
    limit: int = 100,
    days: int = 30,
    sort: str = "date",
) -> Tuple[int, bool, int]:
    """최근 30일 발행 글 수를 세되,
    - 30일 이전 글이 나오면 중단
    - limit(기본 100) 도달하면 즉시 중단하고 over_100=True

    return: (count, over_100, cutoff_yyyymmdd)
    """

    kst = timezone(timedelta(hours=9))
    cutoff = (datetime.now(kst) - timedelta(days=days)).date()
    cutoff_yyyymmdd = int(cutoff.strftime("%Y%m%d"))

    count = 0
    start = 1
    display = 100

    def consume_items(items: List[Dict[str, Any]]) -> Tuple[bool, bool]:
        """return (stop_by_date, reached_limit)"""
        nonlocal count

        for item in items:
            postdate_str = item.get("postdate")  # "YYYYMMDD"
            if not postdate_str or len(str(postdate_str)) != 8:
                continue

            if int(postdate_str) < cutoff_yyyymmdd:
                return True, False

            count += 1
            if count >= limit:
                return True, True

        return False, False

    # 1) 첫 페이지 데이터가 이미 있으면 그걸 먼저 소비
    if first_data is not None:
        items = first_data.get("items", []) or []
        stop_by_date, reached_limit = consume_items(items)
        if reached_limit:
            return limit, True, cutoff_yyyymmdd
        if stop_by_date:
            return count, False, cutoff_yyyymmdd
        start = 101  # 첫 페이지를 처리했으니 다음 페이지

    # 2) 필요하면 다음 페이지들 조회
    while True:
        data = fetch_naver_blog_json(
            query=query,
            client_id=client_id,
            client_secret=client_secret,
            display=display,
            start=start,
            sort=sort,
        )
        items = data.get("items", []) or []
        if not items:
            break

        stop_by_date, reached_limit = consume_items(items)
        if reached_limit:
            return limit, True, cutoff_yyyymmdd
        if stop_by_date:
            break

        start += display
        if start > 1000:  # 안전장치(원하면 조정/삭제)
            break

    return count, False, cutoff_yyyymmdd


def get_blog_count_for_keyword(
    keyword: str,
    client_id: str,
    client_secret: str,
    sort: str = "date",
) -> int:
    """키워드에 대한 최근 30일 블로그 발행수를 반환"""
    try:
        count_30d, over_100, _ = recent_30d_count_capped(
            query=keyword,
            client_id=client_id,
            client_secret=client_secret,
            limit=100,
            days=30,
            sort=sort,
        )
        return count_30d
    except Exception as e:
        print(f"Error getting blog count for '{keyword}': {e}")
        return 0


def main():
    load_dotenv()

    client_id = os.getenv("NAVER_CLIENT_ID", "").strip()
    client_secret = os.getenv("NAVER_CLIENT_SECRET", "").strip()

    if not client_id or not client_secret:
        print("ERROR: .env 또는 환경변수에 NAVER_CLIENT_ID / NAVER_CLIENT_SECRET 이 없습니다.")
        print("프로젝트 폴더에 .env 파일을 만들고 다음을 넣으세요:")
        print("NAVER_CLIENT_ID=...")
        print("NAVER_CLIENT_SECRET=...")
        return

    # Excel 파일 읽기
    excel_path = "result/keywordList_all.xlsx"
    try:
        wb = load_workbook(excel_path)
    except FileNotFoundError:
        print(f"ERROR: {excel_path} 파일을 찾을 수 없습니다.")
        return

    if "removeDuplicate" not in wb.sheetnames:
        print("ERROR: removeDuplicate 시트를 찾을 수 없습니다.")
        return

    ws_remove = wb["removeDuplicate"]

    # 데이터 읽기 (헤더 제외)
    data_rows = []
    for row_idx in range(2, ws_remove.max_row + 1):  # 1행은 헤더
        seed_keyword = ws_remove.cell(row=row_idx, column=1).value
        rel_keyword = ws_remove.cell(row=row_idx, column=2).value
        monthly_total = ws_remove.cell(row=row_idx, column=3).value

        if rel_keyword:  # relKeyword가 있는 행만 처리
            data_rows.append({
                "seed_keyword": seed_keyword or "",
                "rel_keyword": rel_keyword,
                "monthly_total": monthly_total or 0,
            })

    print(f"총 {len(data_rows)}개의 키워드를 처리합니다.")

    # 전체 데이터 처리 (속도가 느리면 위의 테스트 코드로 제한 가능)
    print(f"전체 {len(data_rows)}개의 키워드를 처리합니다.")

    # 각 키워드에 대해 블로그 발행수 계산
    results = []
    for i, row in enumerate(data_rows, 1):
        keyword = row["rel_keyword"]
        print(f"[{i}/{len(data_rows)}] '{keyword}' 처리 중...")

        # API 호출 결과도 함께 표시
        try:
            data1 = fetch_naver_blog_json(
                query=keyword,
                client_id=client_id,
                client_secret=client_secret,
                display=100,
                start=1,
                sort="date",
            )
            api_total = data1.get("total", 0)

            # 최근 30일 카운트 계산 (limit을 높게 설정해서 실제 개수 확인)
            count_30d, over_100, cutoff_yyyymmdd = recent_30d_count_capped(
                query=keyword,
                client_id=client_id,
                client_secret=client_secret,
                first_data=data1,
                limit=1000,  # 더 높은 limit으로 실제 개수 계산
                days=30,
                sort="date",
            )

            blog_count = count_30d if not over_100 else f"{count_30d}+"

            print(f"[{i}/{len(data_rows)}] '{keyword}': API총={api_total}, 30일내={blog_count}")

        except Exception as e:
            print(f"[{i}/{len(data_rows)}] '{keyword}': Error - {e}")
            blog_count = 0

        results.append({
            "seed_keyword": row["seed_keyword"],
            "rel_keyword": row["rel_keyword"],
            "monthly_total": row["monthly_total"],
            "recent30dayblog": blog_count,
        })

    # recent30days 시트 생성/업데이트
    if "recent30days" in wb.sheetnames:
        wb.remove(wb["recent30days"])

    ws_recent = wb.create_sheet("recent30days")

    # 헤더 작성
    headers = ["seed_keyword", "relKeyword", "monthlyTotal", "recent30dayblog"]
    ws_recent.append(headers)

    # 헤더 스타일 적용
    header_font = Font(bold=True)
    for cell in ws_recent[1]:
        cell.font = header_font
        cell.alignment = Alignment(vertical="center")

    # 데이터 작성
    for result in results:
        ws_recent.append([
            result["seed_keyword"],
            result["rel_keyword"],
            result["monthly_total"],
            result["recent30dayblog"],
        ])

    # 컬럼 너비 자동 조정
    for col_idx, col_name in enumerate(headers, start=1):
        max_len = len(col_name)
        for row_idx in range(2, len(results) + 2):
            v = ws_recent.cell(row=row_idx, column=col_idx).value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))
        ws_recent.column_dimensions[chr(64 + col_idx)].width = min(max_len + 2, 60)

    # 필터 및 고정
    ws_recent.freeze_panes = "A2"
    ws_recent.auto_filter.ref = f"A1:{chr(ord('A') + len(headers) - 1)}{len(results) + 1}"

    # 파일 저장
    try:
        wb.save(excel_path)
        print(f"\n완료! '{excel_path}' 파일의 'recent30days' 시트에 결과를 저장했습니다.")
        print(f"총 {len(results)}개의 키워드 처리 완료.")
    except PermissionError:
        print(f"\n경고: '{excel_path}' 파일을 저장할 수 없습니다. 파일이 열려있는지 확인해주세요.")
        print("결과를 텍스트 파일로 저장합니다...")

        # 텍스트 파일로 결과 저장
        txt_path = excel_path.replace('.xlsx', '_recent30days.txt')
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write("seed_keyword\trelKeyword\tmonthlyTotal\trecent30dayblog\n")
            for result in results:
                f.write(f"{result['seed_keyword']}\t{result['rel_keyword']}\t{result['monthly_total']}\t{result['recent30dayblog']}\n")
        print(f"결과를 '{txt_path}' 파일로 저장했습니다.")


if __name__ == "__main__":
    main()