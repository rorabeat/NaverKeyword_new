import os
import re
import sys
import time
import requests
from datetime import datetime, timedelta, timezone
from typing import Optional, Tuple, List, Dict, Any

from dotenv import load_dotenv
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment

NAVER_BLOG_URL = "https://openapi.naver.com/v1/search/blog.json"

# 순차적으로 호출하되, API 요청 사이에 짧게 대기해 429(Too Many Requests)를 방지한다.
REQUEST_INTERVAL_SEC = 0.15


def get_last_processed_row(excel_path: str) -> int:
    """recent30days 시트의 마지막 행 번호를 반환 (헤더 제외)"""
    try:
        wb = load_workbook(excel_path)
        if "recent30days" not in wb.sheetnames:
            return 0

        ws_recent = wb["recent30days"]
        # 헤더가 1행이므로 데이터 행 수 = max_row - 1
        return max(0, ws_recent.max_row - 1)
    except FileNotFoundError:
        return 0
    except Exception as e:
        print(f"Warning: recent30days 시트 확인 중 오류 발생: {e}")
        return 0


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
    """최근 30일 발행 글 수를 세되, 첫 번째 페이지 데이터만으로 계산 (API 호출 1회만)

    - 첫 번째 API 호출 결과만 사용하여 속도 최적화
    - 30일 이전 글이 나오면 중단
    - limit 도달하면 즉시 중단하고 over_limit=True

    return: (count, over_limit, cutoff_yyyymmdd)
    """
    """최근 30일 발행 글 수를 세되,
    - 30일 이전 글이 나오면 중단
    - limit(기본 100) 도달하면 즉시 중단하고 over_100=True

    return: (count, over_100, cutoff_yyyymmdd)
    """

    kst = timezone(timedelta(hours=9))
    cutoff = (datetime.now(kst) - timedelta(days=days)).date()
    cutoff_yyyymmdd = int(cutoff.strftime("%Y%m%d"))

    count = 0

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

    # 첫 페이지 데이터만 처리 (API 호출 1회로 제한)
    if first_data is not None:
        items = first_data.get("items", []) or []
        stop_by_date, reached_limit = consume_items(items)
        if reached_limit:
            return limit, True, cutoff_yyyymmdd
        if stop_by_date:
            return count, False, cutoff_yyyymmdd

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


def _fetch_blog_count(args: Tuple[str, str, str]) -> Tuple[int, int, Optional[Exception]]:
    """키워드 1개의 (blog_count, api_total, error) 를 반환.
    호출 순서대로 순차 처리되며, 429(Too Many Requests)가 나오면 짧게 대기 후 재시도한다.
    """
    keyword, client_id, client_secret = args
    max_retries = 4
    backoff = 0.5

    for attempt in range(max_retries):
        time.sleep(REQUEST_INTERVAL_SEC)
        try:
            data1 = fetch_naver_blog_json(
                query=keyword, client_id=client_id, client_secret=client_secret,
                display=100, start=1, sort="date",
            )
            api_total = data1.get("total", 0)
            count_30d, _over_100, _cutoff = recent_30d_count_capped(
                query=keyword, client_id=client_id, client_secret=client_secret,
                first_data=data1, limit=100, days=30, sort="date",
            )
            return count_30d, api_total, None
        except requests.HTTPError as e:
            status = e.response.status_code if e.response is not None else None
            if status == 429 and attempt < max_retries - 1:
                time.sleep(backoff)
                backoff *= 2
                continue
            return 0, 0, e
        except Exception as e:
            return 0, 0, e

    return 0, 0, RuntimeError(f"'{keyword}' 429 재시도 {max_retries}회 초과")


def main():
    # 이모지가 포함된 print()가 cp949 등 비-UTF-8 콘솔에서 UnicodeEncodeError로
    # 전체 실행을 죽이는 걸 막기 위해, 인코딩 불가능한 문자는 예외 대신 대체 문자로 바꾼다.
    # (GUI에서는 stdout이 Tk 위젯으로 리다이렉트되어 원래 문제 없지만, 콘솔에서 단독
    # 실행할 때를 위한 방어 코드)
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass

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

    # 기존 진행 상황 확인
    last_processed_row = get_last_processed_row(excel_path)
    start_row = last_processed_row + 2  # 헤더(1) + 데이터행 + 1

    # 진행 상황 출력 (GUI 환경에서는 팝업 대신 콘솔 출력)
    print(f"recent30days 시트에 이미 {last_processed_row}개 키워드가 처리되었습니다.")
    print(f"removeDuplicate 시트의 {start_row}행부터 검색을 시작합니다.")
    print("자동으로 진행합니다...")

    # 데이터 읽기 (지정된 행부터 시작)
    data_rows = []
    for row_idx in range(start_row, ws_remove.max_row + 1):  # 지정된 행부터 시작
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
    print(f"(최적화 적용: API 1회 호출, limit=100, 순차 처리, 100개마다 자동 저장)")

    # 엑셀 저장 함수 (중간 저장용)
    def save_results_to_excel(results_to_save, is_final=False):
        """결과를 엑셀 파일에 저장 (중간 저장 또는 최종 저장)"""
        try:
            # 기존 파일 로드 (없으면 새로 생성)
            try:
                wb = load_workbook(excel_path)
            except FileNotFoundError:
                wb = Workbook()
                # 기본 시트 제거
                wb.remove(wb.active)

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
            for result in results_to_save:
                ws_recent.append([
                    result["seed_keyword"],
                    result["rel_keyword"],
                    result["monthly_total"],
                    result["recent30dayblog"],
                ])

            # 컬럼 너비 자동 조정
            for col_idx, col_name in enumerate(headers, start=1):
                max_len = len(col_name)
                for row_idx in range(2, len(results_to_save) + 2):
                    v = ws_recent.cell(row=row_idx, column=col_idx).value
                    if v is None:
                        continue
                    max_len = max(max_len, len(str(v)))
                ws_recent.column_dimensions[chr(64 + col_idx)].width = min(max_len + 2, 60)

            # 필터 및 고정
            ws_recent.freeze_panes = "A2"
            ws_recent.auto_filter.ref = f"A1:{chr(ord('A') + len(headers) - 1)}{len(results_to_save) + 1}"

            # 파일 저장
            wb.save(excel_path)

            if is_final:
                total_time = time.time() - start_time
                print(f"\n🎉 최종 완료! '{excel_path}' 파일의 'recent30days' 시트에 {len(results_to_save)}개 결과를 저장했습니다.")
                print(f"📊 총 소요시간: {total_time/60:.1f}분")
                print(f"⚡ 최적화 적용: API 1회 호출, limit=100, 순차 처리, 100개마다 자동 저장")
            else:
                print(f"💾 중간 저장 완료: {len(results_to_save)}개 결과 저장됨")

        except PermissionError:
            print(f"⚠️  경고: '{excel_path}' 파일을 저장할 수 없습니다. 파일이 열려있는지 확인해주세요.")
        except Exception as e:
            print(f"⚠️  저장 오류: {e}")

    # 각 키워드에 대해 블로그 발행수 계산 (순차 처리)
    results = []
    start_time = time.time()

    try:
        for i, row in enumerate(data_rows, 1):
            keyword = row["rel_keyword"]
            blog_count, api_total, err = _fetch_blog_count((keyword, client_id, client_secret))

            # 100개마다 진행 상황 및 예상 시간 표시
            if i % 100 == 1 or i == len(data_rows):
                elapsed = time.time() - start_time
                avg_time_per_item = elapsed / i
                remaining_items = len(data_rows) - i
                estimated_remaining = remaining_items * avg_time_per_item

                print(f"[{i}/{len(data_rows)}] '{keyword}' 처리 중... "
                      f"(예상 남은 시간: {estimated_remaining/60:.1f}분)")

            if err is None:
                print(f"  └─ '{keyword}' 완료 (API총={api_total}, 30일내={blog_count})")
            else:
                print(f"  └─ '{keyword}' 실패 ({err})")

            results.append({
                "seed_keyword": row["seed_keyword"],
                "rel_keyword": row["rel_keyword"],
                "monthly_total": row["monthly_total"],
                "recent30dayblog": blog_count,
            })

            # 100개마다 엑셀 파일에 중간 저장 (프로그램 중단 시 데이터 보존)
            if i % 100 == 0 or i == len(data_rows):
                save_results_to_excel(results, is_final=(i == len(data_rows)))

    except KeyboardInterrupt:
        print(f"\n⚠️  사용자 요청으로 프로그램을 중단합니다.")
        if results:
            print(f"💾 지금까지 처리된 {len(results)}개 결과를 저장합니다...")
            save_results_to_excel(results, is_final=True)
        raise

    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")
        if results:
            print(f"💾 오류 발생 전까지 처리된 {len(results)}개 결과를 저장합니다...")
            save_results_to_excel(results, is_final=True)
        raise



if __name__ == "__main__":
    main()