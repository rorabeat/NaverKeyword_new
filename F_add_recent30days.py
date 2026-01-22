import os
import re
import time
import requests
from datetime import datetime, timedelta, timezone
from typing import Optional, Tuple, List, Dict, Any
import tkinter as tk
from tkinter import messagebox

from dotenv import load_dotenv
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment

NAVER_BLOG_URL = "https://openapi.naver.com/v1/search/blog.json"


def show_confirmation_popup(message: str) -> bool:
    """사용자에게 확인/취소 팝업을 표시하고 결과를 반환"""
    root = tk.Tk()
    root.withdraw()  # 메인 윈도우 숨김
    root.attributes("-topmost", True)  # 최상위로 표시

    result = messagebox.askyesno("진행 확인", message)

    root.destroy()
    return result


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

    # 기존 진행 상황 확인
    last_processed_row = get_last_processed_row(excel_path)
    start_row = last_processed_row + 2  # 헤더(1) + 데이터행 + 1

    # 팝업으로 진행 확인
    message = f"recent30days 시트에 이미 {last_processed_row}개 키워드가 처리되었습니다.\n\n"
    message += f"removeDuplicate 시트의 {start_row}행부터 검색을 시작합니다.\n\n"
    message += "계속 진행하시겠습니까?"

    if not show_confirmation_popup(message):
        print("사용자가 취소를 선택했습니다. 프로그램을 종료합니다.")
        return

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
    print("(최적화 적용: API 1회 호출, limit=100, 0.04초 지연, 100개마다 자동 저장)")

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
                print(f"⚡ 최적화 적용: API 1회 호출, limit=100, 0.04초 지연, 100개마다 자동 저장")
            else:
                print(f"💾 중간 저장 완료: {len(results_to_save)}개 결과 저장됨")

        except PermissionError:
            print(f"⚠️  경고: '{excel_path}' 파일을 저장할 수 없습니다. 파일이 열려있는지 확인해주세요.")
        except Exception as e:
            print(f"⚠️  저장 오류: {e}")

    # 각 키워드에 대해 블로그 발행수 계산
    results = []
    start_time = time.time()

    try:
        for i, row in enumerate(data_rows, 1):
            keyword = row["rel_keyword"]

            # 100개마다 진행 상황 및 예상 시간 표시
            if i % 100 == 1 or i == len(data_rows):
                elapsed = time.time() - start_time
                avg_time_per_item = elapsed / i
                remaining_items = len(data_rows) - i
                estimated_remaining = remaining_items * avg_time_per_item

                print(f"[{i}/{len(data_rows)}] '{keyword}' 처리 중... "
                      f"(예상 남은 시간: {estimated_remaining/60:.1f}분)")

            # 각 키워드별 진행 표시 (간단하게)
            print(f"  └─ '{keyword}' 처리 중...", end="", flush=True)

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

                # 최근 30일 카운트 계산 (최적화: limit을 100으로 제한)
                count_30d, over_100, cutoff_yyyymmdd = recent_30d_count_capped(
                    query=keyword,
                    client_id=client_id,
                    client_secret=client_secret,
                    first_data=data1,
                    limit=100,  # API 호출 최적화를 위해 100개로 제한
                    days=30,
                    sort="date",
                )

                blog_count = count_30d if not over_100 else f"{count_30d}+"

                # 진행 표시 업데이트
                print(f" 완료 (API총={api_total}, 30일내={blog_count})")

            except Exception as e:
                print(f" 실패 ({e})")
                blog_count = 0

            results.append({
                "seed_keyword": row["seed_keyword"],
                "rel_keyword": row["rel_keyword"],
                "monthly_total": row["monthly_total"],
                "recent30dayblog": blog_count,
            })

            # 100개마다 엑셀 파일에 중간 저장 (프로그램 중단 시 데이터 보존)
            if i % 100 == 0 or i == len(data_rows):
                save_results_to_excel(results, is_final=(i == len(data_rows)))

            # API 호출 최적화를 위한 지연 시간 (네이버 API 제한 고려) - 5배 속도 향상
            time.sleep(0.04)

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