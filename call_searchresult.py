# -*- coding: utf-8 -*-
"""네이버 광고 API를 이용한 검색결과 수집

keywordList_all.xlsx 파일의 sumKeyword_final 시트에서 키워드들을 읽어와서
네이버 광고 API(/keywordstool)를 호출하여 연관 검색어와 검색량 정보를 수집합니다.
결과를 searchResult 시트에 저장합니다.
"""

import os
import time
import hmac
import hashlib
import base64
import re
from datetime import datetime, timezone, timedelta

import requests
from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

# 네이버 광고 API 설정
BASE_URL = "https://api.searchad.naver.com"
ENDPOINT = "/keywordstool"

def make_signature(secret_key: str, timestamp: str, method: str, uri: str) -> str:
    """API 서명 생성"""
    message = f"{timestamp}.{method}.{uri}"
    digest = hmac.new(secret_key.encode("utf-8"), message.encode("utf-8"), hashlib.sha256).digest()
    return base64.b64encode(digest).decode("utf-8")

def build_headers(api_key: str, secret_key: str, customer_id: str, method: str, uri: str) -> dict:
    """API 요청 헤더 생성"""
    timestamp = str(int(time.time() * 1000))
    signature = make_signature(secret_key, timestamp, method, uri)

    return {
        "X-Timestamp": timestamp,
        "X-API-KEY": api_key,
        "X-Customer": customer_id,
        "X-Signature": signature,
        "Content-Type": "application/json; charset=UTF-8",
    }

def call_keywordstool_api(hint_keywords: str, show_detail: int = 1) -> dict:
    """네이버 광고 API 호출"""
    load_dotenv()
    api_key = os.getenv("NAVER_SEARCH_ACCESS_LICENSE_KEY", "").strip()
    secret_key = os.getenv("NAVER_SEARCH_SECRET_KEY", "").strip()
    customer_id = os.getenv("NAVER_SEARCH_CUSTOMER_ID", "").strip()

    if not api_key or not secret_key or not customer_id:
        raise RuntimeError("ENV 누락: .env에 NAVER_SEARCH_ACCESS_LICENSE_KEY / SECRET_KEY / CUSTOMER_ID 를 설정하세요.")

    method = "GET"
    headers = build_headers(api_key, secret_key, customer_id, method, ENDPOINT)

    params = {
        "hintKeywords": hint_keywords,
        "showDetail": str(show_detail),
    }

    url = BASE_URL + ENDPOINT
    r = requests.get(url, headers=headers, params=params, timeout=15)
    r.raise_for_status()
    return r.json()

def collect_search_results():
    """sumKeyword_final 시트의 키워드들에 대해 검색결과 수집"""

    filename = "result/keywordList_all.xlsx"

    if not os.path.exists(filename):
        print(f"[ERROR] 파일이 존재하지 않습니다: {filename}")
        return

    try:
        # 기존 파일 로드
        wb = load_workbook(filename)
        print(f"[INFO] 파일 로드 완료: {filename}")

        # sumKeyword_final 시트 확인
        if "sumKeyword_final" not in wb.sheetnames:
            print("[ERROR] sumKeyword_final 시트가 존재하지 않습니다. 먼저 키워드 필터링을 실행해주세요.")
            return

        ws_final = wb["sumKeyword_final"]
        print(f"[INFO] sumKeyword_final 시트에서 키워드 읽기 중... ({ws_final.max_row}행)")

        # 키워드 수집 (헤더 제외, B열의 키워드만)
        keywords = []
        for row in range(2, ws_final.max_row + 1):
            keyword = ws_final.cell(row=row, column=2).value  # B열
            if keyword and str(keyword).strip():
                keywords.append(str(keyword).strip())

        if not keywords:
            print("[ERROR] sumKeyword_final 시트에 처리할 키워드가 없습니다.")
            return

        print(f"[INFO] 총 {len(keywords)}개 키워드 발견")

        # searchResult 시트 생성
        if "searchResult" in wb.sheetnames:
            wb.remove(wb["searchResult"])
            print("[INFO] 기존 searchResult 시트 제거")

        ws_result = wb.create_sheet("searchResult")
        print("[INFO] searchResult 시트 생성")

        # 헤더 추가
        headers = ["seed_keyword", "relKeyword", "monthlyPcQcCnt", "monthlyMobileQcCnt",
                  "monthlyTotal", "compIdx", "plAvgDepth", "monthlyAvePcClkCnt", "monthlyAveMobileClkCnt"]
        ws_result.append(headers)

        # 헤더 스타일링
        header_font = Font(bold=True)
        for cell in ws_result[1]:
            cell.font = header_font
            cell.alignment = Alignment(vertical="center")

        # 열 너비 설정
        ws_result.column_dimensions["A"].width = 25
        ws_result.column_dimensions["B"].width = 40
        ws_result.column_dimensions["C"].width = 15
        ws_result.column_dimensions["D"].width = 15
        ws_result.column_dimensions["E"].width = 15
        ws_result.column_dimensions["F"].width = 10
        ws_result.column_dimensions["G"].width = 12
        ws_result.column_dimensions["H"].width = 18
        ws_result.column_dimensions["I"].width = 18

        total_processed = 0
        total_rel_keywords = 0

        # 각 키워드에 대해 API 호출
        for idx, seed_keyword in enumerate(keywords, 1):
            print(f"[INFO] 키워드 {idx}/{len(keywords)} 처리 중: {seed_keyword}")

            try:
                # API 호출 (개별 키워드만 전송)
                data = call_keywordstool_api(seed_keyword, show_detail=1)

                # 응답 데이터 파싱
                keyword_list = data.get("keywordList")
                if keyword_list is None and isinstance(data, list):
                    keyword_list = data
                if keyword_list is None:
                    keyword_list = []

                print(f"[INFO] {seed_keyword}: {len(keyword_list)}개 연관 키워드 발견")

                # 각 연관 키워드의 정보를 행으로 추가
                for rel_data in keyword_list:
                    rel_keyword = rel_data.get("relKeyword", "")
                    pc_cnt = rel_data.get("monthlyPcQcCnt", 0) or 0
                    mobile_cnt = rel_data.get("monthlyMobileQcCnt", 0) or 0
                    total_cnt = pc_cnt + mobile_cnt

                    row_data = [
                        seed_keyword,  # A열: 시드 키워드
                        rel_keyword,   # B열: 연관 키워드
                        pc_cnt,        # C열: PC 검색량
                        mobile_cnt,    # D열: 모바일 검색량
                        total_cnt,     # E열: 총 검색량
                        rel_data.get("compIdx", ""),        # F열: 경쟁지수
                        rel_data.get("plAvgDepth", ""),     # G열: 평균 노출 순위
                        rel_data.get("monthlyAvePcClkCnt", ""),     # H열: PC 평균 클릭수
                        rel_data.get("monthlyAveMobileClkCnt", "")  # I열: 모바일 평균 클릭수
                    ]

                    ws_result.append(row_data)

                total_processed += 1
                total_rel_keywords += len(keyword_list)

                # API 호출 간 딜레이 (과도한 호출 방지)
                if idx < len(keywords):  # 마지막이 아니면
                    time.sleep(0.5)

            except requests.HTTPError as e:
                print(f"[WARN] {seed_keyword} - HTTP 오류: {e}")
                # 실패한 경우에도 빈 행 추가 (나중에 확인 가능하도록)
                ws_result.append([seed_keyword, f"API_ERROR: {e}", "", "", "", "", "", "", ""])
            except Exception as e:
                print(f"[WARN] {seed_keyword} - 처리 오류: {e}")
                ws_result.append([seed_keyword, f"ERROR: {e}", "", "", "", "", "", "", ""])

        # 필터링 설정
        max_row = ws_result.max_row
        ws_result.freeze_panes = "A2"
        ws_result.auto_filter.ref = f"A1:I{max_row}"

        # 파일 저장
        wb.save(filename)

        print("[SUCCESS] 검색결과 수집 완료!")
        print(f"   - 처리된 시드 키워드: {total_processed}/{len(keywords)}개")
        print(f"   - 수집된 연관 키워드: {total_rel_keywords}개")
        print(f"   - 저장 파일: {filename}")
        print(f"   - 새 시트: searchResult")

    except Exception as e:
        print(f"[ERROR] 오류 발생: {e}")


def main():
    """메인 함수"""
    print("=== 네이버 광고 API 검색결과 수집 시작 ===")
    collect_search_results()
    print("=== 네이버 광고 API 검색결과 수집 완료 ===")


if __name__ == "__main__":
    main()