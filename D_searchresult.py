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
from urllib.parse import quote_plus, unquote_plus

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
    """네이버 광고 API 호출 (RelKwdStat 기능)"""
    load_dotenv()
    api_key = os.getenv("NAVER_SEARCH_ACCESS_LICENSE_KEY", "").strip()
    secret_key = os.getenv("NAVER_SEARCH_SECRET_KEY", "").strip()
    customer_id = os.getenv("NAVER_SEARCH_CUSTOMER_ID", "").strip()

    if not api_key or not secret_key or not customer_id:
        raise RuntimeError("ENV 누락: .env에 NAVER_SEARCH_ACCESS_LICENSE_KEY / SECRET_KEY / CUSTOMER_ID 를 설정하세요.")

    method = "GET"
    headers = build_headers(api_key, secret_key, customer_id, method, ENDPOINT)

    # 공백을 완전히 제거 (네이버 API 호환성 위해)
    processed_keywords = hint_keywords.replace(' ', '')

    params = {
        "hintKeywords": processed_keywords,
        "showDetail": str(show_detail),
    }

    url = BASE_URL + ENDPOINT
    r = requests.get(url, headers=headers, params=params, timeout=15)

    if r.status_code != 200:
        print(f"[DEBUG] HTTP {r.status_code} 응답: {r.text}")
        r.raise_for_status()

    return r.json()


def call_naver_search_api(keyword: str) -> list:
    """네이버 검색 API를 사용한 연관 키워드 검색 (대안)"""
    load_dotenv()
    client_id = os.getenv("NAVER_CLIENT_ID", "").strip()
    client_secret = os.getenv("NAVER_CLIENT_SECRET", "").strip()

    if not client_id or not client_secret:
        print("[DEBUG] 네이버 검색 API 키 없음, 건너뜀")
        return []

    url = "https://openapi.naver.com/v1/search/encyc.json"
    headers = {
        "X-Naver-Client-Id": client_id,
        "X-Naver-Client-Secret": client_secret
    }

    params = {
        "query": keyword,
        "display": 10  # 최대 10개 결과
    }

    try:
        print(f"[DEBUG] 네이버 검색 API 시도: {keyword}")
        r = requests.get(url, headers=headers, params=params, timeout=15)
        r.raise_for_status()
        result = r.json()

        # 백과사전 검색 결과를 연관 키워드로 변환
        related_keywords = []
        if "items" in result and result["items"]:
            for item in result["items"][:5]:  # 최대 5개로 제한
                title = item.get("title", "").replace("<b>", "").replace("</b>", "").replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&amp;", "&")
                related_keywords.append({
                    "relKeyword": title,
                    "monthlyPcQcCnt": "0",  # 검색 API는 검색량 제공하지 않음
                    "monthlyMobileQcCnt": "0",
                    "compIdx": "",
                    "plAvgDepth": "",
                    "monthlyAvePcClkCnt": "",
                    "monthlyAveMobileClkCnt": "",
                    "monthlyAvePcCtr": "",
                    "monthlyAveMobileCtr": ""
                })

        return related_keywords

    except Exception as e:
        print(f"[DEBUG] 네이버 검색 API 실패: {e}")
        return []

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
            if keyword:
                # 엑셀에서 읽어온 키워드를 엄격하게 정제
                cleaned_keyword = str(keyword).strip()
                # 숨겨진 문자나 특수 공백 제거 (제어 문자, 넌브레이킹 스페이스 등)
                cleaned_keyword = re.sub(r'[\x00-\x1f\x7f-\x9f\xa0]', ' ', cleaned_keyword)
                cleaned_keyword = re.sub(r'\s+', ' ', cleaned_keyword).strip()

                if cleaned_keyword:
                    keywords.append(cleaned_keyword)

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

        # 헤더 추가 (RelKwdStat의 모든 필드 포함)
        headers = ["seed_keyword", "relKeyword", "monthlyPcQcCnt", "monthlyMobileQcCnt",
                  "monthlyTotal", "compIdx", "plAvgDepth", "monthlyAvePcClkCnt", "monthlyAveMobileClkCnt",
                  "monthlyAvePcCtr", "monthlyAveMobileCtr"]
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
        ws_result.column_dimensions["J"].width = 12
        ws_result.column_dimensions["K"].width = 12

        total_processed = 0
        total_rel_keywords = 0

        # 각 키워드에 대해 개별적으로 API 호출 (RelKwdStat 방식)
        for idx, seed_keyword in enumerate(keywords, 1):
            # 너무 긴 키워드만 필터링 (인코딩 문제 방지)
            if len(seed_keyword) > 50:  # 더 긴 길이 허용
                print(f"[SKIP] 키워드 길이 제한으로 제외: {seed_keyword} (길이: {len(seed_keyword)})")
                continue

            print(f"[INFO] 키워드 {idx}/{len(keywords)} 처리 중: {seed_keyword}")

            try:
                # API 호출 (RelKwdStat: 단일 키워드 전송, showDetail=1)
                data = call_keywordstool_api(seed_keyword, show_detail=1)

                # 응답 데이터 파싱 (UnitTest 방식으로 단순화)
                keyword_list = data.get("keywordList")
                if keyword_list is None and isinstance(data, list):
                    keyword_list = data
                if keyword_list is None:
                    keyword_list = []

                print(f"[INFO] {seed_keyword}: {len(keyword_list)}개 연관 키워드 발견")

                # 각 연관 키워드의 정보를 행으로 추가
                for rel_data in keyword_list:
                    rel_keyword = unquote_plus(rel_data.get("relKeyword", ""))  # URL 디코딩

                    # 검색량 처리 (문자열 "< 10"을 9로 변환)
                    pc_cnt_raw = rel_data.get("monthlyPcQcCnt", "")
                    mobile_cnt_raw = rel_data.get("monthlyMobileQcCnt", "")

                    # "< 10" 형식의 문자열을 숫자로 변환
                    def parse_count(value):
                        if isinstance(value, str) and "< 10" in value:
                            return 9  # 10 미만은 9로 처리
                        try:
                            return int(value) if value else 0
                        except (ValueError, TypeError):
                            return 0

                    pc_cnt = parse_count(pc_cnt_raw)
                    mobile_cnt = parse_count(mobile_cnt_raw)
                    total_cnt = pc_cnt + mobile_cnt

                    # RelKwdStat 응답에는 CTR 정보도 포함됨
                    pc_ctr = rel_data.get("monthlyAvePcCtr", "")
                    mobile_ctr = rel_data.get("monthlyAveMobileCtr", "")

                    row_data = [
                        seed_keyword,  # A열: 시드 키워드
                        rel_keyword,   # B열: 연관 키워드
                        pc_cnt,        # C열: PC 검색량
                        mobile_cnt,    # D열: 모바일 검색량
                        total_cnt,     # E열: 총 검색량
                        rel_data.get("compIdx", ""),        # F열: 경쟁지수
                        rel_data.get("plAvgDepth", ""),     # G열: 평균 노출 순위
                        rel_data.get("monthlyAvePcClkCnt", ""),     # H열: PC 평균 클릭수
                        rel_data.get("monthlyAveMobileClkCnt", ""),  # I열: 모바일 평균 클릭수
                        pc_ctr,        # J열: PC CTR (추가)
                        mobile_ctr     # K열: 모바일 CTR (추가)
                    ]

                    ws_result.append(row_data)

                total_processed += 1
                total_rel_keywords += len(keyword_list)
                total_rel_keywords += len(keyword_list)

                # API 호출 간 딜레이 (과도한 호출 방지)
                if idx < len(keywords):  # 마지막이 아니면
                    time.sleep(0.5)

            except requests.HTTPError as e:
                error_msg = f"HTTP_{e.response.status_code}_ERROR"
                if hasattr(e.response, 'text') and e.response.text:
                    try:
                        error_detail = e.response.json()
                        if 'error' in error_detail:
                            error_msg = f"{error_msg}: {error_detail['error']}"
                    except:
                        error_msg = f"{error_msg}: {str(e)}"
                else:
                    error_msg = f"{error_msg}: {str(e)}"

                print(f"[WARN] {seed_keyword} - {error_msg}")

                # 네이버 검색 API로 폴백 시도
                print(f"[INFO] {seed_keyword} - 네이버 검색 API로 재시도")
                try:
                    keyword_list = call_naver_search_api(seed_keyword)
                    if keyword_list:
                        print(f"[INFO] {seed_keyword}: 데이터랩 API로 {len(keyword_list)}개 연관 키워드 발견")
                        # 데이터랩 결과를 엑셀에 추가
                        for rel_data in keyword_list:
                            rel_keyword = rel_data.get("relKeyword", "")

                            pc_cnt = rel_data.get("monthlyPcQcCnt", "0")
                            mobile_cnt = rel_data.get("monthlyMobileQcCnt", "0")
                            total_cnt = str(int(pc_cnt) + int(mobile_cnt)) if pc_cnt.isdigit() and mobile_cnt.isdigit() else "0"

                            row_data = [
                                seed_keyword,
                                rel_keyword,
                                pc_cnt,
                                mobile_cnt,
                                total_cnt,
                                rel_data.get("compIdx", ""),
                                rel_data.get("plAvgDepth", ""),
                                rel_data.get("monthlyAvePcClkCnt", ""),
                                rel_data.get("monthlyAveMobileClkCnt", ""),
                                rel_data.get("monthlyAvePcCtr", ""),
                                rel_data.get("monthlyAveMobileCtr", "")
                            ]
                            ws_result.append(row_data)
                        total_processed += 1
                        total_rel_keywords += len(keyword_list)
                        continue  # 성공했으므로 다음 키워드로
                except Exception as fallback_e:
                    print(f"[WARN] {seed_keyword} - 데이터랩 API도 실패: {fallback_e}")

                # 모두 실패한 경우 빈 행 추가
                ws_result.append([seed_keyword, error_msg, "", "", "", "", "", "", "", "", ""])
            except Exception as e:
                error_msg = f"ERROR: {e}"
                print(f"[WARN] {seed_keyword} - {error_msg}")

                # 네이버 검색 API로 폴백 시도 (연결 오류 등의 일반 예외에서도)
                print(f"[INFO] {seed_keyword} - 네이버 검색 API로 재시도")
                try:
                    keyword_list = call_naver_search_api(seed_keyword)
                    if keyword_list:
                        print(f"[INFO] {seed_keyword}: 데이터랩 API로 {len(keyword_list)}개 연관 키워드 발견")
                        # 데이터랩 결과를 엑셀에 추가
                        for rel_data in keyword_list:
                            rel_keyword = rel_data.get("relKeyword", "")

                            pc_cnt = rel_data.get("monthlyPcQcCnt", "0")
                            mobile_cnt = rel_data.get("monthlyMobileQcCnt", "0")
                            total_cnt = str(int(pc_cnt) + int(mobile_cnt)) if pc_cnt.isdigit() and mobile_cnt.isdigit() else "0"

                            row_data = [
                                seed_keyword,
                                rel_keyword,
                                pc_cnt,
                                mobile_cnt,
                                total_cnt,
                                rel_data.get("compIdx", ""),
                                rel_data.get("plAvgDepth", ""),
                                rel_data.get("monthlyAvePcClkCnt", ""),
                                rel_data.get("monthlyAveMobileClkCnt", ""),
                                rel_data.get("monthlyAvePcCtr", ""),
                                rel_data.get("monthlyAveMobileCtr", "")
                            ]
                            ws_result.append(row_data)
                        total_processed += 1
                        total_rel_keywords += len(keyword_list)
                        continue  # 성공했으므로 다음 키워드로
                except Exception as fallback_e:
                    print(f"[WARN] {seed_keyword} - 데이터랩 API도 실패: {fallback_e}")

                # 모두 실패한 경우 빈 행 추가
                ws_result.append([seed_keyword, error_msg, "", "", "", "", "", "", "", "", ""])

        # 필터링 설정
        max_row = ws_result.max_row
        ws_result.freeze_panes = "A2"
        ws_result.auto_filter.ref = f"A1:K{max_row}"

        # 파일 저장 (권한 문제 처리)
        try:
            wb.save(filename)
            print("[SUCCESS] 검색결과 수집 완료!")
            print(f"   - 처리된 시드 키워드: {total_processed}/{len(keywords)}개")
            print(f"   - 수집된 연관 키워드: {total_rel_keywords}개")
            print(f"   - 저장 파일: {filename}")
            print(f"   - 새 시트: searchResult")
        except PermissionError:
            print("[WARN] 엑셀 파일 저장 실패: 파일이 열려있거나 쓰기 권한이 없습니다.")
            print(f"   - 처리된 시드 키워드: {total_processed}/{len(keywords)}개")
            print(f"   - 수집된 연관 키워드: {total_rel_keywords}개")
            print("   - 결과를 확인하려면 엑셀 파일을 닫고 다시 실행해주세요.")
        except Exception as e:
            print(f"[ERROR] 파일 저장 중 오류: {e}")
            print(f"   - 처리된 시드 키워드: {total_processed}/{len(keywords)}개")
            print(f"   - 수집된 연관 키워드: {total_rel_keywords}개")

    except Exception as e:
        print(f"[ERROR] 오류 발생: {e}")


def main():
    """메인 함수"""
    print("=== 네이버 광고 API 검색결과 수집 시작 ===")
    collect_search_results()
    print("=== 네이버 광고 API 검색결과 수집 완료 ===")


if __name__ == "__main__":
    main()