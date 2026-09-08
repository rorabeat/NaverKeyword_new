# -*- coding: utf-8 -*-
"""Naver Related Keywords (검색광고 API 기반) - Multi Level Extraction

네이버 우측 연관 검색어 서비스가 종료되어, 네이버 검색광고 API(/keywordstool, RelKwdStat)로
연관 키워드를 sub level3 단계까지 추출합니다.
- Level 1: 초기 seed 키워드들
- Level 2: Level 1의 연관 키워드들
- Level 3: Level 2의 연관 키워드들

각 단계마다 rightside_result 시트에 결과를 누적 저장합니다.
(하위 모듈(H_IntegrationGUI, C_sumKeyword)과의 호환을 위해 시트/함수명은 rightside_* 를 유지합니다.)
"""

import os
import re
import time
import random
import hmac
import hashlib
import base64
import logging
from datetime import datetime, timezone, timedelta
from typing import List, Set, Optional, Tuple

import requests
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font


# 설정값들
RESULT_DIR = "result"
LOG_DIR = "result/logs"

AD_API_BASE_URL = "https://api.searchad.naver.com"
AD_API_ENDPOINT = "/keywordstool"  # RelKwdStat list
NAVER_BLOG_SEARCH_URL = "https://openapi.naver.com/v1/search/blog.json"

# seed 1개당 검색광고 API가 수백 개의 연관 키워드를 반환할 수 있어(예: "노트북" -> 879개),
# 결과가 너무 방대해지는 것을 막기 위해 검색량(PC+모바일) 상위 MAX_KEYWORDS_PER_SEED개만
# 남기고(엑셀 저장 포함) 나머지는 버린다.
MAX_KEYWORDS_PER_SEED = 100

# 검색량 상위 몇 개까지 블로그 발행수(최근 30일)를 조회해 "검색은 많고 블로그는 적은"
# 황금키워드 순으로 재정렬할지. 전체를 다 조회하면 seed당 API 호출이 크게 늘어나므로
# 검색량 상위 후보군만 조회한다.
BLOG_CANDIDATE_POOL_SIZE = 50

# 레벨을 그대로 넘기면 다음 레벨의 seed 수가 기하급수적으로 늘어나므로,
# 다음 레벨로는 그중에서도 상위 TOP_N_PER_LEVEL개만 seed로 넘긴다.
TOP_N_PER_LEVEL = 20

# 초기 seed 1개당 전체 레벨(1~max_level)을 통틀어 최종적으로 수집할 연관 키워드 총 개수 상한.
# 도달하는 즉시 이후 레벨/seed 처리를 중단해(불필요한 API 호출과 블로그 조회 생략) 실행 시간을 크게 줄인다.
TOTAL_RELATED_CAP_PER_RUN = 100


def safe_filename(text: str, max_len: int = 40) -> str:
    """파일명으로 안전한 문자열로 변환"""
    text = re.sub(r"[^0-9a-zA-Z가-힣]+", "_", text).strip("_")
    return (text[:max_len] if text else "keyword")


def _make_signature(secret_key: str, timestamp: str, method: str, uri: str) -> str:
    """signature = Base64( HMAC-SHA256(secret_key, f"{timestamp}.{method}.{uri}") )"""
    message = f"{timestamp}.{method}.{uri}"
    digest = hmac.new(secret_key.encode("utf-8"), message.encode("utf-8"), hashlib.sha256).digest()
    return base64.b64encode(digest).decode("utf-8")


def _build_ad_api_headers(api_key: str, secret_key: str, customer_id: str, method: str, uri: str) -> dict:
    timestamp = str(int(time.time() * 1000))
    signature = _make_signature(secret_key, timestamp, method, uri)

    return {
        "X-Timestamp": timestamp,
        "X-API-KEY": api_key,
        "X-Customer": customer_id,
        "X-Signature": signature,
        "Content-Type": "application/json; charset=UTF-8",
    }


def _normalize_hint_keywords(hint_keywords: str) -> str:
    """keywordstool의 hintKeywords는 공백이 포함되면 400(Invalid Parameter)이 발생하므로,
    콤마로 구분된 각 키워드 내부의 공백만 제거한다. (콤마 자체는 유지)
    예) "민음사 빵 책갈피" -> "민음사빵책갈피", "삐에로 남친, 츄라우미 수족관" -> "삐에로남친,츄라우미수족관"
    """
    return ",".join("".join(part.split()) for part in hint_keywords.split(","))


def call_keywordstool(hint_keywords: str, show_detail: int = 1) -> dict:
    """네이버 검색광고 API(/keywordstool)를 호출해 연관 키워드 목록을 가져온다."""
    load_dotenv()
    api_key = os.getenv("NAVER_SEARCH_ACCESS_LICENSE_KEY", "").strip()
    secret_key = os.getenv("NAVER_SEARCH_SECRET_KEY", "").strip()
    customer_id = os.getenv("NAVER_SEARCH_CUSTOMER_ID", "").strip()

    if not api_key or not secret_key or not customer_id:
        raise RuntimeError("ENV 누락: .env에 NAVER_SEARCH_ACCESS_LICENSE_KEY / SECRET_KEY / CUSTOMER_ID 를 설정하세요.")

    method = "GET"
    headers = _build_ad_api_headers(api_key, secret_key, customer_id, method, AD_API_ENDPOINT)

    # hintKeywords는 콤마 구분 최대 5개이며, 각 키워드 내부에 공백이 있으면 400이 발생하므로 제거한다.
    params = {
        "hintKeywords": _normalize_hint_keywords(hint_keywords),
        "showDetail": str(show_detail),
    }

    url = AD_API_BASE_URL + AD_API_ENDPOINT
    r = requests.get(url, headers=headers, params=params, timeout=15)
    r.raise_for_status()
    return r.json()


def _get_recent_30day_blog_count(keyword: str, client_id: str, client_secret: str) -> Optional[int]:
    """키워드의 최근 30일 발행 블로그 수를 반환(최대 100건까지만 카운트). 조회 실패 시 None."""
    headers = {
        "X-Naver-Client-Id": client_id,
        "X-Naver-Client-Secret": client_secret,
    }
    params = {"query": keyword, "display": 100, "start": 1, "sort": "date"}

    try:
        r = requests.get(NAVER_BLOG_SEARCH_URL, headers=headers, params=params, timeout=10)
        r.raise_for_status()
        data = r.json()
    except Exception:
        return None

    kst = timezone(timedelta(hours=9))
    cutoff = int((datetime.now(kst) - timedelta(days=30)).strftime("%Y%m%d"))

    count = 0
    for item in data.get("items", []) or []:
        postdate = item.get("postdate")
        if not postdate or len(str(postdate)) != 8:
            continue
        if int(postdate) < cutoff:
            break  # 최신순(sort=date) 정렬이므로 컷오프 이전 글이 나오면 더 볼 필요 없음
        count += 1

    return count


def extract_related_keywords_from_api(seed: str) -> List[Tuple[str, int, Optional[int], Optional[float]]]:
    """
    검색광고 API 응답에서 연관 키워드(relKeyword) 목록을 추출한다.
    seed 자기 자신은 결과에서 제외하고, 순서를 유지하며 중복을 제거한다.

    검색량 상위 BLOG_CANDIDATE_POOL_SIZE개는 최근 30일 블로그 발행수를 추가로 조회해
    "검색은 많고 블로그는 적은" 황금키워드 점수(golden_score = 검색량 / (블로그수+1)) 순으로
    재정렬한다. (.env에 NAVER_CLIENT_ID/SECRET이 없으면 블로그 조회를 건너뛰고 검색량 순 유지)

    반환값: [(keyword, monthly_search_qc, recent_30day_blog_count, golden_score), ...]
    """
    data = call_keywordstool(seed, show_detail=1)

    keyword_list = data.get("keywordList")
    if keyword_list is None and isinstance(data, list):
        keyword_list = data
    if keyword_list is None:
        keyword_list = []

    seed_norm = seed.strip().replace(" ", "").lower()

    def _qc_to_int(v) -> int:
        # API가 소량 검색량을 "< 10" 같은 문자열로 주는 경우 방어 처리
        try:
            return int(v)
        except (TypeError, ValueError):
            return 0

    rows = []
    seen = set()
    for row in keyword_list:
        kw = str(row.get("relKeyword", "")).strip()
        if not kw:
            continue
        if kw.replace(" ", "").lower() == seed_norm:
            continue  # seed 자기 자신 제외
        if kw in seen:
            continue
        seen.add(kw)
        total_qc = _qc_to_int(row.get("monthlyPcQcCnt")) + _qc_to_int(row.get("monthlyMobileQcCnt"))
        rows.append((kw, total_qc))

    # 1차: 검색량(PC+모바일) 내림차순 정렬
    rows.sort(key=lambda x: x[1], reverse=True)

    load_dotenv()
    client_id = os.getenv("NAVER_CLIENT_ID", "").strip()
    client_secret = os.getenv("NAVER_CLIENT_SECRET", "").strip()

    pool_size = min(BLOG_CANDIDATE_POOL_SIZE, MAX_KEYWORDS_PER_SEED, len(rows))
    pool = rows[:pool_size]

    if not (client_id and client_secret) or not pool:
        # 블로그 API 키가 없으면 검색량 기준 상위 MAX_KEYWORDS_PER_SEED개만 반환
        return [(kw, qc, None, None) for kw, qc in rows[:MAX_KEYWORDS_PER_SEED]]

    # 2차: 검색량 상위 후보군만 블로그 발행수를 조회해 황금키워드 점수로 재정렬
    scored = []
    for kw, total_qc in pool:
        blog_count = _get_recent_30day_blog_count(kw, client_id, client_secret)
        golden_score = (total_qc / (blog_count + 1)) if blog_count is not None else -1.0
        scored.append((kw, total_qc, blog_count, golden_score))
        time.sleep(0.1)  # 블로그 API 과도 호출 방지

    scored.sort(key=lambda x: x[3], reverse=True)
    return scored


def setup_logging() -> logging.Logger:
    """로깅 설정"""
    import os
    os.makedirs(LOG_DIR, exist_ok=True)
    os.makedirs(RESULT_DIR, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_filename = f"{LOG_DIR}/rightside_{timestamp}.log"

    logger = logging.getLogger('rightside')
    logger.setLevel(logging.DEBUG)

    # 파일 핸들러
    fh = logging.FileHandler(log_filename, encoding='utf-8')
    fh.setLevel(logging.DEBUG)

    # 콘솔 핸들러
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)

    # 포맷터
    formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
    fh.setFormatter(formatter)
    ch.setFormatter(formatter)

    logger.addHandler(fh)
    logger.addHandler(ch)

    return logger


def save_to_excel_sheet(seed_keyword: str, related_keywords: List[Tuple[str, int, Optional[int], Optional[float]]], level: int, logger: logging.Logger, is_first_save: bool = False) -> str:
    """keywordList_all 엑셀파일의 naver_ad__result 시트에 결과를 누적해서 저장"""
    import os
    # result 폴더 생성 확인
    os.makedirs(RESULT_DIR, exist_ok=True)
    filename = f"{RESULT_DIR}/keywordList_all.xlsx"

    if is_first_save:
        # 첫 번째 저장: 기존 파일이 있으면 로드, 없으면 새로 생성
        try:
            wb = load_workbook(filename)
            if logger:
                logger.debug(f"기존 엑셀 파일 로드 성공: {filename}")
        except FileNotFoundError:
            # 파일이 없으면 새로 생성
            if logger:
                logger.info(f"엑셀 파일이 없어 새로 생성: {filename}")
            wb = Workbook()
            # 기본 시트 제거 (Workbook 생성 시 자동으로 생성되는 시트)
            if "Sheet" in wb.sheetnames:
                wb.remove(wb["Sheet"])
        except Exception as e:
            if logger:
                logger.error(f"엑셀 파일 로드 중 오류: {e}")
            raise

        # naver_ad__result 시트 확인 및 생성
        if "naver_ad__result" not in wb.sheetnames:
            if logger:
                logger.info("naver_ad__result 시트가 없어 새로 생성")
            ws = wb.create_sheet("naver_ad__result", 0)  # 첫 번째 시트로 생성
            # 헤더 추가
            headers = ["seed", "related_keyword", "level", "monthly_search_qc", "recent_30day_blog_count", "golden_score"]
            ws.append(headers)
            # 헤더 스타일
            header_font = Font(bold=True)
            for cell in ws[1]:
                cell.font = header_font
                cell.alignment = Alignment(vertical="center")
            # 열 너비 설정
            ws.column_dimensions["A"].width = 30
            ws.column_dimensions["B"].width = 50
            ws.column_dimensions["C"].width = 10
            ws.column_dimensions["D"].width = 16
            ws.column_dimensions["E"].width = 20
            ws.column_dimensions["F"].width = 12
        else:
            # 시트가 있으면 해당 시트 사용
            ws = wb["naver_ad__result"]
    else:
        # 기존 파일 열기 (없으면 새로 생성)
        try:
            wb = load_workbook(filename)
            if logger:
                logger.debug(f"기존 엑셀 파일 로드 성공: {filename}")
        except FileNotFoundError:
            # 파일이 없으면 새로 생성
            if logger:
                logger.info(f"엑셀 파일이 없어 새로 생성: {filename}")
            wb = Workbook()
            # 기본 시트 제거 (Workbook 생성 시 자동으로 생성되는 시트)
            if "Sheet" in wb.sheetnames:
                wb.remove(wb["Sheet"])
        except Exception as e:
            if logger:
                logger.error(f"엑셀 파일 로드 중 오류: {e}")
            raise

        # naver_ad__result 시트 확인 및 생성
        if "naver_ad__result" not in wb.sheetnames:
            if logger:
                logger.info("naver_ad__result 시트가 없어 새로 생성")
            ws = wb.create_sheet("naver_ad__result", 0)
            # 헤더 추가
            headers = ["seed", "related_keyword", "level", "monthly_search_qc", "recent_30day_blog_count", "golden_score"]
            ws.append(headers)
            header_font = Font(bold=True)
            for cell in ws[1]:
                cell.font = header_font
                cell.alignment = Alignment(vertical="center")
            ws.column_dimensions["A"].width = 30
            ws.column_dimensions["B"].width = 50
            ws.column_dimensions["C"].width = 10
            ws.column_dimensions["D"].width = 16
            ws.column_dimensions["E"].width = 20
            ws.column_dimensions["F"].width = 12
        else:
            # 시트가 있으면 해당 시트 사용
            ws = wb["naver_ad__result"]

    # 데이터 추가 (마지막 행 다음에)
    if related_keywords:
        # 각 related_keyword마다 별도의 행으로 저장
        for related_keyword, search_qc, blog_count, golden_score in related_keywords:
            ws.append([
                seed_keyword,
                related_keyword,
                level,
                search_qc,
                blog_count if blog_count is not None else "",
                round(golden_score, 2) if golden_score is not None else "",
            ])

    # 필터링 설정 업데이트
    max_row = ws.max_row
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F{max_row}"

    # 파일 저장
    try:
        wb.save(filename)
        if logger:
            logger.info(f"엑셀 파일 저장 성공: {filename}")
    except PermissionError as e:
        error_msg = f"엑셀 파일이 열려있어서 저장할 수 없습니다. 파일을 닫고 다시 시도해주세요: {filename}"
        print(f"❌ {error_msg}")
        if logger:
            logger.error(f"파일 저장 실패 - PermissionError: {error_msg}")
        raise RuntimeError(error_msg) from e
    except Exception as e:
        error_msg = f"엑셀 파일 저장 중 오류 발생: {e}"
        print(f"❌ {error_msg}")
        if logger:
            logger.error(f"파일 저장 실패: {error_msg}")
        raise

    if logger:
        logger.info(f"naver_ad__result 시트 누적 저장: {filename} - seed '{seed_keyword}'의 연관검색어 {len(related_keywords)}개 추가 (레벨 {level})")

    return filename


def extract_multilevel_rightside_related(seeds: List[str], max_level: int = 3, logger: Optional[logging.Logger] = None) -> None:
    """다단계 연관 키워드 추출 (검색광고 API 기반)"""
    try:
        if logger:
            logger.info("=== 다단계 연관 키워드 추출 시작 (검색광고 API) ===")
            logger.info(f"초기 seed 개수: {len(seeds)}")
            logger.info(f"최대 레벨: {max_level}")

        # 레벨별 키워드 저장
        level_keywords = {1: set(seeds)}  # 중복 방지용 set 사용
        all_processed = set()  # 이미 처리한 키워드들
        is_first_save = True  # 첫 번째 저장인지 확인
        total_collected = 0  # 이 실행(초기 seed 1개)에서 지금까지 저장한 연관 키워드 총 개수

        for level in range(1, max_level + 1):
            if total_collected >= TOTAL_RELATED_CAP_PER_RUN:
                if logger:
                    logger.info(f"전체 연관 키워드 상한({TOTAL_RELATED_CAP_PER_RUN}개) 도달로 레벨 {level} 이후 처리를 중단합니다.")
                break

            if logger:
                logger.info(f"=== 레벨 {level} 처리 시작 ===")

            current_level_seeds = list(level_keywords.get(level, []))

            for seed_idx, seed in enumerate(current_level_seeds, 1):
                if total_collected >= TOTAL_RELATED_CAP_PER_RUN:
                    if logger:
                        logger.info(f"전체 연관 키워드 상한({TOTAL_RELATED_CAP_PER_RUN}개) 도달로 남은 seed 처리를 건너뜁니다.")
                    break

                if seed in all_processed:
                    if logger:
                        logger.debug(f"이미 처리된 seed 건너뜀: {seed}")
                    continue

                if logger:
                    logger.info(f"[레벨 {level}] seed {seed_idx}/{len(current_level_seeds)}: {seed}")

                try:
                    # 연관 키워드 추출 (검색광고 API)
                    related_keywords = extract_related_keywords_from_api(seed)

                    # 전체 상한을 넘지 않도록 남은 할당량만큼만 저장
                    remaining_quota = TOTAL_RELATED_CAP_PER_RUN - total_collected
                    if remaining_quota <= 0:
                        related_keywords = []
                    elif len(related_keywords) > remaining_quota:
                        related_keywords = related_keywords[:remaining_quota]

                    if related_keywords:
                        # 콘솔에 결과 출력 (검색량, 최근30일 블로그수, 황금키워드 점수)
                        for related_keyword, search_qc, blog_count, golden_score in related_keywords:
                            print(f"{seed}\t{related_keyword}\t{level}\tqc={search_qc}\tblog30d={blog_count}\tscore={golden_score}")

                        # 실시간으로 엑셀 시트에 누적 저장
                        excel_file = save_to_excel_sheet(seed, related_keywords, level, logger, is_first_save)
                        is_first_save = False  # 첫 번째 저장 이후에는 False
                        total_collected += len(related_keywords)

                        # 로그에 기록
                        if logger:
                            logger.info(f"seed '{seed}' 완료 - 연관 키워드 {len(related_keywords)}개 추출 및 저장 (레벨 {level}, 누적 {total_collected}/{TOTAL_RELATED_CAP_PER_RUN})")
                            logger.info(f"현재 엑셀 파일: {excel_file}")

                        # 다음 레벨을 위한 키워드 수집 (중복 제거)
                        # related_keywords는 황금키워드 점수(검색량은 많고 블로그는 적은 순) 내림차순 정렬되어
                        # 있으므로 상위 N개만 다음 레벨 seed로 사용
                        if level < max_level and total_collected < TOTAL_RELATED_CAP_PER_RUN:
                            next_level_keywords = level_keywords.get(level + 1, set())
                            next_level_keywords.update(kw for kw, _, _, _ in related_keywords[:TOP_N_PER_LEVEL])
                            level_keywords[level + 1] = next_level_keywords
                    else:
                        if logger:
                            logger.warning(f"seed '{seed}' - 연관 키워드 없음")

                except requests.HTTPError as e:
                    if logger:
                        logger.error(f"HTTP 오류 - seed '{seed}': {e}")
                except Exception as e:
                    if logger:
                        logger.error(f"처리 오류 - seed '{seed}': {e}")

                # 처리 완료된 seed 기록
                all_processed.add(seed)

                # 잠시 대기 (서버 부하 방지)
                time.sleep(random.uniform(1.0, 2.0))

        if logger:
            logger.info("=== 다단계 연관 키워드 추출 완료 (검색광고 API) ===")

    except Exception as e:
        if logger:
            logger.error(f"프로그램 실행 중 오류 발생: {e}")
        raise


def main():
    """메인 함수"""
    logger = setup_logging()

    try:
        logger.info("=== 네이버 연관 키워드 다단계 추출기 시작 (검색광고 API) ===")

        # 시드 키워드 입력
        print("시드 키워드를 입력하세요. 여러 개면 콤마(,)로 구분하거나 줄바꿈으로 입력")
        print("입력 끝내려면 빈 줄에서 Enter\n")

        seeds: List[str] = []
        while True:
            line = input("seed > ").strip()
            if not line:
                break
            parts = [p.strip() for p in line.split(",") if p.strip()]
            seeds.extend(parts)

        # 중복 제거하면서 순서 유지
        seen = set()
        seeds = [x for x in seeds if not (x in seen or seen.add(x))]

        if not seeds:
            print("입력된 시드 키워드가 없습니다.")
            return

        print(f"\n처리할 시드 키워드들: {seeds}")
        logger.info(f"입력된 시드 키워드: {seeds}")

        # 다단계 추출 실행
        extract_multilevel_rightside_related(seeds, max_level=3, logger=logger)

    except KeyboardInterrupt:
        logger.info("사용자에 의해 중단됨")
    except Exception as e:
        logger.error(f"프로그램 실행 중 오류 발생: {e}")
        raise


if __name__ == "__main__":
    main()