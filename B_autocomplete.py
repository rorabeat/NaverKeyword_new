# -*- coding: utf-8 -*-
"""Naver Keyword Auto Complete - Multi Level Extraction

이 스크립트는 네이버 검색창 자동완성 키워드를 sub level3 단계까지 추출합니다.
- Level 1: 초기 seed 키워드들
- Level 2: Level 1의 자동완성 검색어들
- Level 3: Level 2의 자동완성 검색어들

각 seed 완료시마다 엑셀 파일에 저장하고 로그에 기록합니다.
"""

try:
    import undetected_chromedriver as uc
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from webdriver_manager.chrome import ChromeDriverManager
    SELENIUM_AVAILABLE = True
except ImportError:
    SELENIUM_AVAILABLE = False

import time
import random
import logging
from datetime import datetime
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from typing import List, Set, Tuple, Optional
import atexit
import gc


# 설정값들
AUTOCOMPLETE_TIMEOUT_SEC = 5
LOG_DIR = "result/logs"
RESULT_DIR = "result"


class AutoCompleteSession:
    """네이버 자동완성 세션 관리 클래스"""

    def __init__(self, headless: bool = True, logger: Optional[logging.Logger] = None):
        if not SELENIUM_AVAILABLE:
            raise RuntimeError("Selenium/undetected_chromedriver 가 설치되지 않았습니다.")

        self.logger = logger
        self._closed = False

        options = uc.ChromeOptions()
        if headless:
            options.add_argument("--headless=new")

        # 안정화 옵션
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-blink-features=AutomationControlled")

        # ChromeDriverManager를 사용하여 호환되는 ChromeDriver 자동 다운로드
        try:
            driver_path = ChromeDriverManager().install()
            self.driver = uc.Chrome(options=options, driver_executable_path=driver_path)
        except Exception as e:
            if self.logger:
                self.logger.warning(f"[AUTO] ChromeDriverManager 실패, 기본 undetected-chromedriver 사용: {e}")
            self.driver = uc.Chrome(options=options)
        self.driver.set_page_load_timeout(20)

        try:
            self.driver.get("https://www.naver.com")
        except Exception as e:
            if self.logger:
                self.logger.warning(f"[AUTO] 네이버 메인 로드 실패(재시도 진행): {e}")
            try:
                self.driver.get("https://www.naver.com")
            except Exception:
                pass

        # 프로세스 종료 시 안전하게 닫기
        atexit.register(self.close)

    def close(self):
        if self._closed:
            return

        self._closed = True
        try:
            if getattr(self, "driver", None):
                try:
                    # Chrome service를 먼저 정지
                    if hasattr(self.driver, 'service') and self.driver.service:
                        try:
                            self.driver.service.stop()
                        except Exception:
                            pass
                    # 드라이버 quit
                    self.driver.quit()
                except (OSError, AttributeError, Exception) as e:
                    if self.logger:
                        self.logger.debug(f"[AUTO] driver.quit() 예외 무시: {type(e).__name__}: {e}")
                    pass
        finally:
            self.driver = None
            gc.collect()

    def get_autocomplete_keywords(self, keyword: str) -> List[str]:
        """단일 키워드의 자동완성 검색어 추출"""
        if self._closed:
            raise RuntimeError("AutoCompleteSession is closed")

        driver = self.driver
        if driver is None:
            raise RuntimeError("AutoCompleteSession driver is closed")

        suggestions: List[str] = []

        try:
            if self.logger:
                self.logger.debug(f"[AUTO] 자동완성 시작: keyword='{keyword}'")

            wait = WebDriverWait(driver, 10)
            search_input = wait.until(EC.element_to_be_clickable((By.ID, "query")))

            search_input.click()
            search_input.clear()

            # 사람처럼 글자별로 랜덤 딜레이를 두면서 입력
            if self.logger:
                self.logger.debug(f"[AUTO] 입력 시작: '{keyword}'")

            for char in keyword:
                search_input.send_keys(char)
                typing_speed = random.uniform(0.1, 0.3)
                time.sleep(typing_speed)

            # 입력 완료 후 자동완성 목록이 갱신될 수 있도록 잠시 대기
            time.sleep(1.0)
            if self.logger:
                self.logger.debug("[AUTO] 입력 완료, 자동완성 대기 중...")

            try:
                if self.logger:
                    self.logger.debug(f"[AUTO] 자동완성 요소 대기 시작 (최대 {AUTOCOMPLETE_TIMEOUT_SEC}초)")

                WebDriverWait(driver, AUTOCOMPLETE_TIMEOUT_SEC).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "li._item[data-keyword]"))
                )

                items = driver.find_elements(By.CSS_SELECTOR, "li._item[data-keyword]")
                suggestions = [
                    item.get_attribute("data-keyword")
                    for item in items
                    if item.get_attribute("data-keyword")
                ]

                if self.logger:
                    self.logger.debug(f"[AUTO] 자동완성 추출 완료: {len(suggestions)}개")

            except Exception as e:
                if self.logger:
                    self.logger.warning(f"[AUTO] 자동완성 요소 대기 실패: {e}")

        except Exception as e:
            if self.logger:
                self.logger.error(f"[AUTO] 자동완성 수집 실패: keyword='{keyword}': {e}")

        return suggestions


def setup_logging() -> logging.Logger:
    """로깅 설정"""
    os.makedirs(LOG_DIR, exist_ok=True)
    os.makedirs(RESULT_DIR, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_filename = f"{LOG_DIR}/autocomplete_{timestamp}.log"

    logger = logging.getLogger('autocomplete')
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


def save_to_excel(seed_keyword: str, autocomplete_keywords: List[str], logger: logging.Logger, is_first_save: bool = False) -> str:
    """결과를 누적해서 하나의 엑셀 파일에 저장"""
    filename = f"{RESULT_DIR}/keywordList_all.xlsx"

    if is_first_save:
        # 첫 번째 저장: 기존 파일이 있으면 로드, 없으면 새로 생성
        try:
            wb = load_workbook(filename)
            if "autocomplete_results" not in wb.sheetnames:
                # 시트가 없으면 새로 생성
                ws = wb.create_sheet("autocomplete_results")
                # 헤더 추가
                headers = ["seed", "autocomplete_keyword"]
                ws.append(headers)
                # 헤더 스타일
                header_font = Font(bold=True)
                for cell in ws[1]:
                    cell.font = header_font
                    cell.alignment = Alignment(vertical="center")
                # 열 너비 설정
                ws.column_dimensions["A"].width = 30
                ws.column_dimensions["B"].width = 50
            else:
                # 시트가 있으면 해당 시트 사용
                ws = wb["autocomplete_results"]
        except FileNotFoundError:
            # 파일이 없으면 새로 생성
            wb = Workbook()
            ws = wb.active
            ws.title = "autocomplete_results"
            # 헤더
            headers = ["seed", "autocomplete_keyword"]
            ws.append(headers)
            # 헤더 스타일
            header_font = Font(bold=True)
            for cell in ws[1]:
                cell.font = header_font
                cell.alignment = Alignment(vertical="center")
            # 열 너비 설정
            ws.column_dimensions["A"].width = 30
            ws.column_dimensions["B"].width = 50
    else:
        # 기존 파일 열기
        try:
            wb = load_workbook(filename)
            if "autocomplete_results" not in wb.sheetnames:
                # 시트가 없으면 새로 생성
                ws = wb.create_sheet("autocomplete_results")
                # 헤더 추가
                headers = ["seed", "autocomplete_keyword"]
                ws.append(headers)
                header_font = Font(bold=True)
                for cell in ws[1]:
                    cell.font = header_font
                    cell.alignment = Alignment(vertical="center")
                ws.column_dimensions["A"].width = 30
                ws.column_dimensions["B"].width = 50
            else:
                # 시트가 있으면 해당 시트 사용
                ws = wb["autocomplete_results"]
        except FileNotFoundError:
            # 파일이 없으면 새로 생성
            wb = Workbook()
            ws = wb.create_sheet("autocomplete_results", 0)

            headers = ["seed", "autocomplete_keyword"]
            ws.append(headers)
            header_font = Font(bold=True)
            for cell in ws[1]:
                cell.font = header_font
                cell.alignment = Alignment(vertical="center")
            ws.column_dimensions["A"].width = 30
            ws.column_dimensions["B"].width = 50

    # 데이터 추가 (마지막 행 다음에)
    for keyword in autocomplete_keywords:
        ws.append([seed_keyword, keyword])

    # 필터링 설정 업데이트
    max_row = ws.max_row
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:B{max_row}"

    # 파일 저장
    try:
        wb.save(filename)
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
        logger.info(f"엑셀 파일 누적 저장: {filename} - seed '{seed_keyword}'의 키워드 {len(autocomplete_keywords)}개 추가")

    return filename


def extract_multilevel_autocomplete(seeds: List[str], max_level: int = 3, logger: Optional[logging.Logger] = None) -> None:
    """다단계 자동완성 키워드 추출"""
    session = AutoCompleteSession(headless=False, logger=logger)

    try:
        if logger:
            logger.info("=== 다단계 자동완성 추출 시작 ===")
            logger.info(f"초기 seed 개수: {len(seeds)}")
            logger.info(f"최대 레벨: {max_level}")

        # 레벨별 키워드 저장
        level_keywords = {1: set(seeds)}  # 중복 방지용 set 사용
        all_processed = set()  # 이미 처리한 키워드들
        is_first_save = True  # 첫 번째 저장인지 확인

        for level in range(1, max_level + 1):
            if logger:
                logger.info(f"=== 레벨 {level} 처리 시작 ===")

            current_level_seeds = list(level_keywords.get(level, []))

            for seed_idx, seed in enumerate(current_level_seeds, 1):
                if seed in all_processed:
                    if logger:
                        logger.debug(f"이미 처리된 seed 건너뜀: {seed}")
                    continue

                if logger:
                    logger.info(f"[레벨 {level}] seed {seed_idx}/{len(current_level_seeds)}: {seed}")

                # 자동완성 키워드 추출
                autocomplete_keywords = session.get_autocomplete_keywords(seed)

                # 결과가 있든 없든 항상 시트에 저장 (결과 없음 표시)
                excel_file = save_to_excel(seed, autocomplete_keywords, logger, is_first_save)
                is_first_save = False  # 첫 번째 저장 이후에는 False

                # 로그에 기록
                if logger:
                    if autocomplete_keywords:
                        logger.info(f"seed '{seed}' 완료 - 자동완성 키워드 {len(autocomplete_keywords)}개 추출 및 저장")
                        logger.info(f"현재 엑셀 파일: {excel_file}")
                    else:
                        logger.warning(f"seed '{seed}' - 자동완성 키워드 없음 (시트에 기록됨)")

                # 다음 레벨을 위한 키워드 수집 (결과가 있을 때만)
                if autocomplete_keywords and level < max_level:
                    next_level_keywords = level_keywords.get(level + 1, set())
                    next_level_keywords.update(autocomplete_keywords)
                    level_keywords[level + 1] = next_level_keywords

                # 처리 완료된 seed 기록
                all_processed.add(seed)

                # 잠시 대기 (서버 부하 방지)
                time.sleep(random.uniform(1.0, 2.0))

        if logger:
            logger.info("=== 다단계 자동완성 추출 완료 ===")

    finally:
        session.close()


def main():
    """메인 함수"""
    logger = setup_logging()

    try:
        logger.info("=== 네이버 자동완성 다단계 추출기 시작 ===")

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

        # 다단계 추출 실행 (헤드리스 모드 비활성화)
        extract_multilevel_autocomplete(seeds, max_level=3, logger=logger)

    except KeyboardInterrupt:
        logger.info("사용자에 의해 중단됨")
    except Exception as e:
        logger.error(f"프로그램 실행 중 오류 발생: {e}")
        raise


if __name__ == "__main__":
    main()