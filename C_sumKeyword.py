# -*- coding: utf-8 -*-
"""키워드 합치기 - autocomplete_results와 rightside_results 시트의 결과를 통합

keywordList_all.xlsx 파일에서:
- autocomplete_results 시트의 B열 (autocomplete_keyword) 데이터
- rightside_results 시트의 B열 (related_keyword) 데이터
를 새로운 sumKeyword 시트에 통합 저장합니다.

A열: 데이터 출처 (autocomplete 또는 rightside)
B열: 키워드 데이터
"""

import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font


def merge_keywords_to_sum_sheet():
    """autocomplete_results와 rightside_results 시트의 키워드를 sumKeyword 시트에 통합"""

    filename = "result/keywordList_all.xlsx"

    if not os.path.exists(filename):
        print(f"[ERROR] 파일이 존재하지 않습니다: {filename}")
        return

    try:
        # 기존 파일 로드
        wb = load_workbook(filename)
        print(f"[INFO] 파일 로드 완료: {filename}")

        # 데이터 수집
        merged_data = []

        # 1. autocomplete_results 시트에서 B열 데이터 수집
        if "autocomplete_results" in wb.sheetnames:
            ws_auto = wb["autocomplete_results"]
            print(f"[INFO] autocomplete_results 시트 처리 중...")

            # 헤더 제외하고 2행부터 데이터 수집
            for row in range(2, ws_auto.max_row + 1):
                keyword = ws_auto.cell(row=row, column=2).value  # B열
                if keyword and str(keyword).strip():
                    merged_data.append(("autocomplete", str(keyword).strip()))
        else:
            print("[WARN] autocomplete_results 시트가 존재하지 않습니다.")

        # 2. rightside_results 시트에서 B열 데이터 수집
        if "rightside_results" in wb.sheetnames:
            ws_right = wb["rightside_results"]
            print(f"[INFO] rightside_results 시트 처리 중...")

            # 헤더 제외하고 2행부터 데이터 수집
            for row in range(2, ws_right.max_row + 1):
                keyword = ws_right.cell(row=row, column=2).value  # B열
                if keyword and str(keyword).strip():
                    merged_data.append(("rightside", str(keyword).strip()))
        else:
            print("[WARN] rightside_results 시트가 존재하지 않습니다.")

        if not merged_data:
            print("[ERROR] 통합할 데이터가 없습니다.")
            return

        # 3. sumKeyword 시트 생성 및 데이터 저장
        if "sumKeyword" not in wb.sheetnames:
            ws_sum = wb.create_sheet("sumKeyword")
            print("[INFO] sumKeyword 시트 생성")
        else:
            ws_sum = wb["sumKeyword"]
            print("[INFO] 기존 sumKeyword 시트 사용")

        # 헤더 추가
        headers = ["source", "keyword"]
        ws_sum.append(headers)

        # 헤더 스타일링
        header_font = Font(bold=True)
        for cell in ws_sum[1]:
            cell.font = header_font
            cell.alignment = Alignment(vertical="center")

        # 열 너비 설정
        ws_sum.column_dimensions["A"].width = 15
        ws_sum.column_dimensions["B"].width = 50

        # 데이터 추가
        for source, keyword in merged_data:
            ws_sum.append([source, keyword])

        # 필터링 설정
        max_row = ws_sum.max_row
        ws_sum.freeze_panes = "A2"
        ws_sum.auto_filter.ref = f"A1:B{max_row}"

        # 파일 저장
        wb.save(filename)

        print("[SUCCESS] 키워드 통합 완료!")
        print(f"   - 총 {len(merged_data)}개 키워드 통합")
        print(f"   - 저장 파일: {filename}")
        print(f"   - 새 시트: sumKeyword")

    except Exception as e:
        print(f"[ERROR] 오류 발생: {e}")


def remove_duplicates_from_sum_sheet():
    """sumKeyword 시트에서 중복 키워드를 제거하여 sumKeyword_exceptDuplicate 시트에 저장"""

    filename = "result/keywordList_all.xlsx"

    if not os.path.exists(filename):
        print(f"[ERROR] 파일이 존재하지 않습니다: {filename}")
        return

    try:
        # 기존 파일 로드
        wb = load_workbook(filename)

        # sumKeyword 시트가 있는지 확인
        if "sumKeyword" not in wb.sheetnames:
            print("[ERROR] sumKeyword 시트가 존재하지 않습니다. 먼저 키워드 통합을 실행해주세요.")
            return

        ws_sum = wb["sumKeyword"]
        print(f"[INFO] sumKeyword 시트에서 데이터 읽기 중... ({ws_sum.max_row}행)")

        # 데이터 수집 (헤더 제외)
        data = []
        for row in range(2, ws_sum.max_row + 1):
            source = ws_sum.cell(row=row, column=1).value  # A열
            keyword = ws_sum.cell(row=row, column=2).value  # B열
            if keyword and str(keyword).strip():
                data.append((str(source).strip(), str(keyword).strip()))

        if not data:
            print("[ERROR] sumKeyword 시트에 처리할 데이터가 없습니다.")
            return

        print(f"[INFO] 총 {len(data)}개 데이터 읽음")

        # 중복 제거 (keyword를 기준으로, 첫 번째로 발견된 source 유지)
        seen_keywords = set()
        deduplicated_data = []

        for source, keyword in data:
            if keyword not in seen_keywords:
                seen_keywords.add(keyword)
                deduplicated_data.append((source, keyword))

        removed_count = len(data) - len(deduplicated_data)
        print(f"[INFO] 중복 제거 완료: {len(deduplicated_data)}개 유지, {removed_count}개 제거")

        # sumKeyword_exceptDuplicate 시트 생성
        if "sumKeyword_exceptDuplicate" in wb.sheetnames:
            # 기존 시트가 있으면 삭제 후 새로 생성
            wb.remove(wb["sumKeyword_exceptDuplicate"])
            print("[INFO] 기존 sumKeyword_exceptDuplicate 시트 제거")

        ws_dedup = wb.create_sheet("sumKeyword_exceptDuplicate")
        print("[INFO] sumKeyword_exceptDuplicate 시트 생성")

        # 헤더 추가
        headers = ["source", "keyword"]
        ws_dedup.append(headers)

        # 헤더 스타일링
        header_font = Font(bold=True)
        for cell in ws_dedup[1]:
            cell.font = header_font
            cell.alignment = Alignment(vertical="center")

        # 열 너비 설정
        ws_dedup.column_dimensions["A"].width = 15
        ws_dedup.column_dimensions["B"].width = 50

        # 중복 제거된 데이터 추가
        for source, keyword in deduplicated_data:
            ws_dedup.append([source, keyword])

        # 필터링 설정
        max_row = ws_dedup.max_row
        ws_dedup.freeze_panes = "A2"
        ws_dedup.auto_filter.ref = f"A1:B{max_row}"

        # 파일 저장
        wb.save(filename)

        print("[SUCCESS] 중복 제거 완료!")
        print(f"   - 원본 데이터: {len(data)}개")
        print(f"   - 중복 제거 후: {len(deduplicated_data)}개")
        print(f"   - 제거된 중복: {removed_count}개")
        print(f"   - 저장 파일: {filename}")
        print(f"   - 새 시트: sumKeyword_exceptDuplicate")

    except Exception as e:
        print(f"[ERROR] 오류 발생: {e}")


def filter_by_required_keywords(required_keywords):
    """sumKeyword_exceptDuplicate 시트에서 필수 키워드가 포함된 데이터만 필터링하여 sumKeyword_final 시트에 저장"""

    filename = "result/keywordList_all.xlsx"

    if not os.path.exists(filename):
        print(f"[ERROR] 파일이 존재하지 않습니다: {filename}")
        return

    try:
        # 기존 파일 로드
        wb = load_workbook(filename)

        # sumKeyword_exceptDuplicate 시트가 있는지 확인
        if "sumKeyword_exceptDuplicate" not in wb.sheetnames:
            print("[ERROR] sumKeyword_exceptDuplicate 시트가 존재하지 않습니다. 먼저 중복 제거를 실행해주세요.")
            return

        ws_dedup = wb["sumKeyword_exceptDuplicate"]
        print(f"[INFO] sumKeyword_exceptDuplicate 시트에서 데이터 읽기 중... ({ws_dedup.max_row}행)")

        # 데이터 수집 및 필터링 (헤더 제외)
        filtered_data = []
        for row in range(2, ws_dedup.max_row + 1):
            source = ws_dedup.cell(row=row, column=1).value  # A열
            keyword = ws_dedup.cell(row=row, column=2).value  # B열

            if keyword and str(keyword).strip():
                keyword_str = str(keyword).strip()

                # 필수 키워드 필터링 로직
                if not required_keywords:
                    # 필수 키워드가 없으면 모든 데이터 유지
                    filtered_data.append((str(source).strip(), keyword_str))
                else:
                    # 필수 키워드가 있으면 필터링 적용
                    contains_required = any(req_kw.lower() in keyword_str.lower() for req_kw in required_keywords)
                    if contains_required:
                        filtered_data.append((str(source).strip(), keyword_str))

        print(f"[INFO] 필수 키워드 필터링 완료: {len(filtered_data)}개 유지")

        if not filtered_data and required_keywords:
            print("[WARN] 필수 키워드가 포함된 데이터가 없습니다. 빈 시트를 생성합니다.")

        # sumKeyword_final 시트 생성
        if "sumKeyword_final" in wb.sheetnames:
            # 기존 시트가 있으면 삭제 후 새로 생성
            wb.remove(wb["sumKeyword_final"])
            print("[INFO] 기존 sumKeyword_final 시트 제거")

        ws_final = wb.create_sheet("sumKeyword_final")
        print("[INFO] sumKeyword_final 시트 생성")

        # 헤더 추가
        headers = ["source", "keyword"]
        ws_final.append(headers)

        # 헤더 스타일링
        header_font = Font(bold=True)
        for cell in ws_final[1]:
            cell.font = header_font
            cell.alignment = Alignment(vertical="center")

        # 열 너비 설정
        ws_final.column_dimensions["A"].width = 15
        ws_final.column_dimensions["B"].width = 50

        # 필터링된 데이터 추가
        for source, keyword in filtered_data:
            ws_final.append([source, keyword])

        # 필터링 설정
        max_row = ws_final.max_row
        ws_final.freeze_panes = "A2"
        ws_final.auto_filter.ref = f"A1:B{max_row}"

        # 파일 저장
        wb.save(filename)

        print("[SUCCESS] 필수 키워드 필터링 완료!")
        print(f"   - 필수 키워드: {required_keywords}")
        print(f"   - 원본 데이터: {ws_dedup.max_row - 1}개")
        print(f"   - 필터링 후: {len(filtered_data)}개")
        print(f"   - 제거된 데이터: {ws_dedup.max_row - 1 - len(filtered_data)}개")
        print(f"   - 저장 파일: {filename}")
        print(f"   - 새 시트: sumKeyword_final")

    except Exception as e:
        print(f"[ERROR] 오류 발생: {e}")


def main():
    """메인 함수"""
    # 필수 키워드 입력 받기
    print("=== 필수 키워드 입력 ===")
    print("최종 결과에 반드시 포함되어야 할 키워드를 입력하세요.")
    print("여러 개는 콤마(,)로 구분, 빈 입력시 모든 데이터 유지")
    print()

    # 테스트용 기본값 설정 (실제 사용시 주석 해제)
    # required_keywords_input = "오키나와"  # 테스트용

    try:
        required_keywords_input = input("필수 키워드 > ").strip()
    except EOFError:
        # 테스트 환경에서는 기본값 사용
        print("[TEST] EOF 감지됨 - 테스트용 기본 키워드 사용")
        required_keywords_input = "오키나와"

    required_keywords = []
    if required_keywords_input:
        # 콤마로 분리하고 공백 제거
        required_keywords = [kw.strip() for kw in required_keywords_input.split(",") if kw.strip()]
        print(f"[INFO] 필수 키워드 설정: {required_keywords}")
    else:
        print("[INFO] 필수 키워드 없음 - 모든 데이터 유지")

    print()
    print("=== 키워드 통합 시작 ===")
    merge_keywords_to_sum_sheet()
    print("=== 키워드 통합 완료 ===")
    print()
    print("=== 중복 제거 시작 ===")
    remove_duplicates_from_sum_sheet()
    print("=== 중복 제거 완료 ===")
    print()
    print("=== 필수 키워드 필터링 시작 ===")
    filter_by_required_keywords(required_keywords)
    print("=== 필수 키워드 필터링 완료 ===")


if __name__ == "__main__":
    main()