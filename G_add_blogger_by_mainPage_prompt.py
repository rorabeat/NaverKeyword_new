#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
G_add_blogger_by_mainPage_prompt.py - 네이버 메인 페이지 검색으로 블로그 아이디 수집 도구

이 스크립트는 사용자가 입력한 여러 개의 키워드에 대해 네이버 메인 페이지 검색을 통해
블로그 아이디들을 수집하여 터미널에 출력합니다.

사용법:
    python G_add_blogger_by_mainPage_prompt.py

기능:
- 여러 개의 키워드 입력 받기
- 각 키워드별 블로그 아이디 수집 (최대 5개씩)
- 수집된 블로그 아이디들을 정리하여 출력
"""

import sys
import os
from typing import List, Dict, Any
from datetime import datetime, timedelta, timezone

# G_add_blogger_by_mainPage.py에서 필요한 함수들 import
from G_add_blogger_by_mainPage import (
    get_top_bloggers_from_main_page,
    get_blogger_data_from_main_page
)

# 환경변수 로드
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    print("경고: python-dotenv가 설치되지 않았습니다. .env 파일에서 환경변수를 로드할 수 없습니다.")


def get_multiple_keywords() -> List[str]:
    """
    터미널에서 여러 개의 키워드를 입력받음
    """
    print("\n" + "="*60)
    print("네이버 메인 페이지 블로그 아이디 수집 도구")
    print("="*60)
    print("사용법:")
    print("- 키워드를 입력하세요 (쉼표로 구분)")
    print("- 예: 오키나와 여행,후쿠오카 맛집,도쿄 관광")
    print("- 입력을 마치려면 빈 줄을 입력하세요")
    print("="*60)

    keywords = []

    while True:
        try:
            user_input = input("\n키워드를 입력하세요: ").strip()

            if not user_input:
                if keywords:
                    break
                else:
                    print("최소 하나의 키워드를 입력해주세요.")
                    continue

            # 쉼표로 구분된 키워드들 분리
            new_keywords = [kw.strip() for kw in user_input.split(',') if kw.strip()]
            keywords.extend(new_keywords)

            print(f"현재까지 입력된 키워드: {keywords}")

        except KeyboardInterrupt:
            print("\n\n입력 취소됨.")
            return []
        except EOFError:
            break

    return keywords


def extract_blogger_ids(bloggers_data: List[Dict[str, Any]]) -> List[str]:
    """
    블로거 데이터에서 블로그 아이디만 추출
    """
    blogger_ids = []

    for blogger_info in bloggers_data:
        blogger_name = blogger_info.get("bloggername", "")
        blogger_link = blogger_info.get("bloggerlink", "")

        # 블로그 링크에서 아이디 추출
        if blogger_link and "blog.naver.com/" in blogger_link:
            blog_id = blogger_link.split("blog.naver.com/")[-1].split("/")[0]
            if blog_id and blog_id != "example":  # 기본값 제외
                blogger_ids.append(blog_id)

    return blogger_ids


def process_keywords(keywords: List[str], debug: bool = False) -> Dict[str, List[str]]:
    """
    여러 키워드에 대해 블로그 아이디들을 수집
    """
    # API 키 확인
    client_id = os.getenv("NAVER_CLIENT_ID", "").strip()
    client_secret = os.getenv("NAVER_CLIENT_SECRET", "").strip()

    if not client_id or not client_secret:
        print("ERROR: .env 파일에 NAVER_CLIENT_ID / NAVER_CLIENT_SECRET이 설정되지 않았습니다.")
        return {}

    print(f"\n총 {len(keywords)}개의 키워드에 대해 블로그 아이디 수집을 시작합니다...")
    print("-" * 50)

    results = {}

    for i, keyword in enumerate(keywords, 1):
        print(f"\n[{i}/{len(keywords)}] '{keyword}' 키워드 처리 중...")

        try:
            # 메인 페이지 검색으로 블로거 정보 수집
            bloggers_data = get_top_bloggers_from_main_page(
                keyword=keyword,
                client_id=client_id,
                client_secret=client_secret,
                top_n=5,  # 최대 5개 블로그 아이디 수집
                debug=debug
            )

            # 블로그 아이디 추출
            blogger_ids = extract_blogger_ids(bloggers_data)

            results[keyword] = blogger_ids

            print(f"  → 수집된 블로그 아이디: {len(blogger_ids)}개")
            if blogger_ids:
                print(f"    {blogger_ids}")

        except Exception as e:
            print(f"  → 오류 발생: {e}")
            results[keyword] = []

    return results


def display_results(results: Dict[str, List[str]]) -> None:
    """
    수집된 결과를 보기 좋게 출력
    """
    print("\n" + "="*80)
    print("블로그 아이디 수집 결과")
    print("="*80)

    total_keywords = len(results)
    total_bloggers = sum(len(ids) for ids in results.values())

    print(f"\n총 키워드 수: {total_keywords}")
    print(f"총 블로그 아이디 수: {total_bloggers}")
    print("-" * 80)

    for keyword, blogger_ids in results.items():
        print(f"\n키워드: '{keyword}'")
        print(f"블로그 아이디 수: {len(blogger_ids)}")

        if blogger_ids:
            print("블로그 아이디들:")
            for i, blog_id in enumerate(blogger_ids, 1):
                print(f"  {i}. {blog_id}")
        else:
            print("  (수집된 블로그 아이디 없음)")

    print("\n" + "="*80)

    # 모든 블로그 아이디를 한 줄로 출력 (쉼표 구분)
    all_blogger_ids = []
    for blogger_ids in results.values():
        all_blogger_ids.extend(blogger_ids)

    if all_blogger_ids:
        print("모든 블로그 아이디 (쉼표 구분):")
        print(", ".join(all_blogger_ids))
        print("="*80)


def save_results_to_file(results: Dict[str, List[str]], filename: str = None) -> None:
    """
    결과를 파일로 저장
    """
    if not filename:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"blogger_ids_{timestamp}.txt"

    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write("네이버 메인 페이지 블로그 아이디 수집 결과\n")
            f.write(f"수집 시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("="*80 + "\n\n")

            for keyword, blogger_ids in results.items():
                f.write(f"키워드: {keyword}\n")
                f.write(f"블로그 아이디 수: {len(blogger_ids)}\n")

                if blogger_ids:
                    f.write("블로그 아이디들:\n")
                    for i, blog_id in enumerate(blogger_ids, 1):
                        f.write(f"  {i}. {blog_id}\n")
                else:
                    f.write("(수집된 블로그 아이디 없음)\n")

                f.write("\n")

            # 모든 블로그 아이디를 한 줄로 저장
            all_blogger_ids = []
            for blogger_ids in results.values():
                all_blogger_ids.extend(blogger_ids)

            if all_blogger_ids:
                f.write("="*80 + "\n")
                f.write("모든 블로그 아이디 (쉼표 구분):\n")
                f.write(", ".join(all_blogger_ids) + "\n")

        print(f"\n결과가 '{filename}' 파일로 저장되었습니다.")

    except Exception as e:
        print(f"파일 저장 중 오류 발생: {e}")


def main():
    """
    메인 함수
    """
    try:
        # 명령줄 인자 확인
        debug_mode = "--debug" in sys.argv
        save_file = "--save" in sys.argv

        # 여러 키워드 입력 받기
        keywords = get_multiple_keywords()

        if not keywords:
            print("키워드가 입력되지 않았습니다. 프로그램을 종료합니다.")
            return

        # 각 키워드에 대해 블로그 아이디 수집
        results = process_keywords(keywords, debug=debug_mode)

        # 결과 출력
        display_results(results)

        # 파일로 저장 옵션
        if save_file or input("\n결과를 파일로 저장하시겠습니까? (y/n): ").lower().startswith('y'):
            save_results_to_file(results)

        print("\n프로그램이 완료되었습니다!")

    except KeyboardInterrupt:
        print("\n\n프로그램이 사용자에 의해 중단되었습니다.")
    except Exception as e:
        print(f"\n오류 발생: {e}")
        if debug_mode:
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    main()