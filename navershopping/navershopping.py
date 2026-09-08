#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
네이버 쇼핑 검색 API를 사용하여 검색어로 상품을 검색하는 스크립트

필요 패키지 설치:
    pip install requests

사용법:
    python navershopping.py "검색어"
"""

import os
import sys
import requests
import json
from typing import Dict, List, Optional
from dotenv import load_dotenv
import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from io import BytesIO


class NaverShoppingSearch:
    """네이버 쇼핑 검색 API 클래스"""

    def __init__(self, client_id: str, client_secret: str):
        """
        네이버 쇼핑 검색 API 초기화

        Args:
            client_id (str): 네이버 개발자 센터에서 발급받은 Client ID
            client_secret (str): 네이버 개발자 센터에서 발급받은 Client Secret
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.base_url = "https://openapi.naver.com/v1/search/shop.json"

        # 세션 설정
        self.session = requests.Session()
        self.session.headers.update({
            'X-Naver-Client-Id': self.client_id,
            'X-Naver-Client-Secret': self.client_secret,
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })

    def search_products(self, query: str, display: int = 10, start: int = 1,
                       sort: str = "sim", filter_type: Optional[str] = None,
                       exclude: Optional[str] = None) -> Dict:
        """
        네이버 쇼핑 검색 API를 호출하여 상품 검색

        Args:
            query (str): 검색어 (UTF-8로 인코딩)
            display (int): 한 번에 표시할 검색 결과 개수 (기본값: 10, 최댓값: 100)
            start (int): 검색 시작 위치 (기본값: 1, 최댓값: 1000)
            sort (str): 정렬 방법 - sim(정확도순), date(날짜순), asc(가격오름차순), dsc(가격내림차순)
            filter_type (str, optional): 검색 결과에 포함할 상품 유형 - naverpay(네이버페이 연동 상품)
            exclude (str, optional): 제외할 상품 유형 - used(중고), rental(렌탈), cbshop(해외직구)

        Returns:
            dict: API 응답 결과를 담은 딕셔너리
        """
        # 파라미터 설정
        params = {
            'query': query,
            'display': min(display, 100),  # 최대 100개
            'start': min(start, 1000),     # 최대 1000페이지
            'sort': sort
        }

        # 선택적 파라미터 추가
        if filter_type:
            params['filter'] = filter_type
        if exclude:
            params['exclude'] = exclude

        try:
            # API 호출
            response = self.session.get(self.base_url, params=params)
            response.raise_for_status()  # HTTP 오류 발생 시 예외 발생

            return response.json()

        except requests.exceptions.RequestException as e:
            print(f"API 요청 중 오류 발생: {e}")
            return {'error': str(e)}

    def print_search_results(self, results: Dict) -> None:
        """
        검색 결과를 읽기 쉽게 출력

        Args:
            results (dict): API 응답 결과
        """
        if 'error' in results:
            print(f"오류: {results['error']}")
            return

        if 'items' not in results:
            print("검색 결과를 찾을 수 없습니다.")
            return

        items = results['items']
        total = results.get('total', 0)

        print(f"\n=== 네이버 쇼핑 검색 결과 ===")
        print(f"총 검색 결과: {total}개")
        print(f"표시된 결과: {len(items)}개")
        print("=" * 50)

        for i, item in enumerate(items, 1):
            print(f"\n{i}. {item.get('title', '제목 없음')}")
            print(f"   가격: {item.get('lprice', '가격 정보 없음')}원")
            print(f"   최저가: {item.get('hprice', '정보 없음')}원")
            print(f"   브랜드: {item.get('brand', '브랜드 정보 없음')}")
            print(f"   제조사: {item.get('maker', '제조사 정보 없음')}")
            print(f"   카테고리: {item.get('category1', '')} > {item.get('category2', '')} > {item.get('category3', '')}")
            print(f"   상품 링크: {item.get('link', '')}")
            print(f"   이미지: {item.get('image', '')}")

            # 몰 정보
            mall_name = item.get('mallName', '')
            if mall_name:
                print(f"   판매처: {mall_name}")

            # 상품 타입 정보
            product_type = item.get('productType', '')
            if product_type:
                print(f"   상품 타입: {product_type}")

    def search_multiple_pages(self, query: str, start_positions: List[int] = None,
                             display: int = 100) -> List[Dict]:
        """
        여러 페이지의 검색 결과를 수집

        Args:
            query (str): 검색어
            start_positions (list): 시작 위치 리스트 (기본값: [1, 101, 201, 301, 401])
            display (int): 한 번에 표시할 검색 결과 개수 (기본값: 100)

        Returns:
            list: 모든 페이지의 검색 결과를 담은 리스트
        """
        if start_positions is None:
            start_positions = [1, 101, 201, 301, 401]

        all_results = []

        for i, start in enumerate(start_positions, 1):
            print(f"페이지 {i}/{len(start_positions)} 검색 중... (start={start})")

            result = self.search_products(
                query=query,
                display=display,
                start=start
            )

            if 'error' in result:
                print(f"페이지 {start} 검색 실패: {result['error']}")
                continue

            if 'items' in result and result['items']:
                all_results.extend(result['items'])
                print(f"  ✓ {len(result['items'])}개 상품 수집")
            else:
                print(f"  ⚠ 페이지 {start}에 검색 결과가 없습니다.")
                break  # 더 이상 결과가 없으면 중단

            # API 호출 간격 조절 (선택사항)
            import time
            time.sleep(0.1)

        return all_results

    def save_to_excel(self, items: List[Dict], query: str, filename: Optional[str] = None) -> str:
        """
        검색 결과를 엑셀 파일로 저장

        Args:
            items (list): 검색 결과 아이템 리스트
            query (str): 검색어
            filename (str, optional): 파일명 (기본값: 자동 생성)

        Returns:
            str: 저장된 파일 경로
        """
        if not items:
            print("저장할 검색 결과가 없습니다.")
            return ""

        # result 폴더 생성 (스크립트 파일이 있는 디렉토리의 하위)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        result_dir = os.path.join(script_dir, 'result')
        os.makedirs(result_dir, exist_ok=True)

        # 파일명 생성 (지정되지 않은 경우)
        if not filename:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            safe_query = query.replace(" ", "_").replace("/", "_").replace("\\", "_")
            filename = f'naver_shopping_{safe_query}_{timestamp}.xlsx'

        filepath = os.path.join(result_dir, filename)

        # 데이터프레임 생성
        data = []
        for idx, item in enumerate(items, 1):
            row = {
                '순번': idx,
                '상품명': item.get('title', '').replace('<b>', '').replace('</b>', ''),
                '가격': item.get('lprice', ''),
                '최저가': item.get('hprice', ''),
                '브랜드': item.get('brand', ''),
                '제조사': item.get('maker', ''),
                '카테고리1': item.get('category1', ''),
                '카테고리2': item.get('category2', ''),
                '카테고리3': item.get('category3', ''),
                '카테고리4': item.get('category4', ''),
                '판매처': item.get('mallName', ''),
                '상품타입': item.get('productType', ''),
                '상품링크': item.get('link', ''),
                '이미지링크': item.get('image', ''),
                '검색어': query,
                '검색일시': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            data.append(row)

        # 데이터프레임 생성
        df = pd.DataFrame(data)

        # 가격 데이터 전처리 (숫자로 변환)
        df['가격_numeric'] = pd.to_numeric(df['가격'], errors='coerce')
        df = df.dropna(subset=['가격_numeric'])  # 가격 정보가 없는 행 제거

        # 가격대별 통계 계산
        price_stats = self._calculate_price_distribution(df)

        try:
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                # 검색 결과 시트
                df.drop('가격_numeric', axis=1).to_excel(writer, sheet_name='검색결과', index=False)

                # 열 너비 자동 조정 (검색결과 시트)
                worksheet = writer.sheets['검색결과']
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)  # 최대 50자
                    worksheet.column_dimensions[column_letter].width = adjusted_width

                # 가격분포 시트 (통계 데이터)
                price_df = pd.DataFrame(list(price_stats.items()), columns=['가격대', '상품수'])
                price_df.to_excel(writer, sheet_name='가격분포_통계', index=False)

                # 가격분포 시트 서식 설정
                stats_sheet = writer.sheets['가격분포_통계']
                stats_sheet.column_dimensions['A'].width = 15
                stats_sheet.column_dimensions['B'].width = 10

                # 그래프 생성 및 삽입
                self._create_price_chart(df, price_stats, writer)

            print(f"총 {len(items)}개의 검색 결과가 엑셀 파일로 저장되었습니다: {filepath}")
            print(f"가격대별 통계 및 그래프가 포함되었습니다.")
            return filepath

        except Exception as e:
            print(f"엑셀 파일 저장 중 오류 발생: {e}")
            return ""

    def _calculate_price_distribution(self, df: pd.DataFrame) -> Dict[str, int]:
        """
        가격대별 상품 분포 계산

        Args:
            df (pd.DataFrame): 상품 데이터프레임

        Returns:
            dict: 가격대별 상품 수
        """
        # 가격대 구간 설정
        bins = [0, 10000, 30000, 50000, 100000, 200000, 500000, float('inf')]
        labels = ['~10,000원', '10,001~30,000원', '30,001~50,000원',
                 '50,001~100,000원', '100,001~200,000원', '200,001~500,000원', '500,000원~']

        # 가격대별로 그룹화
        df['가격대'] = pd.cut(df['가격_numeric'], bins=bins, labels=labels, right=False)

        # 각 가격대별 상품 수 계산
        distribution = df['가격대'].value_counts().sort_index()

        return distribution.to_dict()

    def _create_price_chart(self, df: pd.DataFrame, price_stats: Dict[str, int], writer):
        """
        가격대별 분포 그래프 생성 및 엑셀에 삽입

        Args:
            df (pd.DataFrame): 상품 데이터프레임
            price_stats (dict): 가격대별 통계
            writer: ExcelWriter 객체
        """
        try:
            # matplotlib 설정
            plt.rcParams['font.family'] = 'Malgun Gothic'  # 한글 폰트 설정
            plt.rcParams['axes.unicode_minus'] = False

            # 그래프 생성
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))

            # 1. 막대 그래프
            price_ranges = list(price_stats.keys())
            counts = list(price_stats.values())

            bars = ax1.bar(range(len(price_ranges)), counts, color='skyblue', alpha=0.7)
            ax1.set_title(f'가격대별 상품 분포 (총 {len(df)}개)', fontsize=14, fontweight='bold')
            ax1.set_xlabel('가격대', fontsize=12)
            ax1.set_ylabel('상품 수', fontsize=12)
            ax1.set_xticks(range(len(price_ranges)))
            ax1.set_xticklabels(price_ranges, rotation=45, ha='right')
            ax1.grid(axis='y', alpha=0.3)

            # 막대 위에 값 표시
            for bar, count in zip(bars, counts):
                ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                        f'{count}', ha='center', va='bottom', fontsize=10)

            # 2. 원형 그래프
            ax2.pie(counts, labels=price_ranges, autopct='%1.1f%%', startangle=90)
            ax2.set_title(f'가격대별 비율 (총 {len(df)}개)', fontsize=14, fontweight='bold')
            ax2.axis('equal')  # 원형 유지

            plt.tight_layout()

            # 그래프를 메모리에 저장
            img_buffer = BytesIO()
            plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
            img_buffer.seek(0)
            plt.close()

            # 엑셀에 이미지 삽입
            from openpyxl.drawing.image import Image
            chart_sheet = writer.book.create_sheet('가격분포_그래프')
            img = Image(img_buffer)
            img.width = 800  # 픽셀 단위
            img.height = 400
            chart_sheet.add_image(img, 'A1')

            # 시트 설명 추가
            chart_sheet['A20'] = f"검색어: {df['검색어'].iloc[0] if not df.empty else 'N/A'}"
            chart_sheet['A21'] = f"총 상품 수: {len(df)}개"
            chart_sheet['A22'] = f"생성 일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"

        except Exception as e:
            print(f"그래프 생성 중 오류 발생: {e}")
            # 그래프 생성 실패 시 통계 시트만 유지


def load_api_keys() -> tuple[str, str]:
    """
    .env 파일, 환경변수 또는 설정 파일에서 API 키를 로드

    Returns:
        tuple: (client_id, client_secret)
    """
    # .env 파일 로드 시도
    load_dotenv()

    # 환경변수에서 로드 시도
    client_id = os.getenv('NAVER_CLIENT_ID')
    client_secret = os.getenv('NAVER_CLIENT_SECRET')

    # 환경변수가 없으면 설정 파일에서 로드 시도
    if not client_id or not client_secret:
        config_file = 'naver_api_config.txt'
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                    for line in lines:
                        if line.startswith('CLIENT_ID='):
                            client_id = line.split('=', 1)[1].strip()
                        elif line.startswith('CLIENT_SECRET='):
                            client_secret = line.split('=', 1)[1].strip()
            except Exception as e:
                print(f"설정 파일 읽기 오류: {e}")

    if not client_id or not client_secret:
        print("네이버 API 키가 설정되지 않았습니다.")
        print("다음 방법 중 하나를 사용하세요:")
        print("1. .env 파일 생성:")
        print("   NAVER_CLIENT_ID=your_client_id")
        print("   NAVER_CLIENT_SECRET=your_client_secret")
        print("2. 환경변수 설정: NAVER_CLIENT_ID, NAVER_CLIENT_SECRET")
        print("3. naver_api_config.txt 파일 생성:")
        print("   CLIENT_ID=your_client_id")
        print("   CLIENT_SECRET=your_client_secret")
        sys.exit(1)

    return client_id, client_secret


def main():
    """메인 함수"""
    # 검색어 설정
    if len(sys.argv) >= 2:
        # 명령줄 인자로 검색어가 제공된 경우
        query = sys.argv[1]
    else:
        # 명령줄 인자가 없으면 사용자 입력 받기
        try:
            query = input("검색어를 입력하세요: ").strip()
            if not query:
                print("검색어가 입력되지 않았습니다.")
                sys.exit(1)
        except (EOFError, KeyboardInterrupt):
            print("\n프로그램을 종료합니다.")
            sys.exit(1)

    # API 키 로드
    client_id, client_secret = load_api_keys()

    # 네이버 쇼핑 검색 객체 생성
    searcher = NaverShoppingSearch(client_id, client_secret)

    # 여러 페이지 검색 실행 (display=100, start=1,101,201,301,401)
    print(f"'{query}' 검색 중... (총 5페이지, 각 100개씩)")

    # 시작 위치 설정: 1, 101, 201, 301, 401
    start_positions = [1, 101, 201, 301, 401]

    # 여러 페이지 검색
    all_items = searcher.search_multiple_pages(
        query=query,
        start_positions=start_positions,
        display=100
    )

    if not all_items:
        print("검색 결과를 찾을 수 없습니다.")
        sys.exit(1)

    print(f"\n총 {len(all_items)}개의 상품을 찾았습니다.")

    # 첫 번째 결과만 콘솔에 출력 (선택사항)
    if len(all_items) > 0:
        print("\n=== 첫 번째 상품 샘플 ===")
        sample_result = {'items': [all_items[0]]}
        searcher.print_search_results(sample_result)
        if len(all_items) > 1:
            print(f"... 외 {len(all_items) - 1}개 상품")

    # 엑셀 파일로 저장 (모든 결과)
    excel_file = searcher.save_to_excel(all_items, query)

    if excel_file:
        print(f"\n✓ 검색 결과가 저장되었습니다: {excel_file}")


if __name__ == "__main__":
    # matplotlib 및 패키지 작동 확인
    try:
        import matplotlib.pyplot as plt
        import matplotlib
        print(f"✓ matplotlib {matplotlib.__version__} 작동 중")
    except ImportError as e:
        print(f"✗ matplotlib 오류: {e}")
        exit(1)

    try:
        import pandas as pd
        print(f"✓ pandas {pd.__version__} 작동 중")
    except ImportError as e:
        print(f"✗ pandas 오류: {e}")
        exit(1)

    print("🎉 모든 패키지가 정상적으로 로드되었습니다!")

    # 실제 메인 함수 실행
    main()