"""
H_IntegrationGUI.py - 네이버 키워드 분석 통합 GUI
A~G 모듈을 순차적으로 실행하는 통합 인터페이스

실행 순서:
A_rightside → B_autocomplete → C_sumKeyword → D_searchresult → E_deleteAndPriority → F_add_recent30days → G_add_blogger_by_mainPage
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import sys
import os
import logging
from datetime import datetime
from typing import List, Optional

# 모듈 import
import A_rightside
import B_autocomplete
import C_sumKeyword
import D_searchresult
import E_deleteAndPriority
import F_add_recent30days
import G_add_blogger_by_mainPage


class IntegrationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("네이버 키워드 분석 통합 시스템 v1.0")
        self.root.geometry("900x700")
        self.root.resizable(True, True)

        # 실행 상태 관리
        self.is_running = False
        self.stop_requested = False

        # 로깅 설정
        self.setup_logging()

        # GUI 구성
        self.create_widgets()

    def setup_logging(self):
        """로깅 설정"""
        self.log_text = None  # GUI 생성 후 설정됨

        # 파일 로깅
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_filename = f"result/logs/integration_{timestamp}.log"

        os.makedirs("result/logs", exist_ok=True)

        self.file_logger = logging.getLogger("integration_file")
        self.file_logger.setLevel(logging.INFO)

        file_handler = logging.FileHandler(log_filename, encoding='utf-8')
        file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(file_formatter)
        self.file_logger.addHandler(file_handler)

    def log(self, message: str, level: str = "INFO"):
        """로그 출력"""
        timestamp = datetime.now().strftime("%H:%M:%S")

        # 콘솔 출력
        print(f"[{timestamp}] {message}")

        # 파일 로그
        if level == "INFO":
            self.file_logger.info(message)
        elif level == "ERROR":
            self.file_logger.error(message)
        elif level == "WARNING":
            self.file_logger.warning(message)

        # GUI 로그 (메인 스레드에서 실행)
        if self.log_text:
            self.root.after(0, lambda: self.update_log_text(f"[{timestamp}] {message}\n"))

    def update_log_text(self, text: str):
        """GUI 로그 텍스트 업데이트"""
        if self.log_text:
            self.log_text.insert(tk.END, text)
            self.log_text.see(tk.END)

    def create_widgets(self):
        """GUI 위젯 생성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 타이틀
        title_label = ttk.Label(main_frame, text="네이버 키워드 분석 통합 시스템",
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 20))

        # 입력 프레임
        input_frame = ttk.LabelFrame(main_frame, text="초기 설정", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 10))

        # 시드 키워드 입력
        ttk.Label(input_frame, text="시드 키워드 (쉼표로 구분):").pack(anchor=tk.W)
        self.seed_keywords_var = tk.StringVar()
        seed_entry = ttk.Entry(input_frame, textvariable=self.seed_keywords_var, width=80)
        seed_entry.pack(fill=tk.X, pady=(0, 10))
        seed_entry.insert(0, "오키나와렌트카,오키나와여행")  # 기본값

        # 필수 키워드 입력
        ttk.Label(input_frame, text="필수 키워드 (쉼표로 구분, 선택사항):").pack(anchor=tk.W)
        self.required_keywords_var = tk.StringVar()
        required_entry = ttk.Entry(input_frame, textvariable=self.required_keywords_var, width=80)
        required_entry.pack(fill=tk.X, pady=(0, 10))
        required_entry.insert(0, "오키나와")  # 기본값

        # 진행률 표시
        progress_frame = ttk.LabelFrame(main_frame, text="진행 상황", padding="10")
        progress_frame.pack(fill=tk.X, pady=(0, 10))

        # 진행률 바
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var,
                                          maximum=100, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))

        # 진행 상태 레이블
        self.status_label = ttk.Label(progress_frame, text="준비 완료")
        self.status_label.pack(anchor=tk.W)

        # 현재 단계 표시
        self.current_step_label = ttk.Label(progress_frame, text="")
        self.current_step_label.pack(anchor=tk.W)

        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))

        # 시작 버튼
        self.start_button = ttk.Button(button_frame, text="분석 시작",
                                     command=self.start_analysis, style="Accent.TButton")
        self.start_button.pack(side=tk.LEFT, padx=(0, 10))

        # 중단 버튼
        self.stop_button = ttk.Button(button_frame, text="중단",
                                    command=self.stop_analysis, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT)

        # 로그 프레임
        log_frame = ttk.LabelFrame(main_frame, text="실행 로그", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)

        # 로그 텍스트 영역
        self.log_text = scrolledtext.ScrolledText(log_frame, height=20, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 버튼 스타일 설정
        style = ttk.Style()
        style.configure("Accent.TButton", font=("Arial", 10, "bold"))

    def parse_keywords(self, keyword_string: str) -> List[str]:
        """키워드 문자열을 리스트로 변환"""
        if not keyword_string.strip():
            return []

        keywords = []
        for kw in keyword_string.split(","):
            kw = kw.strip()
            if kw:
                keywords.append(kw)
        return keywords

    def start_analysis(self):
        """분석 시작"""
        if self.is_running:
            return

        # 입력 검증
        seed_keywords = self.parse_keywords(self.seed_keywords_var.get())
        if not seed_keywords:
            messagebox.showerror("입력 오류", "시드 키워드를 최소 1개 이상 입력해주세요.")
            return

        # 상태 변경
        self.is_running = True
        self.stop_requested = False
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

        # 로그 초기화
        self.log_text.delete(1.0, tk.END)
        self.log("=== 네이버 키워드 분석 통합 시스템 시작 ===")
        self.log(f"시드 키워드: {seed_keywords}")

        required_keywords = self.parse_keywords(self.required_keywords_var.get())
        if required_keywords:
            self.log(f"필수 키워드: {required_keywords}")
        else:
            self.log("필수 키워드: 없음 (모든 데이터 유지)")

        # 백그라운드에서 분석 실행
        analysis_thread = threading.Thread(target=self.run_analysis,
                                         args=(seed_keywords, required_keywords))
        analysis_thread.daemon = True
        analysis_thread.start()

    def stop_analysis(self):
        """분석 중단"""
        if not self.is_running:
            return

        self.stop_requested = True
        self.log("사용자가 분석 중단을 요청했습니다. 현재 단계 완료 후 중단됩니다...")
        self.status_label.config(text="중단 요청됨")

    def run_analysis(self, seed_keywords: List[str], required_keywords: List[str]):
        """분석 실행"""
        try:
            total_steps = 7
            current_step = 0

            # A 단계: 우측 연관 검색어 추출
            current_step += 1
            self.update_progress(current_step, total_steps, "A단계: 우측 연관 검색어 추출")
            if not self.run_module_a(seed_keywords):
                return

            # B 단계: 자동완성 검색어 추출
            current_step += 1
            self.update_progress(current_step, total_steps, "B단계: 자동완성 검색어 추출")
            if not self.run_module_b(seed_keywords):
                return

            # C 단계: 키워드 통합
            current_step += 1
            self.update_progress(current_step, total_steps, "C단계: 키워드 데이터 통합")
            if not self.run_module_c(required_keywords):
                return

            # D 단계: 검색 결과 처리
            current_step += 1
            self.update_progress(current_step, total_steps, "D단계: 검색 결과 처리")
            if not self.run_module_d():
                return

            # E 단계: 데이터 필터링
            current_step += 1
            self.update_progress(current_step, total_steps, "E단계: 데이터 필터링 및 우선순위")
            if not self.run_module_e():
                return

            # F 단계: 최근 30일 데이터 추가
            current_step += 1
            self.update_progress(current_step, total_steps, "F단계: 최근 30일 블로그 데이터 추가")
            if not self.run_module_f():
                return

            # G 단계: 블로거 정보 관리
            current_step += 1
            self.update_progress(current_step, total_steps, "G단계: 블로거 정보 관리")
            if not self.run_module_g():
                return

            # 완료
            self.update_progress(total_steps, total_steps, "모든 단계 완료!")
            self.log("=== 모든 분석 단계가 성공적으로 완료되었습니다! ===")

            # 완료 메시지
            self.root.after(0, lambda: messagebox.showinfo("완료",
                "네이버 키워드 분석이 성공적으로 완료되었습니다!\n\n"
                "결과 파일: result/keywordList_all.xlsx"))

        except Exception as e:
            self.log(f"분석 중 오류 발생: {str(e)}", "ERROR")
            self.root.after(0, lambda: messagebox.showerror("오류",
                f"분석 중 오류가 발생했습니다:\n{str(e)}"))

        finally:
            # 상태 초기화
            self.root.after(0, self.reset_ui)

    def update_progress(self, current: int, total: int, status: str):
        """진행률 업데이트"""
        progress = (current / total) * 100
        self.root.after(0, lambda: self.progress_var.set(progress))
        self.root.after(0, lambda: self.status_label.config(text=f"{current}/{total} 단계 진행중"))
        self.root.after(0, lambda: self.current_step_label.config(text=status))
        self.log(f"진행률: {current}/{total} - {status}")

    def run_module_a(self, seed_keywords: List[str]) -> bool:
        """A 모듈 실행"""
        try:
            self.log("A_rightside 모듈 실행 시작...")

            # monkey patch로 입력 우회
            original_input = __builtins__.input

            def mock_input(prompt):
                if "seed >" in prompt:
                    # 첫 번째 호출에서 모든 키워드 반환
                    mock_input.call_count += 1
                    if mock_input.call_count == 1:
                        return ", ".join(seed_keywords)
                    else:
                        return ""  # 빈 입력으로 종료
                return original_input(prompt)

            mock_input.call_count = 0
            __builtins__.input = mock_input

            try:
                A_rightside.main()
                self.log("A_rightside 모듈 실행 완료")
                return True
            finally:
                __builtins__.input = original_input

        except Exception as e:
            self.log(f"A_rightside 모듈 실행 실패: {str(e)}", "ERROR")
            return False

    def run_module_b(self, seed_keywords: List[str]) -> bool:
        """B 모듈 실행"""
        try:
            self.log("B_autocomplete 모듈 실행 시작...")

            # monkey patch로 입력 우회
            original_input = __builtins__.input

            def mock_input(prompt):
                if "seed >" in prompt:
                    mock_input.call_count += 1
                    if mock_input.call_count == 1:
                        return ", ".join(seed_keywords)
                    else:
                        return ""
                return original_input(prompt)

            mock_input.call_count = 0
            __builtins__.input = mock_input

            try:
                B_autocomplete.main()
                self.log("B_autocomplete 모듈 실행 완료")
                return True
            finally:
                __builtins__.input = original_input

        except Exception as e:
            self.log(f"B_autocomplete 모듈 실행 실패: {str(e)}", "ERROR")
            return False

    def run_module_c(self, required_keywords: List[str]) -> bool:
        """C 모듈 실행"""
        try:
            self.log("C_sumKeyword 모듈 실행 시작...")

            # monkey patch로 입력 우회
            original_input = __builtins__.input

            def mock_input(prompt):
                if "필수 키워드 >" in prompt:
                    return ", ".join(required_keywords)
                return original_input(prompt)

            __builtins__.input = mock_input

            try:
                C_sumKeyword.main()
                self.log("C_sumKeyword 모듈 실행 완료")
                return True
            finally:
                __builtins__.input = original_input

        except Exception as e:
            self.log(f"C_sumKeyword 모듈 실행 실패: {str(e)}", "ERROR")
            return False

    def run_module_d(self) -> bool:
        """D 모듈 실행"""
        try:
            self.log("D_searchresult 모듈 실행 시작...")
            D_searchresult.main()
            self.log("D_searchresult 모듈 실행 완료")
            return True
        except Exception as e:
            self.log(f"D_searchresult 모듈 실행 실패: {str(e)}", "ERROR")
            return False

    def run_module_e(self) -> bool:
        """E 모듈 실행"""
        try:
            self.log("E_deleteAndPriority 모듈 실행 시작...")
            E_deleteAndPriority.main()
            self.log("E_deleteAndPriority 모듈 실행 완료")
            return True
        except Exception as e:
            self.log(f"E_deleteAndPriority 모듈 실행 실패: {str(e)}", "ERROR")
            return False

    def run_module_f(self) -> bool:
        """F 모듈 실행"""
        try:
            self.log("F_add_recent30days 모듈 실행 시작...")
            F_add_recent30days.main()
            self.log("F_add_recent30days 모듈 실행 완료")
            return True
        except Exception as e:
            self.log(f"F_add_recent30days 모듈 실행 실패: {str(e)}", "ERROR")
            return False

    def run_module_g(self) -> bool:
        """G 모듈 실행"""
        try:
            self.log("G_add_blogger_by_mainPage 모듈 실행 시작...")

            # 명령줄 인자로 첫 번째 시드 키워드 전달
            original_argv = sys.argv.copy()
            sys.argv = ["G_add_blogger_by_mainPage.py", seed_keywords[0]]

            try:
                G_add_blogger_by_mainPage.main()
                self.log("G_add_blogger_by_mainPage 모듈 실행 완료")
                return True
            finally:
                sys.argv = original_argv

        except Exception as e:
            self.log(f"G_add_blogger_by_mainPage 모듈 실행 실패: {str(e)}", "ERROR")
            return False

    def reset_ui(self):
        """UI 상태 초기화"""
        self.is_running = False
        self.stop_requested = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.status_label.config(text="준비 완료")
        self.current_step_label.config(text="")


def main():
    """메인 함수"""
    root = tk.Tk()
    app = IntegrationGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()