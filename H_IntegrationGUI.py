"""
H_IntegrationGUI.py - 네이버 키워드 분석 통합 GUI
A~G 모듈을 순차적으로 실행하는 통합 인터페이스

실행 순서:
A_rightside → B_autocomplete → C_sumKeyword → D_searchresult → E_deleteAndPriority → F_add_recent30days → F_sort_result_1 → G_add_blogger_by_mainPage
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import sys
import os
from typing import List, Optional

# 모듈 import
import A_rightside
import B_autocomplete
import C_sumKeyword
import D_searchresult
import E_deleteAndPriority
import F_add_recent30days
import F_sort_result_1
import G_add_blogger_by_mainPage


class MockInput:
    """GUI에서 모듈 input을 모의하는 클래스"""
    def __init__(self, seed_keyword: str = "", required_keywords: List[str] = None):
        self.seed_keyword = seed_keyword  # 단일 키워드
        self.required_keywords = required_keywords or []
        self.seed_call_count = 0
        self.required_call_count = 0

    def __call__(self, prompt):
        if "seed >" in prompt:
            self.seed_call_count += 1
            if self.seed_call_count == 1:
                # 첫 번째 호출: 단일 시드 키워드 반환
                return self.seed_keyword
            else:
                # 두 번째 호출: 빈 입력으로 종료
                return ""
        elif "필수 키워드 >" in prompt:
            self.required_call_count += 1
            if self.required_call_count == 1:
                # 필수 키워드들 반환
                return ", ".join(self.required_keywords)
        return input(prompt)  # 실제 input으로 폴백


class Logger:
    """GUI 텍스트 위젯과 파일에 동시에 출력하는 로거 클래스"""
    def __init__(self, text_widget=None, log_file="log.txt"):
        self.text_widget = text_widget
        self.log_file = log_file
        self.buffer = ""

        # 로그 파일 초기화
        try:
            with open(self.log_file, 'w', encoding='utf-8') as f:
                f.write("=== 네이버 키워드 분석 로그 시작 ===\n")
        except Exception as e:
            print(f"로그 파일 초기화 실패: {e}")

    def write(self, text):
        """텍스트를 GUI와 파일에 동시에 출력"""
        # 버퍼에 추가
        self.buffer += text

        # 모든 완전한 줄 처리
        lines = self.buffer.split('\n')
        self.buffer = lines.pop() if lines and not self.buffer.endswith('\n') else ''

        # 완전한 줄들을 처리
        for line in lines:
            line += '\n'

            # GUI에 출력 (스레드 안전하게)
            if self.text_widget:
                self.text_widget.after(0, lambda l=line: self._append_to_text_widget(l))

            # 파일에 출력
            try:
                with open(self.log_file, 'a', encoding='utf-8') as f:
                    f.write(line)
            except Exception as e:
                # 파일 쓰기 실패 시 GUI에 에러 표시
                if self.text_widget:
                    error_msg = f"[로그 파일 쓰기 오류: {e}]\n"
                    self.text_widget.after(0, lambda: self._append_to_text_widget(error_msg))

    def flush(self):
        """버퍼 비우기"""
        if self.buffer:
            # 남은 버퍼 내용 처리
            if self.text_widget:
                self.text_widget.after(0, lambda: self._append_to_text_widget(self.buffer))

            # 파일에 출력
            try:
                with open(self.log_file, 'a', encoding='utf-8') as f:
                    f.write(self.buffer)
            except Exception as e:
                if self.text_widget:
                    error_msg = f"[로그 파일 쓰기 오류: {e}]\n"
                    self.text_widget.after(0, lambda: self._append_to_text_widget(error_msg))

        self.buffer = ""

    def _append_to_text_widget(self, text):
        """GUI 텍스트 위젯에 텍스트 추가 (메인 스레드에서 실행)"""
        if self.text_widget:
            self.text_widget.insert(tk.END, text)
            self.text_widget.see(tk.END)  # 자동 스크롤


class IntegrationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("네이버 키워드 분석 통합 시스템 v1.0")
        self.root.geometry("900x700")
        self.root.resizable(True, True)

        # 실행 상태 관리
        self.is_running = False
        self.stop_requested = False

        # GUI 구성
        self.create_widgets()

        # Logger 초기화 및 sys.stdout 대체
        self.logger = Logger(text_widget=self.log_text)
        self.original_stdout = sys.stdout
        sys.stdout = self.logger


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

        # 필수 키워드 입력
        ttk.Label(input_frame, text="필수 키워드 (쉼표로 구분, 선택사항):").pack(anchor=tk.W)
        self.required_keywords_var = tk.StringVar()
        required_entry = ttk.Entry(input_frame, textvariable=self.required_keywords_var, width=80)
        required_entry.pack(fill=tk.X, pady=(0, 10))

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

        # 로그 표시 영역
        log_frame = ttk.LabelFrame(main_frame, text="실행 로그", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 스크롤 가능한 텍스트 위젯
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=15,
                                                 font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))

        # 통합 실행 버튼들
        integrated_frame = ttk.LabelFrame(button_frame, text="통합 실행", padding="5")
        integrated_frame.pack(side=tk.LEFT, padx=(0, 10))

        # 시작 버튼
        self.start_button = ttk.Button(integrated_frame, text="전체 분석 시작",
                                     command=self.start_analysis, style="Accent.TButton")
        self.start_button.pack(side=tk.LEFT, padx=(0, 5))

        # 중단 버튼
        self.stop_button = ttk.Button(integrated_frame, text="중단",
                                    command=self.stop_analysis, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT)

        # 개별 실행 버튼들
        individual_frame = ttk.LabelFrame(button_frame, text="개별 모듈 실행", padding="5")
        individual_frame.pack(side=tk.LEFT)

        # A~G 버튼들을 3줄로 배치
        row1_frame = ttk.Frame(individual_frame)
        row1_frame.pack(fill=tk.X, pady=(0, 2))

        row2_frame = ttk.Frame(individual_frame)
        row2_frame.pack(fill=tk.X, pady=(0, 2))

        row3_frame = ttk.Frame(individual_frame)
        row3_frame.pack(fill=tk.X)

        # 1줄: A, B, C
        ttk.Button(row1_frame, text="A (우측연관)", width=10,
                  command=lambda: self.run_individual_module('A')).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(row1_frame, text="B (자동완성)", width=10,
                  command=lambda: self.run_individual_module('B')).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(row1_frame, text="C (키워드통합)", width=10,
                  command=lambda: self.run_individual_module('C')).pack(side=tk.LEFT)

        # 2줄: D, E, F
        ttk.Button(row2_frame, text="D (검색결과)", width=10,
                  command=lambda: self.run_individual_module('D')).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(row2_frame, text="E (필터링)", width=10,
                  command=lambda: self.run_individual_module('E')).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(row2_frame, text="F (최근데이터)", width=10,
                  command=lambda: self.run_individual_module('F')).pack(side=tk.LEFT)

        # 3줄: F_sort, G
        ttk.Button(row3_frame, text="F_sort (정렬)", width=10,
                  command=lambda: self.run_individual_module('F_sort')).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(row3_frame, text="G (블로거)", width=10,
                  command=lambda: self.run_individual_module('G')).pack(side=tk.LEFT)


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

        required_keywords = self.parse_keywords(self.required_keywords_var.get())

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
        self.status_label.config(text="중단 요청됨")

    def run_individual_module(self, module_name: str):
        """개별 모듈 실행"""
        if self.is_running:
            messagebox.showwarning("경고", "현재 다른 작업이 실행 중입니다.")
            return

        # 입력 검증
        seed_keywords = self.parse_keywords(self.seed_keywords_var.get())
        required_keywords = self.parse_keywords(self.required_keywords_var.get())

        # 모듈별 입력 요구사항 확인
        if module_name in ['A', 'B', 'G'] and not seed_keywords:
            messagebox.showerror("입력 오류", f"{module_name} 모듈은 시드 키워드가 필요합니다.")
            return
        if module_name == 'C' and not required_keywords:
            messagebox.showwarning("입력 확인", "C 모듈에 필수 키워드가 설정되지 않았습니다.\n모든 데이터를 유지합니다.")

        # 실행 시작
        self.is_running = True
        self.stop_requested = False
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

        # 백그라운드에서 모듈 실행
        module_thread = threading.Thread(target=self.run_single_module,
                                       args=(module_name, seed_keywords, required_keywords))
        module_thread.daemon = True
        module_thread.start()

    def run_single_module(self, module_name: str, seed_keywords: List[str], required_keywords: List[str]):
        """단일 모듈 실행"""
        try:
            module_map = {
                'A': ('A_rightside', lambda: self.run_module_a(seed_keywords[0] if seed_keywords else "")),
                'B': ('B_autocomplete', lambda: self.run_module_b(seed_keywords[0] if seed_keywords else "")),
                'C': ('C_sumKeyword', lambda: self.run_module_c(required_keywords)),
                'D': ('D_searchresult', lambda: self.run_module_d()),
                'E': ('E_deleteAndPriority', lambda: self.run_module_e()),
                'F': ('F_add_recent30days', lambda: self.run_module_f()),
                'F_sort': ('F_sort_result_1', lambda: self.run_module_f_sort()),
                'G': ('G_add_blogger_by_mainPage', lambda: self.run_module_g(seed_keywords[0] if seed_keywords else ""))
            }

            module_title, module_func = module_map[module_name]

            self.update_progress(0, 1, f"{module_name}단계: {module_title} 실행중")

            success = module_func()

            if success:
                self.update_progress(1, 1, f"{module_name}단계 완료")

                success_msg = f"{module_name} 모듈이 성공적으로 실행되었습니다!"
                messagebox.showinfo("완료", success_msg)
            else:
                messagebox.showerror("실패", f"{module_name} 모듈 실행에 실패했습니다.")

        except Exception as e:
            messagebox.showerror("오류", f"{module_name} 모듈 실행 중 오류가 발생했습니다:\n{str(e)}")

        finally:
            # UI 상태 초기화
            self.root.after(0, self.reset_ui)

    def run_analysis(self, seed_keywords: List[str], required_keywords: List[str]):
        """분석 실행 - 각 키워드별로 A~G 단계 순차 실행"""
        try:
            total_keywords = len(seed_keywords)
            total_steps_per_keyword = 8
            total_steps = total_keywords * total_steps_per_keyword
            current_step = 0

            for keyword_idx, seed_keyword in enumerate(seed_keywords):
                keyword_display = f"[{keyword_idx + 1}/{total_keywords}] {seed_keyword}"

                # A 단계: 우측 연관 검색어 추출
                current_step += 1
                self.update_progress(current_step, total_steps, f"{keyword_display} - A단계: 우측 연관 검색어 추출")
                if not self.run_module_a(seed_keyword):
                    return

                # B 단계: 자동완성 검색어 추출
                current_step += 1
                self.update_progress(current_step, total_steps, f"{keyword_display} - B단계: 자동완성 검색어 추출")
                if not self.run_module_b(seed_keyword):
                    return

                # C 단계: 키워드 통합
                current_step += 1
                self.update_progress(current_step, total_steps, f"{keyword_display} - C단계: 키워드 데이터 통합")
                if not self.run_module_c(required_keywords):
                    return

                # D 단계: 검색 결과 처리
                current_step += 1
                self.update_progress(current_step, total_steps, f"{keyword_display} - D단계: 검색 결과 처리")
                if not self.run_module_d():
                    return

                # E 단계: 데이터 필터링
                current_step += 1
                self.update_progress(current_step, total_steps, f"{keyword_display} - E단계: 데이터 필터링 및 우선순위")
                if not self.run_module_e():
                    return

                # F 단계: 최근 30일 데이터 추가
                current_step += 1
                self.update_progress(current_step, total_steps, f"{keyword_display} - F단계: 최근 30일 블로그 데이터 추가")
                if not self.run_module_f():
                    return

                # F_sort 단계: 최근 30일 데이터 정렬 및 분석
                current_step += 1
                self.update_progress(current_step, total_steps, f"{keyword_display} - F_sort단계: 최근 30일 데이터 정렬 및 분석")
                if not self.run_module_f_sort():
                    return

                # G 단계: 블로거 정보 관리
                current_step += 1
                self.update_progress(current_step, total_steps, f"{keyword_display} - G단계: 블로거 정보 관리")
                if not self.run_module_g(seed_keyword):
                    return

                # 각 키워드 완료 메시지
                self.root.after(0, lambda k=seed_keyword: messagebox.showinfo("키워드 완료",
                    f"'{k}' 키워드 분석이 완료되었습니다!\n\n"
                    f"결과 파일: keywordResult_{k}.xlsx"))

            # 모든 키워드 완료
            self.update_progress(total_steps, total_steps, "모든 키워드 분석 완료!")

            # 최종 완료 메시지
            self.root.after(0, lambda: messagebox.showinfo("완료",
                f"총 {total_keywords}개 키워드 분석이 성공적으로 완료되었습니다!\n\n"
                "결과 파일들은 프로젝트 루트에 keywordResult_*.xlsx 형태로 저장되었습니다."))

        except Exception as e:
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

    def run_module_a(self, seed_keyword: str) -> bool:
        """A 모듈 실행"""
        if self.stop_requested:
            return False

        try:
            # monkey patch로 입력 우회
            original_input = __builtins__.input
            mock_input = MockInput(seed_keyword)
            __builtins__.input = mock_input

            try:
                A_rightside.main()
                return True
            finally:
                __builtins__.input = original_input

        except KeyboardInterrupt:
            return False
        except Exception as e:
            return False

    def run_module_b(self, seed_keyword: str) -> bool:
        """B 모듈 실행"""
        if self.stop_requested:
            return False

        # Selenium 사용 가능 여부 확인
        if not hasattr(B_autocomplete, 'SELENIUM_AVAILABLE') or not B_autocomplete.SELENIUM_AVAILABLE:
            return True  # 실패로 처리하지 않고 건너뛰기

        try:
            # monkey patch로 입력 우회
            original_input = __builtins__.input
            mock_input = MockInput(seed_keyword)
            __builtins__.input = mock_input

            try:
                B_autocomplete.main()
                return True
            finally:
                __builtins__.input = original_input

        except KeyboardInterrupt:
            return False
        except Exception as e:
            return False

    def run_module_c(self, required_keywords: List[str]) -> bool:
        """C 모듈 실행"""
        if self.stop_requested:
            return False

        try:
            # monkey patch로 입력 우회
            original_input = __builtins__.input
            mock_input = MockInput(required_keywords=required_keywords)
            __builtins__.input = mock_input

            try:
                C_sumKeyword.main()
                return True
            finally:
                __builtins__.input = original_input

        except KeyboardInterrupt:
            return False
        except Exception as e:
            return False

    def run_module_d(self) -> bool:
        """D 모듈 실행"""
        if self.stop_requested:
            return False

        try:
            D_searchresult.main()
            return True
        except Exception as e:
            return False

    def run_module_e(self) -> bool:
        """E 모듈 실행"""
        if self.stop_requested:
            return False

        try:
            E_deleteAndPriority.main()
            return True
        except Exception as e:
            return False

    def run_module_f(self) -> bool:
        """F 모듈 실행"""
        if self.stop_requested:
            return False

        try:
            F_add_recent30days.main()
            return True
        except Exception as e:
            return False

    def run_module_f_sort(self) -> bool:
        """F_sort 모듈 실행"""
        if self.stop_requested:
            return False

        try:
            F_sort_result_1.main()
            return True
        except Exception as e:
            return False

    def run_module_g(self, seed_keyword: str) -> bool:
        """G 모듈 실행"""
        if self.stop_requested:
            return False

        try:
            # 명령줄 인자로 시드 키워드 전달
            original_argv = sys.argv.copy()
            sys.argv = ["G_add_blogger_by_mainPage.py", seed_keyword]

            try:
                G_add_blogger_by_mainPage.main()

                # G 단계 완료 후 파일명 변경 및 기존 파일 삭제
                self.rename_and_cleanup_file(seed_keyword)

                return True
            finally:
                sys.argv = original_argv

        except Exception as e:
            return False

    def rename_and_cleanup_file(self, seed_keyword: str):
        """G 단계 완료 후 파일명 변경 및 정리"""
        try:
            excel_path = "result/keywordList_all.xlsx"
            new_filename = f"keywordResult_{seed_keyword}.xlsx"

            # 파일 존재 확인
            if os.path.exists(excel_path):
                # 새 파일명으로 변경
                os.rename(excel_path, new_filename)
                print(f"파일명이 '{new_filename}'으로 변경되었습니다.")
            else:
                print(f"경고: {excel_path} 파일을 찾을 수 없습니다.")

        except Exception as e:
            print(f"파일명 변경 중 오류 발생: {e}")


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

    # 창이 닫힐 때 sys.stdout 복원
    def on_closing():
        if hasattr(app, 'original_stdout'):
            sys.stdout = app.original_stdout
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

    # mainloop 종료 후 sys.stdout 복원 (안전장치)
    if hasattr(app, 'original_stdout'):
        sys.stdout = app.original_stdout


if __name__ == "__main__":
    main()