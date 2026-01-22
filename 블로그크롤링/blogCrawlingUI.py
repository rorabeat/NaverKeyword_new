#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
네이버 블로그 크롤링 GUI 애플리케이션

이 애플리케이션은 네이버 블로그 크롤링을 위한 GUI 인터페이스를 제공합니다.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import os
import sys
import webbrowser
from pathlib import Path

# blogcontentsClawring.py에서 필요한 함수들 import
try:
    from blogcontentsClawring import get_blog_posts_by_keyword, crawl_multiple_posts, BLOGGER_SEARCH_AVAILABLE
except ImportError:
    messagebox.showerror("오류", "blogcontentsClawring.py 파일을 찾을 수 없습니다.")
    sys.exit(1)


class BlogCrawlerGUI:
    """네이버 블로그 크롤링 GUI 클래스"""

    def __init__(self, root):
        self.root = root
        self.root.title("네이버 블로그 크롤링 도구")
        self.root.geometry("600x500")
        self.root.resizable(True, True)

        # 아이콘 설정 (선택사항)
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass

        self.create_widgets()
        self.center_window()

    def create_widgets(self):
        """GUI 위젯 생성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 제목
        title_label = ttk.Label(main_frame, text="네이버 블로그 크롤링 도구",
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # 키워드 입력 섹션
        keyword_frame = ttk.LabelFrame(main_frame, text="검색 설정", padding="10")
        keyword_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))

        ttk.Label(keyword_frame, text="키워드:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.keyword_var = tk.StringVar()
        self.keyword_entry = ttk.Entry(keyword_frame, textvariable=self.keyword_var, width=40)
        self.keyword_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=5)

        ttk.Label(keyword_frame, text="크롤링할 갯수:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.count_var = tk.IntVar(value=5)
        self.count_spinbox = tk.Spinbox(keyword_frame, from_=1, to=50, textvariable=self.count_var, width=10)
        self.count_spinbox.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=5)

        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=(10, 20))

        self.start_button = ttk.Button(button_frame, text="크롤링 시작",
                                     command=self.start_crawling, style="Accent.TButton")
        self.start_button.grid(row=0, column=0, padx=(0, 10))

        self.stop_button = ttk.Button(button_frame, text="중지", command=self.stop_crawling, state="disabled")
        self.stop_button.grid(row=0, column=1, padx=(0, 10))

        self.open_result_button = ttk.Button(button_frame, text="결과 폴더 열기",
                                          command=self.open_result_folder, state="disabled")
        self.open_result_button.grid(row=0, column=2)

        # 진행 상황 표시 영역
        progress_frame = ttk.LabelFrame(main_frame, text="진행 상황", padding="10")
        progress_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))

        # 진행바
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        # 로그 텍스트 영역
        self.log_text = tk.Text(progress_frame, height=15, wrap=tk.WORD, state="disabled")
        scrollbar = ttk.Scrollbar(progress_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.log_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))

        # 상태 표시줄
        self.status_var = tk.StringVar()
        self.status_var.set("준비 완료")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))

        # 그리드 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        keyword_frame.columnconfigure(1, weight=1)
        progress_frame.columnconfigure(0, weight=1)
        progress_frame.rowconfigure(1, weight=1)

        # 스타일 설정
        style = ttk.Style()
        style.configure("Accent.TButton", font=("Arial", 10, "bold"))

        # 단축키 설정
        self.root.bind('<Return>', lambda e: self.start_crawling())
        self.root.bind('<Escape>', lambda e: self.stop_crawling())

        # 크롤링 상태
        self.is_crawling = False
        self.crawl_thread = None

    def center_window(self):
        """윈도우를 화면 중앙에 배치"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def log_message(self, message):
        """로그 메시지 추가"""
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        self.root.update_idletasks()

    def clear_log(self):
        """로그 영역 초기화"""
        self.log_text.config(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state="disabled")

    def start_crawling(self):
        """크롤링 시작"""
        keyword = self.keyword_var.get().strip()
        count = self.count_var.get()

        if not keyword:
            messagebox.showwarning("입력 오류", "키워드를 입력해주세요.")
            self.keyword_entry.focus()
            return

        if not (1 <= count <= 50):
            messagebox.showwarning("입력 오류", "크롤링 갯수는 1-50 사이로 입력해주세요.")
            return

        # 크롤링 시작 확인
        if not messagebox.askyesno("확인", f"'{keyword}' 키워드로 {count}개의 블로그를 크롤링하시겠습니까?"):
            return

        # UI 상태 변경
        self.is_crawling = True
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.open_result_button.config(state="disabled")
        self.keyword_entry.config(state="disabled")
        self.count_spinbox.config(state="disabled")

        self.status_var.set("크롤링 준비 중...")
        self.progress_var.set(0)
        self.clear_log()

        # 크롤링 스레드 시작
        self.crawl_thread = threading.Thread(target=self.run_crawling, args=(keyword, count))
        self.crawl_thread.daemon = True
        self.crawl_thread.start()

    def stop_crawling(self):
        """크롤링 중지"""
        if self.is_crawling:
            self.is_crawling = False
            self.status_var.set("크롤링 중지 중...")
            self.log_message("사용자에 의해 크롤링이 중지되었습니다.")

    def run_crawling(self, keyword, count):
        """크롤링 실행 (별도 스레드에서 실행)"""
        try:
            self.status_var.set("블로그 검색 중...")
            self.log_message(f"키워드 '{keyword}'로 블로그 검색을 시작합니다...")

            # 블로그 검색
            post_urls = get_blog_posts_by_keyword(keyword, count)

            if not post_urls:
                self.root.after(0, lambda: messagebox.showwarning("검색 결과", f"'{keyword}' 키워드로 검색된 블로그가 없습니다."))
                self.root.after(0, self.reset_ui)
                return

            self.log_message(f"검색 완료: {len(post_urls)}개의 블로그를 찾았습니다.")
            for i, url in enumerate(post_urls, 1):
                self.log_message(f"  {i}. {url}")

            self.status_var.set("크롤링 중...")
            self.progress_var.set(10)

            # 크롤링 실행
            crawl_multiple_posts(post_urls, keyword)

            if self.is_crawling:  # 중지되지 않은 경우
                self.root.after(0, self.crawling_completed)

        except Exception as e:
            error_msg = f"크롤링 중 오류 발생: {str(e)}"
            self.log_message(error_msg)
            self.root.after(0, lambda: messagebox.showerror("오류", error_msg))
            self.root.after(0, self.reset_ui)

    def crawling_completed(self):
        """크롤링 완료 처리"""
        self.status_var.set("크롤링 완료")
        self.progress_var.set(100)
        self.log_message("\n크롤링이 완료되었습니다!")
        self.log_message("결과 폴더를 확인해주세요.")

        # 결과 폴더 열기 버튼 활성화
        self.open_result_button.config(state="normal")

        # 성공 메시지 표시
        messagebox.showinfo("완료", "블로그 크롤링이 완료되었습니다!\n\n결과 폴더 열기 버튼을 클릭하여 결과를 확인하세요.")

        self.reset_ui()

    def reset_ui(self):
        """UI를 초기 상태로 리셋"""
        self.is_crawling = False
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.keyword_entry.config(state="normal")
        self.count_spinbox.config(state="normal")
        self.status_var.set("준비 완료")
        self.progress_var.set(0)

    def open_result_folder(self):
        """결과 폴더 열기"""
        script_dir = Path(__file__).parent
        result_dir = script_dir / "result"
        if result_dir.exists():
            try:
                # Windows에서 폴더 열기
                if os.name == 'nt':
                    os.startfile(str(result_dir))
                # macOS에서 폴더 열기
                elif os.name == 'posix':
                    if sys.platform == 'darwin':  # macOS
                        os.system(f'open "{result_dir}"')
                    else:  # Linux
                        os.system(f'xdg-open "{result_dir}"')
                else:
                    # 크로스 플랫폼 방식
                    webbrowser.open(str(result_dir))
            except Exception as e:
                messagebox.showerror("오류", f"결과 폴더를 열 수 없습니다: {e}")
        else:
            messagebox.showwarning("경고", "결과 폴더가 존재하지 않습니다.")


def main():
    """메인 함수"""
    root = tk.Tk()
    app = BlogCrawlerGUI(root)

    # 창 닫기 이벤트 처리
    def on_closing():
        if app.is_crawling:
            if messagebox.askyesno("확인", "크롤링이 진행 중입니다. 정말 종료하시겠습니까?"):
                app.stop_crawling()
                root.destroy()
        else:
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()