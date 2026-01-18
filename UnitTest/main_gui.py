import tkinter as tk
from tkinter import ttk, messagebox
import threading
import pandas as pd
import os
import webbrowser
from datetime import datetime
import traceback

from related_keywords import get_related_keywords
from keyword_stats import get_keyword_stats_from_keywordtool


class KeywordAnalyzerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("네이버 키워드 분석기 (최근 30일 블로그 발행량)")
        self.root.geometry("900x520")
        self.file_path = None

        frm = ttk.Frame(root, padding=10)
        frm.pack(fill="x")

        ttk.Label(frm, text="키워드").grid(row=0, column=0, padx=5)
        self.entry_keyword = ttk.Entry(frm, width=26)
        self.entry_keyword.grid(row=0, column=1, padx=5)

        ttk.Label(frm, text="연관 키워드 수").grid(row=0, column=2, padx=5)
        self.entry_limit = ttk.Entry(frm, width=8)
        self.entry_limit.insert(0, "20")
        self.entry_limit.grid(row=0, column=3, padx=5)

        ttk.Label(frm, text="최근 N일").grid(row=0, column=4, padx=5)
        self.entry_days = ttk.Entry(frm, width=6)
        self.entry_days.insert(0, "30")
        self.entry_days.grid(row=0, column=5, padx=5)

        self.btn_run = ttk.Button(frm, text="분석 시작", command=self.run)
        self.btn_run.grid(row=0, column=6, padx=10)

        self.progress = ttk.Progressbar(root, orient="horizontal", length=860, mode="determinate")
        self.progress.pack(pady=5)

        # 탭: 결과(표) / 로그
        self.tabs = ttk.Notebook(root)
        self.tabs.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_result = ttk.Frame(self.tabs)
        self.tab_log = ttk.Frame(self.tabs)
        self.tabs.add(self.tab_result, text="결과")
        self.tabs.add(self.tab_log, text="로그")

        # 결과 표
        cols = ("NO", "키워드", "PC", "Mobile", "검색수합계", "월간 블로그 발행", "포화도")
        self.tree = ttk.Treeview(self.tab_result, columns=cols, show="headings")
        for c in cols:
            self.tree.heading(c, text=c)
            if c in ("NO",):
                self.tree.column(c, width=50, anchor="center")
            elif c == "키워드":
                self.tree.column(c, width=240, anchor="w")
            else:
                self.tree.column(c, width=120, anchor="e")

        vsb = ttk.Scrollbar(self.tab_result, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree.bind("<Double-1>", self.open_keyword)

        # 로그
        self.text_log = tk.Text(self.tab_log, wrap="word")
        self.text_log.pack(fill="both", expand=True)

        bottom = ttk.Frame(root)
        bottom.pack(fill="x", padx=10, pady=(0, 10))

        self.btn_open = ttk.Button(bottom, text="📂 결과 엑셀 열기", command=self.open_excel, state="disabled")
        self.btn_open.pack(side="right")

    # --- UI 업데이트는 메인 스레드에서만 ---
    def log(self, msg: str):
        def _append():
            self.text_log.insert(tk.END, msg + "\n")
            self.text_log.see(tk.END)
        self.root.after(0, _append)

    def set_progress(self, value: float):
        self.root.after(0, lambda v=value: self.progress.configure(value=v))

    def set_buttons(self, run_state: str, open_state: str):
        def _set():
            self.btn_run.config(state=run_state)
            self.btn_open.config(state=open_state)
        self.root.after(0, _set)

    def open_excel(self):
        if self.file_path and os.path.exists(self.file_path):
            try:
                os.startfile(os.path.abspath(self.file_path))  # Windows
            except Exception:
                webbrowser.open(os.path.abspath(self.file_path))
        else:
            messagebox.showerror("오류", "결과 파일이 존재하지 않습니다.")

    def open_keyword(self, event=None):
        item = self.tree.selection()
        if not item:
            return
        values = self.tree.item(item[0]).get("values", [])
        if len(values) < 2:
            return
        keyword = values[1]
        webbrowser.open(f"https://search.naver.com/search.naver?query={keyword}")

    def run(self):
        keyword = self.entry_keyword.get().strip()
        if not keyword:
            messagebox.showerror("입력 오류", "키워드를 입력하세요.")
            return

        try:
            limit = int(self.entry_limit.get())
        except Exception:
            messagebox.showerror("입력 오류", "연관 키워드 수를 숫자로 입력하세요.")
            return

        try:
            days = int(self.entry_days.get())
            if days <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror("입력 오류", "최근 N일은 1 이상의 숫자로 입력하세요.")
            return

        self.tree.delete(*self.tree.get_children())
        self.text_log.delete("1.0", tk.END)
        self.set_progress(0)
        self.set_buttons("disabled", "disabled")
        self.tabs.select(self.tab_log)

        t = threading.Thread(target=self._run_worker, args=(keyword, limit, days), daemon=True)
        t.start()

    def _run_worker(self, keyword: str, limit: int, days: int):
        try:
            self.log(f"🔍 '{keyword}' 연관 키워드 수집 중...")
            keywordtool_data = get_related_keywords(keyword, limit)

            if not keywordtool_data:
                self.log("⚠️ 연관 키워드를 찾을 수 없습니다.")
                self.set_buttons("normal", "disabled")
                return

            self.log(f"총 {len(keywordtool_data)}개 수집 완료.\n검색량 컷 적용 후 최근 {days}일 블로그 발행량 조회...\n")
            total = len(keywordtool_data)

            def progress_callback(current, _total):
                pct = (current / total) * 100 if total else 0
                self.set_progress(pct)

            stats = get_keyword_stats_from_keywordtool(
                keywordtool_data,
                progress_callback=progress_callback,
                min_total_search=1000,
                recent_days=days,
            )

            df = pd.DataFrame(stats)
            required = {"PC", "Mobile", "검색수합계", "월간 블로그 발행", "포화도"}

            if df.empty or not required.issubset(df.columns):
                self.log("⚠️ 조건에 맞는 결과가 없습니다.")
                self.set_buttons("normal", "disabled")
                return

            df = df.sort_values("포화도", ascending=True)

            # 표 채우기
            def _fill_table():
                self.tree.delete(*self.tree.get_children())
                for i, row in enumerate(df.itertuples(index=False), start=1):
                    self.tree.insert(
                        "",
                        "end",
                        values=(
                            i,
                            getattr(row, "키워드"),
                            f"{getattr(row, 'PC'):,}",
                            f"{getattr(row, 'Mobile'):,}",
                            f"{getattr(row, '검색수합계'):,}",
                            f"{getattr(row, '월간_블로그_발행'):,}",
                            f"{getattr(row, '포화도')}%",
                        ),
                    )

                self.tabs.select(self.tab_result)

            self.root.after(0, _fill_table)

            # 엑셀 저장
            os.makedirs("result", exist_ok=True)
            now = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.file_path = os.path.join("result", f"{now}_{keyword}_recent{days}d.xlsx")
            df.to_excel(self.file_path, index=False)

            self.set_progress(100)
            self.log(f"\n✅ 분석 완료! 결과 파일: {self.file_path}")
            self.set_buttons("normal", "normal")
            self.root.after(0, lambda: messagebox.showinfo("완료", "분석이 완료되었습니다."))

        except Exception as e:
            err_text = "".join(traceback.format_exception(type(e), e, e.__traceback__))
            self.log("\n🚨 오류 발생!\n" + err_text)
            self.set_buttons("normal", "disabled")
            err_msg = str(e)
            self.root.after(0, lambda m=err_msg: messagebox.showerror("오류", m))


if __name__ == "__main__":
    root = tk.Tk()
    app = KeywordAnalyzerApp(root)
    root.mainloop()
