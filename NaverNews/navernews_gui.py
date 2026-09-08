#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
import subprocess
import logging
from datetime import datetime
from typing import List, Tuple

# GUI imports
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QSpinBox, QPushButton, QTextEdit, QProgressBar,
                             QMessageBox, QGroupBox, QSplitter, QFrame)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QIcon, QTextCursor

# 크롤러 모듈 import
from navernews import main as crawler_main


class CrawlerThread(QThread):
    """크롤링 작업을 백그라운드에서 실행하는 스레드"""
    progress_signal = pyqtSignal(str)  # 진행상황 메시지
    finished_signal = pyqtSignal(bool, str)  # 완료 시그널 (성공여부, 메시지)
    log_signal = pyqtSignal(str)  # 로그 메시지

    def __init__(self, keyword: str, count: int):
        super().__init__()
        self.keyword = keyword
        self.count = count
        self.result_folder = ""

    def run(self):
        """크롤링 작업 실행"""
        try:
            # 로그 캡처를 위한 핸들러 설정
            class GUILogHandler(logging.Handler):
                def emit(self, record):
                    msg = self.format(record)
                    self.thread.log_signal.emit(msg)

            # 핸들러에 스레드 참조 추가
            GUILogHandler.thread = self

            # 기존 로거에 핸들러 추가
            logger = logging.getLogger()
            gui_handler = GUILogHandler()
            gui_handler.setFormatter(logging.Formatter('%(asctime)s | %(levelname)s | %(message)s',
                                                     datefmt='%H:%M:%S'))
            logger.addHandler(gui_handler)
            logger.setLevel(logging.INFO)

            self.progress_signal.emit(f"크롤링 시작: '{self.keyword}' 키워드로 {self.count}개 기사 수집")

            # 명령줄 인수 설정 (크롤러에서 인식할 수 있도록)
            original_argv = sys.argv.copy()
            sys.argv = ['navernews.py', self.keyword, str(self.count)]

            try:
                result = crawler_main()
                if result == 0:
                    # 키워드 폴더명 생성 (navernews.py의 로직과 동일)
                    safe_keyword = re.sub(r'[^\w\s가-힣]', '', self.keyword)[:30].strip()
                    if not safe_keyword:
                        safe_keyword = "unknown_keyword"
                    self.result_folder = os.path.join("results", safe_keyword)
                    self.progress_signal.emit("크롤링 완료!")
                    self.finished_signal.emit(True, f"크롤링이 성공적으로 완료되었습니다!\n저장 폴더: {self.result_folder}")
                else:
                    self.progress_signal.emit("크롤링 실패")
                    self.finished_signal.emit(False, "크롤링 중 오류가 발생했습니다.")
            finally:
                # 원래 argv 복원
                sys.argv = original_argv
                # 핸들러 제거
                logger.removeHandler(gui_handler)

        except Exception as e:
            self.log_signal.emit(f"오류 발생: {str(e)}")
            self.finished_signal.emit(False, f"오류 발생: {str(e)}")


class NaverNewsCrawlerGUI(QMainWindow):
    """네이버 뉴스 크롤러 GUI 메인 윈도우"""

    def __init__(self):
        super().__init__()
        self.crawler_thread = None
        self.result_folder = ""
        self.check_playwright_browser()
        self.init_ui()

    def init_ui(self):
        """UI 초기화"""
        self.setWindowTitle("네이버 뉴스 크롤러 v2.0")
        self.setGeometry(100, 100, 900, 700)
        self.setWindowIcon(QIcon())  # 아이콘 설정 (필요시)

        # 중앙 위젯
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 메인 레이아웃
        main_layout = QVBoxLayout(central_widget)

        # 타이틀
        title_label = QLabel("네이버 뉴스 크롤러")
        title_label.setFont(QFont("Arial", 10, QFont.Bold))  # 크롤링설정과 동일한 크기
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # 입력 그룹
        input_group = QGroupBox("크롤링 설정")
        input_layout = QHBoxLayout()

        # 키워드 입력
        keyword_layout = QVBoxLayout()
        keyword_label = QLabel("키워드:")
        keyword_label.setFont(QFont("Arial", 10, QFont.Bold))  # 크롤링설정과 동일한 크기
        self.keyword_input = QLineEdit()
        self.keyword_input.setPlaceholderText("크롤링할 키워드를 입력하세요")
        self.keyword_input.setMinimumWidth(300)
        self.keyword_input.setMinimumHeight(30)  # 높이도 적당히 조정
        self.keyword_input.setFont(QFont("Arial", 30))  # 크롤링설정과 동일한 크기로 통일
        keyword_layout.addWidget(keyword_label)
        keyword_layout.addWidget(self.keyword_input)
        input_layout.addLayout(keyword_layout)

        # 갯수 입력
        count_layout = QVBoxLayout()
        count_label = QLabel("수집 개수:")
        count_label.setFont(QFont("Arial", 10, QFont.Bold))  # 이미 10pt
        self.count_spinbox = QSpinBox()
        self.count_spinbox.setRange(1, 100)
        self.count_spinbox.setValue(10)
        self.count_spinbox.setSuffix(" 개")
        count_layout.addWidget(count_label)
        count_layout.addWidget(self.count_spinbox)
        input_layout.addLayout(count_layout)

        # 버튼들
        button_layout = QVBoxLayout()

        self.start_button = QPushButton("크롤링 시작")
        self.start_button.setFont(QFont("Arial", 10, QFont.Bold))  # 크롤링설정과 동일한 크기
        self.start_button.setMinimumHeight(35)
        self.start_button.clicked.connect(self.start_crawling)
        button_layout.addWidget(self.start_button)

        self.stop_button = QPushButton("중지")
        self.stop_button.setEnabled(False)
        self.stop_button.setFont(QFont("Arial", 10))  # 크롤링설정과 동일한 크기
        self.stop_button.setMinimumHeight(35)
        self.stop_button.clicked.connect(self.stop_crawling)
        button_layout.addWidget(self.stop_button)

        self.open_folder_button = QPushButton("결과 폴더 열기")
        self.open_folder_button.setEnabled(False)
        self.open_folder_button.setFont(QFont("Arial", 10))  # 크롤링설정과 동일한 크기
        self.open_folder_button.setMinimumHeight(35)
        self.open_folder_button.clicked.connect(self.open_result_folder)
        button_layout.addWidget(self.open_folder_button)

        input_layout.addLayout(button_layout)
        input_group.setLayout(input_layout)
        main_layout.addWidget(input_group)

        # 진행바
        progress_layout = QVBoxLayout()
        progress_label = QLabel("진행상황:")
        progress_label.setFont(QFont("Arial", 10, QFont.Bold))  # 크롤링설정과 동일한 크기
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setMinimumHeight(20)
        progress_layout.addWidget(progress_label)
        progress_layout.addWidget(self.progress_bar)
        main_layout.addLayout(progress_layout)

        # 로그 표시 영역
        log_group = QGroupBox("로그")
        log_layout = QVBoxLayout()

        self.log_text = QTextEdit()
        self.log_text.setFont(QFont("Consolas", 10))  # 크롤링설정과 동일한 크기 (Consolas지만 10pt)
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(350)
        self.log_text.setMinimumHeight(300)
        log_layout.addWidget(self.log_text)

        # 로그 클리어 버튼
        clear_log_button = QPushButton("로그 지우기")
        clear_log_button.setFont(QFont("Arial", 10))  # 크롤링설정과 동일한 크기로 통일
        clear_log_button.setMinimumHeight(30)
        clear_log_button.clicked.connect(self.clear_log)
        log_layout.addWidget(clear_log_button)

        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)

        # 상태바
        self.status_bar = self.statusBar()
        self.status_bar.showMessage("준비됨")

        # 스타일시트 적용
        self.apply_stylesheet()

    def check_playwright_browser(self):
        """Playwright 브라우저 설치 확인 및 자동 설치"""
        try:
            # Playwright 초기화 시도
            from playwright.sync_api import sync_playwright
            playwright = sync_playwright().start()
            # 테스트 브라우저 실행 시도
            browser = playwright.chromium.launch(headless=True, args=["--headless", "--disable-gpu", "--no-sandbox"])
            browser.close()
            playwright.stop()
            return True
        except Exception as e:
            # 브라우저 설치 필요
            reply = QMessageBox.question(
                self, '브라우저 설치 필요',
                'Playwright 브라우저가 설치되지 않았습니다.\n'
                '크롤링을 위해 브라우저를 설치하시겠습니까?\n\n'
                f'오류: {str(e)}',
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )

            if reply == QMessageBox.Yes:
                self.install_playwright_browser()
            else:
                QMessageBox.warning(self, '경고', '브라우저가 설치되지 않아 프로그램을 사용할 수 없습니다.')
                sys.exit(1)

    def install_playwright_browser(self):
        """Playwright 브라우저 자동 설치"""
        try:
            from PyQt5.QtWidgets import QProgressDialog
            from PyQt5.QtCore import Qt

            progress = QProgressDialog("Playwright 브라우저 설치 중...", "취소", 0, 0, self)
            progress.setWindowModality(Qt.WindowModal)
            progress.setAutoReset(False)
            progress.setAutoClose(False)
            progress.show()

            import subprocess
            import sys

            # playwright install chromium 실행
            process = subprocess.run(
                [sys.executable, "-m", "playwright", "install", "chromium"],
                capture_output=True,
                text=True
            )

            progress.close()

            if process.returncode == 0:
                QMessageBox.information(self, '완료', '브라우저 설치가 완료되었습니다!')
                return True
            else:
                error_msg = process.stderr if process.stderr else "알 수 없는 오류"
                QMessageBox.critical(self, '설치 실패',
                                   f'브라우저 설치에 실패했습니다:\n{error_msg}')
                return False

        except Exception as e:
            progress.close()
            QMessageBox.critical(self, '오류', f'브라우저 설치 중 오류 발생:\n{str(e)}')
            return False

    def apply_stylesheet(self):
        """스타일시트 적용"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }

            QGroupBox {
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: 1ex;
            }

            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 10px 0 10px;
            }

            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 12px 18px;   /* ✅ 눌림 영역 커짐 */
                border-radius: 6px;
                font-size: 12px;      /* ✅ 글자 조금 더 큼 */
                min-height: 42px;     /* ✅ 버튼 높이 확보 */
                min-width: 130px;     /* ✅ 오른쪽 버튼 폭 통일 */
            }

            QPushButton:hover {
                background-color: #45a049;
            }

            QPushButton:pressed {
                background-color: #3e8e41;
            }

            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }

            QLineEdit, QSpinBox {
                padding: 6px;
                border: 1px solid #cccccc;
                border-radius: 4px;
                font-size: 10px;
            }

            QSpinBox {
                font-size: 18px;
                min-height: 44px;
                min-width: 120px;
                padding: 8px;
                padding-right: 26px; /* 화살표 영역 확보 */
            }

            QLineEdit {
                font-size: 22px;
                min-height: 44px;
                padding: 10px;
            }

            QLineEdit:focus, QSpinBox:focus {
                border-color: #4CAF50;
            }

            QProgressBar {
                border: 1px solid #cccccc;
                border-radius: 4px;
                text-align: center;
            }

            QProgressBar::chunk {
                background-color: #4CAF50;
            }

            QTextEdit {
                border: 1px solid #cccccc;
                border-radius: 4px;
                font-family: 'Consolas', monospace;
                font-size: 10px;  /* 크롤링설정과 동일한 크기 */
            }


        """)

    def start_crawling(self):
        """크롤링 시작"""
        keyword = self.keyword_input.text().strip()
        count = self.count_spinbox.value()

        if not keyword:
            QMessageBox.warning(self, "입력 오류", "키워드를 입력해주세요.")
            return

        # UI 상태 변경
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.open_folder_button.setEnabled(False)
        self.keyword_input.setEnabled(False)
        self.count_spinbox.setEnabled(False)

        # 로그 초기화
        self.log_text.clear()
        self.progress_bar.setValue(0)
        self.status_bar.showMessage("크롤링 중...")

        # 크롤링 스레드 시작
        self.crawler_thread = CrawlerThread(keyword, count)
        self.crawler_thread.progress_signal.connect(self.update_progress)
        self.crawler_thread.log_signal.connect(self.add_log)
        self.crawler_thread.finished_signal.connect(self.crawling_finished)
        self.crawler_thread.start()

    def stop_crawling(self):
        """크롤링 중지"""
        if self.crawler_thread and self.crawler_thread.isRunning():
            self.crawler_thread.terminate()
            self.crawling_finished(False, "사용자에 의해 중단되었습니다.")

    def update_progress(self, message: str):
        """진행상황 업데이트"""
        self.status_bar.showMessage(message)

        # 진행바 업데이트 (단순 추정)
        if "시작" in message:
            self.progress_bar.setValue(10)
        elif "링크 수집" in message:
            self.progress_bar.setValue(30)
        elif "처리 중" in message:
            # 현재 진행률 계산 (로그에서 추출)
            if "/" in message:
                try:
                    parts = message.split("/")
                    if len(parts) == 2:
                        current = int(parts[0].split()[-1])
                        total = int(parts[1].split()[0])
                        progress = 30 + (current / total) * 60
                        self.progress_bar.setValue(int(progress))
                except:
                    pass
        elif "완료" in message:
            self.progress_bar.setValue(100)

    def add_log(self, message: str):
        """로그 추가"""
        self.log_text.append(message)
        # 자동 스크롤
        cursor = self.log_text.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.log_text.setTextCursor(cursor)

    def crawling_finished(self, success: bool, message: str):
        """크롤링 완료 처리"""
        # UI 상태 복원
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.keyword_input.setEnabled(True)
        self.count_spinbox.setEnabled(True)

        if success:
            self.open_folder_button.setEnabled(True)
            self.result_folder = getattr(self.crawler_thread, 'result_folder', '')
            QMessageBox.information(self, "완료", message)
        else:
            QMessageBox.warning(self, "오류", message)

        self.status_bar.showMessage("준비됨" if success else "오류 발생")
        self.progress_bar.setValue(100 if success else 0)

    def open_result_folder(self):
        """결과 폴더 열기"""
        if self.result_folder:
            try:
                # 절대 경로로 변환
                folder_path = os.path.abspath(self.result_folder)

                # 폴더가 존재하는지 확인
                if not os.path.exists(folder_path):
                    QMessageBox.warning(self, "오류", f"결과 폴더가 존재하지 않습니다:\n{folder_path}")
                    return

                if os.name == 'nt':  # Windows
                    os.startfile(folder_path)
                elif os.name == 'posix':  # macOS, Linux
                    subprocess.run(['xdg-open', folder_path])
                self.add_log(f"결과 폴더 열기: {folder_path}")
                QMessageBox.information(self, "완료", f"결과 폴더를 열었습니다:\n{folder_path}")
            except Exception as e:
                QMessageBox.warning(self, "오류", f"폴더를 열 수 없습니다: {str(e)}")
        else:
            QMessageBox.warning(self, "오류", "결과 폴더가 설정되지 않았습니다.")

    def clear_log(self):
        """로그 지우기"""
        self.log_text.clear()
        self.add_log("로그가 초기화되었습니다.")

    def closeEvent(self, event):
        """윈도우 닫기 이벤트"""
        if self.crawler_thread and self.crawler_thread.isRunning():
            reply = QMessageBox.question(
                self, '확인',
                "크롤링이 진행 중입니다. 정말 종료하시겠습니까?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )

            if reply == QMessageBox.Yes:
                self.stop_crawling()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


def main():
    """메인 함수"""
    app = QApplication(sys.argv)

    # 한글 지원
    app.setFont(QFont("맑은 고딕", 10))  # 크롤링설정과 동일한 크기로 통일

    # GUI 실행
    window = NaverNewsCrawlerGUI()
    window.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()