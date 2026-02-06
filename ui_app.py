import sys
import os
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QPushButton,
    QLabel,
    QLineEdit,
    QTextEdit,
    QFileDialog,
    QVBoxLayout,
    QHBoxLayout,
    QMessageBox,
)
from PySide6.QtCore import Qt, QThread, Signal

# src를 import 경로에 추가 (엔진 모듈 사용)
BASE_DIR = Path(__file__).parent
SRC_DIR = BASE_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.append(str(SRC_DIR))

from tools.engine import (  # type: ignore
    connect_document,
    get_current_document_path,
    rewrite_current_document,
)


class RewriteWorker(QThread):
    """긴 작업(전체 재작성)을 UI와 분리해서 돌리는 스레드."""

    log_signal = Signal(str)
    done_signal = Signal(bool)

    def __init__(self, mode: str = "rewrite"):
        super().__init__()
        self.mode = mode

    def run(self):
        try:
            self.log_signal.emit(f"[INFO] 전체 문서 재작성을 시작합니다 (mode={self.mode}).")
            rewrite_current_document(self.mode)  # 엔진 호출
            self.log_signal.emit("[INFO] 전체 문서 재작성 작업이 종료되었습니다.")
            self.done_signal.emit(True)
        except Exception as e:
            self.log_signal.emit(f"[ERROR] 재작성 중 예외 발생: {e}")
            self.done_signal.emit(False)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("HwpInlineAI v0.2 (Python GUI)")
        self.setMinimumSize(800, 500)

        # ---- 상단: 파일 선택 + 연결 상태 ----
        self.path_label = QLabel("선택된 파일:")
        self.path_edit = QLineEdit()
        self.path_edit.setReadOnly(True)

        self.browse_button = QPushButton("한글 파일 선택")
        self.connect_button = QPushButton("연결")

        self.status_label = QLabel("상태: 연결 안 됨")

        path_layout = QHBoxLayout()
        path_layout.addWidget(self.path_label)
        path_layout.addWidget(self.path_edit)
        path_layout.addWidget(self.browse_button)
        path_layout.addWidget(self.connect_button)

        # ---- 가운데: 챗봇 / 액션 영역 (v0: 로그 + 입력 + 버튼) ----
        self.chat_log = QTextEdit()
        self.chat_log.setReadOnly(True)

        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText("명령이나 메시지를 입력하세요. (v0에서는 전체 다듬기만 동작)")
        self.send_button = QPushButton("보내기 (전체 다듬기)")
        self.send_button.setEnabled(False)  # 연결 전에는 비활성화

        input_layout = QHBoxLayout()
        input_layout.addWidget(self.input_edit)
        input_layout.addWidget(self.send_button)

        main_layout = QVBoxLayout()
        main_layout.addLayout(path_layout)
        main_layout.addWidget(self.status_label)
        main_layout.addWidget(QLabel("대화 / 로그:"))
        main_layout.addWidget(self.chat_log)
        main_layout.addLayout(input_layout)

        self.setLayout(main_layout)

        # 시그널 연결
        self.browse_button.clicked.connect(self.on_browse_clicked)
        self.connect_button.clicked.connect(self.on_connect_clicked)
        self.send_button.clicked.connect(self.on_send_clicked)

        self.worker: RewriteWorker | None = None

    # ---- 유틸 ----
    def log(self, message: str):
        self.chat_log.append(message)
        self.chat_log.ensureCursorVisible()

    def set_connected_ui(self, connected: bool):
        if connected:
            path = get_current_document_path() or "(알 수 없음)"
            self.status_label.setText(f"상태: 연결됨 → {path}")
            self.send_button.setEnabled(True)
            self.connect_button.setEnabled(False)
        else:
            self.status_label.setText("상태: 연결 안 됨")
            self.send_button.setEnabled(False)
            self.connect_button.setEnabled(True)

    # ---- 슬롯 ----
    def on_browse_clicked(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "한글 파일 선택",
            "",
            "HWP Files (*.hwp);;All Files (*)",
        )
        if not file_path:
            return

        self.path_edit.setText(file_path)
        self.log(f"[INFO] 선택된 파일: {file_path}")

    def on_connect_clicked(self):
        path = self.path_edit.text().strip()
        if not path:
            QMessageBox.warning(self, "경고", "먼저 한글 파일을 선택해주세요.")
            return

        if not os.path.exists(path):
            QMessageBox.critical(self, "에러", f"파일을 찾을 수 없습니다:\n{path}")
            return

        try:
            self.log(f"[INFO] 한글에 문서를 연결합니다: {path}")
            connect_document(path, visible=True)
            self.log("[INFO] 문서 연결 완료.")
            self.set_connected_ui(True)
        except Exception as e:
            self.log(f"[ERROR] 문서 연결 실패: {e}")
            QMessageBox.critical(self, "연결 실패", str(e))
            self.set_connected_ui(False)

    def on_send_clicked(self):
        # v0: 입력 내용은 아직 사용하지 않고, 전체 다듬기만 실행
        text = self.input_edit.text().strip()
        if text:
            self.log(f"[사용자] {text}")
        else:
            self.log("[사용자] (전체 다듬기 실행)")

        # 확인 대화상자
        reply = QMessageBox.question(
            self,
            "확인",
            "현재 연결된 문서 전체를 AI로 재작성할까요?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        self.send_button.setEnabled(False)
        self.log("[INFO] 전체 다듬기 작업을 시작합니다...")

        self.worker = RewriteWorker(mode="rewrite")
        self.worker.log_signal.connect(self.log)
        self.worker.done_signal.connect(self.on_worker_done)
        self.worker.start()

    def on_worker_done(self, success: bool):
        if success:
            QMessageBox.information(self, "완료", "문서 전체 재작성이 완료되었습니다.")
        else:
            QMessageBox.warning(self, "실패", "작업 중 오류가 발생했습니다. 로그를 확인해주세요.")
        self.send_button.setEnabled(True)


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
