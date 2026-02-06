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
    QSplitter,
    QFrame,
    QMessageBox,
)
from PySide6.QtCore import Qt

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


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("HwpInlineAI (HWP + AI Editor)")
        self.setMinimumSize(960, 540)

        # ---- 좌측 패널: 파일 / 상태 / 액션 ----
        left_frame = QFrame(objectName="LeftPanel")
        left_layout = QVBoxLayout()
        left_layout.setContentsMargins(12, 12, 12, 12)
        left_layout.setSpacing(10)

        self.app_title = QLabel("HwpInlineAI")
        self.app_title.setObjectName("AppTitle")

        self.path_label = QLabel("연결된 파일 없음")
        self.path_label.setWordWrap(True)

        self.status_label = QLabel("● 상태: 연결 안 됨")
        self.status_label.setObjectName("StatusLabel")

        self.path_edit = QLineEdit()
        self.path_edit.setReadOnly(True)
        self.path_edit.setPlaceholderText("한글 파일을 선택하세요...")

        self.browse_button = QPushButton("파일 선택")
        self.browse_button.setObjectName("SecondaryButton")
        self.connect_button = QPushButton("한글과 연결")
        self.connect_button.setObjectName("PrimaryButton")

        left_layout.addWidget(self.app_title)
        left_layout.addSpacing(8)
        left_layout.addWidget(self.path_label)
        left_layout.addWidget(self.status_label)
        left_layout.addSpacing(8)
        left_layout.addWidget(self.path_edit)
        left_layout.addWidget(self.browse_button)
        left_layout.addWidget(self.connect_button)
        left_layout.addSpacing(16)

        self.send_button = QPushButton("전체 문서 다듬기")
        self.send_button.setObjectName("PrimaryButton")
        self.send_button.setEnabled(False)
        self.sel_get_button = QPushButton("선택 영역 가져오기")
        self.sel_get_button.setObjectName("SecondaryButton")
        self.sel_get_button.setEnabled(False)
        self.sel_rewrite_button = QPushButton("선택 영역 다듬기")
        self.sel_rewrite_button.setObjectName("SecondaryButton")
        self.sel_rewrite_button.setEnabled(False)

        left_layout.addWidget(self.send_button)
        left_layout.addWidget(self.sel_get_button)
        left_layout.addWidget(self.sel_rewrite_button)
        left_layout.addStretch(1)

        left_frame.setLayout(left_layout)

        # ---- 우측 패널: 대화 / 로그 / 입력 ----
        right_frame = QFrame()
        right_layout = QVBoxLayout()
        right_layout.setContentsMargins(12, 12, 12, 12)
        right_layout.setSpacing(8)

        # 선택 정보 라벨 (GPT 스타일: 입력 위에 현재 선택 상태 한 줄 표시)
        self.selection_label = QLabel("선택: 없음")

        # 대화 로그
        self.chat_log = QTextEdit()
        self.chat_log.setReadOnly(True)

        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText("명령이나 메시지를 입력하고 Enter를 누르세요...")

        input_row = QHBoxLayout()
        input_row.addWidget(self.input_edit)

        right_layout.addWidget(QLabel("대화 / 로그:"))
        right_layout.addWidget(self.chat_log, stretch=1)
        right_layout.addWidget(self.selection_label)
        right_layout.addLayout(input_row)

        right_frame.setLayout(right_layout)

        # ---- 메인 Splitter ----
        splitter = QSplitter()
        splitter.addWidget(left_frame)
        splitter.addWidget(right_frame)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([260, 700])

        main_layout = QVBoxLayout()
        main_layout.addWidget(splitter)
        self.setLayout(main_layout)

        # 선택/참조 상태
        self.last_selection_text: str = ""  # 가장 최근 선택 텍스트

        # 시그널 연결
        self.browse_button.clicked.connect(self.on_browse_clicked)
        self.connect_button.clicked.connect(self.on_connect_clicked)
        self.send_button.clicked.connect(self.on_send_clicked)
        self.sel_get_button.clicked.connect(self.on_sel_get_clicked)
        self.sel_rewrite_button.clicked.connect(self.on_sel_rewrite_clicked)

    # ---- 유틸 ----
    def log(self, message: str):
        self.chat_log.append(message)
        self.chat_log.ensureCursorVisible()

    def set_connected_ui(self, connected: bool):
        if connected:
            path = get_current_document_path() or "(알 수 없음)"
            short = path if len(path) < 60 else "..." + path[-57:]
            self.path_label.setText(f"연결된 파일:\n{short}")
            self.status_label.setText("● 상태: 연결됨")
            self.send_button.setEnabled(True)
            self.sel_get_button.setEnabled(True)
            self.sel_rewrite_button.setEnabled(True)
            self.connect_button.setEnabled(False)
        else:
            self.path_label.setText("연결된 파일 없음")
            self.status_label.setText("● 상태: 연결 안 됨")
            self.send_button.setEnabled(False)
            self.sel_get_button.setEnabled(False)
            self.sel_rewrite_button.setEnabled(False)
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

        # 메인 스레드에서 직접 엔진 호출 (CLI와 동일 조건)
        try:
            self.log("[INFO] 전체 문서 재작성을 시작합니다 (mode=rewrite).")
            rewrite_current_document("rewrite")
            self.log("[INFO] 전체 문서 재작성 작업이 종료되었습니다.")
            QMessageBox.information(self, "완료", "문서 전체 재작성이 완료되었습니다.")
        except Exception as e:
            self.log(f"[ERROR] 재작성 중 예외 발생: {e}")
            QMessageBox.warning(self, "실패", f"작업 중 오류가 발생했습니다:\n{e}")
        finally:
            self.send_button.setEnabled(True)

    def on_sel_get_clicked(self):
        """현재 한글에서 선택된 영역 텍스트를 가져와 선택 상태 라벨을 갱신."""
        from tools.engine import get_selection_text_via_clipboard  # type: ignore

        try:
            sel_text = get_selection_text_via_clipboard()
            if sel_text:
                self.last_selection_text = sel_text
                length = len(sel_text)

                # 커서 위치 메타데이터 가져오기 (문단/오프셋 정보)
                from tools.engine import get_cursor_position_meta  # type: ignore
                pos = get_cursor_position_meta()
                if pos:
                    para_id = pos.get("para_id")
                    char_pos = pos.get("char_pos")
                    self.selection_label.setText(
                        f"선택: 문단 ID {para_id}, 시작 오프셋 {char_pos}, {length}글자"
                    )
                else:
                    self.selection_label.setText(f"선택: {length}글자")

                self.log("[INFO] 선택 영역을 참조로 설정했습니다.")
            else:
                self.last_selection_text = ""
                self.selection_label.setText("선택: 없음")
                self.log("[INFO] 선택된 텍스트가 없거나 가져오지 못했습니다.")
        except Exception as e:
            self.log(f"[ERROR] 선택 영역 가져오기 실패: {e}")

    def on_sel_rewrite_clicked(self):
        """선택된 영역만 다듬기 v0: 선택 텍스트를 AI에 보내고 결과를 다시 붙임."""
        from tools.engine import (
            get_selection_text_via_clipboard,
            apply_text_to_selection_via_clipboard,
            _call_ai_server,
        )  # type: ignore

        instr = self.input_edit.text().strip()
        if instr:
            self.log(f"[사용자-선택] {instr}")

        try:
            # 마지막 선택 텍스트가 있으면 그걸 쓰고, 없으면 다시 시도
            sel_text = self.last_selection_text or get_selection_text_via_clipboard()
            if not sel_text:
                self.log("[INFO] 선택된 텍스트가 없어서 다듬기를 수행하지 않습니다.")
                return

            prompt_parts = []
            prompt_parts.append("너는 한글(HWP) 문서 일부를 다듬는 한국어 편집 어시스턴트야.")
            prompt_parts.append("\n[수정 대상 텍스트]\n" + sel_text)
            if instr:
                prompt_parts.append("\n[사용자 요청]\n" + instr)
            else:
                prompt_parts.append("\n[사용자 요청]\n위 텍스트를 의미는 유지하면서 자연스럽게 다듬어줘.")

            full_prompt = "\n".join(prompt_parts)

            self.log("[INFO] 선택 영역 다듬기 시작...")

            new_text = _call_ai_server(full_prompt, mode="rewrite")
            apply_text_to_selection_via_clipboard(new_text)
            self.log("[INFO] 선택 영역 다듬기 완료.")
        except Exception as e:
            self.log(f"[ERROR] 선택 영역 다듬기 실패: {e}")


def main():
    app = QApplication(sys.argv)

    # 라이트 테마 (GPT 스타일에 가까운 밝은 UI)
    app.setStyleSheet(
        """
        QWidget {
            background-color: #FFFFFF;
            color: #202124;
            font-family: 'Segoe UI', sans-serif;
            font-size: 10pt;
        }
        QFrame#LeftPanel {
            background-color: #F4F4F7;
            border-right: 1px solid #E0E0E0;
        }
        QLabel#AppTitle {
            font-size: 16pt;
            font-weight: 600;
        }
        QLabel#StatusLabel {
            color: #5F6368;
        }
        QLineEdit {
            background-color: #FFFFFF;
            border: 1px solid #DADCE0;
            border-radius: 6px;
            padding: 6px 8px;
        }
        QTextEdit {
            background-color: #FFFFFF;
            border: 1px solid #E0E0E0;
            border-radius: 6px;
        }
        QPushButton {
            background-color: #F1F3F4;
            border: 1px solid #DADCE0;
            border-radius: 6px;
            padding: 6px 10px;
        }
        QPushButton:hover {
            background-color: #E8EAED;
        }
        QPushButton:disabled {
            background-color: #F8F9FA;
            color: #BDC1C6;
            border-color: #E0E0E0;
        }
        QPushButton#PrimaryButton {
            background-color: #1A73E8;
            border-color: #1A73E8;
            color: white;
        }
        QPushButton#PrimaryButton:hover {
            background-color: #4285F4;
        }
        QPushButton#SecondaryButton {
            background-color: #F1F3F4;
        }
        QSplitter::handle {
            background-color: #E0E0E0;
            width: 1px;
        }
        """
    )

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
