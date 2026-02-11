import sys
import os
import datetime
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
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QFont, QColor, QTextCursor

# src를 import 경로에 추가 (엔진 모듈 사용)
BASE_DIR = Path(__file__).parent
SRC_DIR = BASE_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.append(str(SRC_DIR))

try:
    from tools.engine import (  # type: ignore
        cancel_table_modification,
        connect_document,
        finalize_table_modification,
        get_last_table_preview_cells,
        get_current_document_path,
        preview_current_table_modification,
        rewrite_current_document,
        smart_fill_table_from_json,
        text_to_table_json,
    )
except ImportError:
    # 엔진이 없는 환경에서도 UI는 뜨도록 예외 처리
    pass


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("HwpInlineAI — Modern HWP Editor")
        self.setMinimumSize(1000, 650)

        # ---- 좌측 패널: 파일 / 상태 / 액션 ----
        left_frame = QFrame(objectName="LeftPanel")
        left_layout = QVBoxLayout()
        left_layout.setContentsMargins(20, 20, 20, 20)
        left_layout.setSpacing(12)

        self.app_title = QLabel("HwpInlineAI")
        self.app_title.setObjectName("AppTitle")
        
        self.status_container = QFrame(objectName="StatusContainer")
        status_box = QVBoxLayout(self.status_container)
        status_box.setContentsMargins(10, 10, 10, 10)
        
        self.status_label = QLabel("○ Disconnected")
        self.status_label.setObjectName("StatusLabel")
        
        self.path_label = QLabel("연결된 파일 없음")
        self.path_label.setObjectName("PathLabel")
        self.path_label.setWordWrap(True)
        
        status_box.addWidget(self.status_label)
        status_box.addWidget(self.path_label)

        self.path_edit = QLineEdit()
        self.path_edit.setReadOnly(True)
        self.path_edit.setPlaceholderText("한글 파일을 선택하세요...")

        btn_row = QHBoxLayout()
        self.browse_button = QPushButton("📂 파일 선택")
        self.browse_button.setObjectName("SecondaryButton")
        self.connect_button = QPushButton("🔗 연결")
        self.connect_button.setObjectName("PrimaryButton")
        btn_row.addWidget(self.browse_button)
        btn_row.addWidget(self.connect_button)

        # 액션 그룹
        self.actions_label = QLabel("DOCUMENT ACTIONS")
        self.actions_label.setObjectName("GroupLabel")
        
        self.send_button = QPushButton("✨ 전체 문서 다듬기")
        self.send_button.setObjectName("ActionButton")
        self.send_button.setEnabled(False)
        
        self.sel_get_button = QPushButton("🔍 선택 영역 가져오기")
        self.sel_get_button.setObjectName("ActionButton")
        self.sel_get_button.setEnabled(False)
        
        self.sel_rewrite_button = QPushButton("📝 선택 영역 다듬기")
        self.sel_rewrite_button.setObjectName("ActionButton")
        self.sel_rewrite_button.setEnabled(False)

        self.table_label = QLabel("TABLE TOOLS")
        self.table_label.setObjectName("GroupLabel")

        self.sel_to_table_button = QPushButton("📊 선택 → 표 생성")
        self.sel_to_table_button.setObjectName("ActionButton")
        self.sel_to_table_button.setEnabled(False)

        self.table_fill_button = QPushButton("📥 입력 → 표 채우기")
        self.table_fill_button.setObjectName("ActionButton")
        self.table_fill_button.setEnabled(False)

        self.table_preview_button = QPushButton("👁️ 수정 미리보기")
        self.table_preview_button.setObjectName("ActionButton")
        self.table_preview_button.setEnabled(False)

        left_layout.addWidget(self.app_title)
        left_layout.addSpacing(10)
        left_layout.addWidget(self.status_container)
        left_layout.addSpacing(10)
        left_layout.addWidget(self.path_edit)
        left_layout.addLayout(btn_row)
        
        left_layout.addSpacing(20)
        left_layout.addWidget(self.actions_label)
        left_layout.addWidget(self.send_button)
        left_layout.addWidget(self.sel_get_button)
        left_layout.addWidget(self.sel_rewrite_button)
        
        left_layout.addSpacing(15)
        left_layout.addWidget(self.table_label)
        left_layout.addWidget(self.sel_to_table_button)
        left_layout.addWidget(self.table_fill_button)
        left_layout.addWidget(self.table_preview_button)
        left_layout.addStretch(1)

        left_frame.setLayout(left_layout)

        # ---- 우측 패널: 대화 / 로그 / 입력 ----
        right_frame = QFrame()
        right_layout = QVBoxLayout()
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(0)

        # 채팅 헤더
        header_frame = QFrame(objectName="HeaderPanel")
        header_layout = QHBoxLayout(header_frame)
        header_layout.setContentsMargins(20, 15, 20, 15)
        header_layout.addWidget(QLabel("Assistant Logs", objectName="HeaderText"))
        header_layout.addStretch(1)
        self.selection_label = QLabel("Current selection: None", objectName="SelectionText")
        header_layout.addWidget(self.selection_label)

        # 대화 로그
        self.chat_log = QTextEdit()
        self.chat_log.setReadOnly(True)
        self.chat_log.setObjectName("ChatLog")

        # 입력 영역
        input_container = QFrame(objectName="InputContainer")
        input_container_layout = QVBoxLayout(input_container)
        input_container_layout.setContentsMargins(20, 15, 20, 20)
        
        self.input_edit = QLineEdit()
        self.input_edit.setObjectName("MainInput")
        self.input_edit.setPlaceholderText("무엇을 도와드릴까요? 명령이나 메시지를 입력하세요...")
        self.input_edit.setFixedHeight(50)
        self.input_edit.returnPressed.connect(self.on_input_enter)

        input_container_layout.addWidget(self.input_edit)

        # 인라인 승인/거절 패널 (표 미리보기 후 표시)
        self.preview_action_frame = QFrame(objectName="PreviewPanel")
        preview_layout = QHBoxLayout()
        preview_layout.setContentsMargins(20, 12, 20, 12)
        preview_layout.setSpacing(15)

        self.preview_action_label = QLabel("✨ 표 수정 미리보기 생성됨")
        self.preview_action_label.setObjectName("PreviewLabel")

        self.inline_apply_button = QPushButton("적용하기")
        self.inline_apply_button.setObjectName("ApplyButton")
        self.inline_apply_button.setEnabled(False)

        self.inline_cancel_button = QPushButton("취소")
        self.inline_cancel_button.setObjectName("CancelButton")
        self.inline_cancel_button.setEnabled(False)

        preview_layout.addWidget(self.preview_action_label, stretch=1)
        preview_layout.addWidget(self.inline_apply_button)
        preview_layout.addWidget(self.inline_cancel_button)
        self.preview_action_frame.setLayout(preview_layout)
        self.preview_action_frame.setVisible(False)

        right_layout.addWidget(header_frame)
        right_layout.addWidget(self.chat_log, stretch=1)
        right_layout.addWidget(self.preview_action_frame)
        right_layout.addWidget(input_container)

        right_frame.setLayout(right_layout)

        # ---- 메인 Splitter ----
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left_frame)
        splitter.addWidget(right_frame)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([300, 700])
        splitter.setHandleWidth(1)

        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.addWidget(splitter)
        self.setLayout(main_layout)

        # 시그널 연결
        self.browse_button.clicked.connect(self.on_browse_clicked)
        self.connect_button.clicked.connect(self.on_connect_clicked)
        self.send_button.clicked.connect(self.on_send_clicked)
        self.sel_get_button.clicked.connect(self.on_sel_get_clicked)
        self.sel_rewrite_button.clicked.connect(self.on_sel_rewrite_clicked)
        self.sel_to_table_button.clicked.connect(self.on_sel_to_table_clicked)
        self.table_fill_button.clicked.connect(self.on_table_fill_clicked)
        self.table_preview_button.clicked.connect(self.on_table_preview_clicked)
        self.inline_apply_button.clicked.connect(self.on_table_apply_clicked)
        self.inline_cancel_button.clicked.connect(self.on_table_cancel_clicked)

        self.log("[SYSTEM] HwpInlineAI v1.1 — Ready.")

    # ---- 유틸 ----
    def log(self, message: str):
        now = datetime.datetime.now().strftime("%H:%M:%S")
        
        color = "#E8EAED"
        if "[ERROR]" in message: color = "#F28B82"
        elif "[INFO]" in message: color = "#8AB4F8"
        elif "[SYSTEM]" in message: color = "#9AA0A6"
        elif "[사용자]" in message: color = "#D2E3FC"

        styled_msg = f'<p style="margin-bottom: 8px;"><span style="color: #5F6368; font-family: monospace;">[{now}]</span> <span style="color: {color};">{message}</span></p>'
        self.chat_log.append(styled_msg)
        self.chat_log.moveCursor(QTextCursor.End)

    def set_connected_ui(self, connected: bool):
        if connected:
            path = get_current_document_path() or "(알 수 없음)"
            filename = os.path.basename(path)
            self.path_label.setText(filename)
            self.status_label.setText("● Connected")
            self.status_label.setStyleSheet("color: #81C995; font-weight: bold;")
            
            for btn in [self.send_button, self.sel_get_button, self.sel_rewrite_button, 
                        self.sel_to_table_button, self.table_fill_button, self.table_preview_button]:
                btn.setEnabled(True)
            self.connect_button.setEnabled(False)
        else:
            self.path_label.setText("연결된 파일 없음")
            self.status_label.setText("○ Disconnected")
            self.status_label.setStyleSheet("color: #9AA0A6;")
            for btn in [self.send_button, self.sel_get_button, self.sel_rewrite_button, 
                        self.sel_to_table_button, self.table_fill_button, self.table_preview_button]:
                btn.setEnabled(False)
            self.connect_button.setEnabled(True)

    # ---- 슬롯 ----
    def on_input_enter(self):
        text = self.input_edit.text().strip()
        if not text: return
        self.log(f"[사용자] {text}")
        self.input_edit.clear()
        # 여기에 추후 대화형 처리 로직 추가 가능

    def on_browse_clicked(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "한글 파일 선택", "", "HWP Files (*.hwp);;All Files (*)")
        if file_path:
            self.path_edit.setText(file_path)
            self.log(f"[INFO] 파일 선택됨: {os.path.basename(file_path)}")

    def on_connect_clicked(self):
        path = self.path_edit.text().strip()
        if not path: return
        try:
            connect_document(path, visible=True)
            self.log("[INFO] 한글 문서 연결 성공.")
            self.set_connected_ui(True)
        except Exception as e:
            self.log(f"[ERROR] 연결 실패: {e}")
            self.set_connected_ui(False)

    def on_send_clicked(self):
        reply = QMessageBox.question(self, "확인", "전체 문서를 AI로 다듬으시겠습니까?", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                self.log("[INFO] 전체 문서 재작성 시작...")
                rewrite_current_document("rewrite")
                self.log("[INFO] 완료되었습니다.")
            except Exception as e:
                self.log(f"[ERROR] 실패: {e}")

    def on_sel_get_clicked(self):
        from tools.engine import get_selection_text_via_clipboard, get_cursor_position_meta
        try:
            sel_text = get_selection_text_via_clipboard()
            if sel_text:
                self.last_selection_text = sel_text
                length = len(sel_text)
                
                # 커서 위치 메타데이터 가져오기
                pos = get_cursor_position_meta()
                if pos:
                    para_id = pos.get("para_id")
                    char_pos = pos.get("char_pos")
                    self.selection_label.setText(f"📍 선택됨: 문단 {para_id}, 위치 {char_pos} ({length}자)")
                else:
                    self.selection_label.setText(f"📍 선택됨: {length}자")
                
                self.log("[INFO] 선택 영역 텍스트 캡처 완료.")
            else:
                self.selection_label.setText("📍 선택: 없음")
                self.log("[INFO] 선택된 영역이 없습니다.")
        except Exception as e:
            self.log(f"[ERROR] 가져오기 실패: {e}")

    def on_sel_rewrite_clicked(self):
        from tools.engine import apply_text_to_selection_via_clipboard, _call_ai_server
        try:
            if not getattr(self, 'last_selection_text', None): return
            self.log("[INFO] 선택 영역 다듬기 중...")
            new_text = _call_ai_server(f"다듬어줘:\n{self.last_selection_text}", mode="rewrite")
            apply_text_to_selection_via_clipboard(new_text)
            self.log("[INFO] 완료.")
        except Exception as e:
            self.log(f"[ERROR] 실패: {e}")

    def on_sel_to_table_clicked(self):
        from tools.engine import apply_planned_table_action
        try:
            if not getattr(self, 'last_selection_text', None): return
            self.log("[INFO] 표 생성 계획 중...")
            msg = apply_planned_table_action(self.last_selection_text, "")
            self.log(f"[INFO] 결과: {msg}")
        except Exception as e:
            self.log(f"[ERROR] 실패: {e}")

    def on_table_fill_clicked(self):
        raw_text = self.input_edit.text().strip() # 입력창 내용 사용
        if not raw_text: 
            self.log("[INFO] 입력창에 데이터를 입력해주세요.")
            return
        try:
            json_str = text_to_table_json(raw_text)
            msg = smart_fill_table_from_json(json_str, has_header=True)
            self.log(f"[INFO] 표 채우기: {msg}")
        except Exception as e:
            self.log(f"[ERROR] 실패: {e}")

    def on_table_preview_clicked(self):
        instr = self.input_edit.text().strip()
        if not instr: return
        try:
            self.log(f"[INFO] 미리보기 생성 중: {instr}")
            msg = preview_current_table_modification(instr)
            if "Error" not in msg:
                self.preview_action_frame.setVisible(True)
                self.inline_apply_button.setEnabled(True)
                self.inline_cancel_button.setEnabled(True)
            self.log(f"[INFO] {msg}")
        except Exception as e:
            self.log(f"[ERROR] 실패: {e}")

    def on_table_apply_clicked(self):
        try:
            msg = finalize_table_modification()
            self.log(f"[INFO] 적용 완료: {msg}")
        finally:
            self.preview_action_frame.setVisible(False)

    def on_table_cancel_clicked(self):
        try:
            msg = cancel_table_modification()
            self.log(f"[INFO] 취소됨: {msg}")
        finally:
            self.preview_action_frame.setVisible(False)


def main():
    app = QApplication(sys.argv)
    
    # Modern Dark Theme StyleSheet
    app.setStyleSheet("""
        QWidget {
            background-color: #202124;
            color: #E8EAED;
            font-family: 'Segoe UI', 'Malgun Gothic', sans-serif;
            font-size: 10pt;
        }
        QFrame#LeftPanel {
            background-color: #2D2E31;
            border-right: 1px solid #3C4043;
        }
        QLabel#AppTitle {
            font-size: 18pt;
            font-weight: bold;
            color: #8AB4F8;
            margin-bottom: 5px;
        }
        QFrame#StatusContainer {
            background-color: #35363A;
            border-radius: 8px;
            border: 1px solid #3C4043;
        }
        QLabel#StatusLabel {
            font-size: 9pt;
            font-weight: bold;
        }
        QLabel#PathLabel {
            font-size: 8pt;
            color: #9AA0A6;
        }
        QLabel#GroupLabel {
            font-size: 8pt;
            font-weight: bold;
            color: #9AA0A6;
            padding-left: 2px;
            margin-top: 5px;
        }
        QLineEdit {
            background-color: #35363A;
            border: 1px solid #5F6368;
            border-radius: 6px;
            padding: 8px;
            color: #E8EAED;
        }
        QLineEdit:focus {
            border: 1px solid #8AB4F8;
        }
        QLineEdit#MainInput {
            background-color: #303134;
            border: 1px solid #5F6368;
            border-radius: 25px;
            padding-left: 20px;
            font-size: 11pt;
        }
        QTextEdit#ChatLog {
            background-color: #202124;
            border: none;
            padding: 15px;
            font-size: 10pt;
            line-height: 1.5;
        }
        QPushButton {
            background-color: #3C4043;
            border: 1px solid #5F6368;
            border-radius: 6px;
            padding: 8px 15px;
            color: #E8EAED;
            font-weight: 500;
        }
        QPushButton:hover {
            background-color: #4F5256;
        }
        QPushButton:pressed {
            background-color: #5F6368;
        }
        QPushButton:disabled {
            color: #5F6368;
            background-color: #2D2E31;
        }
        QPushButton#PrimaryButton {
            background-color: #8AB4F8;
            color: #202124;
            border: none;
        }
        QPushButton#PrimaryButton:hover {
            background-color: #AECBFA;
        }
        QPushButton#ActionButton {
            text-align: left;
            padding-left: 15px;
            background-color: transparent;
            border: 1px solid transparent;
        }
        QPushButton#ActionButton:hover {
            background-color: #3C4043;
            border: 1px solid #5F6368;
        }
        QFrame#HeaderPanel {
            background-color: #202124;
            border-bottom: 1px solid #3C4043;
        }
        QLabel#HeaderText {
            font-weight: bold;
            color: #9AA0A6;
        }
        QLabel#SelectionText {
            color: #8AB4F8;
            font-size: 9pt;
        }
        QFrame#PreviewPanel {
            background-color: #1A73E8;
            border-radius: 0px;
        }
        QLabel#PreviewLabel {
            color: white;
            font-weight: bold;
        }
        QPushButton#ApplyButton {
            background-color: white;
            color: #1A73E8;
            border: none;
        }
        QPushButton#CancelButton {
            background-color: transparent;
            color: white;
            border: 1px solid white;
        }
        QSplitter::handle {
            background-color: #3C4043;
        }
    """)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
