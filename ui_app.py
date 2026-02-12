import sys
import os
from pathlib import Path
import datetime

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
from PySide6.QtGui import QTextCursor

# src를 import 경로에 추가
BASE_DIR = Path(__file__).parent
SRC_DIR = BASE_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.append(str(SRC_DIR))

from tools.engine import (  # type: ignore
    connect_document,
    get_current_document_path,
    rewrite_current_document,
    smart_fill_table_from_json,
    text_to_table_json,
    ensure_connected,
    get_selection_text_via_clipboard,
    get_cursor_position_meta,
    apply_planned_table_action,
    create_selection_changeset,
    preview_selection_changeset,
    create_table_changeset,
    preview_table_changeset,
    approve_changeset,
    reject_changeset,
    get_changeset_diff_summary,
)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("HwpInlineAI (HWP + AI Editor)")
        self.setMinimumSize(1000, 600)

        # 상태 관리 변수
        self.last_selection_text: str = ""
        self._modification_mode: str = None  # 'table' 또는 'selection'
        self._current_changeset_id: str = ""

        # ---- UI 구성 ----
        self.init_ui()
        
        # 시그널 연결
        self.connect_signals()

        self.log("[SYSTEM] HwpInlineAI v1.2 — 준비 완료.")

    def init_ui(self):
        # 메인 레이아웃
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        splitter = QSplitter(Qt.Horizontal)

        # ---- 좌측 패널 ----
        left_frame = QFrame(objectName="LeftPanel")
        left_layout = QVBoxLayout(left_frame)
        left_layout.setContentsMargins(15, 20, 15, 20)
        left_layout.setSpacing(12)

        self.app_title = QLabel("HwpInlineAI", objectName="AppTitle")
        
        # 상태 표시창
        status_box = QFrame(objectName="StatusContainer")
        status_box_layout = QVBoxLayout(status_box)
        self.status_label = QLabel("○ Disconnected", objectName="StatusLabel")
        self.path_label = QLabel("연결된 파일 없음", objectName="PathLabel")
        status_box_layout.addWidget(self.status_label)
        status_box_layout.addWidget(self.path_label)

        # 파일 연결부
        self.path_edit = QLineEdit(placeholderText="한글 파일 경로...")
        self.browse_button = QPushButton("📂 파일 선택")
        self.connect_button = QPushButton("🔗 한글 연결", objectName="PrimaryButton")

        # 기능 버튼 그룹
        group_doc = QLabel("📄 문서 전체", objectName="GroupLabel")
        self.send_button = QPushButton("전체 문서 다듬기")
        
        group_sel = QLabel("🎯 선택 영역", objectName="GroupLabel")
        self.sel_get_button = QPushButton("선택 영역 가져오기")
        self.sel_rewrite_button = QPushButton("✨ 선택 영역 다듬기", objectName="PrimaryButton")
        self.sel_to_table_button = QPushButton("📊 선택 → 표 생성")

        group_table = QLabel("📅 표 제어", objectName="GroupLabel")
        self.table_fill_button = QPushButton("📥 입력 → 표 채우기")
        self.table_preview_button = QPushButton("🔍 표 수정 미리보기")

        # 레이아웃 배치
        left_layout.addWidget(self.app_title)
        left_layout.addWidget(status_box)
        left_layout.addSpacing(10)
        left_layout.addWidget(self.path_edit)
        left_layout.addWidget(self.browse_button)
        left_layout.addWidget(self.connect_button)
        left_layout.addSpacing(15)
        
        left_layout.addWidget(group_doc)
        left_layout.addWidget(self.send_button)
        left_layout.addSpacing(5)
        
        left_layout.addWidget(group_sel)
        left_layout.addWidget(self.sel_get_button)
        left_layout.addWidget(self.sel_rewrite_button)
        left_layout.addWidget(self.sel_to_table_button)
        left_layout.addSpacing(5)
        
        left_layout.addWidget(group_table)
        left_layout.addWidget(self.table_fill_button)
        left_layout.addWidget(self.table_preview_button)
        
        left_layout.addStretch(1)

        # ---- 우측 패널 ----
        right_frame = QFrame()
        right_layout = QVBoxLayout(right_frame)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(0)

        # 헤더 (선택 정보 표시)
        header_panel = QFrame(objectName="HeaderPanel")
        header_layout = QHBoxLayout(header_panel)
        self.selection_label = QLabel("📍 선택: 없음", objectName="SelectionText")
        header_layout.addWidget(self.selection_label)
        
        # Diff 요약 패널
        self.diff_summary = QTextEdit(objectName="DiffSummary")
        self.diff_summary.setReadOnly(True)
        self.diff_summary.setMaximumHeight(140)
        self.diff_summary.setPlaceholderText("변경 요약이 여기에 표시됩니다.")

        # 채팅 로그
        self.chat_log = QTextEdit(objectName="ChatLog")
        self.chat_log.setReadOnly(True)

        # 입력창 구역
        input_container = QFrame(objectName="InputContainer")
        input_layout = QVBoxLayout(input_container)
        self.input_edit = QLineEdit(objectName="MainInput", placeholderText="AI에게 시킬 내용을 입력하세요 (Enter)...")
        input_layout.addWidget(self.input_edit)

        # 승인/거절 패널 (숨김 상태)
        self.preview_action_frame = QFrame(objectName="PreviewPanel")
        preview_layout = QHBoxLayout(self.preview_action_frame)
        self.preview_action_label = QLabel("변경 사항을 확인하세요.")
        self.inline_apply_button = QPushButton("✅ 승인", objectName="ApplyButton")
        self.inline_cancel_button = QPushButton("❌ 거절", objectName="CancelButton")
        preview_layout.addWidget(self.preview_action_label, stretch=1)
        preview_layout.addWidget(self.inline_apply_button)
        preview_layout.addWidget(self.inline_cancel_button)
        self.preview_action_frame.setVisible(False)

        right_layout.addWidget(header_panel)
        right_layout.addWidget(self.diff_summary)
        right_layout.addWidget(self.chat_log, stretch=1)
        right_layout.addWidget(self.preview_action_frame)
        right_layout.addWidget(input_container)

        splitter.addWidget(left_frame)
        splitter.addWidget(right_frame)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([280, 720])
        
        main_layout.addWidget(splitter)

    def connect_signals(self):
        self.browse_button.clicked.connect(self.on_browse_clicked)
        self.connect_button.clicked.connect(self.on_connect_clicked)
        self.send_button.clicked.connect(self.on_send_clicked)
        self.sel_get_button.clicked.connect(self.on_sel_get_clicked)
        self.sel_rewrite_button.clicked.connect(self.on_sel_rewrite_clicked)
        self.sel_to_table_button.clicked.connect(self.on_sel_to_table_clicked)
        self.table_fill_button.clicked.connect(self.on_table_fill_clicked)
        self.table_preview_button.clicked.connect(self.on_table_preview_clicked)
        self.inline_apply_button.clicked.connect(self.on_apply_clicked)
        self.inline_cancel_button.clicked.connect(self.on_cancel_clicked)
        self.input_edit.returnPressed.connect(self.on_sel_rewrite_clicked)

    def log(self, message: str):
        now = datetime.datetime.now().strftime("%H:%M:%S")
        color = "#E8EAED"
        if "[ERROR]" in message: color = "#F28B82"
        elif "[INFO]" in message: color = "#8AB4F8"
        elif "[SYSTEM]" in message: color = "#9AA0A6"
        
        styled_msg = f'<p style="margin-bottom: 8px;"><span style="color: #5F6368;">[{now}]</span> <span style="color: {color};">{message}</span></p>'
        self.chat_log.append(styled_msg)
        self.chat_log.moveCursor(QTextCursor.End)

    def render_diff_summary(self, diff: dict):
        if not diff:
            self.diff_summary.setPlainText("변경 요약 없음")
            return

        kind = diff.get("kind", "unknown")
        if kind == "text":
            lines = [
                f"[TEXT] before={diff.get('chars_before', 0)} / after={diff.get('chars_after', 0)}",
                f"added={diff.get('chars_added', 0)}, removed={diff.get('chars_removed', 0)}",
            ]
            for i, s in enumerate(diff.get("sample_spans", [])[:5], start=1):
                lines.append(f"{i}. {s.get('tag')} | -{s.get('old','')} | +{s.get('new','')}")
            self.diff_summary.setPlainText("\n".join(lines))
            return

        if kind == "table":
            lines = [f"[TABLE] changed_cells={diff.get('changed_cells', 0)}"]
            for i, c in enumerate(diff.get("sample_cells", [])[:10], start=1):
                lines.append(f"{i}. (r{c.get('row')}, c{c.get('col')}): '{c.get('old','')}' -> '{c.get('new','')}'")
            self.diff_summary.setPlainText("\n".join(lines))
            return

        self.diff_summary.setPlainText(str(diff))

    def set_connected_ui(self, connected: bool):
        if connected:
            path = get_current_document_path() or "(알 수 없음)"
            self.path_label.setText(os.path.basename(path))
            self.status_label.setText("● Connected")
            self.status_label.setStyleSheet("color: #81C995; font-weight: bold;")
            self.connect_button.setEnabled(False)
            btns = [self.send_button, self.sel_get_button, self.sel_rewrite_button, 
                    self.sel_to_table_button, self.table_fill_button, self.table_preview_button]
            for b in btns: b.setEnabled(True)
        else:
            self.path_label.setText("연결된 파일 없음")
            self.status_label.setText("○ Disconnected")
            self.status_label.setStyleSheet("color: #9AA0A6;")
            self.connect_button.setEnabled(True)
            btns = [self.send_button, self.sel_get_button, self.sel_rewrite_button, 
                    self.sel_to_table_button, self.table_fill_button, self.table_preview_button]
            for b in btns: b.setEnabled(False)

    # ---- 슬롯 함수들 ----
    def on_browse_clicked(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "한글 파일 선택", "", "HWP Files (*.hwp);;All Files (*)")
        if file_path:
            self.path_edit.setText(file_path)

    def on_connect_clicked(self):
        path = self.path_edit.text().strip()
        if not path: return
        try:
            connect_document(path)
            self.log(f"[INFO] 문서 연결 성공: {os.path.basename(path)}")
            self.set_connected_ui(True)
        except Exception as e:
            self.log(f"[ERROR] 연결 실패: {e}")

    def on_sel_get_clicked(self):
        try:
            sel_text = get_selection_text_via_clipboard()
            if sel_text:
                self.last_selection_text = sel_text
                pos = get_cursor_position_meta()
                if pos:
                    self.selection_label.setText(f"📍 선택됨: 문단 {pos['para_id']}, 위치 {pos['char_pos']} ({len(sel_text)}자)")
                else:
                    self.selection_label.setText(f"📍 선택됨: {len(sel_text)}자")
                self.log("[INFO] 선택 영역 텍스트를 가져왔습니다.")
            else:
                self.log("[INFO] 선택된 영역이 없습니다.")
        except Exception as e:
            self.log(f"[ERROR] 가져오기 실패: {e}")

    def on_sel_rewrite_clicked(self):
        instr = self.input_edit.text().strip()
        try:
            sel_text = get_selection_text_via_clipboard()
            if not sel_text:
                self.log("[INFO] 다듬을 텍스트를 먼저 드래그하여 선택해 주세요.")
                return

            self.log("[INFO] AI가 문장을 다듬고 있습니다 (미리보기 모드)...")

            cs_id = create_selection_changeset(instr)
            preview_selection_changeset(cs_id)

            self._current_changeset_id = cs_id
            self._modification_mode = "selection"
            self.render_diff_summary(get_changeset_diff_summary(cs_id))

            self.preview_action_frame.setVisible(True)
            self.preview_action_label.setText("문장에서 변경 사항(빨강/초록)을 확인하세요.")
            self.log(f"[INFO] 미리보기가 생성되었습니다. 승인 또는 거절을 선택하세요. (id={cs_id[:8]})")

        except Exception as e:
            self.log(f"[ERROR] 실패: {e}")

    def on_table_preview_clicked(self):
        instr = self.input_edit.text().strip()
        if not instr:
            self.log("[INFO] 표를 어떻게 수정할지 입력창에 적어주세요.")
            return
        try:
            self.log(f"[INFO] 표 수정 계획 중: {instr}")
            cs_id = create_table_changeset(instr)
            msg = preview_table_changeset(cs_id)
            self._current_changeset_id = cs_id
            self._modification_mode = "table"
            self.render_diff_summary(get_changeset_diff_summary(cs_id))
            self.preview_action_frame.setVisible(True)
            self.preview_action_label.setText("표 수정 미리보기가 준비되었습니다.")
            self.log(f"[INFO] {msg} (id={cs_id[:8]})")
        except Exception as e:
            self.log(f"[ERROR] 실패: {e}")

    def on_apply_clicked(self):
        try:
            if not self._current_changeset_id:
                self.log("[INFO] 적용할 변경안이 없습니다.")
                return
            msg = approve_changeset(self._current_changeset_id)
            self.log(f"[INFO] {msg}")
        except Exception as e:
            self.log(f"[ERROR] 적용 실패: {e}")
        finally:
            self.preview_action_frame.setVisible(False)
            self._modification_mode = None
            self._current_changeset_id = ""
            self.diff_summary.setPlainText("변경 요약 없음")

    def on_cancel_clicked(self):
        try:
            if not self._current_changeset_id:
                self.log("[INFO] 취소할 변경안이 없습니다.")
                return
            msg = reject_changeset(self._current_changeset_id)
            self.log(f"[INFO] {msg}")
        except Exception as e:
            self.log(f"[ERROR] 취소 실패: {e}")
        finally:
            self.preview_action_frame.setVisible(False)
            self._modification_mode = None
            self._current_changeset_id = ""
            self.diff_summary.setPlainText("변경 요약 없음")

    # 단순한 기능들
    def on_send_clicked(self):
        if QMessageBox.question(self, "확인", "전체 문서를 AI로 다듬으시겠습니까?", QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            try:
                self.log("[INFO] 전체 문서 재작성 시작...")
                rewrite_current_document("rewrite")
                self.log("[INFO] 완료.")
            except Exception as e: self.log(f"[ERROR] 실패: {e}")

    def on_sel_to_table_clicked(self):
        sel_text = get_selection_text_via_clipboard()
        if not sel_text: return
        try:
            self.log("[INFO] 표 생성 중...")
            msg = apply_planned_table_action(sel_text, self.input_edit.text())
            self.log(f"[INFO] 결과: {msg}")
        except Exception as e: self.log(f"[ERROR] 실패: {e}")

    def on_table_fill_clicked(self):
        raw_text = self.input_edit.text().strip()
        if not raw_text: return
        try:
            json_str = text_to_table_json(raw_text)
            msg = smart_fill_table_from_json(json_str)
            self.log(f"[INFO] 표 채우기: {msg}")
        except Exception as e: self.log(f"[ERROR] 실패: {e}")


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
            padding: 5px;
        }
        QLabel#StatusLabel { font-size: 9pt; font-weight: bold; }
        QLabel#PathLabel { font-size: 8pt; color: #9AA0A6; }
        QLabel#GroupLabel {
            font-size: 8pt; font-weight: bold; color: #8AB4F8;
            margin-top: 15px; text-transform: uppercase;
        }
        QLineEdit {
            background-color: #35363A; border: 1px solid #5F6368;
            border-radius: 6px; padding: 8px; color: #E8EAED;
        }
        QLineEdit:focus { border: 1px solid #8AB4F8; }
        QLineEdit#MainInput {
            background-color: #303134; border-radius: 20px;
            padding: 10px 20px; font-size: 10.5pt;
        }
        QTextEdit#DiffSummary {
            background-color: #1B1C1F;
            border-bottom: 1px solid #3C4043;
            padding: 10px 14px;
            font-size: 9pt;
        }
        QTextEdit#ChatLog {
            background-color: #202124; border: none;
            padding: 20px; line-height: 1.6;
        }
        QPushButton {
            background-color: #3C4043; border: 1px solid #5F6368;
            border-radius: 6px; padding: 8px; color: #E8EAED;
        }
        QPushButton:hover { background-color: #4F5256; }
        QPushButton#PrimaryButton {
            background-color: #8AB4F8; color: #202124; border: none; font-weight: bold;
        }
        QPushButton#PrimaryButton:hover { background-color: #AECBFA; }
        QFrame#HeaderPanel {
            background-color: #202124; border-bottom: 1px solid #3C4043;
            padding: 8px 20px;
        }
        QLabel#SelectionText { color: #8AB4F8; font-size: 9pt; }
        QFrame#PreviewPanel {
            background-color: #3367D6; padding: 10px 20px;
        }
        QLabel#PreviewLabel { color: white; font-weight: bold; }
        QPushButton#ApplyButton {
            background-color: #81C995; color: #202124; border: none; font-weight: bold; min-width: 80px;
        }
        QPushButton#CancelButton {
            background-color: #F28B82; color: #202124; border: none; font-weight: bold; min-width: 80px;
        }
        QSplitter::handle { background-color: #3C4043; }
    """)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
