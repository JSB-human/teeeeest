import sys
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QPushButton,
    QHBoxLayout,
    QVBoxLayout,
    QLabel,
)
from PySide6.QtCore import Qt, Signal


class FloatingApproveDialog(QWidget):
    # 사용자의 선택을 외부로 알릴 시그널
    choice_made = Signal(str)  # "approve" or "cancel"

    def __init__(self, instruction="AI가 제안한 수정을 적용할까요?"):
        super().__init__()
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setStyleSheet("""
            QWidget {
                background-color: #2c3e50;
                color: white;
                border-radius: 10px;
                font-family: 'Malgun Gothic';
            }
            QLabel {
                padding: 10px;
                font-size: 14px;
            }
            QPushButton#approve {
                background-color: #27ae60;
                border-radius: 5px;
                padding: 8px;
                font-weight: bold;
            }
            QPushButton#cancel {
                background-color: #c0392b;
                border-radius: 5px;
                padding: 8px;
            }
        """)

        layout = QVBoxLayout()

        label = QLabel(instruction)
        layout.addWidget(label)

        btn_layout = QHBoxLayout()

        self.btn_approve = QPushButton("승인 (적용)")
        self.btn_approve.setObjectName("approve")
        self.btn_approve.clicked.connect(self.on_approve)

        self.btn_cancel = QPushButton("거절 (취소)")
        self.btn_cancel.setObjectName("cancel")
        self.btn_cancel.clicked.connect(self.on_cancel)

        btn_layout.addWidget(self.btn_approve)
        btn_layout.addWidget(self.btn_cancel)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

        # 화면 우측 하단에 배치
        screen = QApplication.primaryScreen().geometry()
        self.move(screen.width() - 350, screen.height() - 200)

    def on_approve(self):
        self.choice_made.emit("approve")
        self.close()

    def on_cancel(self):
        self.choice_made.emit("cancel")
        self.close()


def show_approve_dialog(instruction="AI 수정을 적용하시겠습니까?"):
    app = QApplication.instance() or QApplication(sys.argv)
    dialog = FloatingApproveDialog(instruction)
    dialog.show()

    # 이 함수는 별도 프로세스나 스레드에서 실행되어야 메인 서버가 멈추지 않습니다.
    return dialog, app
