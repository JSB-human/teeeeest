import sys
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QPushButton,
    QHBoxLayout,
    QVBoxLayout,
    QLabel,
)
from PySide6.QtCore import Qt


class FloatingApproveDialog(QWidget):
    def __init__(self, instruction):
        super().__init__()
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint | Qt.Tool)
        self.setStyleSheet("""
            QWidget { background-color: #2c3e50; color: white; border-radius: 10px; font-family: 'Malgun Gothic'; border: 2px solid #34495e; }
            QLabel { padding: 15px; font-size: 14px; }
            QPushButton#approve { background-color: #27ae60; border-radius: 5px; padding: 10px; font-weight: bold; min-width: 80px; }
            QPushButton#cancel { background-color: #c0392b; border-radius: 5px; padding: 10px; min-width: 80px; }
        """)

        layout = QVBoxLayout()
        layout.addWidget(QLabel(instruction))

        btn_layout = QHBoxLayout()
        btn_approve = QPushButton("승인 (적용)")
        btn_approve.setObjectName("approve")
        btn_approve.clicked.connect(self.approve)

        btn_cancel = QPushButton("거절 (취소)")
        btn_cancel.setObjectName("cancel")
        btn_cancel.clicked.connect(self.cancel)

        btn_layout.addWidget(btn_approve)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)
        self.setLayout(layout)

        screen = QApplication.primaryScreen().geometry()
        self.move(screen.width() - 400, screen.height() - 250)

    def approve(self):
        sys.exit(0)  # 승인 성공 코드

    def cancel(self):
        sys.exit(1)  # 거절 코드


if __name__ == "__main__":
    app = QApplication(sys.argv)
    msg = sys.argv[1] if len(sys.argv) > 1 else "AI 수정을 적용하시겠습니까?"
    dialog = FloatingApproveDialog(msg)
    dialog.show()
    sys.exit(app.exec())
