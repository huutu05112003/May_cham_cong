from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import QSize
import sys

class MyApp(QWidget):
    def __init__(self):
        super().__init__()

        # Tạo nút bấm
        self.search_button = QPushButton(self)
        self.search_button.setText('Tìm kiếm')
        self.search_button.setIcon(QIcon('search.jfif'))  # Đường dẫn đến icon
        self.search_button.setIconSize(QSize(24, 24))  # Kích thước icon

        # Đặt layout
        layout = QVBoxLayout()
        layout.addWidget(self.search_button)
        self.setLayout(layout)

        # Thiết lập cửa sổ chính
        self.setWindowTitle('Tìm kiếm')
        self.setGeometry(100, 100, 300, 200)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec())
