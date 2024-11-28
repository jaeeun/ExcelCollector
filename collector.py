import sys
from functions import collect, parse_input
from PyQt5.QtCore import QSettings
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QTextEdit, QLineEdit, QPushButton, QLabel


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel data Collector")
        self.setGeometry(100, 100, 400, 300)

        # 중앙 위젯과 레이아웃 생성
        central_widget = QWidget()
        layout = QVBoxLayout()

        # QLineEdit 추가
        self.folder_line_edit = QLineEdit(self)
        self.folder_line_edit.setPlaceholderText("엑셀 파일이 있는 폴더 이름")
        layout.addWidget(self.folder_line_edit)

        # QTextEdit 추가
        self.locations_text_edit = QTextEdit(self)
        self.locations_text_edit.setPlaceholderText("행,열을 쓰시오(띄어쓰기 유의)\n(예시)\nA1 A2 AB42")
        layout.addWidget(self.locations_text_edit)

        # QLineEdit 추가
        self.outfile_line_edit = QLineEdit(self)
        self.outfile_line_edit.setPlaceholderText("저장할 파일 이름")
        layout.addWidget(self.outfile_line_edit)

        # Submit 버튼 추가
        self.submit_button = QPushButton("데이터 모으기", self)
        self.submit_button.clicked.connect(self.submit_action)
        layout.addWidget(self.submit_button)

        # 상태 표시용
        self.state_label = QLabel("상태", self)
        self.state_label.setStyleSheet("color: red;")  # 글씨 색상 설정
        layout.addWidget(self.state_label)

        # 중앙 위젯 설정
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        self.load_data()

    def submit_action(self):
        folder_path = self.folder_line_edit.text()
        locations = self.locations_text_edit.toPlainText()
        out_file = self.outfile_line_edit.text()

        print(f"folder_path: {folder_path}\nlocations:{locations}\nout_file:{out_file}")
        locations = parse_input(locations)
        # print(locations)
        collect(folder_path, locations, out_file, self.state_label)
        self.save_data()

    def save_data(self):
        """QLineEdit 값을 QSettings에 저장"""
        settings = QSettings("MyApp", "ExcelCollector")

        settings.setValue("folder_path", self.folder_line_edit.text())
        settings.setValue("locations", self.locations_text_edit.toPlainText())
        settings.setValue("out_file", self.outfile_line_edit.text())

        print("QSettings 저장 완료",)
        print("folder_path : ", self.folder_line_edit.text())
        print("locations : ", self.locations_text_edit.toPlainText())
        print("out_file : ", self.outfile_line_edit.text())

    def load_data(self):
        print("QSettings 로딩 완료",)
        settings = QSettings("MyApp", "ExcelCollector")

        folder_path = settings.value("folder_path", "")
        self.folder_line_edit.setText(folder_path)

        locations = settings.value("locations", "")
        self.locations_text_edit.setText(locations)

        out_file = settings.value("out_file", "")
        self.outfile_line_edit.setText(out_file)

# 앱 실행
app = QApplication(sys.argv)
window = MyWindow()
window.show()
app.exec()

