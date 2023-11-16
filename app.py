from PyQt6.QtWidgets import QApplication, QVBoxLayout, QMainWindow, QPushButton, QLabel, QLineEdit, QToolTip, QWidget

import sys
import pandas as pd
from math import isnan


def fill_blanks(FILENAME):
    df = pd.read_excel(FILENAME)
    dates = df['Date'].tolist()
    fps = df['FP'].tolist()
    fys = df['FY'].tolist()
    voucher_categories = df['Voucher Category'].tolist()
    voucher_no = df['Voucher No.'].tolist()

    def fill_blanks(list):
        current_element = list
        for index, value in enumerate(list):
            # if it is a datetime, isnan will fail and except branch is triggered.
            try:
                if isnan(value):
                    list[index] = current_element
                else:
                    current_element = value
            except:
                current_element = value
        return list

    new_dates = fill_blanks(dates)
    new_fps = fill_blanks(fps)
    new_fys = fill_blanks(fys)
    new_voucher_categories = fill_blanks(voucher_categories)
    new_voucher_no = fill_blanks(voucher_no)

    df['Date'] = new_dates
    df['FP'] = new_fps
    df['FY'] = new_fys
    df['Voucher Category'] = new_voucher_categories
    df['Voucher No.'] = new_voucher_no

    df.to_excel('output.xlsx')


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Kingdee Export Formatting Tool")
        self.username = ""

        self.label = QLabel()
        self.input = QLineEdit()
        self.input.textChanged.connect(self.label.setText)
        self.input.textChanged.connect(self.on_text_change)

        layout = QVBoxLayout()
        layout.addWidget(self.input)
        layout.addWidget(self.label)

        self.button = QPushButton("Fill Blanks")
        self.button.setFixedSize(150, 40)
        self.button.setCheckable(True)
        self.button.clicked.connect(self.on_button_clicked)

        layout.addWidget(self.button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        self.setFixedSize(400, 400)

    def on_button_clicked(self):
        print("boom", self.username)
        fill_blanks(self.username)

    def on_text_change(self):
        self.username = self.input.text()
        print(self.username)


app = QApplication([])

window = MainWindow()
window.show()

app.exec()
