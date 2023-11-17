from PyQt6.QtWidgets import QApplication, QVBoxLayout, QMainWindow, QPushButton, QLabel, QLineEdit, QToolTip, QWidget
from PyQt6 import QtGui

import sys
import pandas as pd
from math import isnan


def fill_blanks(FILENAME, OUTPUT_FILENAME):
    df = pd.read_excel(FILENAME)
    columns = df.columns
    dates = df[columns[0]].tolist()
    fps = df[columns[1]].tolist()
    fys = df[columns[2]].tolist()
    voucher_categories = df[columns[3]].tolist()
    voucher_no = df[columns[4]].tolist()

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

    df[columns[0]] = new_dates
    df[columns[1]] = new_fps
    df[columns[2]] = new_fys
    df[columns[3]] = new_voucher_categories
    df[columns[4]] = new_voucher_no

    try:
        df.to_excel(OUTPUT_FILENAME, index=False)
    except:
        print("Please ensure there is a file extension like .xlsx in the output file name.")
        print("Also, make sure the output file is not open, or the same name as the input file")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Kingdee Export Formatting Tool")
        self.username = ""
        self.output_filename = ""

        self.title = QLabel("Input File Name (with extension like .xlsx):")
        self.output_filename_title = QLabel(
            "Output File Name (with extension like .xlsx):")

        self.label = QLabel()
        self.input = QLineEdit()
        self.input.textChanged.connect(self.label.setText)
        self.input.textChanged.connect(self.on_text_change)

        self.output_filename_label = QLabel()
        self.output_filename_input = QLineEdit()
        self.output_filename_input.textChanged.connect(
            self.output_filename_label.setText)
        self.output_filename_input.textChanged.connect(
            self.on_output_filename_change)

        layout = QVBoxLayout()

        # input file
        layout.addWidget(self.title)
        layout.addWidget(self.input)
        layout.addWidget(self.label)

        # output file
        layout.addWidget(self.output_filename_title)
        layout.addWidget(self.output_filename_input)
        layout.addWidget(self.output_filename_label)

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
        fill_blanks(self.username, self.output_filename)

    def on_text_change(self):
        self.username = self.input.text()
        print(self.username)

    def on_output_filename_change(self):
        self.output_filename = self.output_filename_input.text()
        print(self.output_filename)


app = QApplication([])


window = MainWindow()
window.setWindowIcon(QtGui.QIcon('kingdee.ico'))
window.show()

app.exec()
