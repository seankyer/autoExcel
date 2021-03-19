# Script by Sean Kyer github.com/seankyer
# Version 0.0.1 2021-03-18
# Project Description:
#     Create fields for users to input variable data and have Python input it directly into generated excel file

import os
import sys
import xlsxwriter
import random
from PySide6 import QtCore, QtWidgets, QtGui


# Qt GUI
class MyWidget(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.home_path = os.path.join(os.environ["HOMEPATH"], "Desktop")
        self.textEdit = QtWidgets.QLineEdit("C:\\Users\\seana\\Desktop")
        self.instructionForTextEdit = QtWidgets.QLabel("Enter Desired Save Path")
        self.button = QtWidgets.QPushButton("Generate Excel Sheet!")
        self.text = QtWidgets.QLabel("Auto Excel", alignment=QtCore.Qt.AlignCenter)

        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.instructionForTextEdit)
        self.layout.addWidget(self.textEdit)
        self.layout.addWidget(self.text)
        self.layout.addWidget(self.button)

        self.button.clicked.connect(build_excel())


# Builds the excel sheet
# Args:
# Outputs: xlsx file to designated directory.
def build_excel():
    workbook = xlsxwriter.Workbook("C:\\Users\\seana\\Desktop")  # test directory
    worksheet = workbook.add_worksheet()

    # Style Stuff:
    bold = workbook.add_format({'bold': True})  # Adds bold specification
    worksheet.set_column('A:A', 20)  # Widens first column

    # Add a bold format to use to highlight cells.

    # Write some simple text.
    worksheet.write('A1', 'Hello')

    # Text with formatting.
    worksheet.write('A2', 'World', bold)

    # Write some numbers, with row/column notation.
    worksheet.write(2, 0, 123)
    worksheet.write(3, 0, 123.456)

    workbook.close()


if __name__ == '__main__':
    print("Launching Excel Writer")
    app = QtWidgets.QApplication([])

    widget = MyWidget()
    widget.resize(800, 600)
    widget.show()

    sys.exit(app.exec_())
