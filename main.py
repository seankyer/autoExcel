# Script by Sean Kyer github.com/seankyer
# Version 0.0.1 2021-03-18
# Project Description:
#     Create fields for users to input variable data and have Python input it directly into generated excel file

import os
import sys
import xlsxwriter
from PySide6 import QtCore, QtWidgets, QtGui


# Qt GUI
class MyWidget(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.home_path = os.path.join(os.environ["HOMEPATH"], "Desktop")
        self.file_path_entry = QtWidgets.QLineEdit(self.home_path)
        self.file_name_entry = QtWidgets.QLineEdit("test.xlsx")
        self.instruction_text = QtWidgets.QLabel("Enter Desired Save Path")
        self.button = QtWidgets.QPushButton("Generate Excel Sheet!")

        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.instruction_text)
        self.layout.addWidget(self.file_path_entry)
        self.layout.addWidget(self.file_name_entry)
        self.layout.addWidget(self.button)

        # Save path specified by user
        self.savePath = os.path.join(self.file_path_entry.text(), self.file_name_entry.text())

        # Call to build_excel
        self.button.clicked.connect(build_excel(self.savePath))


# Builds the excel sheet
# Args:
#   save_path: The filepath, including name of file and .xlsx suffix
# Outputs:
#   xlsx file to designated directory.
def build_excel(save_path):
    workbook = xlsxwriter.Workbook(save_path)  # test directory
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
