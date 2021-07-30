# Script by Sean Kyer https://github.com/seankyer/autoExcel
# Version 1.0.0 2021-03-24
# Project Description:
#   Assist user in creating tediously repetitive excel sheets where data repeats, but is also variable. A user will
#   input their prefix and suffix, along with desired repetition range.
#
#   An example would be if you need to create a job that has name tags for products with unique identifiers but also
#   common names. AH-1 -> AH-99 would need to be created in a spread sheet. For the same job, BH-5 -> BH 509 must be
#   created. Instead of manually entering or 'with excel', you can input your list of variations and python
#   will construct the excel sheet for you. Saving potentially 15-35 minutes per job.

# To update UI:
#   Cmd Code:
#   python -m PyQt5.uic.pyuic -x C:\Users\seana\PycharmProjects\autoExcel\auto_excel_ui.ui -o C:\Users\seana\PycharmProjects\autoExcel\auto_excel_ui.py
#   WARNING: This will overwrite auto_excel_ui.py and all changes will be lost! Copy and paste updated ui to main.py
#            to avoid unwanted loss of work!

import os
import sys
import xlsxwriter
from PyQt5 import QtCore, QtGui, QtWidgets

EXECUTION_LIST = []


class ExcelItem:
    def __init__(self, prefix="", min_max=[], suffix=""):
        self.prefix = prefix
        self.min_max = min_max
        self.suffix = suffix


# Initial setup of GUI
class ui_dialog(object):
    def setupUi(self, dialog):
        dialog.setObjectName("Dialog")
        dialog.resize(375, 516)
        dialog.setFixedSize(375, 516)
        dialog.setMinimumSize(QtCore.QSize(0, 0))
        self.list_of_executions = QtWidgets.QListWidget(dialog)
        self.list_of_executions.setGeometry(QtCore.QRect(10, 110, 351, 191))
        self.list_of_executions.setAlternatingRowColors(True)
        self.list_of_executions.setObjectName("list_of_executions")
        self.file_path_entry = QtWidgets.QLineEdit(dialog)
        self.file_path_entry.setGeometry(QtCore.QRect(10, 30, 201, 22))
        self.file_path_entry.setObjectName("file_path_entry")
        self.directory_label = QtWidgets.QLabel(dialog)
        self.directory_label.setGeometry(QtCore.QRect(10, 10, 71, 16))
        font = QtGui.QFont()
        font.setBold(True)
        self.directory_label.setFont(font)
        self.directory_label.setObjectName("directory_label")
        self.file_name_entry = QtWidgets.QLineEdit(dialog)
        self.file_name_entry.setGeometry(QtCore.QRect(212, 30, 151, 22))
        self.file_name_entry.setObjectName("file_name_entry")
        self.file_name_label = QtWidgets.QLabel(dialog)
        self.file_name_label.setGeometry(QtCore.QRect(210, 10, 61, 16))
        font = QtGui.QFont()
        font.setBold(True)
        self.file_name_label.setFont(font)
        self.file_name_label.setObjectName("file_name_label")
        self.path_name_label = QtWidgets.QLabel(dialog)
        self.path_name_label.setGeometry(QtCore.QRect(10, 60, 351, 20))
        self.path_name_label.setObjectName("path_name_label")
        self.data_label = QtWidgets.QLabel(dialog)
        self.data_label.setGeometry(QtCore.QRect(170, 90, 49, 16))
        font = QtGui.QFont()
        font.setBold(True)
        self.data_label.setFont(font)
        self.data_label.setObjectName("data_label")
        self.prefix_entry = QtWidgets.QLineEdit(dialog)
        self.prefix_entry.setGeometry(QtCore.QRect(10, 370, 113, 22))
        self.prefix_entry.setObjectName("prefix_entry")
        self.suffix_entry = QtWidgets.QLineEdit(dialog)
        self.suffix_entry.setGeometry(QtCore.QRect(250, 370, 113, 22))
        self.suffix_entry.setObjectName("suffix_entry")
        self.min_spinBox = QtWidgets.QSpinBox(dialog)
        self.min_spinBox.setGeometry(QtCore.QRect(140, 370, 42, 22))
        self.min_spinBox.setMinimum(1)
        self.min_spinBox.setMaximum(9999)
        self.min_spinBox.setObjectName("min_spinBox")
        self.max_spinBox = QtWidgets.QSpinBox(dialog)
        self.max_spinBox.setGeometry(QtCore.QRect(190, 370, 42, 22))
        self.max_spinBox.setMinimum(1)
        self.max_spinBox.setMaximum(9999)
        self.max_spinBox.setObjectName("max_spinBox")
        self.prefix_label = QtWidgets.QLabel(dialog)
        self.prefix_label.setGeometry(QtCore.QRect(10, 350, 49, 16))
        font = QtGui.QFont()
        font.setBold(True)
        self.prefix_label.setFont(font)
        self.prefix_label.setObjectName("prefix_label")
        self.range_label = QtWidgets.QLabel(dialog)
        self.range_label.setGeometry(QtCore.QRect(140, 350, 49, 16))
        font = QtGui.QFont()
        font.setBold(True)
        self.range_label.setFont(font)
        self.range_label.setObjectName("range_label")
        self.suffix_label = QtWidgets.QLabel(dialog)
        self.suffix_label.setGeometry(QtCore.QRect(250, 350, 49, 16))
        font = QtGui.QFont()
        font.setBold(True)
        self.suffix_label.setFont(font)
        self.suffix_label.setObjectName("suffix_label")
        self.remove_selection_button = QtWidgets.QPushButton(dialog)
        self.remove_selection_button.setGeometry(QtCore.QRect(228, 310, 140, 24))
        self.remove_selection_button.setObjectName("remove_selection_button")
        self.clear_button = QtWidgets.QPushButton(dialog)
        self.clear_button.setGeometry(QtCore.QRect(91, 400, 90, 24))
        self.clear_button.setObjectName("clear_button")
        self.add_input_button = QtWidgets.QPushButton(dialog)
        self.add_input_button.setGeometry(QtCore.QRect(187, 400, 90, 24))
        self.add_input_button.setObjectName("add_input_button")
        self.error_info_label = QtWidgets.QLabel(dialog)
        self.error_info_label.setGeometry(QtCore.QRect(10, 460, 351, 51))
        font = QtGui.QFont()
        font.setUnderline(False)
        self.error_info_label.setFont(font)
        self.error_info_label.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.error_info_label.setText("")
        self.error_info_label.setObjectName("error_info_label")
        self.generate_excel_button = QtWidgets.QPushButton(dialog)
        self.generate_excel_button.setGeometry(QtCore.QRect(85, 430, 200, 24))
        self.generate_excel_button.setObjectName("generate_excel_button")

        self.retranslateUi(dialog)
        QtCore.QMetaObject.connectSlotsByName(dialog)

    # More GUI setup, setting texts and linking functions
    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Auto Excel"))
        self.home_path = "/Users/prepress-2/Desktop"  # Static user desktop
        self.file_path_entry.setText(_translate("Dialog", self.home_path))
        self.directory_label.setText(_translate("Dialog", "Save Path:"))
        self.file_name_entry.setText(_translate("Dialog", "test.xlsx"))
        self.save_path = os.path.join(self.file_path_entry.text(), self.file_name_entry.text())  # Initial instantiation
        self.file_name_label.setText(_translate("Dialog", "File Name:"))
        self.path_name_label.setText(_translate("Dialog", self.save_path))
        self.data_label.setText(_translate("Dialog", "Data"))
        self.prefix_label.setText(_translate("Dialog", "Prefix:"))
        self.range_label.setText(_translate("Dialog", "Range:"))
        self.suffix_label.setText(_translate("Dialog", "Suffix:"))
        self.remove_selection_button.setText(_translate("Dialog", "Remove Selection"))
        self.clear_button.setText(_translate("Dialog", "Clear Input"))
        self.add_input_button.setText(_translate("Dialog", "Add Input"))
        self.generate_excel_button.setText(_translate("Dialog", "Generate Excel File"))

        # Function Hookup
        self.file_path_entry.textChanged.connect(self.update_filepath)
        self.file_name_entry.textChanged.connect(self.update_filepath)
        self.add_input_button.clicked.connect(self.add_input)
        self.remove_selection_button.clicked.connect(self.remove_item)
        self.clear_button.clicked.connect(self.clear_input)
        self.generate_excel_button.clicked.connect(self.generate_excel)

    # Backend of script:

    # After checking base case parameters, builds sheet. If an error is found, post error message
    def generate_excel(self):
        if not EXECUTION_LIST:
            self.error_info_label.setText("No items to build!")
            return
        if os.path.isfile(self.save_path):
            self.error_info_label.setText("Error: File name already exists!")
            return
        if not self.file_name_entry.text().endswith(".xlsx"):
            self.error_info_label.setText("Ensure file name ends with '.xlsx'")
            return
        if not os.path.isdir(self.file_path_entry.text()):
            self.error_info_label.setText("File Path is invalid")
            return
        if os.path.isdir(self.file_path_entry.text()) and self.file_name_entry.text().endswith(".xlsx"):
            self.error_info_label.setText("")
            self.build_excel()

    # Called every key-stroke when editing file_path_entry or file_name_entry
    def update_filepath(self):
        self.save_path = os.path.join(self.file_path_entry.text(), self.file_name_entry.text())
        self.path_name_label.setText(self.save_path)

    # Sets inputs to default values for new object
    def clear_input(self):
        self.suffix_entry.setText("")
        self.prefix_entry.setText("")
        self.max_spinBox.setValue(1)
        self.min_spinBox.setValue(1)

    # Adds the specified input to the back of EXECUTION_LIST and to the back of list_of_executions
    def add_input(self):
        if self.min_spinBox.value() > self.max_spinBox.value():
            self.error_info_label.setText("Error: Range start must be less than range end!")
            return
        prefix = self.prefix_entry.text()
        min_max = list(range(self.min_spinBox.value(), self.max_spinBox.value() + 1))
        suffix = self.suffix_entry.text()
        item = ExcelItem(prefix, min_max, suffix)
        EXECUTION_LIST.append(item)
        item_name = (item.prefix + str(item.min_max[0]) + "-" + str(item.min_max[len(item.min_max) - 1]) + item.suffix)
        self.list_of_executions.addItem(item_name)

    # Removes selected item by index from list_of_executions and by same index from EXECUTION_LIST
    def remove_item(self):
        try:
            index = self.list_of_executions.row(self.list_of_executions.selectedItems()[0])
        except IndexError:
            self.error_info_label.setText("Error: No selection made!")
            return
        self.list_of_executions.takeItem(self.list_of_executions.row(self.list_of_executions.selectedItems()[0]))
        EXECUTION_LIST.pop(index)
        self.error_info_label.setText("")

    # Loops through EXECUTION_LIST, each item found increases column. For each given item, the item is repeated built as
    # many times as is specified by the min_max, increasing row each iteration
    def build_excel(self):

        workbook = xlsxwriter.Workbook(self.save_path)  # test directory
        worksheet = workbook.add_worksheet()

        column = -1

        for item in EXECUTION_LIST:
            column = column + 1
            row = 0
            for count in range(item.min_max[0], item.min_max[len(item.min_max) - 1] + 1):
                text = item.prefix + str(count) + item.suffix
                worksheet.write(row, column, text)
                row = row + 1

        workbook.close()
        print("Excel file generated at " + self.save_path)


if __name__ == '__main__':
    print("Launching Auto Excel")
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = ui_dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
