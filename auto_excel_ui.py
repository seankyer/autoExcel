# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'C:\Users\seana\PycharmProjects\autoExcel\auto_excel_ui.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_dialog(object):
    def setupUi(self, dialog):
        dialog.setObjectName("dialog")
        dialog.resize(375, 516)
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
        self.remove_selection_button.setGeometry(QtCore.QRect(244, 310, 111, 24))
        self.remove_selection_button.setObjectName("remove_selection_button")
        self.clear_button = QtWidgets.QPushButton(dialog)
        self.clear_button.setGeometry(QtCore.QRect(100, 400, 75, 24))
        self.clear_button.setObjectName("clear_button")
        self.add_input_button = QtWidgets.QPushButton(dialog)
        self.add_input_button.setGeometry(QtCore.QRect(190, 400, 75, 24))
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
        self.generate_excel_button.setGeometry(QtCore.QRect(100, 430, 161, 24))
        self.generate_excel_button.setObjectName("generate_excel_button")

        self.retranslateUi(dialog)
        QtCore.QMetaObject.connectSlotsByName(dialog)

    def retranslateUi(self, dialog):
        _translate = QtCore.QCoreApplication.translate
        dialog.setWindowTitle(_translate("dialog", "Auto Excel"))
        self.file_path_entry.setText(_translate("dialog", "testing"))
        self.directory_label.setText(_translate("dialog", "Save Path:"))
        self.file_name_entry.setText(_translate("dialog", "test.xlsx"))
        self.file_name_label.setText(_translate("dialog", "File Name:"))
        self.path_name_label.setText(_translate("dialog", "Auto Value will Be Variable"))
        self.data_label.setText(_translate("dialog", "Data"))
        self.prefix_label.setText(_translate("dialog", "Prefix:"))
        self.range_label.setText(_translate("dialog", "Range:"))
        self.suffix_label.setText(_translate("dialog", "Suffix:"))
        self.remove_selection_button.setText(_translate("dialog", "Remove Selection"))
        self.clear_button.setText(_translate("dialog", "Clear Input"))
        self.add_input_button.setText(_translate("dialog", "Add Input"))
        self.generate_excel_button.setText(_translate("dialog", "Generate Excel File"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    dialog = QtWidgets.QDialog()
    ui = Ui_dialog()
    ui.setupUi(dialog)
    dialog.show()
    sys.exit(app.exec_())
