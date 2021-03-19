# Script by Sean Kyer github.com/seankyer
# Version 0.0.1 2021-03-18
# Project Description:
#     Create fields for users to input variable data and have Python input it directly into generated excel file

import os
import sys
import xlsxwriter


# Builds the excel sheet
# Args:
# Outputs: xlsx file to designated directory.
def build_excel():
    workbook = xlsxwriter.Workbook("C:/Users/seana/Desktop/test.xlsx")  # test directory
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
    build_excel()
