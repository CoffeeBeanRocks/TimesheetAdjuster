# @Author: Ethan Meyers
# Email: ewm1230@gmail.com
# Phone#: 847-212-2264
# @Date: 06/13/2022

# This script takes a path to an excel file
# and outputs an excel sheet that filters all names
# that are not on the acceptable 1099 driver list

import pandas as pd
from openpyxl import load_workbook
import openpyxl
from openpyxl.styles import numbers, PatternFill


def deleteRows(FilePath):
    # Use pandas library to get the data from the excel sheet
    df = pd.read_excel(FilePath, sheet_name='Import', header=1)
    df2 = pd.read_excel(FilePath, sheet_name='Sheet3', header=0)
    df = df[df['Login'].isin(df2['1099 Drivers'])]

    # Use pandas library to remove elements not in the 1099 drivers list
    ExcelWorkbook = load_workbook(FilePath)
    writer = pd.ExcelWriter(FilePath, engine='openpyxl')
    writer.book = ExcelWorkbook
    df.to_excel(writer, sheet_name='Output')

    writer.save()
    writer.close()

def formatCols(FilePath):
    workbook = openpyxl.load_workbook(FilePath)
    worksheet = workbook['Output']

    # Using openpyxl to format the new excel sheet
    for cell in worksheet['E']:
        cell.number_format = numbers.FORMAT_DATE_XLSX14
    for cell in worksheet['H']:
        cell.number_format = numbers.FORMAT_DATE_XLSX14

    for cell in worksheet['F']:
        cell.number_format = numbers.FORMAT_DATE_TIME1
    for cell in worksheet['I']:
        cell.number_format = numbers.FORMAT_DATE_TIME1

    worksheet.column_dimensions['C'].width = 20
    worksheet.column_dimensions['D'].width = 20
    worksheet.column_dimensions['E'].width = 20
    worksheet.column_dimensions['F'].width = 20
    worksheet.column_dimensions['G'].width = 20
    worksheet.column_dimensions['H'].width = 20
    worksheet.column_dimensions['I'].width = 20

    # Alternates filling each row based off of a new username
    yellowFill = PatternFill(start_color='00FFFF00', end_color='00FFFF00', fill_type='solid')
    currentName = worksheet['C2'].value
    fill = True
    for i in range(2, worksheet.max_row+1):
        if worksheet.cell(row=i, column=3).value == currentName and fill:
            for j in range(3, worksheet.max_column+1):
                worksheet.cell(row=i, column=j).fill = yellowFill
        else:
            if worksheet.cell(row=i, column=3).value != currentName:
                fill = not fill
                currentName = worksheet.cell(row=i, column=3).value
                if fill:
                    for j in range(3, worksheet.max_column+1):
                        worksheet.cell(row=i, column=j).fill = yellowFill

    workbook.save(FilePath)

if __name__ == '__main__':
    print("Enter the absolute file path of the excel file: ")
    path = input()
    if '"' in path:
        path = path.replace('"', '')
    deleteRows(path)
    formatCols(path)
    print("Task Completed!")


