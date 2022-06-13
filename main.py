# @Author: Ethan Meyers
# Email: ewm1230@gmail.com
# Phone#: 847-212-2264
# @Date: 06/13/2022

# This script takes a path to an excel file
# and outputs an excel sheet that filters all names
# that are not on the acceptable 1099 driver list

import pandas as pd
import xlsxwriter
from openpyxl import load_workbook
import openpyxl
from openpyxl.styles import numbers


def deleteRows(FilePath):
    df = pd.read_excel(FilePath, sheet_name='Import', header=1)
    df2 = pd.read_excel(FilePath, sheet_name='Sheet3', header=0)
    df = df[df['Login'].isin(df2['1099 Drivers'])]

    ExcelWorkbook = load_workbook(FilePath)
    writer = pd.ExcelWriter(FilePath, engine='openpyxl')
    writer.book = ExcelWorkbook
    df.to_excel(writer, sheet_name='Data')

    writer.save()
    writer.close()

def formatCols(FilePath):
    # df = pd.read_excel(r'C:\Users\emeyers\Desktop\Excel File.xlsx', sheet_name='Data', header=0)
    workbook = openpyxl.load_workbook(FilePath)
    worksheet = workbook['Data']

    # worksheet.column_dimensions['E'].number_format = numbers.FORMAT_DATE_XLSX14
    # worksheet.column_dimensions['H'].number_format = numbers.FORMAT_DATE_XLSX14

    worksheet.column_dimensions['C'].width = 20
    worksheet.column_dimensions['D'].width = 20
    worksheet.column_dimensions['E'].width = 20
    worksheet.column_dimensions['F'].width = 20
    worksheet.column_dimensions['G'].width = 20
    worksheet.column_dimensions['H'].width = 20
    worksheet.column_dimensions['I'].width = 20

    workbook.save(r'C:\Users\emeyers\Desktop\Excel File.xlsx')

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print("Enter the absolute file path of the excel file: ")
    path = input()
    if '"' in path:
        path = path.replace('"', '')
    deleteRows(path)
    formatCols(path)
    print("Finished!")


