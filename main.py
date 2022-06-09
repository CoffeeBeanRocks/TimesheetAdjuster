import pandas as pd
from openpyxl import load_workbook

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

# def formatCols():
#     df = pd.read_excel(r'C:\Users\emeyers\Desktop\Excel File.xlsx', sheet_name='Data', header=0)
#     workbook = xlsxwriter.Workbook(r'C:\Users\emeyers\Desktop\Excel File.xlsx')
#     worksheet = workbook.add_worksheet(name='Data')
#
#     format3 = workbook.add_format({'num_format': 'mm/dd/yy'})
#     worksheet.set_column(4, 4, 20, format3)
#

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print("Enter the absolute file path of the excel file: ")
    path = input()
    deleteRows(path)


