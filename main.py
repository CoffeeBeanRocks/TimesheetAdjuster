# @Author: Ethan Meyers
# Email: ewm1230@gmail.com
# Phone#: 847-212-2264
# @Date: 06/13/2022

# This script takes a path to an Excel file
# and outputs an Excel sheet that filters all names
# that are on the 1099 driver list

import pandas as pd
from openpyxl import load_workbook
import openpyxl
from openpyxl.styles import numbers, PatternFill


def get1099Drivrs():
    logins = [
        'cculp',
        'jlorden',
        'fherron',
        'ekotlarz',
        'ssetterlund',
        'kjohnson',
        'rgrover',
        'mmartin',
        'gkozlowski',
        'zkaczmarczyk',
        'bmihaylov',
        'kslupek',
        'ibuzinskis',
        'PWALAWSKI',
        'VDIMITROV',
        'mrudzinski',
        'aszymanski',
        'awiech',
        'mwiechetek',
        'awolak',
        'mzareba',
        'shicks',
        'sutterback',
        'ffrench',
        'jgrzesiak',
        'bgal',
        'jmitchell',
        'kpopek',
        'mlooney',
        'sgorczyca',
        'jhaynie',
        'epetrov',
        'lbronikowski',
        'rgranados',
        'kpodstawka',
        'rbanas',
        'khadera',
        'rpettis',
        'aplecki',
        'kwasowicz',
        'kbryja',
        'jsadkowski',
        'jlukacs',
        'miwaniec',
        'tcachro',
        'rsingh',
        'dgornikowski',
        'michaelj',
        'lbalinski',
        'rkokot',
        'dzajac',
        'sspear',
        'tbrown',
        'sbliznakov',
        'jfoxx',
        'skahlon',
        'bmontgomery',
        'jhernandez',
        'dbooker',
        'tvarela',
        'rwyrick',
        'rwade',
        'hhernandez',
        'jrzepka',
        'dfidowski',
        'nlewis',
        'mholder',
        'jchavez',
        'brobbins',
        'smartinez',
        'rpetraitis',
        'saddison',
        'tthompson',
        'eagbenyadzi',
        'iasenov',
        'ccortes',
        'omara',
        'jramirez',
        'aradon',
        'apatino',
        'awilliams',
        'psoja',
        'ccardona'
    ]
    drivers = pd.DataFrame(logins, columns=['1099 Drivers'])
    return drivers

def colored(r, g, b, text):
    return "\033[38;2;{};{};{}m{} \033[38;2;255;255;255m".format(r, g, b, text)

def deleteRows(FilePath):
    # Gets data from Excel sheet and removes elements that are in the 1099 drivers list
    df = pd.read_excel(FilePath, sheet_name='Duty Time', header=8)
    #df2 = pd.read_excel(FilePath, sheet_name='Sheet1', header=0)
    df2 = get1099Drivrs()
    df = df[~df['Login'].isin(df2['1099 Drivers'])]

    # Removes empty columns
    df.replace("", "NaN", inplace=True)
    df.dropna(subset=['Login'], inplace=True)

    # with pd.option_context('display.max_rows', None, 'display.max_columns', None):
    #     print(df)

    #Print DF
    writer = pd.ExcelWriter(FilePath, engine='openpyxl')  # TODO: REPLACES FILE!!!!
    df.to_excel(writer, sheet_name='Output')
    writer.save()


def formatCols(FilePath):
    workbook = openpyxl.load_workbook(FilePath)
    worksheet = workbook['Output']

    # Using openpyxl to format the new Excel sheet
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
    fill = False
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
    print("Starting HOS Timecard Creation, this may take a few moments...")
    try:
        deleteRows(path)
        formatCols(path)
        print("Task Completed!")
    except Exception as e:
        if "Errno 13" in str(e):
            print(colored(255, 0, 0, "Please close the file and try again!"))
        else:
            print(colored(255, 0, 0, "An unknown error occurred:"), e)
    input("Press enter to finish: ")


