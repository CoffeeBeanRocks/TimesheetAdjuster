# @Author: Ethan Meyers
# Email: ewm1230@gmail.com
# Phone#: 847-212-2264
# @Date: 06/13/2022

# This script takes a path to an Excel file
# and outputs an Excel sheet that filters all names
# that are on the 1099 driver list

import os
import sys
import pandas as pd
import pickle
from openpyxl import load_workbook
import numpy as np
import openpyxl
from openpyxl.styles import numbers, PatternFill
from datetime import datetime, timedelta

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.environ.get("_MEIPASS2",os.path.abspath("."))

    return os.path.join(base_path, relative_path)

# TODO: Update W4 from new excel sheet, then write that list to the txt
def excelToTxt():
    print("The sheet name for the list with the drivers must be named \"Sheet1\" and there should be a header at A1 titled \"W4 Drivers\"")
    FilePath = input("Enter the path to the W4 drivers here: ")
    if '"' in FilePath:
        FilePath = FilePath.replace('"', '')
    df = pd.read_excel(FilePath, sheet_name='Sheet1', header=0)
    #np.savetxt(resource_path('usernames.txt'), df['W4 Drivers'].values, fmt='%s')
    with open(resource_path('mypickle.pk'), 'wb') as p:
        pickle.dump(df, p)

def printW4():
    with open(resource_path('mypickle.pk'), 'rb') as p:
        print(pickle.load(p))

def deleteW4():
    name = input("Enter the name to be removed: ")
    with open("usernames.txt", "r") as f:
        lines = f.readlines()
    with open("usernames.txt", "w") as f:
        for line in lines:
            if line.strip("\n").lower() != name.lower():
                f.write(line)
    printW4()

def addW4():
    with open('usernames.txt', 'a') as f:
        f.write(input("Enter New Name: ")+"\n")
        f.close()
    printW4()

def getW4():
    # TODO: Add JSON to add and subtract names
    logins = [
        'MADAMS',
        'MBATTEN ',
        'ACOBBS',
        'TDANIELUK',
        'VDAVIS',
        'FDOMINGUEZ',
        'JFRAGASSI',
        'AGILFORD',
        'CGILFORD',
        'HGLICKSON',
        'GHERNANDEZ',
        'JKANIEWSKI',
        'LKANIEWSKI',
        'HMADRID',
        'PMELHORN',
        'MNOWAKOWSKI',
        'OMARA',
        'MREED ',
        'KRIBAJ',
        'HSALGADO',
        'FSANCHEZ ',
        'CTHOMPSON',
        'VVALAC',
        'PWOJDILA',
        'RCOVER',
        'RGROVER',
        'JMARTIN',
        'JRUDE',
        'PWHALEN'
    ]
    for i in range(len(logins)):
        logins[i] = logins[i].lower()
    drivers = pd.DataFrame(logins, columns=['W4 Drivers'])
    return drivers

def deleteRows(FilePath):
    # Gets data from Excel sheet and removes elements that are in the 1099 drivers list
    df = pd.read_excel(FilePath, sheet_name='Duty Time', header=8)
    # df2 = pd.read_excel(FilePath, sheet_name='Sheet1', header=0)
    df2 = getW4()
    df = df[df['Login'].str.lower().isin(df2['W4 Drivers'])]

    # Removes empty columns
    df.replace("", "NaN", inplace=True)
    df.dropna(subset=['Login'], inplace=True)

    # Removes mystery total column
    del df['Total']
    del df['Unnamed: 0']

    # with pd.option_context('display.max_rows', None, 'display.max_columns', None):
    #     print(df)

    # Print DF
    writer = pd.ExcelWriter(FilePath, engine='openpyxl')
    df.to_excel(writer, sheet_name='Output')
    writer.save()


def formatCols(FilePath):
    workbook = openpyxl.load_workbook(FilePath)
    worksheet = workbook['Output']

    # Using openpyxl to format the new Excel sheet

    # Format Source Code: https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/styles/numbers.html
    for cell in worksheet['D']:
        cell.number_format = numbers.FORMAT_DATE_XLSX14
    for cell in worksheet['G']:
        cell.number_format = numbers.FORMAT_DATE_XLSX14

    for cell in worksheet['E']:
        cell.number_format = numbers.FORMAT_DATE_TIME2
    for cell in worksheet['H']:
        cell.number_format = numbers.FORMAT_DATE_TIME2

    worksheet.column_dimensions['B'].width = 20
    worksheet.column_dimensions['C'].width = 20
    worksheet.column_dimensions['D'].width = 20
    worksheet.column_dimensions['E'].width = 20
    worksheet.column_dimensions['F'].width = 20
    worksheet.column_dimensions['G'].width = 20
    worksheet.column_dimensions['H'].width = 20

    # Alternates filling each row based off of a new username
    yellowFill = PatternFill(start_color='00FFFF00', end_color='00FFFF00', fill_type='solid')
    currentName = worksheet['B2'].value
    colIndex = 2
    fill = False
    for i in range(2, worksheet.max_row+1):
        if worksheet.cell(row=i, column=colIndex).value == currentName and fill:
            for j in range(colIndex, worksheet.max_column+1):
                worksheet.cell(row=i, column=j).fill = yellowFill
        else:
            if worksheet.cell(row=i, column=colIndex).value != currentName:
                fill = not fill
                currentName = worksheet.cell(row=i, column=colIndex).value
                if fill:
                    for j in range(colIndex, worksheet.max_column+1):
                        worksheet.cell(row=i, column=j).fill = yellowFill

    # Adds time total column
    redFill = PatternFill(start_color='00FC0303', end_color='00FC0303', fill_type='solid')
    orangeFill = PatternFill(start_color='00FAAC02', end_color='00FAAC02', fill_type='solid')
    currentName = worksheet['B2'].value
    startIndex = 2
    fill = False
    for i in range(2, worksheet.max_row + 2):
        if worksheet.cell(row=i, column=colIndex).value != currentName:
            currentName = worksheet.cell(row=i, column=colIndex).value
            total = timedelta(hours=0)
            for j in range(startIndex, i):
                s1 = str(worksheet.cell(row=j, column=5).value)
                startTime = datetime.strptime(s1, '%Y-%m-%d %H:%M:%S')
                s2 = str(worksheet.cell(row=j, column=8).value)
                endTime = datetime.strptime(s2, '%Y-%m-%d %H:%M:%S')
                total = total + (endTime-startTime)
            worksheet.cell(row=i - 1, column=9).value = total
            if fill:
                worksheet.cell(row=i - 1, column=9).fill = orangeFill
            else:
                worksheet.cell(row=i - 1, column=9).fill = redFill
            fill = not fill
            startIndex = i

    workbook.save(FilePath)

if __name__ == '__main__':
    print("Instructions: ", "Before running the program make sure the relevant Excel file is closed", "1) Find the timecard sheet in file-explorer or on the desktop", "2) Right click the file and select, \"Copy\"", "3) Right click in file-explorer or on the desktop and select, \"Paste\"", "4) Hold shift then right click the copied file and select, \"Copy as path\"", sep='\n')
    print("5) You've just copied the file-path to your clipboard, press \"CTRL V\" and paste the path below. Then press enter")
    path = input("Paste on this line here: ")
    if '"' in path:
        path = path.replace('"', '')
    elif "-v" == path:
        printW4()
        # input("\nPress enter to finish: ")
        sys.exit("Finished!")
    elif "-e" == path:
        excelToTxt()
        #input("\nPress enter to finish: ")
        sys.exit("Finished!")
    try:
        print("Formatting File, Please Wait!")
        deleteRows(path)
        formatCols(path)
        print("Task Completed!")
    except Exception as e:
        if "Errno 13" in str(e):
            print("Please close the file and try again!")
        else:
            print("An unknown error occurred:", e)
    input("\nPress enter to finish: ")


