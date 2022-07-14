# @Author: Ethan Meyers
# @Email: ewm1230@gmail.com
# @Phone: 847-212-2264
# @Date: 06/13/2022

# This script takes a path to an Excel file
# and outputs an Excel sheet that filters all names
# that are on the 1099 driver list

import os
import sys
import pandas as pd
from os.path import exists
import openpyxl
from openpyxl.styles import numbers, PatternFill
from datetime import datetime, timedelta


class Data:
    dir_path = '%s\\HOSFilter\\' % os.environ['APPDATA']
    w2Path = '%sDrivers.xlsx' % dir_path


def loadData():
    dir_path = '%s\\HOSFilter\\' % os.environ['APPDATA']
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)

    file_path = '%sDrivers.xlsx' % dir_path
    if not exists(file_path):
        print('List of W2 Drivers not found, please enter new list!')
        copyXLSX(input("Enter new W2 list: "))


def copyXLSX(FilePath):
    if '"' in FilePath:
        FilePath = FilePath.replace('"', '')
    df = pd.read_excel(FilePath, sheet_name='Sheet1', header=0)
    writer = pd.ExcelWriter(Data.w2Path, engine='openpyxl')
    df.columns = ["W2 Drivers"]
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()


def getW2():
    if not exists(Data.w2Path):
        loadData()

    drivers = pd.read_excel(Data.w2Path, sheet_name='Sheet1', converters={'Hours': str}, header=0)
    drivers['W2 Drivers'] = drivers['W2 Drivers'].str.lower()
    return drivers


def deleteRows(FilePath):
    df = pd.read_excel(FilePath, sheet_name='Summary', header=7, index_col=False, date_parser="%H:%M:%S")
    df.drop('Miles', axis=1, inplace=True)
    df.drop('Date', axis=1, inplace=True)
    df = df[df['Full Driver Name'].str.contains("Total") == True]
    for i in range(0, len(df.index)):
        index = df.iloc[i]['Full Driver Name'].index('Total')
        df.iloc[i]['Full Driver Name'] = df.iloc[i]['Full Driver Name'][0:index-1].upper()
    df.drop(df.tail(1).index, inplace=True)

    print(df)

    writer = pd.ExcelWriter(FilePath, engine='openpyxl')
    df.to_excel(writer, sheet_name='Output', index=False)
    writer.save()


if __name__ == '__main__':
    loadData()
    deleteRows(r"C:\Users\emeyers\Desktop\GigiPayroll - Copy.xlsx")
    # print("Instructions: ", "Before running the program make sure the relevant Excel file is closed", "1) Find the timecard sheet in file-explorer or on the desktop", "2) Right click the file and select, \"Copy\"", "3) Right click in file-explorer or on the desktop and select, \"Paste\"", "4) Hold shift then right click the copied file and select, \"Copy as path\"", sep='\n')
    # print("5) You've just copied the file-path to your clipboard, press \"CTRL V\" and paste the path below. Then press enter")
    # path = input("Paste on this line here: ")
    # if '"' in path:
    #     path = path.replace('"', '')
    # elif "-v" == path:
    #     print(getW2())
    #     input("\nPress enter to finish: ")
    #     sys.exit("Finished!")
    # elif "-c" == path:
    #     print('W2 list amendment mode, please enter new list!')
    #     print('All data must be on a sheet titled "Sheet1" and have a header at "A0" followed by the data')
    #     copyXLSX(input("Enter new W2 list: "))
    #     input("\nPress enter to finish: ")
    #     sys.exit("Finished!")
    # try:
    #     print("Formatting File, Please Wait!")
    #     deleteRows(path)
    #     # formatCols(path)
    #     print("Task Completed!")
    # except Exception as e:
    #     if "Errno 13" in str(e):
    #         print("Please close the file and try again!")
    #     else:
    #         print("An unknown error occurred:", e)
    # input("\nPress enter to finish: ")