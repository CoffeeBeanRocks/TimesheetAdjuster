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
from datetime import datetime

# TODO: Create new file for output

# @Var dir_Path: Path to the folder containing the necessary files for this program
# @Var w2Path: Path to the file containing the names of the W2 drivers
# @Description: Class for necessary global variables
class Data:
    dir_path = '%s\\HOSFilter\\' % os.environ['APPDATA']
    w2Path = '%sDrivers.xlsx' % dir_path


# @Description: Verifies that the proper directories and files are in place
def loadData():
    # Creating proper directory
    if not os.path.exists(Data.dir_path):
        os.makedirs(Data.dir_path)

    # Checking for necessary W2 list and if not found, prompts the user to enter a new one
    if not exists(Data.w2Path):
        print('List of W2 Drivers not found, please enter new list!')
        copyXLSX(input("Enter new W2 list: "))


# @Param filePath: File path to the new Drivers.xlsx file provided by the user
# @Description: Fills the drivers.xlsx file with names given from a different file
def copyXLSX(filePath: str):
    # Reading input W2 file into dataframe
    if '"' in filePath:
        filePath = filePath.replace('"', '')
    df = pd.read_excel(filePath, sheet_name='Sheet1', header=0)

    # Saving dataframe to different W2 file
    writer = pd.ExcelWriter(Data.w2Path, engine='openpyxl')
    df.columns = ["W2 Drivers"]
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()


# @Return: Dataframe of all the names of the W2 Drivers
# @Description: Finds the file of Excel drivers and returns a dataframe of all their names
def getW2() -> pd.DataFrame:
    # Checking for list of W2, if not found then user must enter a new one
    if not exists(Data.w2Path):
        loadData()

    # Reading list of names from local W2 file
    drivers = pd.read_excel(Data.w2Path, sheet_name='Sheet1', header=0)
    drivers['W2 Drivers'] = drivers['W2 Drivers'].str.upper()
    return drivers


# @Param filePath: File path to the user's input file
# @Description: Overwrites input file and outputs all the names and their clock in/out if they're a W2 Driver
def deleteRows(filePath: str):
    # Copying data from .xlsx into a dataframe
    df = pd.read_excel(filePath, sheet_name='Summary', header=7, index_col=False, dtype={'Hours': str})
    df.drop('Miles', axis=1, inplace=True)
    df.drop('Date', axis=1, inplace=True)
    df = df[df['Full Driver Name'].str.contains("Total") == True]

    # Formatting time column to fit proper time format
    for i in range(0, len(df.index)):
        indexOfTotal = df.iloc[i]['Full Driver Name'].index('Total')
        df.iloc[i]['Full Driver Name'] = df.iloc[i]['Full Driver Name'][0:indexOfTotal-1].upper()
        try:
            dt = datetime.strptime(df.iloc[i]['Hours'], '%Y-%m-%d %H:%M:%S')
            df.iloc[i]['Hours'] = '{:02d}:{:02d}:{:02d}'.format((24 * dt.day)+dt.hour, dt.minute, dt.second)
        except ValueError:
            pass

    # Filtering data
    df.drop(df.tail(1).index, inplace=True)
    df2 = getW2()
    # TODO: Comparison with excel file isn't always accurate (Hex: 0d0a [end of line] may be cause)
    df = df[df['Full Driver Name'].str.upper().isin(df2['W2 Drivers'])]

    # Writing filtered data to output file
    writer = pd.ExcelWriter(filePath, engine='openpyxl')
    df.to_excel(writer, sheet_name='Output', index=False)
    writer.save()
    print('Finished')


if __name__ == '__main__':
    loadData()
    print("Instructions: ", "Before running the program make sure the relevant Excel file is closed",
          "1) Find the timecard sheet in file-explorer or on the desktop",
          "2) Right click the file and select, \"Copy\"",
          "3) Right click in file-explorer or on the desktop and select, \"Paste\"",
          "4) Hold shift then right click the copied file and select, \"Copy as path\"", sep='\n')
    print("5) You've just copied the file-path to your clipboard, press \"CTRL V\" and paste "
          "the path below. Then press enter")
    path = input("Paste on this line here: ")
    # path = r"C:\Users\emeyers\Desktop\GigiPayroll - Copy.xlsx"  # TODO: Remove test path
    if '"' in path:
        path = path.replace('"', '')
    elif "-v" == path:
        print(getW2())
        input("\nPress enter to finish: ")
        sys.exit("Finished!")
    elif "-c" == path:
        print('W2 list amendment mode, please enter new list!')
        print('All data must be on a sheet titled "Sheet1" and have a header at "A0" followed by the data')
        copyXLSX(input("Enter new W2 list: "))
        input("\nPress enter to finish: ")
        sys.exit("Finished!")
    try:
        print("Formatting File, Please Wait!")
        deleteRows(path)
        print("Task Completed!")
    except Exception as e:
        if "Errno 13" in str(e):
            print("Please close the file and try again!")
        else:
            print("The following error has occurred which could not be resolved:", e)
    input("\nPress enter to finish: ")