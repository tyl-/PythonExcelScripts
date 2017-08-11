#!/usr/bin/env python
import shutil
import os
from tkinter import *


# Global variables
months = ['01 - January', '02 - February', '03 - March', '04 - April', '05 - May', '06 - June', '07 - July', '08 - August', '09 - September', '10 - October', '11 - November', '12 - December']
# Change year as needed
year = 2017
var1 = 0
var2 = 0
currentPath = os.path.abspath(os.curdir).replace('\\', '/') + '/'
templatePath = currentPath + "Templates/"
templatePathMonthly = templatePath + "Monthly/"
templatePathSpecific = templatePath + "SpecificReports/"
dailyDir = currentPath + "Daily Report " + str(year) + "/"
monthlyDir = currentPath + "Monthly Report " + str(year) + "/"
specificDir = currentPath + "Specific Report " + str(year) + "/"
salesDir = specificDir + "Sales " + str(year) + "/"
tipDir = specificDir + "Tips " + str(year) + "/"
numCustDir = specificDir + "NumberOfCustomer " + str(year) + "/"


def makeDirectories():
    print("Creating directories...")
    if not os.path.exists(dailyDir):
        os.makedirs(dailyDir)
        print(" \"" + dailyDir + "\"")
    if not os.path.exists(monthlyDir):
        os.makedirs(monthlyDir)
        print(" \"" + monthlyDir + "\"")
    for month in months:
        dailyMonthDir = dailyDir + month + "/"
        if not os.path.exists(dailyMonthDir):
            os.makedirs(dailyMonthDir)
            print(" \"" + dailyMonthDir + "\"")
    if not os.path.exists(specificDir):
        os.makedirs(specificDir)
        print(" \"" + specificDir + "\"")
    if not os.path.exists(salesDir):
        os.makedirs(salesDir)
        print(" \"" + salesDir + "\"")
    if not os.path.exists(tipDir):
        os.makedirs(tipDir)
        print(" \"" + tipDir + "\"")
    if not os.path.exists(numCustDir):
        os.makedirs(numCustDir)
        print(" \"" + numCustDir + "\"")
    print("Completed creating directories.")

def makeMonthlySheets():
    print("Generating monthly spreadsheets...")
    for month in months:
        if (overwrite == 1 or os.path.isfile((monthlyDir + month + " " + str(year) + ".xlsx")) == False):
            if(month[:2] == "01" or month[:2] == "03" or month[:2] == "05" or month[:2] == "07" or month[:2] == "08" or month[:2] == "10" or month[:2] == "12"):
                shutil.copy(templatePathMonthly + "01 - January 2015.xlsx", monthlyDir + month + " " + str(year) + ".xlsx")
                print(" \"\\" + month + " " + str(year) + ".xlsx\"")
            elif(month[:2] == "04" or month[:2] == "06" or month[:2] == "09" or month[:2] == "11"):
                shutil.copy(templatePathMonthly + "04 - April 2016.xlsx", monthlyDir + month + " " + str(year) + ".xlsx")
                print(" \"\\" + month + " " + str(year) + ".xlsx\"")
            elif(month[:2] == "02"):
                if(year % 4 == 0):
                    shutil.copy(templatePathMonthly + "02 - February 2017 L.xlsx", monthlyDir + month + " " + str(year) + ".xlsx")
                    print(" \"\\" + month + " " + str(year) + ".xlsx\"")
                else:
                    shutil.copy(templatePathMonthly + "02 - February 2018 NL.xlsx", monthlyDir + month + " " + str(year) + ".xlsx")
                    print(" \"\\" + month + " " + str(year) + ".xlsx\"")
            else:
                print("Error: Unknown month")
        else:
            print(" Skipping \"\\" + month + " " + str(year) + ".xlsx\"")
    print("Completed monthly spreadsheets.")

def makeDailySheets():
    print("Generating daily spreadsheets...")
    for month in months:
        dailyMonthDir = dailyDir + month + "/"
        for num in range(1,10):
            if(overwrite == 1 or os.path.isfile((dailyMonthDir + month[:2] + "-0" + str(num) + "-" + str(year)[2:] + ".xlsx")) == False):
                shutil.copy(templatePath + "01-01-00.xlsx", dailyMonthDir + month[:2] + "-0" + str(num) + "-" + str(year)[2:] + ".xlsx")
                print(" \"\\" + month[:2] + "-0" + str(num) + "-" + str(year)[2:] + ".xlsx\"")
            else:
                print(" Skipping \"\\" + month[:2] + "-0" + str(num) + "-" + str(year)[2:] + ".xlsx\"")
        if(month[:2] == "01" or month[:2] == "03" or month[:2] == "05" or month[:2] == "07" or month[:2] == "08" or month[:2] == "10" or month[:2] == "12"):
            for num in range(10,32):
                if(overwrite == 1 or os.path.isfile((dailyMonthDir + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx")) == False):
                    shutil.copy(templatePath + "01-01-00.xlsx", dailyMonthDir + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx")
                    print(" \"\\" + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx\"")
                else:
                    print(" Skipping \"\\" + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx\"")
        elif(month[:2] == "04" or month[:2] == "06" or month[:2] == "09" or month[:2] == "11"):
            for num in range(10,31):
                if(overwrite == 1 or os.path.isfile((dailyMonthDir + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx")) == False):
                    shutil.copy(templatePath + "01-01-00.xlsx", dailyMonthDir + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx")
                    print(" \"\\" + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx\"")
                else:
                    print(" Skipping \"\\" + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx\"")
        elif(month[:2] == "02"):
            if(year % 4 == 0):
                for num in range(10,30):
                    if(overwrite == 1 or os.path.isfile((dailyMonthDir + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx")) == False):
                        shutil.copy(templatePath + "01-01-00.xlsx", dailyMonthDir + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx")
                        print(" \"\\" + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx\"")
                    else:
                        print(" Skipping \"\\" + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx\"")
            else:
                for num in range(10,29):
                    if(overwrite == 1 or os.path.isfile((dailyMonthDir + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx")) == False):
                        shutil.copy(templatePath + "01-01-00.xlsx", dailyMonthDir + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx")
                        print(" \"\\" + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx\"")
                    else:
                        print(" Skipping \"\\" + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx\"")
        else:
            print("Error: Unknown month")
    print("Completed daily spreadsheets.")

def makeSpecificReports():
    print("Generating specfic report spreadsheets...")
    if(overwrite == 1 or os.path.isfile(salesDir + "Sales.xlsx") == False):
        shutil.copy(templatePathSpecific + "Sales.xlsx", salesDir + "Sales.xlsx")
        print(" \"\\Sales.xlsx\"")
    else:
        print(" Skipping \"\\Sales.xlsx\"")
    if(overwrite == 1 or os.path.isfile(tipDir + "Tips.xlsx") == False):
        shutil.copy(templatePathSpecific + "Tips.xlsx", tipDir + "Tips.xlsx")
        print(" \"\\Tips.xlsx\"")
    else:
        print(" Skipping \"\\Tips.xlsx\"")
    if(overwrite == 1 or os.path.isfile(numCustDir + "NumberOfCustomer.xlsx") == False):
        shutil.copy(templatePathSpecific + "NumberOfCustomer.xlsx", numCustDir + "NumberOfCustomer.xlsx")
        print(" \"\\NumberOfCustomer.xlsx\"")
    else:
        print(" Skipping \"\\NumberOfCustomer.xlsx\"")
    print("Completed specific report spreadsheets.")

def makeUpdateScripts():
    print("Generating click_me_to_update scripts...")
    if(overwrite == 1 or os.path.isfile(monthlyDir + "click_me_to_update.py") == False):
        shutil.copy(currentPath + "monthlyupdate.py", monthlyDir + "click_me_to_update.py")
        print(" \"" + monthlyDir + "click_me_to_update.py\"")
    else:
        print(" Skipping \"" + monthlyDir + "click_me_to_update.py\"")
        
    if(overwrite == 1 or os.path.isfile(salesDir + "click_me_to_update.py") == False):
        shutil.copy(currentPath + "salesupdate.py", salesDir + "click_me_to_update.py")
        print(" \"" + salesDir + "click_me_to_update.py\"")
    else:
        print(" Skipping \"" + salesDir + "click_me_to_update.py\"")
        
    if(overwrite == 1 or os.path.isfile(tipDir + "click_me_to_update.py") == False):
        shutil.copy(currentPath + "tipsupdate.py", tipDir + "click_me_to_update.py")
        print(" \"" + tipDir + "click_me_to_update.py\"")
    else:
        print(" Skipping \"" + tipDir + "click_me_to_update.py\"")
        
    if(overwrite == 1 or os.path.isfile(numCustDir + "click_me_to_update.py") == False):
        shutil.copy(currentPath + "customerupdate.py", numCustDir + "click_me_to_update.py")
        print(" \"" + numCustDir + "click_me_to_update.py\"")
    else:
        print(" Skipping \"" + numCustDir + "click_me_to_update.py\"")
    print("Completed click_me_to_update scripts.")

def updatePaths():
    global year
    global dailyDir
    global monthlyDir
    global specificDir
    global salesDir
    global tipDir
    global numCustDir
    global overwrite
    year = var1.get()
    overwrite = var2.get()
    dailyDir = currentPath + "Daily Report " + str(year) + "/"
    monthlyDir = currentPath + "Monthly Report " + str(year) + "/"
    specificDir = currentPath + "Specific Report " + str(year) + "/"
    salesDir = specificDir + "Sales " + str(year) + "/"
    tipDir = specificDir + "Tips " + str(year) + "/"
    numCustDir = specificDir + "NumberOfCustomer " + str(year) + "/"
    
def makeThings():
    updatePaths()
    makeDirectories()
    makeMonthlySheets()
    makeDailySheets()
    makeSpecificReports()
    makeUpdateScripts()
    print("Completed generating all files.")
    
def main():
    global var1
    global var2
    
    master = Tk()
    master.title("File Creator")
    var1 = IntVar(master)
    var1.set(year)
    var2 = IntVar(master)
    var2.set(0)

    Label(master, text="Year: ", width=10).grid(row=1, column=0, sticky=E, pady=20)
    OptionMenu(master, var1, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030).grid(row=1, sticky=W, column= 1, pady=10)
    Checkbutton(master, text="Overwrite", variable=var2).grid(row=2, columnspan=2, sticky=W+E+N+S, pady=5)
    Button(master, text="Make files", command=makeThings).grid(row=3, columnspan=2, sticky=W+E+N+S, pady=5)
    mainloop()

if __name__ == "__main__":
    main()
