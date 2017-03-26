#!/usr/bin/env python
import shutil
import os

# Global variables
months = ['01 - January', '02 - February', '03 - March', '04 - April', '05 - May', '06 - June', '07 - July', '08 - August', '09 - September', '10 - October', '11 - November', '12 - December']
# Change year as needed
year = 2017
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
    print("Finished creating directories.")

def makeMonthlySheets():
    print("Generating monthly spreadsheets from monthly template...")
    for month in months:
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
    print("Finished creating monthly spreadsheets.")

def makeDailySheets():
    print("Generating daily spreadsheets from daily template...")
    for month in months:
        dailyMonthDir = dailyDir + month + "/"
        for num in range(1,10):
            shutil.copy(templatePath + "01-01-00.xlsx", dailyMonthDir + month[:2] + "-0" + str(num) + "-" + str(year)[2:] + ".xlsx")
            print(" \"\\" + month[:2] + "-0" + str(num) + "-" + str(year)[2:] + ".xlsx\"")
        if(month[:2] == "01" or month[:2] == "03" or month[:2] == "05" or month[:2] == "07" or month[:2] == "08" or month[:2] == "10" or month[:2] == "12"):
            for num in range(10,32):
                shutil.copy(templatePath + "01-01-00.xlsx", dailyMonthDir + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx")
                print(" \"\\" + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx\"")
        elif(month[:2] == "04" or month[:2] == "06" or month[:2] == "09" or month[:2] == "11"):
            for num in range(10,31):
                shutil.copy(templatePath + "01-01-00.xlsx", dailyMonthDir + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx")
                print(" \"\\" + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx\"")
        elif(month[:2] == "02"):
            if(year % 4 == 0):
                for num in range(10,30):
                    shutil.copy(templatePath + "01-01-00.xlsx", dailyMonthDir + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx")
                    print(" \"\\" + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx\"")
            else:
                for num in range(10,29):
                    shutil.copy(templatePath + "01-01-00.xlsx", dailyMonthDir + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx")
                    print(" \"\\" + month[:2] + "-" + str(num) + "-" + str(year)[2:] + ".xlsx\"")
        else:
            print("Error: Unknown month")

def makeSpecificReports():
    print("Generating specfic report spreadsheets from the template...")
    shutil.copy(templatePathSpecific + "Sales.xlsx", salesDir + "Sales.xlsx")
    print(" \"\\Sales.xlsx\"")
    shutil.copy(templatePathSpecific + "Tips.xlsx", tipDir + "Tips.xlsx")
    print(" \"\\Tips.xlsx\"")
    shutil.copy(templatePathSpecific + "NumberOfCustomer.xlsx", numCustDir + "NumberOfCustomer.xlsx")
    print(" \"\\NumberOfCustomer.xlsx\"")

def makeUpdateScripts():
    print("Generating specfic click_me_to_update scripts...")
    shutil.copy(currentPath + "monthlyupdate.py", monthlyDir + "click_me_to_update.py")
    print(" \"" + monthlyDir + "click_me_to_update.py\"")
    shutil.copy(currentPath + "salesupdate.py", salesDir + "click_me_to_update.py")
    print(" \"" + salesDir + "click_me_to_update.py\"")
    shutil.copy(currentPath + "tipsupdate.py", tipDir + "click_me_to_update.py")
    print(" \"" + tipDir + "click_me_to_update.py\"")
    shutil.copy(currentPath + "customerupdate.py", numCustDir + "click_me_to_update.py")
    print(" \"" + numCustDir + "click_me_to_update.py\"")

def main():
    makeDirectories()
    makeMonthlySheets()
    makeDailySheets()
    makeSpecificReports()
    makeUpdateScripts()
    print("Finished generating files")

if __name__ == "__main__":
    main()
