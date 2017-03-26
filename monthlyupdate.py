#!/usr/bin/python

from openpyxl import load_workbook
import os

# Global variables
months = ["01 - January", "02 - February", "03 - March", "04 - April", "05 - May", "06 - June", "07 - July", "08 - August", "09 - September", "10 - October", "11 - November", "12 - December"]
currentPath = os.path.abspath(os.curdir).replace('\\', '/') + '/'
dailyDir = os.path.abspath(os.path.join(os.pardir,os.pardir)).replace('\\', '/') + '/'

def main():
    print currentPath
    print dailyDir
    

if __name__ == "__main__":
    main()
            
