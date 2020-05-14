# -*- coding: utf-8 -*-
"""
Created on Thu May 14 12:10:44 2020

@author: ahmadhmirza
"""
from sys import exit
import os
import datetime
import xlrd
import csv
from optparse import OptionParser
SCRIPT_PATH = os.path.dirname(__file__)

def readExcelFile(xlFile,sheetIndex):
    try:
        workbook=xlrd.open_workbook(xlFile,'r')
        workSheet = workbook.sheet_by_index(sheetIndex)
        sheetData = []
        rowData   = []
        for rowIndex in range(1,workSheet.nrows):
            rowData = []
            for colIndex in range(1,workSheet.ncols):
                rowData.append(workSheet.cell_value(rowIndex,colIndex))
            sheetData.append(rowData)
        return sheetData
    except Exception as e:
        print(str(e))
        return False

def writeCsv(data,outFile):
    try:
        with open(outFile, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerows(data)
        return True
    except Exception as e:
        print(str(e))
        return False
        
def main():
    outFileName_default = "auto_generated.csv"
    outFile = os.path.join(SCRIPT_PATH,outFileName_default)
    usage = "usage: %prog arg1<input xlsx file path> arg2<sheet to be converted(index)> [options]"
    parser = OptionParser(usage=usage)

    parser.add_option("-o", "--output",dest="outFile",
                      help="Full path for output csv file", metavar="FilePath",
                      default=outFile)
    
    (options, args)  = parser.parse_args()
    outFile          = options.outFile # File path
    
    if len(args) != 2:
        print("Incorrect number of arguments provided!")
        exit(1)
    else:    
        inputXL_file = args[0]
        sheet_index  = args[1]
     
    try:
        sheetIndex = int(sheet_index)
    except Exception as e:
        print(str(e))
        exit(1)
    sheetData = readExcelFile(inputXL_file,sheetIndex)
    
    if sheetData != False:
        if writeCsv(sheetData,outFile):
            print("Conversion operation successful")
            exit(0)
        else:
            print("Errors encountered while converting excel to csv")
            exit(1)
    
if __name__ == '__main__':
    main()