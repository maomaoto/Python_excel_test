#!/usr/bin/env python 
# -*- coding: utf-8 -*- 
from easyExcel import *

"""
Analyze cpk data in excel file
"""

def read_sheet_names(xlBook):
    """
    xlBook -> list of sheet names
    """
    sheet_names = []
    for i in range(1,xlBook.Sheets.Count+1):
        #print(xlBook.Sheets(i).Name)
        sheet_names.append(xlBook.Sheets(i).Name)
    return sheet_names

def read_cpk_test_items(cpk_sheet):
    """
    cpk_sheet -> list of all test itmes

    Read column 1 of cpk_sheet to obtain all test items
    """
    test_items = []
    for rw in cpk_sheet.Cells.Rows:
        if cpk_sheet.Cells(rw.Row,1).Value:
            test_items.append(cpk_sheet.Cells(rw.Row,1).Value)
        else: break

    test_items = test_items[2:]
    return test_items

def read_test_items_data_sheet(data_sheet):
    """
    data_sheet -> list of test itmes is data_sheet

    Read row 2 of data_sheet to obtain test items contain in this sheet
    Return sheet name and a list of test items
    If Cells(1,1) value is None, return None
    """
    test_items = []
    if data_sheet.Cells(1,1).Value == None:
        return None
    for col in data_sheet.Cells.Columns:
        if data_sheet.Cells(2, col.Column).Value:
            test_items.append(data_sheet.Cells(2, col.Column).Value)

    return test_items
            
# variables
test_items = []  # Contain all test items
sheet_names = []  # Contain all sheet names
    
exl = easyExcel("D:\git_folder\Python_excel_test\Test_WCDMA.xlsx")
exl.xlApp.Visible = 1

sheet_names = read_sheet_names(exl.xlBook)
print(sheet_names)

test_items = read_all_test_items(exl.xlBook.Sheets("OK_Cpk"))

test_items_each_sheet = {}
for i in sheet_names:
    print(exl.xlBook.Sheets(i).Cells(2,1))
    if 

#exl.close()
