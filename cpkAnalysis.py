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
        else: break
    return test_items

def read_SN_data(item_name, xlBook, items_each_sheet, n=10):
    """
    (item_name, xlBook, items_each_sheet, n=10) -> dict{SN: value}
    Input item_name and data(xlBook, items_each_sheet)
    Output a dict with n pairs of SN vs. value of item_name(first finded)
    If item_name not in items_each_sheet return None
    """
    dict_SN_data = {}
    for sheet_name in items_each_sheet.keys():
        if item_name in items_each_sheet[sheet_name]:
            for i in range(n):
                dict_SN_data[xlBook.Sheets(sheet_name).Cells(3+i,1).Value] =\
                    xlBook.Sheets(sheet_name).Cells(3+i,(items_each_sheet[sheet_name].index(item_name))+1).Value
            return dict_SN_data

    return None

def get_number_sample_data_sheet(data_sheet):
    """
    data_sheet -> 
    """
    
# variables
test_items = []  # Contain all test items
sheet_names = []  # Contain all sheet names
    
exl = easyExcel("D:\git_folder\Python_excel_test\Test_WCDMA.xlsx")
exl.xlApp.Visible = 1

sheet_names = read_sheet_names(exl.xlBook)
print(sheet_names)

test_items = read_cpk_test_items(exl.xlBook.Sheets("OK_Cpk"))

items_each_sheet = {}
for name in sheet_names:
    #print(exl.xlBook.Sheets(name).Cells(2,1))
    if "Cpk" not in name:
        items = read_test_items_data_sheet(exl.xlBook.Sheets(name))
        if items != None: items_each_sheet[name] = items

for col in exl.xlBook.Sheets('SPE1_TestReport').Columns("A:C"):
    print(exl.xlBook.Sheets('SPE1_TestReport').Cells(2,col.Column).Value)
for col in exl.xlBook.Sheets('SPE1_TestReport').UsedRange.Columns:
    print(exl.xlBook.Sheets('SPE1_TestReport').Cells(2,col.Column).Value)
#exl.close()
