#!/usr/bin/env python 
# -*- coding: utf-8 -*- 
from win32com.client import Dispatch 
import win32com.client 
class easyExcel: 
    """A utility to make it easier to get at Excel.  Remembering 
    to save the data is your problem, as is  error handling. 
    Operates on one workbook at a time.""" 
    def __init__(self, filename=None): 
        self.xlApp = win32com.client.Dispatch('Excel.Application') 
        if filename: 
            self.filename = filename 
            self.xlBook = self.xlApp.Workbooks.Open(filename) 
        else: 
            self.xlBook = self.xlApp.Workbooks.Add() 
            self.filename = ''  
    def save(self, newfilename=None): 
        if newfilename: 
            self.filename = newfilename 
            self.xlBook.SaveAs(newfilename) 
        else: 
            self.xlBook.Save()    
    def close(self): 
        self.xlBook.Close(SaveChanges=0) 
        del self.xlApp 
    def getCell(self, sheet, row, col): 
        "Get value of one cell" 
        sht = self.xlBook.Worksheets(sheet) 
        return sht.Cells(row, col).Value 
    def setCell(self, sheet, row, col, value): 
        "set value of one cell" 
        sht = self.xlBook.Worksheets(sheet) 
        sht.Cells(row, col).Value = value 
    def getRange(self, sheet, row1, col1, row2, col2): 
        "return a 2d array (i.e. tuple of tuples)" 
        sht = self.xlBook.Worksheets(sheet) 
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value 
    def addPicture(self, sheet, pictureName, Left, Top, Width, Height): 
        "Insert a picture in sheet" 
        sht = self.xlBook.Worksheets(sheet) 
        sht.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height) 
    def cpSheet(self, before): 
        "copy sheet" 
        shts = self.xlBook.Worksheets 
        shts(1).Copy(None,shts(1)) 


if __name__ == "__main__": 
    """
    PNFILE = r'c:\screenshot.bmp' 
    xls = easyExcel(r'D:\test.xls') 
    xls.addPicture('Sheet1', PNFILE, 20,20,1000,1000) 
    xls.cpSheet('Sheet1') 
    xls.save() 
    xls.close()
    """
    exl = easyExcel("D:\git_folder\Python_excel_test\Test_WCDMA.xlsx")
    exl.xlBook.Sheets.Count
    for i in range(1,exl.xlBook.Sheets.Count+1):
        print(exl.xlBook.Sheets(i).Name)

    test_item = []
    
    for rw in exl.xlBook.Sheets("OK_Cpk").Cells.Rows:
        if exl.xlBook.Sheets("OK_Cpk").Cells(rw.Row,1).Value:
            test_item.append(exl.xlBook.Sheets("OK_Cpk").Cells(rw.Row,1).Value)
        else: break
