import os
import tkinter as tk
import tkinter.ttk as ttk
import webbrowser
from tkinter.scrolledtext import ScrolledText
import time
import pyautogui
import openpyxl
from datetime import date
import datetime
import os
import shutil
import pygubu
from pathlib import Path
from pygubuPolished import id_list
# idList = []
# print("Enter in the ids here:")
# while True:
#     s = input()
#     if s != '':
#         idList.append(s)
#     else:
#         break;
idList = id_list

def send2GSTDB_inner(id_list):
    GSTDBsheetPath = "C:\\Users\kschwartz\Documents\GA-SCDB-Search-helper_realTEST.xlsm"
    GA_wb = openpyxl.load_workbook(GSTDBsheetPath)
    GA_sheet = GA_wb['main_sheet']
    GA_list = []
    def wipePreviousEntries():
        for rowOfCellObjects in GA_sheet['B2':'B51']:
            for cellObj in rowOfCellObjects:
                if cellObj.value != None:
                    cellObj.value = ''
        for rowOfCellObjects in GA_sheet['D2':'D51']:
            for cellObj in rowOfCellObjects:
                if cellObj.value != None:
                    cellObj.value = ''
        for rowOfCellObjects in GA_sheet['F2':'F51']:
            for cellObj in rowOfCellObjects:
                if cellObj.value != None:
                    cellObj.value = ''
        for rowOfCellObjects in GA_sheet['J2':'J12']:
            for cellObj in rowOfCellObjects:
                if cellObj.value != None:
                    cellObj.value = ''
        for rowOfCellObjects in GA_sheet['L2':'L12']:
            for cellObj in rowOfCellObjects:
                if cellObj.value != None:
                    cellObj.value = ''
    def addNewEntries():
        end_point = str(len(idList)+1)
        print(f"end point is {end_point}")
        columnEnd = "B" + end_point
        print(f"Column end is {columnEnd}")
        sheet2list = GA_sheet['B2':columnEnd]
        for rowOfCellObjects in sheet2list:
            # print(sheet2list.index(rowOfCellObjects))
            for cellObj in rowOfCellObjects:
                # print(cellObj.coordinate, cellObj.value)
                    cellObj.value = str(idList[sheet2list.index(rowOfCellObjects)])
    wipePreviousEntries()
    addNewEntries()
    GA_wb.save("C:\\Users\kschwartz\Documents\GA-SCDB-Search-helper_realTEST.xls")
    GA_wb.close()
    os.startfile("C:\\Users\kschwartz\Documents\GA-SCDB-Search-helper_realTEST.xls")
    return
send2GSTDB_inner(id_list)
