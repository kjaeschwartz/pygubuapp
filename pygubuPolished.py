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

#TODO MAKE A GSTDB BUTTON THAT CAN POST ALL IDS TO THE GSTDB SHEET
#(ugh ive had some issues with posting to excel sheets before)
PROJECT_PATH = os.path.abspath(os.path.dirname(__file__))
PROJECT_UI = os.path.join(PROJECT_PATH, "searchHelperUI_1.ui")
# I CANT PUT ANYTHING OUTSIDE OF THE APP CLASS BECAUSE STUFF AFTER THAT CLASS WILL ONLY RUN ONCE THE PROGRAM CLOSES.
id_list = []


def get_new_tab():
    pass

class Searchhelperui1App:
    def __init__(self, master=None):
        # build ui
        self.main = ttk.Frame(master)
        self.header = ttk.Label(self.main)
        self.mainLabel = tk.IntVar(value='ID Exporter')
        self.header.configure(anchor='ne', borderwidth='2', compound='top', font='{Arial CYR} 14 {bold}')
        self.header.configure(foreground='#276a70', relief='flat', state='disabled', style='Toolbutton')
        self.header.configure(takefocus=True, text='ID Exporter', textvariable=self.mainLabel)
        self.header.grid(column='0', row='0')
        self.leftFrame = ttk.Frame(self.main)
        self.ID_entry = ScrolledText(self.leftFrame)
        self.ID_entry.configure(autoseparators='true', background='#f7fcfd', blockcursor='true',
                                borderwidth='1')
        self.ID_entry.configure(height='12', highlightbackground='#69c4cb', highlightthickness='1')
        self.ID_entry.configure(setgrid='false', tabstyle='wordprocessor', takefocus=False, undo='true')
        self.ID_entry.configure(width='20')
        self.ID_entry.pack(side='top')
        self.enterButton = tk.Button(self.leftFrame)
        self.enterButton.configure(background='#8ceaea', justify='left', relief='raised', text='Enter',
                                   command=self.retrieve_input)
        self.enterButton.pack(ipadx='13', padx='10', pady='10', side='top')
        self.leftFrame.configure(height='200', padding='10', relief='flat', width='200')
        self.leftFrame.grid(column='0', row='1')
        self.RightFrame = tk.Frame(self.main)
        self.inputATMbutton = tk.Button(self.RightFrame)
        self.inputATMbutton.configure(background='#8ceaea', foreground='#030a07', justify='left', padx='22')
        self.inputATMbutton.configure(relief='raised', text='Input to ATM', command=self.open_ids_ATM)
        self.inputATMbutton.grid(column='0', pady='7', row='0')
        self.sendGSTDBsheetButton = tk.Button(self.RightFrame)
        self.sendGSTDBsheetButton.configure(background='#8ceaea', compound='top', text='Send to GSTDB sheet',command=self.send2GSTDB_func)
        self.sendGSTDBsheetButton.grid(column='0', pady='7', row='1')
        self.makeGSTDBfoldersButton = tk.Button(self.RightFrame)
        self.makeGSTDBfoldersButton.configure(background='#8ceaea', justify='left', text='Make GSTDB folders',command=self.make_GSTDB_folders)
        self.makeGSTDBfoldersButton.grid(column='0', pady='7', row='2')
        self.CTsearchSetupButton = tk.Button(self.RightFrame)
        self.CTsearchSetupButton.configure(background='#8ceaea', cursor='arrow', justify='left', padx='13')
        self.CTsearchSetupButton.configure(relief='raised', text='CT search setup')
        self.CTsearchSetupButton.grid(column='0', pady='7', row='3')
        self.RightFrame.configure(height='170', highlightbackground='#bcb5e6',
                                  highlightcolor='#a8bef2')
        self.RightFrame.configure(padx='20', pady='20', takefocus=False, width='200')
        self.RightFrame.grid(column='2', row='1', sticky='n')
        self.main.columnconfigure('2', pad='0')
        self.main.configure(borderwidth='2', height='400', relief='raised', takefocus=True)
        self.main.configure(width='600')
        self.main.pack(side='top')

        # Main widget
        self.mainwindow = self.main

    def retrieve_input(self):
        id = self.ID_entry.get("1.0", 'end-1c')
        idListprototype = id.splitlines()
        # print(idListprototype)
        if len(id_list) != 0:
            id_list.clear()
            # print(f"id_list =={id_list}")
        for j in idListprototype:
            id_list.append(j)
        print(id_list)
        return id_list

    class open_ids():
        def make_chrome_window2(self):
            fw = pyautogui.getWindowsWithTitle('ATM - Google Chrome')
            pyautogui.scroll(200)
            if len(fw) == 0:
                print("l is 0")
                webbrowser.open('https://atm.accuratebackground.com/atm/login.jsp')
                fw = pyautogui.getWindowsWithTitle('Vendor Login | Accurate Background - Google Chrome')
            fw = fw[0]
            fw.width = 974
            fw.topleft = (953, 0)

        def get_new_tab(self):
            webbrowser.open('https://atm.accuratebackground.com/atm/findSearch.html')
            time.sleep(1.1)

        def search_id_fetch(self, ids_list):
            self.make_chrome_window2()
            # for h in range(0,len(ids_list)):
            for h in id_list:
                time.sleep(.05)
                self.get_new_tab()
                findsearch = ["enter_id_box.png", "search_press_box.png"]
                time.sleep(.05)
                enter_id_box = (1340, 405)
                press_search_button = (1625, 405)
                pyautogui.click(enter_id_box)
                time.sleep(.05)
                pyautogui.typewrite(h)
                time.sleep(.15)
                pyautogui.click(press_search_button)

    # auto_start()
    def open_ids_ATM(self):
        ids_list = self.retrieve_input()
        self.open_ids().search_id_fetch(ids_list)
        pyautogui.hotkey('ctrl', '2', interval=.07)
        # Testimport = "test succeeded"
    #
    def make_GSTDB_folders(self):
        exec(open('C:\\Users\kschwartz\PycharmProjects\pythonProject\make_GSCDB_folders.py').read())

    def send2GSTDB_func(self):
        if len(id_list) != 0:
            GSTDBsheetPath = "C:\\Users\kschwartz\Documents\GA-SCDB-Search-helper_realTEST.xlsm"
            GA_wb = openpyxl.load_workbook(GSTDBsheetPath)
            GA_sheet = GA_wb['main_sheet']
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
                end_point = str(len(id_list) + 1)
                print(f"end point is {end_point}")
                columnEnd = "B" + end_point
                print(f"Column end is {columnEnd}")
                sheet2list = GA_sheet['B2':columnEnd]
                for rowOfCellObjects in sheet2list:
                    for cellObj in rowOfCellObjects:
                        # print(cellObj.coordinate, cellObj.value)
                        cellObj.value = str(id_list[sheet2list.index(rowOfCellObjects)])

            wipePreviousEntries()
            addNewEntries()
            GA_wb.save("C:\\Users\kschwartz\Documents\GA-SCDB-Search-helper_realTEST.xls")
            GA_wb.close()
            os.startfile("C:\\Users\kschwartz\Documents\GA-SCDB-Search-helper_realTEST.xls")
            return
    # def send2GSTDB_sheet(self):
    #     exec(open("C:\\Users\kschwartz\PycharmProjects\\tkinterproject\send2GSTDB.py").read())
    #     # ids_list = self.retrieve_input()
    #     # if ids_list != []:
    #     #     send2GSTDB.send2GSTDB_inner(id_list)
    #     #     print("sent to GSTDB sheet")
    #     # else:
    #     #     print("input ids first")
    def run(self):
        self.mainwindow.mainloop()


if __name__ == '__main__':
    root = tk.Tk()
    app = Searchhelperui1App(root)
    app.run()
