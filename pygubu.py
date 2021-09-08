import os
import tkinter as tk
import tkinter.ttk as ttk
import pygubu

PROJECT_PATH = os.path.abspath(os.path.dirname(__file__))
PROJECT_UI = os.path.join(PROJECT_PATH, "pygubutest.ui")


class PygubutestApp:
    def __init__(self, master=None):
        # build ui
        self.main = ttk.Frame(master)
        self.label1 = ttk.Label(self.main)
        self.mainLabel = tk.StringVar(value='ID Exporter')
        self.label1.configure(compound='top', font='{Bahnschrift} 12 {}', style='Toolbutton', takefocus=False)
        self.label1.configure(text='ID Exporter', textvariable=self.mainLabel)
        self.label1.pack(side='top')
        self.inner = ttk.Frame(self.main)
        self.ID_entry = tk.Text(self.inner)
        self.ID_entry.configure(autoseparators='false', cursor='arrow', exportselection='true', font='TkDefaultFont')
        self.ID_entry.configure(height='10', insertborderwidth='2', padx='20', pady='10')
        self.ID_entry.configure(relief='raised', width='50')
        _text_ = '''insert IDS'''
        self.ID_entry.insert('0.0', _text_)
        self.ID_entry.pack(side='top')
        self.enter = ttk.Button(self.inner, text = "Enter", command = self.retrieve_input)
        self.enter.configure(text='Enter',)
        self.enter.pack(pady='6', side='top')
        self.inner.configure(height='200', padding='10', relief='flat', width='200')
        self.inner.pack(side='top')
        self.main.configure(height='400', relief='flat', takefocus=True, width='400')
        self.main.pack(side='top')

        # Main widget
        self.mainwindow = self.main

    def retrieve_input(self):
        id = self.ID_entry.get("1.0", 'end-1c')
        print(id)
    def run(self):
        self.mainwindow.mainloop()

if __name__ == '__main__':
    root = tk.Tk()
    app = PygubutestApp(root)
    app.run()
