# command_properties.py
import tkinter as tk
import tk
from tkinter import messagebox
import pygubu

# define the function callbacks
def on_button1_click():
    messagebox.showinfo('Message', 'You clicked Button 1')

def on_button2_click():
    messagebox.showinfo('Message', 'You clicked Button 2')

def on_button3_click():
    messagebox.showinfo('Message', 'You clicked Button 3')


class MyApplication(pygubu.TkApplication):

    def _create_ui(self):
        #1: Create a builder
        self.builder = builder = pygubu.Builder()

        #2: Load an ui file
        builder.add_from_file('command_properties.ui')

        #3: Create the widget using self.master as parent
        self.mainwindow = builder.get_object('mainwindow', self.master)

        # Configure callbacks
        callbacks = {
            'on_button1_clicked': on_button1_click,
            'on_button2_clicked': on_button2_click,
            'on_button3_clicked': on_button3_click
        }

        builder.connect_callbacks(callbacks)


if __name__ == '__main__':
    root = tk.Tk()
    app = MyApplication(root)
    app.run()
