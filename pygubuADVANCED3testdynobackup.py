import os
import tkinter as tk
import tkinter.ttk as ttk
import pygubu


PROJECT_PATH = "C:\\Users\kschwartz\PycharmProjects\\tkinterproject"
#os.path.abspath(os.path.dirname(__file__))
PROJECT_UI = os.path.join(PROJECT_PATH, "pygubuADVANCED3_msgbox.ui")


class Pygubuadvanced3MsgboxApp:
    def __init__(self):
        self.about_dialog = None
        self.builder = builder = pygubu.Builder()
        # builder.add_resource_path(PROJECT_PATH)
        builder.add_from_file(PROJECT_UI)
        self.mainwindow = builder.get_object('mainFrame')

        self.mainLabel = None
        builder.import_variables(self, ['mainLabel'])

        builder.connect_callbacks(self)

    def retrieve_input(self):
        # id = self.builder.tkinterscrolledtext['ID_entry'].get("1.0", 'end-1c')
        id = self.builder.widgets.tkinterscrolledtex['ID_entry'].get("1.0", 'end-1c')
        idListprototype = id.splitlines()
        # print(idListprototype)
        if len(id_list) != 0:
            id_list.clear()
            # print(f"id_list =={id_list}")
        for j in idListprototype:
            id_list.append(j)
        print(id_list)
        return id_list

    def run(self):
        self.mainwindow.mainloop()


if __name__ == '__main__':
    # root = tk.Tk()
    app = Pygubuadvanced3MsgboxApp()
    app.run()

