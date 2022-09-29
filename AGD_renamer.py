from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from openpyxl import load_workbook
from functions import main
import ntpath
import os

'''
    AGD renamer GUI.
'''

#Configuration of the GUI main frame

root = Tk()
root.title("AGD parts files renamer")

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
root.resizable(FALSE,FALSE)
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

#Paths variables

path1 = 'Please choose a path to the workbook file.'
path2 = 'Please choose a path to the parts files.'

#Choosing the excel parts list file

def search_workbook():
    global path1
    path1 = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All Files", "*")))
    if len(path1) > 60:
        path = '.../' + ntpath.basename(path1)
    else:
        path = path1
    label1.config(text=path)

#Choosing the directory with parts files

def search_directory():
    global path2
    path2 = filedialog.askdirectory()+'/'
    if len(path2) > 60:
        list = path2.split('/')
        path = '.../' + list[-2]
    else:
        path = path2
    label2.config(text=path)

#Starting a main program

def run_button():
    try:
        #function main ->  returns quantity of files names changed(img or png, to, dxf)
        func = main(load_workbook(path1),path2)
        func

        if len(os.listdir(path2)) == 0:
            messagebox.showwarning(message="You don't have any files in selected directory!")
        elif func == (0,0,0):
            messagebox.showwarning(message="Unfortunately, this directory don't have any files fitting with the parts list!")
        else:
            messagebox.showinfo(message=f'Done! Changed names: .img/.png: {func[0]}, .stp: {func[1]}, .dxf: {func[2]}')
            root.destroy()
    except:
        messagebox.showerror(message='Input Error. Check your inputs and try again!')

def close_button():
    root.destroy()


ttk.Label(mainframe, text = "Welcome to AGD Renamer App!", font=12).grid(column=1, row=1, sticky=E+W, padx=10, pady=10)

ttk.Label(mainframe, text = "Please, choose the Parts List file path in .xlsx format:", font=12).grid(column=1, row=2, sticky=W, padx=10, pady=10)
global label1
label1 = ttk.Label(mainframe, text = path1, background='white')
label1.grid(column=1, row=3, sticky=E, padx=10, pady=10)
ttk.Button(mainframe, text="Search", command=search_workbook).grid(column=2, row=3, sticky=E, padx=5, pady=10)


ttk.Label(mainframe, text = "Please, choose the parts files directory:", font=12).grid(column=1, row=4, sticky=W, padx=10, pady=10)
global label2
label2 = ttk.Label(mainframe, text = path2, background='white')
label2.grid(column=1, row=5, sticky=E, padx=10, pady=10)
ttk.Button(mainframe, text="Search", command=search_directory).grid(column=2, row=5, sticky=E, padx=5, pady=10)

ttk.Button(mainframe, text="Close", command=close_button).grid(column=1, row=6, sticky=E, padx=5, pady=10)
ttk.Button(mainframe, text="Run", command=run_button).grid(column=2, row=6, sticky=E, padx=5, pady=5)

root.mainloop()