import tkinter
import os
from tkinter import filedialog

root = tkinter.Tk()
filez = filedialog.askopenfilenames(parent=root,title='Choose a file')
for file in filez:
    files = file.replace('(1)','')
    os.rename(file,files)
    