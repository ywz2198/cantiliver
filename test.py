import tkinter
from tkinter import filedialog

root = tkinter.Tk()
filez = filedialog.askopenfilenames(parent=root,title='Choose a file')
print( root.tk.splitlist(filez))