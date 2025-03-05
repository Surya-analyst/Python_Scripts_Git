import os
from tkinter import Tk, filedialog, simpledialog, messagebox
from os import listdir
from datetime import date
from pprint import pprint
from os.path import isfile, join

root = Tk()  # pointing root to Tk() to use it as Tk() in program.
root.withdraw()  # Hides small tkinter window.
root.attributes('-topmost', True)  # Opened windows will be active. above all windows despite of selection.
cwd = filedialog.askdirectory().replace('/', '\\')  # Returns opened path as str

os.chdir(cwd)

folder_path = '.'
onlyfiles = [join(folder_path, f) for f in listdir(folder_path) if isfile(join(folder_path, f)) and not f.split('.')[-1] in ['py','ipynb']]

for entry in onlyfiles:
    lines = open(entry, 'r').readlines() 
    del lines[-1]
    open(entry, 'w').writelines(lines)