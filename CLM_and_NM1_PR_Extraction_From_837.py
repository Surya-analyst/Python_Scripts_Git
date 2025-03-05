import os
from tkinter import Tk, filedialog, simpledialog, messagebox
import re

root = Tk()  # pointing root to Tk() to use it as Tk() in program.
root.withdraw()  # Hides small tkinter window.
root.attributes('-topmost', True)  # Opened windows will be active. above all windows despite of selection.
cwd = filedialog.askdirectory().replace('/', '\\')  # Returns opened path as str

if cwd != "":
    path = os.path.join(cwd, "extracted_file").replace('/', '\\')
    os.chdir(cwd)
    try:
        os.stat(path)
    except:
        os.mkdir(path)
    for filename in os.listdir(cwd):
        if os.path.isfile(os.path.join(cwd, filename)):
            with open(filename, 'r') as s:
                with open(os.path.join(path, filename), 'w') as f:
                    for line in s:
                        if line.startswith('CLM*'):
                            f.write(line.rstrip())
                            f.write("\n")
                        elif line.startswith('NM1*PR*2*XXX XXXX XXXX*****PI*XXX-XXXXXXXXXX'):
                            f.write(line.rstrip())
                            f.write("\n")
                         
    messagebox.showinfo('Info', 'Process completed!')

else:
    messagebox.showinfo('Info', 'User Terminated!')





