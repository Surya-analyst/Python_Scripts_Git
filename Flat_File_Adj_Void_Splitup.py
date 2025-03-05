import os
from tkinter import Tk, filedialog, simpledialog, messagebox

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
            with open(os.path.join(path, filename), 'w') as sur:
                with open(os.path.join(path, filename.replace(".","_Adj.")), 'w') as sur_a:
                    with open(os.path.join(path, filename.replace(".","_Void.")), 'w') as sur_void:
                        with open(os.path.join(path, filename.replace(".","_Others.")), 'w') as sur_others:
                            with open(filename, 'r') as f:
                                header = f.readline()
                                sur.write(header)
                                sur_a.write(header)
                                sur_void.write(header)
                                sur_others.write(header)
                                for line in f:
                                    if line.split("|")[62] == "1":
                                        sur.write(line)
                                    elif line.split("|")[62] == "7":
                                        sur_a.write(line)
                                    elif line.split("|")[62]== "8":
                                        sur_void.write(line)
                                    else:
                                        sur_others.write(line)

    for filename in os.listdir(path):
        if os.path.isfile(os.path.join(path, filename)):
            counter = 0
            with open(os.path.join(path, filename), 'r') as sur:
                for line in sur:
                    counter += 1
            if counter == 1:
                os.remove(os.path.join(path, filename))

    messagebox.showinfo('Info', 'Process completed!')
else:
    messagebox.showinfo('Info', 'User Terminated!')

