import os
from tkinter import Tk, filedialog, simpledialog, messagebox

root = Tk()  # pointing root to Tk() to use it as Tk() in program.
root.withdraw()  # Hides small tkinter window.
root.attributes('-topmost', True)  # Opened windows will be active. above all windows despite of selection.
cwd = filedialog.askdirectory().replace('/', '\\')  # Returns opened path as str

if cwd != "":
    path = os.path.join(cwd, "converted_file").replace('/', '\\')
    os.chdir(cwd)
    try:
        os.stat(path)
    except:
        os.mkdir(path)

# to break line into multiple segments if it's not already segmented for 837 file which use "~" as segment separator

    for filename in os.listdir(cwd):
        if os.path.isfile(os.path.join(cwd, filename)):
                t = open(filename, 'r')
                t_contents = t.read()
                if t_contents.find('\n') == -1:
                    t_contents = t_contents.replace("~", "~\n")
                else:
                    pass
                f = open(filename,'w')
                f.write(t_contents)
                f.close()

# to replace CLM ID if not present already with immediate REF segment value

    for filename in os.listdir(cwd):
        if os.path.isfile(os.path.join(cwd, filename)):
            with open(os.path.join(path, filename), 'w') as sur:
                 with open(filename,'r') as f:
                    for line in f:
                        if line.startswith('CLM'):
                            line = line.split("*")
                                to_replace = f.readline()
                                if to_replace.startswith("REF"):
                                    t = to_replace.split("*")[2].split("~")[0]
                                    line[1] = t
                                    line = "*".join(line)
                                    sur.write(line)
                                    sur.write(to_replace)
                                else:
                                    sur.write("*".join(line))
                                    sur.write(to_replace)
                        else:
                            sur.write(line)
    messagebox.showinfo('Info', 'Process completed!')
    
else:
    messagebox.showinfo('Info', 'User Terminated!')