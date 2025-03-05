import os
from tkinter import Tk, filedialog, simpledialog, messagebox

root = Tk()  # pointing root to Tk() to use it as Tk() in program.
root.withdraw()  # Hides small tkinter window.
root.attributes('-topmost', True)  # Opened windows will be active. above all windows despite of selection.
cwd = filedialog.askdirectory().replace('/', '\\')  # Returns opened path as str

clm_pos = 0
sub_element_dlm = ">"

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
                        if line.startswith('ISA'):
                            sub_element_dlm = line.split("*")[16][-3]
                            sur.write(line)
                        elif line.startswith('CLM'):
                            clm_pos = line.split("*")[5].split(sub_element_dlm)[0]
                            sur.write(line)
                        elif line.startswith('SV1'):
                            sv1_pos = line.split("*")
                            if sv1_pos[5] == clm_pos:
                                sv1_pos[5] = ''
                                sv1_pos = '*'.join(sv1_pos)
                                sur.write(sv1_pos)
                            else:
                                sur.write(line)
                        else:
                            sur.write(line)
    messagebox.showinfo('Info', 'Process completed!')
    
else:
    messagebox.showinfo('Info', 'User Terminated!')
    
