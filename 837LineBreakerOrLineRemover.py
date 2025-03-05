import os
from tkinter import Tk, filedialog, simpledialog, messagebox

root = Tk()  # pointing root to Tk() to use it as Tk() in program.
root.withdraw()  # Hides small tkinter window.
root.attributes('-topmost', True)  # Opened windows will be active. above all windows despite of selection.
cwd = filedialog.askdirectory().replace('/', '\\')  # Returns opened path as str

if cwd != "":
    path = os.path.join(cwd, "converted_file").replace('/', '\\')
    print(cwd)
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    user_input = simpledialog.askstring(title='Provide input "Y/N" only!',
                                        prompt="Type 'Y' to add line break in 837 file and 'N' to remove line break "
                                               "in 837 file")

    os.chdir(cwd)

    if user_input == "y" or user_input == "Y":
        try:
            os.stat(path)
        except:
            os.mkdir(path)

        for filename in os.listdir(cwd):
            if os.path.isfile(os.path.join(cwd, filename)):
                with open(os.path.join(path, filename), 'w') as f:
                    t = open(filename, 'r')
                    t_contents = t.read()
                    if t_contents.find('\n') == -1:
                        t_contents = t_contents.replace("~", "~\n")
                    else:
                        pass
                    f.write(t_contents)
            else:
                messagebox.showinfo('Info', 'Process completed!')

    elif user_input == "n" or user_input == 'N':
        try:
            os.stat(path)
        except:
            os.mkdir(path)

        for filename in os.listdir(cwd):
            if os.path.isfile(os.path.join(cwd, filename)):
                with open(os.path.join(path, filename), 'w') as f:
                    t = open(filename, 'r')
                    t_contents = t.read()
                    if t_contents.find('\n') != -1:
                        t_contents = t_contents.replace("~\n", "~")
                    else:
                        pass
                    f.write(t_contents)
            else:
                messagebox.showinfo('Info', 'Process completed!')

    else:
        messagebox.showinfo('Info', 'User Terminated!')

else:
    messagebox.showinfo('Info', 'User Terminated!')
