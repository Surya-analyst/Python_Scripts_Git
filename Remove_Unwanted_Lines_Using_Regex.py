import customtkinter as tk
from customtkinter import filedialog
from tkinter import messagebox
import re
def browse_file():
    file_path = filedialog.askopenfilename(title = "Select csv/text file", filetypes = (("CSV Files", "*.csv"),("Text Files", "*.txt")))
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)

def regex_execute():
    file_path = file_entry.get()
    pattern = regex_entry.get()
    
    if not file_path:
        messagebox. showwarning ("Warning", "Please select a CSV file first!")
        return
    if not pattern:
        messagebox. showwarning ("Warning", "Please select a Regex pattern!")
        return
    try:
        with open (file_path,"r" ,encoding="utf-16le") as file: # to open the source file in read mode to execute regex replacement #encoding="utf-16le"
            filecontents = file.read()
            filecontents = filecontents.replace("\n\n"," ") # to replace empty new lines and merge with previous lines
            filecontents = re.sub(pattern," ", filecontents) # to replace and merge the lines not starting with 2 with previous lines
        with open (file_path, "w" ,encoding="utf-8") as source_file: # to open the file selected by user in write mode for overwiting
            source_file.write(filecontents)
            
        messagebox.showinfo("Completed", "Python Successfully Compleled Removing Unwanted LineBreaks")
        
    except Exception as e:
        messagebox.showerror("Error", f"An error occured: {e}")
            
tk.set_appearance_mode("dark") # Modes: system (default), light, dark tk. set_default color theme ("green") # Themes: blue (default), dark-blue, green
root = tk.CTk() # create Ck window like you do with the Tk window
root.title("Regex Search and Replace unwanted lines in CSV files") 
root.geometry("870x100")

tk.CTkLabel(root, text = "Select CSV File:").grid(row=0, column=0, padx=5, pady=10, sticky="w")
file_entry = tk.CTkEntry(root, width=600)
file_entry.grid(row=0, column=1, padx=5, pady=10, sticky ="w")
tk. CTkButton (root, text = "Browse", command=browse_file).grid(row=0, column=2, padx=5, pady=10)

tk.CTkLabel(root, text = "Regex Pattern:").grid(row=1, column=0, padx=5, pady=10, sticky ="w")
regex_entry = tk.CTkEntry(root, width=600)
regex_entry.grid(row=1, column=1, padx=5, pady=10, sticky ="w") 
regex_entry.insert(0, r'\n(?![2])')
tk.CTkButton (root, text = "Execute", command=regex_execute).grid(row=1, column=2, padx=5, pady=10)

root.mainloop()