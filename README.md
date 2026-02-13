```
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

# Variables to be used outside GUI
file = None
separator = None
output_name = None
l = []

def run_gui():
    global file, separator, output_name, l

    def browse_file():
        global file
        file = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv")]
        )
        if file:
            file_label.config(text=file)

            # Read only header row to get column names
            try:
                df = pd.read_csv(file, sep=sep_entry.get(), nrows=0)
                listbox.delete(0, tk.END)
                for col in df.columns:
                    listbox.insert(tk.END, col)
            except Exception as e:
                messagebox.showerror("Error", f"Could not read file:\n{e}")

    def submit():
        global separator, output_name, l

        separator = sep_entry.get()
        output_name = output_entry.get()

        selected_indices = listbox.curselection()
        l = [listbox.get(i) for i in selected_indices]

        root.destroy()  # Close GUI

    root = tk.Tk()
    root.title("CSV Processor")

    # File Selection
    tk.Button(root, text="Select CSV File", command=browse_file).pack(pady=5)
    file_label = tk.Label(root, text="No file selected")
    file_label.pack()

    # Separator Entry
    tk.Label(root, text="Separator:").pack()
    sep_entry = tk.Entry(root)
    sep_entry.insert(0, ",")  # default separator
    sep_entry.pack()

    # Output Name Entry
    tk.Label(root, text="Output File Name:").pack()
    output_entry = tk.Entry(root)
    output_entry.pack()

    # Column Selection (Multi-select)
    tk.Label(root, text="Select Columns:").pack()
    listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, width=50, height=10)
    listbox.pack()

    # Submit Button
    tk.Button(root, text="Submit", command=submit).pack(pady=10)

    root.mainloop()


# Run GUI
run_gui()

# Variables available here
print("File:", file)
print("Separator:", separator)
print("Output Name:", output_name)
print("Selected Columns List (l):", l)
