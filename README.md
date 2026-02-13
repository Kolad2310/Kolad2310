```
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

# Variables to use outside GUI
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

        root.destroy()

    root = tk.Tk()
    root.title("CSV Duplicate Drop Tool")

    # Row 0 - File selection
    tk.Button(root, text="Select CSV File", command=browse_file).grid(row=0, column=0, padx=5, pady=5)
    file_label = tk.Label(root, text="No file selected", anchor="w")
    file_label.grid(row=0, column=1, columnspan=2, sticky="w")

    # Row 1 - Separator
    tk.Label(root, text="Separator:").grid(row=1, column=0, sticky="e", padx=5)
    sep_entry = tk.Entry(root, width=10)
    sep_entry.insert(0, ",")
    sep_entry.grid(row=1, column=1, sticky="w")

    # Row 2 - Output Name
    tk.Label(root, text="Output File Name:").grid(row=2, column=0, sticky="e", padx=5)
    output_entry = tk.Entry(root, width=25)
    output_entry.grid(row=2, column=1, sticky="w")

    # Row 3 - Column Selection Label
    tk.Label(root, text="Select Columns:").grid(row=3, column=0, sticky="ne", padx=5, pady=5)

    # Multi-select Listbox
    listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, width=40, height=10)
    listbox.grid(row=3, column=1, padx=5, pady=5)

    # Description next to listbox
    desc_label = tk.Label(
        root,
        text="Subset to drop duplicate columns",
        fg="blue",
        justify="left"
    )
    desc_label.grid(row=3, column=2, sticky="nw", padx=5)

    # Submit button
    tk.Button(root, text="Submit", command=submit).grid(row=4, column=1, pady=10)

    root.mainloop()


# Run GUI
run_gui()

# Variables available outside
print("File:", file)
print("Separator:", separator)
print("Output Name:", output_name)
print("Selected Columns List (l):", l)
