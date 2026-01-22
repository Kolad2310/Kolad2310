```
import tkinter as tk
from tkinter import filedialog, messagebox

# -------------------- GUI SETUP --------------------
root = tk.Tk()
root.title("Template Generation – Input Selector")
root.geometry("700x520")
root.configure(bg="#f4f6f8")

font_label = ("Segoe UI", 10)
font_entry = ("Segoe UI", 9)
font_button = ("Segoe UI", 10, "bold")

# -------------------- VARIABLES --------------------
omnia_file = tk.StringVar()
legal_entity_file = tk.StringVar()
cost_summary_file = tk.StringVar()

output_template_folder = tk.StringVar()
value_output_folder = tk.StringVar()
validation_folder = tk.StringVar()

# -------------------- FUNCTIONS --------------------
def select_file(var):
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx *.xlsm *.xls")]
    )
    if file_path:
        var.set(file_path)

def select_folder(var):
    folder_path = filedialog.askdirectory()
    if folder_path:
        var.set(folder_path)

def submit():
    # Save values into Python variables
    selections = {
        "Omnia Extract": omnia_file.get(),
        "Legal Entity Extract": legal_entity_file.get(),
        "Cost Summary Walk": cost_summary_file.get(),
        "Output Template Folder": output_template_folder.get(),
        "Value Output Folder": value_output_folder.get(),
        "Validation Folder": validation_folder.get()
    }

    # Basic validation
    if not all(selections.values()):
        messagebox.showerror("Missing Selection", "Please select all files and folders.")
        return

    # At this point, variables are ready for use
    messagebox.showinfo("Success", "All inputs captured successfully!")

    # 🔹 Example: print or pass to next function
    for k, v in selections.items():
        print(f"{k}: {v}")

# -------------------- UI LAYOUT --------------------
tk.Label(root, text="Input & Output Configuration",
         font=("Segoe UI", 14, "bold"),
         bg="#f4f6f8").pack(pady=15)

frame = tk.Frame(root, bg="#ffffff", padx=20, pady=20)
frame.pack(fill="both", expand=True, padx=30)

def add_selector(row, label, var, is_file=True):
    tk.Label(frame, text=label, font=font_label, bg="#ffffff").grid(row=row, column=0, sticky="w", pady=8)
    tk.Entry(frame, textvariable=var, font=font_entry, width=50).grid(row=row, column=1, padx=10)
    tk.Button(
        frame,
        text="Browse",
        font=font_button,
        bg="#0078D7",
        fg="white",
        command=lambda: select_file(var) if is_file else select_folder(var)
    ).grid(row=row, column=2)

# ---- File Selectors ----
add_selector(0, "Omnia Extract File", omnia_file, True)
add_selector(1, "Legal Entity Extract File", legal_entity_file, True)
add_selector(2, "Cost Summary Walk File", cost_summary_file, True)

# ---- Folder Selectors ----
add_selector(3, "Output Template Folder", output_template_folder, False)
add_selector(4, "Value Version Output Folder", value_output_folder, False)
add_selector(5, "Validation Output Folder", validation_folder, False)

# ---- Submit Button ----
tk.Button(
    root,
    text="Submit",
    font=("Segoe UI", 11, "bold"),
    bg="#28a745",
    fg="white",
    width=15,
    command=submit
).pack(pady=20)

root.mainloop()
