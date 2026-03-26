```
def browse_folder():
    folder = filedialog.askdirectory()
    if not folder:
        return

    files = get_excel_files(folder)

    print("DEBUG Files:", files)

    if not files:
        messagebox.showerror("Error", "No Excel files found")
        return

    sheet_window = tk.Toplevel(root)
    sheet_window.title("Select Sheets")

    canvas = tk.Canvas(sheet_window)
    scrollbar = tk.Scrollbar(sheet_window, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas)

    scroll_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    selection_vars = {}

    for file in files:
        try:
            sheets = pd.ExcelFile(file).sheet_names
        except:
            sheets = []

        selection_vars[file] = {}

        # 🔹 Row frame (file + sheets in one line)
        row_frame = tk.Frame(scroll_frame)
        row_frame.pack(fill="x", pady=5)

        # File name (left)
        tk.Label(
            row_frame,
            text=os.path.basename(file),
            width=30,
            anchor="w",
            font=("Arial", 9, "bold")
        ).pack(side="left", padx=5)

        # Sheets (horizontal)
        for sheet in sheets:
            var = tk.BooleanVar()

            # ✅ Preselect only "Income Sub."
            if sheet.strip().lower() == "income sub.":
                var.set(True)
            else:
                var.set(False)

            chk = tk.Checkbutton(row_frame, text=sheet, variable=var)
            chk.pack(side="left", padx=5)

            selection_vars[file][sheet] = var

    def submit():
        selection_dict = {}

        for file, sheets in selection_vars.items():
            selected = [s for s, v in sheets.items() if v.get()]
            if selected:
                selection_dict[file] = selected

        if not selection_dict:
            messagebox.showerror("Error", "Select at least one sheet")
            return

        sheet_window.destroy()
        process_files(folder, selection_dict)

    tk.Button(sheet_window, text="Submit", command=submit).pack(pady=10)
