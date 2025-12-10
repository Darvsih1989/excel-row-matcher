import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

def count_matching_rows(file_path):
    try:
        wb = openpyxl.load_workbook(file_path)
    except FileNotFoundError:
        messagebox.showerror("Error", f"The file {file_path} was not found.")
        return

    ws = wb.active
    count = 0
    output_text.delete(1.0, tk.END)  # Clear previous output

    output_text.insert(tk.END, "Matching Rows:\n")
    output_text.insert(tk.END, "-" * 50 + "\n")

    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        col_b, col_c, col_d = row[1], row[2], row[3]

        try:
            col_c = int(col_c) if col_c is not None else 0
        except ValueError:
            col_c = 0

        if (
            isinstance(col_d, str) and "WS-C2960-24TT-L" in col_d and
            col_c >= 20 and
            isinstance(col_b, str) and "atm" not in col_b.lower() and "bje" not in col_b.lower()
        ):
            count += 1
            output_text.insert(tk.END, f"Row {row_idx}: B={col_b}, C={col_c}, D={col_d}\n")

    output_text.insert(tk.END, "-" * 50 + "\n")
    output_text.insert(tk.END, f"Total Matching Rows: {count}\n")

def open_file():
    file_path = filedialog.askopenfilename(
        title="Select an Excel file",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
    )
    if file_path:
        count_matching_rows(file_path)

# Create the main window
root = tk.Tk()
root.title("Excel Row Matcher")

# Create a button to open the file dialog
open_button = tk.Button(root, text="Open Excel File", command=open_file)
open_button.pack(pady=10)

# Create a scrolled text area to display the output
output_text = scrolledtext.ScrolledText(root, width=80, height=20)
output_text.pack(padx=10, pady=10)

# Start the GUI event loop
root.mainloop()

# This line will not be reached until the GUI is closed
input("Press Enter to exit...")