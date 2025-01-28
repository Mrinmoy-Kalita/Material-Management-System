import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd

# File paths for Excel sheets
INDENT_RECORDS_FILE = "material_indent_records.xlsx"
DEMAND_RECORDS_FILE = "demand_records.xlsx"
SEARCH_RESULTS_FILE = "search_results.xlsx"

# Standard Colors
BG_COLOR = "#f7f7f7"
BUTTON_COLOR = "#4CAF50"
BUTTON_TEXT_COLOR = "white"
HIGHLIGHT_COLOR = "#ff9800"
TEXT_COLOR = "#333333"
FRAME_COLOR = "#ffffff"
HEADING_COLOR = "#3E50B4"

# Function to Export Data to Excel
def export_to_excel(data, file_path, columns):
    try:
        df = pd.DataFrame(data, columns=columns)
        for col in columns:
            if 'date' in col.lower():
                df[col] = pd.to_datetime(df[col]).dt.strftime('%Y-%m-%d')
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Success", f"Data exported to {file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to export data: {e}")

# Function to Add Material Indent Record
def add_material_indent():
    def save_record():
        try:
            new_record = pd.DataFrame([{
                "Indent Date": indent_date_var.get(),
                "Employee Name": employee_name_var.get(),
                "Items": items_var.get(),
                "No of Items": no_of_items_var.get(),
                "Description": description_var.get(),
                "Estimated Value": estimated_value_var.get()
            }])

            try:
                df = pd.read_excel(INDENT_RECORDS_FILE)
            except FileNotFoundError:
                df = pd.DataFrame(columns=new_record.columns)

            df = pd.concat([df, new_record], ignore_index=True)
            df.to_excel(INDENT_RECORDS_FILE, index=False)
            messagebox.showinfo("Success", "Record added successfully to Excel!")
            window.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save record: {e}")

    window = tk.Toplevel()
    window.title("Add Material Indent Record")
    window.geometry("500x500")
    window.configure(bg=BG_COLOR)

    frame = tk.Frame(window, bg=FRAME_COLOR, padx=20, pady=20, relief="groove", bd=2)
    frame.pack(pady=20)

    tk.Label(frame, text="Add Material Indent Record", font=("Arial", 16, "bold"), fg=HEADING_COLOR, bg=FRAME_COLOR).grid(row=0, column=0, columnspan=2, pady=10)

    indent_date_var = tk.StringVar()
    employee_name_var = tk.StringVar()
    items_var = tk.StringVar()
    no_of_items_var = tk.IntVar()
    description_var = tk.StringVar()
    estimated_value_var = tk.DoubleVar()

    fields = [
        ("Indent Date (YYYY-MM-DD):", indent_date_var),
        ("Employee Name:", employee_name_var),
        ("Items:", items_var),
        ("No of Items:", no_of_items_var),
        ("Description:", description_var),
        ("Estimated Value:", estimated_value_var)
    ]

    for i, (label, var) in enumerate(fields):
        tk.Label(frame, text=label, font=("Arial", 12), bg=FRAME_COLOR, fg=TEXT_COLOR).grid(row=i + 1, column=0, sticky="w", pady=5)
        tk.Entry(frame, textvariable=var, font=("Arial", 12), width=30).grid(row=i + 1, column=1, pady=5)

    tk.Button(frame, text="Save", command=save_record, font=("Arial", 12, "bold"), bg=BUTTON_COLOR, fg=BUTTON_TEXT_COLOR, padx=20, pady=5).grid(row=len(fields) + 1, column=0, columnspan=2, pady=20)

# Function to Add Demand Record
def add_demand_record():
    def save_record():
        try:
            new_record = pd.DataFrame([{
                "Indent No": indent_no_var.get(),
                "Demand Date": demand_date_var.get(),
                "Item Name": item_name_var.get(),
                "Quantity Issued": quantity_issued_var.get(),
                "Actual Price": actual_price_var.get()
            }])

            try:
                df = pd.read_excel(DEMAND_RECORDS_FILE)
            except FileNotFoundError:
                df = pd.DataFrame(columns=new_record.columns)

            df = pd.concat([df, new_record], ignore_index=True)
            df.to_excel(DEMAND_RECORDS_FILE, index=False)
            messagebox.showinfo("Success", "Record added successfully to Excel!")
            window.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save record: {e}")

    window = tk.Toplevel()
    window.title("Add Demand Record")
    window.geometry("500x500")
    window.configure(bg=BG_COLOR)

    frame = tk.Frame(window, bg=FRAME_COLOR, padx=20, pady=20, relief="groove", bd=2)
    frame.pack(pady=20)

    tk.Label(frame, text="Add Demand Record", font=("Arial", 16, "bold"), fg=HEADING_COLOR, bg=FRAME_COLOR).grid(row=0, column=0, columnspan=2, pady=10)

    indent_no_var = tk.IntVar()
    demand_date_var = tk.StringVar()
    item_name_var = tk.StringVar()
    quantity_issued_var = tk.IntVar()
    actual_price_var = tk.DoubleVar()

    fields = [
        ("Indent No:", indent_no_var),
        ("Demand Date (YYYY-MM-DD):", demand_date_var),
        ("Item Name:", item_name_var),
        ("Quantity Issued:", quantity_issued_var),
        ("Actual Price:", actual_price_var)
    ]

    for i, (label, var) in enumerate(fields):
        tk.Label(frame, text=label, font=("Arial", 12), bg=FRAME_COLOR, fg=TEXT_COLOR).grid(row=i + 1, column=0, sticky="w", pady=5)
        tk.Entry(frame, textvariable=var, font=("Arial", 12), width=30).grid(row=i + 1, column=1, pady=5)

    tk.Button(frame, text="Save", command=save_record, font=("Arial", 12, "bold"), bg=BUTTON_COLOR, fg=BUTTON_TEXT_COLOR, padx=20, pady=5).grid(row=len(fields) + 1, column=0, columnspan=2, pady=20)

# Function to Search Material Indent Records
def search_material_indent():
    def perform_search():
        try:
            search_field = search_field_var.get()
            search_value = search_value_var.get()

            demand_df = pd.read_excel(DEMAND_RECORDS_FILE)
            indent_df = pd.read_excel(INDENT_RECORDS_FILE)

            if search_field == "Indent No":
                results = demand_df[demand_df["Indent No"].astype(str).str.contains(search_value, case=False, na=False)]
            elif search_field == "Issued Date":
                results = demand_df[demand_df["Demand Date"].astype(str).str.contains(search_value, case=False, na=False)]

            if not results.empty:
                results["Remaining Items"] = results.apply(
                    lambda row: max(
                        indent_df[indent_df["Items"].str.contains(row["Item Name"], case=False, na=False)]["No of Items"].sum() - row["Quantity Issued"],
                        0
                    ), axis=1
                )

            export_to_excel(results.values.tolist(), SEARCH_RESULTS_FILE, results.columns)

            for row in results.itertuples(index=False):
                tree.insert("", "end", values=row)
        except Exception as e:
            messagebox.showerror("Error", f"Search failed: {e}")

    window = tk.Toplevel()
    window.title("Search Material Indent Records")
    window.geometry("700x500")
    window.configure(bg=BG_COLOR)

    frame = tk.Frame(window, bg=FRAME_COLOR, padx=20, pady=20, relief="groove", bd=2)
    frame.pack(pady=20, fill="both", expand=True)

    tk.Label(frame, text="Search Material Indent Records", font=("Arial", 16, "bold"), fg=HEADING_COLOR, bg=FRAME_COLOR).grid(row=0, column=0, columnspan=2, pady=10)

    search_field_var = tk.StringVar(value="Indent No")
    search_value_var = tk.StringVar()

    tk.Label(frame, text="Search Field:", font=("Arial", 12), bg=FRAME_COLOR, fg=TEXT_COLOR).grid(row=1, column=0, pady=5, sticky="w")
    ttk.Combobox(frame, textvariable=search_field_var, values=["Indent No", "Issued Date"], font=("Arial", 12)).grid(row=1, column=1, pady=5)

    tk.Label(frame, text="Search Value:", font=("Arial", 12), bg=FRAME_COLOR, fg=TEXT_COLOR).grid(row=2, column=0, pady=5, sticky="w")
    tk.Entry(frame, textvariable=search_value_var, font=("Arial", 12), width=30).grid(row=2, column=1, pady=5)

    tree = ttk.Treeview(frame, columns=["Indent No", "Demand Date", "Item Name", "Quantity Issued", "Actual Price", "Remaining Items"], show="headings")
    tree.grid(row=3, column=0, columnspan=2, pady=10, sticky="nsew")
    frame.grid_rowconfigure(3, weight=1)
    frame.grid_columnconfigure(1, weight=1)

    for col in tree["columns"]:
        tree.heading(col, text=col, anchor="w")
        tree.column(col, width=100, anchor="w")

    tk.Button(frame, text="Search", command=perform_search, font=("Arial", 12, "bold"), bg=BUTTON_COLOR, fg=BUTTON_TEXT_COLOR, padx=20).grid(row=4, column=0, pady=10)
    tk.Button(frame, text="Clear", command=lambda: tree.delete(*tree.get_children()), font=("Arial", 12, "bold"), bg=HIGHLIGHT_COLOR, fg=BUTTON_TEXT_COLOR, padx=20).grid(row=4, column=1, pady=10)

# Main Application Window
root = tk.Tk()
root.title("Material Management System")
root.geometry("400x500")
root.configure(bg=BG_COLOR)

frame = tk.Frame(root, bg=FRAME_COLOR, padx=20, pady=20, relief="groove", bd=2)
frame.pack(pady=20, fill="both", expand=True)

heading_label = tk.Label(frame, text="Material Management System", font=("Arial", 18, "bold"), fg=HEADING_COLOR, bg=FRAME_COLOR)
heading_label.pack(pady=20)

tk.Button(frame, text="Add Material Indent", command=add_material_indent, font=("Arial", 12, "bold"), bg=BUTTON_COLOR, fg=BUTTON_TEXT_COLOR, padx=20, pady=5).pack(pady=10)
tk.Button(frame, text="Add Demand Record", command=add_demand_record, font=("Arial", 12, "bold"), bg=BUTTON_COLOR, fg=BUTTON_TEXT_COLOR, padx=20, pady=5).pack(pady=10)
tk.Button(frame, text="Search Material Indent", command=search_material_indent, font=("Arial", 12, "bold"), bg=BUTTON_COLOR, fg=BUTTON_TEXT_COLOR, padx=20, pady=5).pack(pady=10)

root.mainloop()
