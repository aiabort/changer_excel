import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import sqlite3
import os

class ExcelDatabaseApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Database Viewer")
        self.root.geometry("800x600")

        self.filepath = None
        self.df = None
        self.db_path = None
        self.conn = None
        self.table_name = "excel_data"

        # Load Excel Button
        self.load_button = tk.Button(root, text="Load Excel File", command=self.load_file)
        self.load_button.pack(pady=10)

        # Search Frame
        self.search_frame = tk.Frame(root)
        self.search_frame.pack(pady=5, fill=tk.X)

        tk.Label(self.search_frame, text="Search Column:").pack(side=tk.LEFT, padx=5)
        self.search_column_var = tk.StringVar()
        self.search_column_cb = ttk.Combobox(self.search_frame, textvariable=self.search_column_var, state="readonly")
        self.search_column_cb.pack(side=tk.LEFT, padx=5)

        tk.Label(self.search_frame, text="Search Term:").pack(side=tk.LEFT, padx=5)
        self.search_entry = tk.Entry(self.search_frame)
        self.search_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        self.search_button = tk.Button(self.search_frame, text="Search", command=self.search_data)
        self.search_button.pack(side=tk.RIGHT, padx=5)

        self.clear_button = tk.Button(self.search_frame, text="Clear Search", command=self.clear_search)
        self.clear_button.pack(side=tk.RIGHT, padx=5)

        # Treeview for displaying data
        self.tree_frame = tk.Frame(root)
        self.tree_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(self.tree_frame)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Scrollbars
        self.vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=self.vsb.set)

        self.hsb = ttk.Scrollbar(root, orient="horizontal", command=self.tree.xview)
        self.hsb.pack(fill=tk.X)
        self.tree.configure(xscrollcommand=self.hsb.set)

        # Save Button
        self.save_button = tk.Button(root, text="Save Filtered Data to Excel", command=self.save_file, state=tk.DISABLED)
        self.save_button.pack(pady=10)

        # Bind column header click for sorting
        self.tree.heading("#0", text="", command=lambda: self.sort_column("#0"))
        self.sort_order = {}

    def load_file(self):
        filetypes = [('Excel files', '*.xlsx *.xls')]
        filepath = filedialog.askopenfilename(title="Open Excel file", filetypes=filetypes)
        if not filepath:
            return

        try:
            self.df = pd.read_excel(filepath)
            self.filepath = filepath
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{e}")
            return

        if self.df.empty:
            messagebox.showinfo("Info", "The loaded Excel file has no data.")
            return

        # Create database
        self.db_path = filepath.replace('.xlsx', '.db').replace('.xls', '.db')
        self.conn = sqlite3.connect(self.db_path)
        self.df.to_sql(self.table_name, self.conn, if_exists='replace', index=False)

        # Setup Treeview
        self.setup_treeview()

        # Populate search column combobox
        self.search_column_cb['values'] = list(self.df.columns)
        if self.df.columns:
            self.search_column_var.set(self.df.columns[0])

        self.save_button.config(state=tk.NORMAL)

    def setup_treeview(self):
        # Clear existing columns
        for col in self.tree['columns']:
            self.tree.heading(col, text="")
            self.tree.column(col, width=0)

        # Set new columns
        self.tree['columns'] = list(self.df.columns)
        for col in self.df.columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_column(c))
            self.tree.column(col, width=100, anchor='w')

        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Insert data
        for index, row in self.df.iterrows():
            values = [row[col] for col in self.df.columns]
            self.tree.insert("", tk.END, values=values)

    def sort_column(self, col):
        if col not in self.sort_order:
            self.sort_order[col] = False
        self.sort_order[col] = not self.sort_order[col]

        # Get data from tree
        data = [(self.tree.set(child, col), child) for child in self.tree.get_children('')]

        # Sort data
        data.sort(reverse=self.sort_order[col])

        # Rearrange items in sorted positions
        for index, (val, child) in enumerate(data):
            self.tree.move(child, '', index)

        # Reverse sort next time
        self.sort_order[col] = not self.sort_order[col]

    def search_data(self):
        column = self.search_column_var.get()
        term = self.search_entry.get().strip()
        if not column or not term:
            messagebox.showwarning("Warning", "Please select a column and enter a search term.")
            return

        # Query database
        query = f'SELECT * FROM {self.table_name} WHERE "{column}" LIKE ?'
        cursor = self.conn.cursor()
        cursor.execute(query, ('%' + term + '%',))
        rows = cursor.fetchall()

        # Clear tree
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Insert filtered data
        for row in rows:
            self.tree.insert("", tk.END, values=row)

    def clear_search(self):
        self.search_entry.delete(0, tk.END)
        self.setup_treeview()

    def save_file(self):
        filetypes = [('Excel files', '*.xlsx')]
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=filetypes, title="Save filtered Excel file")
        if not save_path:
            return

        # Get current data from tree
        data = []
        for child in self.tree.get_children():
            data.append(self.tree.item(child)['values'])

        if not data:
            messagebox.showwarning("Warning", "No data to save.")
            return

        # Create DataFrame
        filtered_df = pd.DataFrame(data, columns=self.df.columns)

        try:
            filtered_df.to_excel(save_path, index=False)
            messagebox.showinfo("Success", f"Filtered file saved successfully to:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file:\n{e}")

    def __del__(self):
        if self.conn:
            self.conn.close()

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelDatabaseApp(root)
    root.mainloop()
