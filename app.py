import tkinter as tk
from tkinter import filedialog, messagebox
import pandpas as pd

class ExcelFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Column Filter")
        self.root.geometry("400x400")

        self.filepath = None
        self.df = None
        self.columns_vars = []

        self.load_button = tk.Button(root, text="Load Excel File", command=self.load_file)
        self.load_button.pack(pady=10)

        self.columns_frame = tk.Frame(root)
        self.columns_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        self.save_button = tk.Button(root, text="Save Filtered File", command=self.save_file, state=tk.DISABLED)
        self.save_button.pack(pady=10)

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

        # Clear previous columns checkboxes
        for widget in self.columns_frame.winfo_children():
            widget.destroy()
        self.columns_vars = []

        # Create checkbuttons for columns
        if self.df.empty:
            messagebox.showinfo("Info", "The loaded Excel file has no data.")
            return

        tk.Label(self.columns_frame, text="Select columns to keep:").pack(anchor='w')
        for col in self.df.columns:
            var = tk.BooleanVar(value=True)
            cb = tk.Checkbutton(self.columns_frame, text=col, variable=var)
            cb.pack(anchor='w')
            self.columns_vars.append((col, var))

        self.save_button.config(state=tk.NORMAL)

    def save_file(self):
        selected_columns = [col for col, var in self.columns_vars if var.get()]
        if not selected_columns:
            messagebox.showwarning("Warning", "Please select at least one column to save.")
            return

        filetypes = [('Excel files', '*.xlsx')]
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=filetypes, title="Save filtered Excel file")
        if not save_path:
            return

        try:
            self.df.to_excel(save_path, columns=selected_columns, index=False)
            messagebox.showinfo("Success", f"Filtered file saved successfully to:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelFilterApp(root)
    root.mainloop()
