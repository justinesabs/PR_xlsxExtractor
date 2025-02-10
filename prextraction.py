import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import pyperclip
from datetime import datetime
import io
import openpyxl

class SimpleFileCopier:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Simple Excel Data Copier")
        self.window.geometry("600x400")
        
        self.expected_columns = [
            '12-DigitBarcode', 'StockNo', 'Item Description', 'SuppCode',
            'BatchDate', 'Quantity', 'Price', 'SuppCode_BatchDate'
        ]
        
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        instructions = """
        How to use:
        1. Click 'Select Source File' to choose your Excel/CSV file
        2. Click 'Copy Data' to copy the data to clipboard
        3. Click 'Paste to File' to paste the copied data into another Excel file
        """
        ttk.Label(main_frame, text=instructions, justify=tk.LEFT).pack(pady=10)
        
        self.file_label = ttk.Label(main_frame, text="No file selected", wraplength=500)
        self.file_label.pack(pady=5)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="Select Source File", command=self.select_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Copy Data", command=self.copy_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Paste to File", command=self.paste_data).pack(side=tk.LEFT, padx=5)
        
        preview_frame = ttk.LabelFrame(main_frame, text="Data Preview", padding="5")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.preview = tk.Text(preview_frame, height=10, wrap=tk.NONE)
        scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview.yview)
        self.preview.configure(yscrollcommand=scrollbar.set)
        
        self.preview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.status = ttk.Label(main_frame, text="Ready")
        self.status.pack(pady=5)
        
        self.current_file = None
        self.copied_data = None

    def select_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.current_file = file_path
            self.file_label.config(text=f"Selected file: {file_path}")
            self.show_preview()
            self.status.config(text="File loaded successfully")

    def show_preview(self):
        try:
            if self.current_file.endswith('.csv'):
                df = pd.read_csv(self.current_file)
            else:
                df = pd.read_excel(self.current_file)
            
            if list(df.columns) != self.expected_columns:
                df.columns = self.expected_columns
            
            self.preview.delete(1.0, tk.END)
            self.preview.insert(tk.END, df.head().to_string(index=False, header=False))
            
        except Exception as e:
            self.preview.delete(1.0, tk.END)
            self.preview.insert(tk.END, f"Error previewing file: {str(e)}")

    def copy_data(self):
        if not self.current_file:
            messagebox.showwarning("Warning", "Please select a file first")
            return
            
        try:
            if self.current_file.endswith('.csv'):
                df = pd.read_csv(self.current_file)
            else:
                df = pd.read_excel(self.current_file)
            
            if list(df.columns) != self.expected_columns:
                df.columns = self.expected_columns
            
            df = df.fillna("")
            
            current_date = datetime.now().strftime("%m%y")
            if current_date[0] != '0':
                current_date = '0' + current_date
            
            df['BatchDate'] = current_date
            
            self.copied_data = df
            
            data_str = df.to_csv(sep="\t", index=False, header=False)
            pyperclip.copy(data_str)
            
            self.preview.delete(1.0, tk.END)
            self.preview.insert(tk.END, "Copied Data Preview (without headers):\n")
            self.preview.insert(tk.END, df.to_string(index=False, header=False))
            
            self.status.config(text=f"Data copied successfully! BatchDate: {current_date}")
            messagebox.showinfo("Success", "Data copied to clipboard!")
            
        except Exception as e:
            self.status.config(text=f"Error: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def save_to_excel_without_headers(self, df, filename):
        values = df.values
        
        wb = openpyxl.Workbook()
        ws = wb.active
        
        for row_idx, row in enumerate(values, 1):
            for col_idx, value in enumerate(row, 1):
                ws.cell(row=row_idx, column=col_idx, value=value)
        
        wb.save(filename)

    def paste_data(self):
        try:
            target_file = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xls")]
            )
            
            if not target_file:
                return
            
            clipboard_data = pyperclip.paste()
            
            new_data = pd.read_csv(
                io.StringIO(clipboard_data),
                sep="\t",
                header=None,
                names=self.expected_columns,
                dtype=str
            )
            
            try:
                target_df = pd.read_excel(target_file, header=None)
                if not target_df.empty:
                    target_df.columns = self.expected_columns
            except Exception:
                target_df = pd.DataFrame(columns=self.expected_columns)
            
            combined_df = pd.concat([target_df, new_data], ignore_index=True)
            
            self.save_to_excel_without_headers(combined_df, target_file)
            
            self.preview.delete(1.0, tk.END)
            self.preview.insert(tk.END, "Pasted Data Preview (without headers):\n")
            self.preview.insert(tk.END, combined_df.tail().to_string(index=False, header=False))
            
            self.status.config(text=f"Data pasted successfully to {target_file}")
            messagebox.showinfo("Success", "Data pasted successfully!")
            
        except Exception as e:
            self.status.config(text=f"Error: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}\n\nPlease make sure you've copied data first.")

    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = SimpleFileCopier()
    app.run()