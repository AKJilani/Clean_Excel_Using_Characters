import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import string

class ExcelCleanerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Unicode Cleaner")
        self.root.geometry("550x320")

        self.file_path = ""

        tk.Button(root, text="Select Excel File", command=self.load_file).pack(pady=8)

        tk.Label(root, text="Sheet Name").pack()
        self.sheet_entry = tk.Entry(root, width=45)
        self.sheet_entry.pack()

        tk.Label(root, text="Allowed Special Characters (space separated)").pack()
        self.char_entry = tk.Entry(root, width=45)
        self.char_entry.pack()

        tk.Button(root, text="Start Cleaning", command=self.start_cleaning,
                  bg="green", fg="white").pack(pady=10)

        self.progress = ttk.Progressbar(root, length=450, mode='determinate')
        self.progress.pack(pady=10)

    def load_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        messagebox.showinfo("Selected File", self.file_path)

    def build_allowed_set(self, extra_chars):
        # always include a-z automatically
        base_chars = set(string.ascii_lowercase)

        # add user input chars (converted to lowercase)
        extra = set(ch.lower() for ch in extra_chars.split())

        return base_chars.union(extra)

    def is_row_valid(self, text, allowed_set):
        text = str(text).lower()

        for ch in text:
            if ch.strip() == "":
                continue
            if ch not in allowed_set:
                return False
        return True

    def start_cleaning(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please select Excel file")
            return

        sheet_name = self.sheet_entry.get()
        allowed_input = self.char_entry.get()

        try:
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, engine="openpyxl")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        allowed_set = self.build_allowed_set(allowed_input)

        total_rows = len(df)
        self.progress["maximum"] = total_rows

        cleaned_rows = []

        for i, row in df.iterrows():
            row_text = " ".join(str(v) for v in row.values if pd.notna(v))

            if self.is_row_valid(row_text, allowed_set):
                cleaned_rows.append(row)

            self.progress["value"] = i + 1
            self.root.update_idletasks()

        cleaned_df = pd.DataFrame(cleaned_rows)

        save_path = self.file_path.replace(".xlsx", "_cleaned.xlsx")
        cleaned_df.to_excel(save_path, index=False, sheet_name=sheet_name)

        messagebox.showinfo("Done", f"Clean file saved:\n{save_path}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelCleanerApp(root)
    root.mainloop()