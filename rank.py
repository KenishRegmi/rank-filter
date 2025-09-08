import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pdfplumber

class ApplicantFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Applicant Filter (Rank Sorted)")

        # Load button
        self.load_btn = tk.Button(root, text="Load File", command=self.load_file)
        self.load_btn.pack(pady=5)

        # Filter controls
        control_frame = tk.Frame(root)
        control_frame.pack(pady=5)

        tk.Label(control_frame, text="Select Preference (p1–p9):").grid(row=0, column=0, padx=5)
        self.pref_choice = ttk.Combobox(control_frame, values=[f"p{i}" for i in range(1, 10)], state="readonly")
        self.pref_choice.grid(row=0, column=1, padx=5)
        self.pref_choice.current(0)

        tk.Label(control_frame, text="Value(s):").grid(row=0, column=2, padx=5)
        self.value_entry = tk.Entry(control_frame, width=15)
        self.value_entry.grid(row=0, column=3, padx=5)

        self.filter_btn = tk.Button(control_frame, text="Apply Filter", command=self.apply_filter)
        self.filter_btn.grid(row=0, column=4, padx=5)

        # Checkbox for searching all preferences
        self.search_all = tk.BooleanVar()
        self.search_all_check = tk.Checkbutton(root, text="Search in all preferences (p1–p9)", variable=self.search_all)
        self.search_all_check.pack(pady=5)

        # Export button
        self.export_btn = tk.Button(root, text="Export Filtered Results", command=self.export_results)
        self.export_btn.pack(pady=5)

        # Results table with S.N column
        self.tree = ttk.Treeview(
            root,
            columns=("S.N","Rank", "Applicant Name", "Gender", "District",
                     "p1","p2","p3","p4","p5","p6","p7","p8","p9"),
            show="headings"
        )
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=90, anchor="center")
        self.tree.pack(fill="both", expand=True, pady=10)

        self.df = None
        self.filtered = None

    # -------- File Loader --------
    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[
            ("Excel files", "*.xlsx"),
            ("CSV files", "*.csv"),
            ("PDF files", "*.pdf")
        ])
        if not file_path:
            return
        try:
            if file_path.endswith(".csv"):
                self.df = pd.read_csv(file_path)
            elif file_path.endswith(".xlsx"):
                self.df = pd.read_excel(file_path)
            elif file_path.endswith(".pdf"):
                self.df = self.extract_from_pdf(file_path)
            else:
                messagebox.showerror("Error", "Unsupported file format.")
                return
            messagebox.showinfo("Success", "File loaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{e}")

    # -------- PDF Extractor --------
    def extract_from_pdf(self, pdf_path):
        """Extract tables from PDF using pdfplumber"""
        all_data = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    all_data.extend(table)

        # First row is header
        df = pd.DataFrame(all_data[1:], columns=all_data[0])

        # Try converting numeric columns
        for col in ["Rank","p1","p2","p3","p4","p5","p6","p7","p8","p9"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")
        return df

    # -------- Filtering --------
    def apply_filter(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Please load a file first.")
            return

        try:
            # Allow multiple values like "11,12,13"
            values = [int(v.strip()) for v in self.value_entry.get().split(",") if v.strip()]
        except ValueError:
            messagebox.showwarning("Warning", "Please enter valid number(s), separated by commas.")
            return

        if self.search_all.get():
            # Search across all preferences (p1–p9)
            conditions = False
            for col in [f"p{i}" for i in range(1, 10) if f"p{i}" in self.df.columns]:
                conditions |= self.df[col].isin(values)
            self.filtered = self.df[conditions].sort_values(by="Rank")
        else:
            # Search only in selected preference
            pref = self.pref_choice.get()
            self.filtered = self.df[self.df[pref].isin(values)].sort_values(by="Rank")

        # Clear old data
        for row in self.tree.get_children():
            self.tree.delete(row)

        # Insert new data with numbering
        for idx, (_, row) in enumerate(self.filtered.iterrows(), start=1):
            self.tree.insert("", "end", values=(idx, row["Rank"], row["Applicant Name"], row["Gender"], row["District"],
                                                row.get("p1"), row.get("p2"), row.get("p3"), row.get("p4"), row.get("p5"),
                                                row.get("p6"), row.get("p7"), row.get("p8"), row.get("p9")))

    # -------- Export Results --------
    def export_results(self):
        if self.filtered is None or self.filtered.empty:
            messagebox.showwarning("Warning", "No filtered results to export.")
            return

        # Add S.N column for export
        export_df = self.filtered.copy().reset_index(drop=True)
        export_df.insert(0, "S.N", range(1, len(export_df)+1))

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if not file_path:
            return
        try:
            if file_path.endswith(".csv"):
                export_df.to_csv(file_path, index=False)
            else:
                export_df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"Filtered results saved to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ApplicantFilterApp(root)
    root.mainloop()
