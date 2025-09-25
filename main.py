import customtkinter as ctk
from tkinter import filedialog, ttk
import pandas as pd
from datetime import datetime

class ARTool(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Accounts Receivable Analysis Tool")
        self.geometry("1200x650")
        ctk.set_appearance_mode("dark")

        # --- Main Frame ---
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=15, pady=15)

        # --- Header ---
        self.header_frame = ctk.CTkFrame(self.main_frame, corner_radius=10)
        self.header_frame.pack(fill="x", pady=(0, 15))

        self.load_button = ctk.CTkButton(
            self.header_frame, 
            text="📂 Load Excel File", 
            command=self.load_excel_data, 
            width=160
        )
        self.load_button.pack(side="left", padx=15, pady=15)

        self.filename_label = ctk.CTkLabel(
            self.header_frame, 
            text="No file loaded", 
            text_color="gray", 
            anchor="w"
        )
        self.filename_label.pack(side="left", padx=10)

        # --- Summary ---
        self.summary_frame = ctk.CTkFrame(self.main_frame, corner_radius=10)
        self.summary_frame.pack(fill="x", pady=10, padx=5)

        self.summary_label = ctk.CTkLabel(
            self.summary_frame, 
            text="📊 Summary", 
            font=ctk.CTkFont(size=18, weight="bold")
        )
        self.summary_label.grid(row=0, column=0, columnspan=2, sticky="w", padx=15, pady=(15, 5))

        self.aging_buckets_label = ctk.CTkLabel(
            self.summary_frame, text="", font=ctk.CTkFont(size=14), justify="left"
        )
        self.aging_buckets_label.grid(row=1, column=0, sticky="w", padx=25, pady=10)

        self.cash_inflow_label = ctk.CTkLabel(
            self.summary_frame, text="", font=ctk.CTkFont(size=14), justify="left"
        )
        self.cash_inflow_label.grid(row=1, column=1, sticky="w", padx=25, pady=10)

        # --- Treeview for Data Display ---
        self.tree_frame = ctk.CTkFrame(self.main_frame, corner_radius=10)
        self.tree_frame.pack(fill="both", expand=True, pady=(15, 0), padx=5)

        # Add Scrollbars
        self.tree_scroll_y = ctk.CTkScrollbar(self.tree_frame)
        self.tree_scroll_y.pack(side="right", fill="y")

        self.tree_scroll_x = ctk.CTkScrollbar(self.tree_frame, orientation="horizontal")
        self.tree_scroll_x.pack(side="bottom", fill="x")

        # --- Treeview Styling ---
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Treeview",
            background="#2a2d2e",
            foreground="white",
            rowheight=28,
            fieldbackground="#2a2d2e",
            font=("Segoe UI", 12)
        )
        style.map("Treeview", background=[("selected", "#22559b")])

        style.configure(
            "Treeview.Heading",
            background="#3e4142",
            foreground="white",
            font=("Segoe UI", 13, "bold")
        )

        self.tree = ttk.Treeview(
            self.tree_frame,
            columns=("InvoiceID", "CustomerName", "InvoiceDate", "Amount", "Age", "AgingBucket"),
            show="headings",
            yscrollcommand=self.tree_scroll_y.set,
            xscrollcommand=self.tree_scroll_x.set
        )
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        self.tree_scroll_y.configure(command=self.tree.yview)
        self.tree_scroll_x.configure(command=self.tree.xview)

        # --- Treeview Headers ---
        headers = {
            "InvoiceID": "Invoice ID",
            "CustomerName": "Customer Name",
            "InvoiceDate": "Invoice Date",
            "Amount": "Amount",
            "Age": "Age (Days)",
            "AgingBucket": "Aging Bucket"
        }
        for col, text in headers.items():
            self.tree.heading(col, text=text, anchor="center")
            self.tree.column(col, anchor="center", width=150, stretch=True)

        # --- Tag for overdue items ---
        self.tree.tag_configure("overdue", background="#8B0000", foreground="white")

    def load_excel_data(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return
        self.filename_label.configure(text=file_path.split("/")[-1])
        self.process_data(file_path)

    def process_data(self, file_path):
        df = pd.read_excel(file_path)
        df['InvoiceDate'] = pd.to_datetime(df['InvoiceDate'])
        df['Age'] = (datetime.now() - df['InvoiceDate']).dt.days

        bins = [0, 30, 60, 90, float('inf')]
        labels = ['0-30 Days', '31-60 Days', '61-90 Days', '90+ Days']
        df['AgingBucket'] = pd.cut(df['Age'], bins=bins, labels=labels, right=False)

        # Clear old data
        for i in self.tree.get_children():
            self.tree.delete(i)

        # Insert rows
        for _, row in df.iterrows():
            tags = ("overdue",) if row['Age'] > 30 else ()
            self.tree.insert(
                "", "end",
                values=(
                    row['InvoiceID'],
                    row['CustomerName'],
                    row['InvoiceDate'].strftime('%Y-%m-%d'),
                    f"${row['Amount']:,.2f}",
                    row['Age'],
                    row['AgingBucket']
                ),
                tags=tags
            )

        self.update_summary(df)

    def update_summary(self, df):
        # Aging Buckets
        aging_summary = df.groupby('AgingBucket')['Amount'].sum().reset_index()
        aging_text = "Aging Buckets:\n"
        for _, row in aging_summary.iterrows():
            aging_text += f"  {row['AgingBucket']}: ${row['Amount']:,.2f}\n"
        self.aging_buckets_label.configure(text=aging_text)

        # Cash Inflow
        inflow_30_days = df[df['AgingBucket'] == '0-30 Days']['Amount'].sum()
        cash_inflow_text = (
            "Projected Cash Inflows (next 30 days):\n"
            f"  Expected from '0-30 Days' bucket: ${inflow_30_days:,.2f}"
        )
        self.cash_inflow_label.configure(text=cash_inflow_text)

if __name__ == "__main__":
    app = ARTool()
    app.mainloop()