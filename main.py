
import customtkinter as ctk
from tkinter import filedialog, ttk
import pandas as pd
from datetime import datetime

class ARTool(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Accounts Receivable Analysis Tool")
        self.geometry("1200x600")

        # --- Main Frame ---
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # --- Header ---
        self.header_frame = ctk.CTkFrame(self.main_frame)
        self.header_frame.pack(fill="x", padx=10, pady=10)

        self.load_button = ctk.CTkButton(self.header_frame, text="Load Excel File", command=self.load_excel_data)
        self.load_button.pack(side="left", padx=10)

        self.filename_label = ctk.CTkLabel(self.header_frame, text="No file loaded", text_color="gray")
        self.filename_label.pack(side="left", padx=10)

        # --- Summary ---
        self.summary_frame = ctk.CTkFrame(self.main_frame)
        self.summary_frame.pack(fill="x", padx=10, pady=10)

        self.summary_label = ctk.CTkLabel(self.summary_frame, text="Summary:", font=("Arial", 16, "bold"))
        self.summary_label.pack(side="left", padx=10)

        self.aging_buckets_label = ctk.CTkLabel(self.summary_frame, text="")
        self.aging_buckets_label.pack(side="left", padx=20)
        
        self.cash_inflow_label = ctk.CTkLabel(self.summary_frame, text="")
        self.cash_inflow_label.pack(side="left", padx=20)

        # --- Treeview for Data Display ---
        self.tree_frame = ctk.CTkFrame(self.main_frame)
        self.tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.tree = ttk.Treeview(self.tree_frame, columns=("InvoiceID", "CustomerName", "InvoiceDate", "Amount", "Age", "AgingBucket"), show="headings")
        self.tree.pack(fill="both", expand=True)

        # --- Treeview Headers ---
        self.tree.heading("InvoiceID", text="Invoice ID")
        self.tree.heading("CustomerName", text="Customer Name")
        self.tree.heading("InvoiceDate", text="Invoice Date")
        self.tree.heading("Amount", text="Amount")
        self.tree.heading("Age", text="Age (Days)")
        self.tree.heading("AgingBucket", text="Aging Bucket")

        # --- Treeview Styling ---
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2a2d2e", foreground="white", fieldbackground="#2a2d2e", borderwidth=0)
        style.map('Treeview', background=[('selected', '#22559b')])
        style.configure("Treeview.Heading", background="#565b5e", foreground="white", relief="flat")
        style.map("Treeview.Heading", background=[('active', '#3484F0')])

        # --- Tag for overdue items ---
        self.tree.tag_configure("overdue", background="#8B0000")


    def load_excel_data(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return

        self.filename_label.configure(text=file_path.split("/")[-1])
        self.process_data(file_path)

    def process_data(self, file_path):
        df = pd.read_excel(file_path)
        
        # --- Data Processing ---
        df['InvoiceDate'] = pd.to_datetime(df['InvoiceDate'])
        df['Age'] = (datetime.now() - df['InvoiceDate']).dt.days
        
        bins = [0, 30, 60, 90, float('inf')]
        labels = ['0-30 Days', '31-60 Days', '61-90 Days', '90+ Days']
        df['AgingBucket'] = pd.cut(df['Age'], bins=bins, labels=labels, right=False)

        # --- Update Treeview ---
        for i in self.tree.get_children():
            self.tree.delete(i)

        for _, row in df.iterrows():
            tags = ()
            if row['Age'] > 30: # Example: flag invoices older than 30 days
                tags = ("overdue",)
            
            self.tree.insert("", "end", values=(
                row['InvoiceID'],
                row['CustomerName'],
                row['InvoiceDate'].strftime('%Y-%m-%d'),
                f"${row['Amount']:,.2f}",
                row['Age'],
                row['AgingBucket']
            ), tags=tags)

        # --- Update Summary ---
        self.update_summary(df)

    def update_summary(self, df):
        # --- Aging Buckets Summary ---
        aging_summary = df.groupby('AgingBucket')['Amount'].sum().reset_index()
        aging_text = "Aging Buckets:\n"
        for _, row in aging_summary.iterrows():
            aging_text += f"  {row['AgingBucket']}: ${row['Amount']:,.2f}\n"
        self.aging_buckets_label.configure(text=aging_text)

        # --- Cash Inflow Projection ---
        # Simple projection: assumes amounts in buckets are collected in the next period
        cash_inflow_text = "Projected Cash Inflows (next 30 days):\n"
        inflow_30_days = df[df['AgingBucket'] == '0-30 Days']['Amount'].sum()
        cash_inflow_text += f"  Expected from '0-30 Days' bucket: ${inflow_30_days:,.2f}"
        self.cash_inflow_label.configure(text=cash_inflow_text)

if __name__ == "__main__":
    app = ARTool()
    app.mainloop()