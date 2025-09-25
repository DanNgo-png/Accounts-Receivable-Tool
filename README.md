# Accounts Receivable Analysis Tool

A user-friendly desktop application built with Python to help businesses analyze their accounts receivable data. This tool provides a clear overview of outstanding invoices, calculates aging buckets, highlights overdue items, and offers a simple cash inflow projection.

![image](https://github.com/user-attachments/assets/b8849788-b413-4318-8f56-621f92e7681c)

## Features

-   **Load Excel Data:** Easily import your accounts receivable ledger from an `.xlsx` or `.xls` file.
-   **Automated Aging Analysis:** Automatically calculates the age of each invoice and categorizes them into standard aging buckets (0-30, 31-60, 61-90, 90+ days).
-   **Interactive Data Grid:** View, sort, and review all your invoices in a clean, table-based format.
-   **Visual Highlighting:** Overdue invoices (older than 30 days) are highlighted in red for immediate attention.
-   **Financial Summary:** Get a quick overview of the total amount outstanding in each aging bucket.
-   **Cash Inflow Projection:** A simple forecast of expected cash collections in the next 30 days based on current receivables.

## Prerequisites

Before you run the application, make sure you have Python installed on your system. You will also need to install the following libraries:

-   `customtkinter`
-   `pandas`
-   `openpyxl` (required by pandas to read `.xlsx` files)

## Installation & Usage

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/your-username/Accounts-Receivable-Tool.git
    cd Accounts-Receivable-Tool
    ```

2.  **Install the required libraries:**
    ```bash
    pip install customtkinter pandas openpyxl
    ```

3.  **Run the application:**
    ```bash
    python main.py
    ```

4.  **Load your data:**
    -   Click the **"Load Excel File"** button.
    -   Select your accounts receivable Excel file.
    -   The tool will process the data and display the analysis.

## Input File Format

The tool requires an Excel file (`.xlsx` or `.xls`) with specific columns for proper analysis. Please ensure your file contains the following headers:

-   `InvoiceID`: A unique identifier for each invoice.
-   `CustomerName`: The name of the customer.
-   `InvoiceDate`: The date the invoice was issued (e.g., `YYYY-MM-DD`).
-   `Amount`: The total amount of the invoice.

### Sample `receivables.xlsx`:

| InvoiceID | CustomerName  | InvoiceDate | Amount    |
| :-------- | :------------ | :---------- | :-------- |
| INV001    | Alpha Corp    | 2024-04-15  | 1500.00   |
| INV002    | Bravo LLC     | 2024-05-01  | 250.75    |
| INV003    | Charlie Inc   | 2024-05-20  | 800.50    |
| INV004    | Delta Co      | 2024-03-10  | 3200.00   |
| INV005    | Alpha Corp    | 2024-02-05  | 5000.00   |
