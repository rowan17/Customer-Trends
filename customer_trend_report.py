import pyodbc
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime

# --- Configuration ---
# MSSQL connection details
server = '192.168.68.54'
database = 'pcay'
username = 'paradise'
password = 'paradise1'

# Number of years for the report
YEARS_TO_REPORT = 6

# --- Dynamic Year Calculation ---
current_date = datetime.now()
current_year = current_date.year
day_of_year = current_date.timetuple().tm_yday
years = list(range(current_year - YEARS_TO_REPORT + 1, current_year + 1))
years_str = ", ".join(map(str, years))

# --- Dynamic SQL Query Generation ---
# Build the parts of the query that depend on the years
purchase_sum_cols = []
for year in years:
    purchase_sum_cols.append(f"SUM(CASE WHEN YEAR(COALESCE(TxnDate, ShipDate)) = {year} THEN InvoiceLineAmount ELSE 0 END) AS Purchased{year}")

purchase_sum_statement = ",\n        ".join(purchase_sum_cols)

final_select_cols = []
for year in years:
    final_select_cols.append(f"cs.Purchased{year}")

final_select_statement = ",\n    ".join(final_select_cols)

query = f'''
WITH CustomerSales AS (
    SELECT
        CustomerRefListID AS CustomerID,
        {purchase_sum_statement},
        SUM(InvoiceLineAmount) AS TotalSales
    FROM dbo.InvoiceLine
    WHERE
        CustomerRefListID IS NOT NULL
        AND YEAR(COALESCE(TxnDate, ShipDate)) IN ({years_str})
        AND DATEPART(dy, COALESCE(TxnDate, ShipDate)) <= {day_of_year}
    GROUP BY CustomerRefListID
)
SELECT
    c.Name AS CustomerName,
    {final_select_statement},
    cs.TotalSales,
    CONCAT_WS(' ', c.FirstName, c.LastName) AS ContactName,
    c.Email,
    c.Phone AS PhoneNumber,
    CONCAT_WS(', ', c.BillAddressAddr1, c.BillAddressAddr2, c.BillAddressCity, c.BillAddressState, c.BillAddressPostalCode) AS BillingAddress
FROM CustomerSales cs
JOIN dbo.Customer c ON cs.CustomerID = c.ListID
WHERE cs.TotalSales > 0
ORDER BY c.Name;
'''

# --- Database Connection and Query Execution ---
try:
    conn_str = (
        f'DRIVER={{ODBC Driver 18 for SQL Server}};'
        f'SERVER={server};'
        f'DATABASE={database};'
        f'UID={username};'
        f'PWD={password};'
        f'TrustServerCertificate=yes;'
    )
    conn = pyodbc.connect(conn_str)
    
    # Run query and fetch results
    results = pd.read_sql(query, conn)
    
    # --- Report Generation ---
    print(f'\nCustomer Purchase Report for the Last {YEARS_TO_REPORT} Years (All Companies)')
    print('-' * 120)
    print(results.to_string(index=False))
    print('-' * 120)

    # Calculate and print summary statistics
    print(f'\nSummary:')
    for year in years:
        col_name = f'Purchased{year}'
        total_sales = results[col_name].sum()
        print(f'Total Sales {year}: ${total_sales:,.2f}')

    # --- Excel Export and Formatting ---
    excel_path = 'customer_change_report_last_6_years.xlsx'
    
    # Insert empty columns for spacing
    contact_name_index = results.columns.get_loc('ContactName')
    results.insert(contact_name_index + 1, ' ', '')

    results.to_excel(excel_path, index=False)

    # Load workbook for formatting
    wb = load_workbook(excel_path)
    ws = wb.active

    # --- Excel Formatting ---
    wb = load_workbook(excel_path)
    ws = wb.active

    # Define fills
    yellow_fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
    orange_fill = PatternFill(start_color='FFFFA500', end_color='FFFFA500', fill_type='solid')

    # Find the header names to locate columns dynamically
    headers = [cell.value for cell in ws[1]]
    
    try:
        # Get the column index for TotalSales
        total_sales_col_idx = headers.index('TotalSales') + 1
        total_sales_col_letter = chr(ord('A') + total_sales_col_idx - 1)
    except ValueError:
        total_sales_col_idx = -1 # TotalSales column not found

    # Find the columns for the years
    year_columns_indices = [i for i, h in enumerate(headers) if h and h.startswith('Purchased')]

    for row in range(2, ws.max_row + 1):
        # Determine which cells to highlight
        highlight_row = False
        purchases = {}
        for col_idx in year_columns_indices:
            cell_value = ws.cell(row=row, column=col_idx + 1).value
            if isinstance(cell_value, (int, float)):
                purchases[col_idx + 1] = cell_value
                if cell_value > 2000:
                    highlight_row = True

        # Apply highlights in the correct order
        if highlight_row:
            for col in range(1, len(headers) + 1):
                ws.cell(row=row, column=col).fill = orange_fill
        
        if purchases:
            max_year_col_idx = max(purchases, key=purchases.get)
            ws.cell(row=row, column=max_year_col_idx).fill = yellow_fill

    wb.save(excel_path)
    print(f'\nReport saved to {excel_path}.')

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    if 'conn' in locals() and conn:
        conn.close()
