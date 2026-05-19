import pyodbc
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# MSSQL connection details
server = 'QBSERVER'
database = 'pcay'
username = 'paradise'
password = 'paradise1'

# Connect to MSSQL
conn_str = (
    f'DRIVER={{ODBC Driver 18 for SQL Server}};'
    f'SERVER={server};'
    f'DATABASE={database};'
    f'UID={username};'
    f'PWD={password};'
    f'TrustServerCertificate=yes;'
)
conn = pyodbc.connect(conn_str)

# All companies, showing 2023-2025 totals
query = '''
WITH AllCustomerIDs AS (
    -- First, get a distinct list of all customer IDs that have any activity
    -- (transaction or shipment) in the target years. This prevents missing customers.
    SELECT DISTINCT CustomerRefListID
    FROM dbo.InvoiceLine
    WHERE 
        (YEAR(ShipDate) IN (2023, 2024, 2025) OR YEAR(TxnDate) IN (2023, 2024, 2025))
        AND CustomerRefListID IS NOT NULL
),
AllCustomers AS (
    -- Now, join the definitive list of IDs with the Customer table to get their names.
    SELECT
        cid.CustomerRefListID AS CustomerID,
        c.Name AS CustomerName
    FROM AllCustomerIDs cid
    JOIN dbo.Customer c ON cid.CustomerRefListID = c.ListID
),
YearlyTotals AS (
    -- This CTE calculates the total purchases for each customer per year based on transaction date.
    SELECT 
        CustomerRefListID AS CustomerID,
        YEAR(TxnDate) AS PurchaseYear,
        SUM(InvoiceLineAmount) AS TotalPurchased
    FROM dbo.InvoiceLine
    WHERE TxnDate IS NOT NULL AND YEAR(TxnDate) IN (2023, 2024, 2025)
    GROUP BY CustomerRefListID, YEAR(TxnDate)
)
SELECT
    ac.CustomerID,
    ac.CustomerName,
    COALESCE(y2023.TotalPurchased, 0) AS Purchased2023,
    COALESCE(y2024.TotalPurchased, 0) AS Purchased2024,
    COALESCE(y2025.TotalPurchased, 0) AS Purchased2025
FROM AllCustomers ac
LEFT JOIN YearlyTotals y2023 ON ac.CustomerID = y2023.CustomerID AND y2023.PurchaseYear = 2023
LEFT JOIN YearlyTotals y2024 ON ac.CustomerID = y2024.CustomerID AND y2024.PurchaseYear = 2024
LEFT JOIN YearlyTotals y2025 ON ac.CustomerID = y2025.CustomerID AND y2025.PurchaseYear = 2025
ORDER BY ac.CustomerName;
'''

# Run query and fetch results
results = pd.read_sql(query, conn)

# Calculate summary statistics first
total_2023 = results['Purchased2023'].sum()
total_2024 = results['Purchased2024'].sum()
total_2025 = results['Purchased2025'].sum()

# Create a copy of results for display
display_results = results.copy()

# Format the display values
display_results['Purchased2023'] = display_results['Purchased2023'].apply(lambda x: f'${x:,.2f}')
display_results['Purchased2024'] = display_results['Purchased2024'].apply(lambda x: f'${x:,.2f}')
display_results['Purchased2025'] = display_results['Purchased2025'].apply(lambda x: f'${x:,.2f}')

# Display results in a formatted table
print('\nCustomer Purchase Report 2023-2025 (All Companies)')
print('-' * 120)
print(results.to_string(index=False))
print('-' * 120)

print(f'\nSummary:')
print(f'Total Sales 2023: ${total_2023:,.2f}')
print(f'Total Sales 2024: ${total_2024:,.2f}')
print(f'Total Sales 2025: ${total_2025:,.2f}')

# Save to Excel for easy desktop viewing
excel_path = 'customer_trend_report_comprehensive.xlsx'
results.to_excel(excel_path, index=False)

# Magnitude-based color coding for change (2024→2025)
wb = load_workbook(excel_path)
ws = wb.active

def clamp(val):
    return max(0, min(255, int(val)))

def get_red_shade(magnitude, max_magnitude):
    intensity = clamp(200 + 55 * (magnitude / max_magnitude)) if max_magnitude > 0 else 255
    hex_color = f'FF{intensity:02X}C7CE'
    if len(hex_color) != 8:
        hex_color = 'FFC7CECE'
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type='solid')

def get_green_shade(magnitude, max_magnitude):
    intensity = clamp(200 + 55 * (magnitude / max_magnitude)) if max_magnitude > 0 else 255
    hex_color = f'FFC6EF{intensity:02X}'
    if len(hex_color) != 8:
        hex_color = 'FFC6EFCE'
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type='solid')

# Highlight for Top Year
yellow_fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
for row in range(2, ws.max_row + 1):
    # Highlight Top Year
    purchases = {
        'C': ws[f'C{row}'].value,
        'D': ws[f'D{row}'].value,
        'E': ws[f'E{row}'].value
    }
    
    # Find the max purchase value, ignoring non-numeric values
    numeric_purchases = {k: v for k, v in purchases.items() if isinstance(v, (int, float))}
    if numeric_purchases:
        max_year_col = max(numeric_purchases, key=numeric_purchases.get)
        ws[f'{max_year_col}{row}'].fill = yellow_fill

wb.save(excel_path)
print('\nReport saved to customer_trend_report_comprehensive.xlsx.')
