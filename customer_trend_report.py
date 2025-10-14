import pyodbc
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# MSSQL connection details
server = '192.168.68.54'
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

# All companies, showing 2023-2025 totals and change from 2024 to 2025
query = '''
WITH YearlyTotals AS (
    SELECT 
        c.ListID AS CustomerID,
        c.Name AS CustomerName,
        YEAR(il.TxnDate) AS PurchaseYear,
        SUM(il.InvoiceLineAmount) AS TotalPurchased
    FROM dbo.Customer c
    JOIN dbo.InvoiceLine il ON c.ListID = il.CustomerRefListID
    WHERE il.TxnDate IS NOT NULL
    GROUP BY c.ListID, c.Name, YEAR(il.TxnDate)
),
Change2025 AS (
    SELECT 
        yt.CustomerID,
        yt.CustomerName,
        COALESCE(y2023.TotalPurchased, 0) AS Purchased2023,
        COALESCE(y2024.TotalPurchased, 0) AS Purchased2024,
        yt.TotalPurchased AS Purchased2025,
        (yt.TotalPurchased - COALESCE(y2024.TotalPurchased, 0)) AS Change2025
    FROM YearlyTotals yt
    LEFT JOIN YearlyTotals y2024 ON yt.CustomerID = y2024.CustomerID AND y2024.PurchaseYear = 2024
    LEFT JOIN YearlyTotals y2023 ON yt.CustomerID = y2023.CustomerID AND y2023.PurchaseYear = 2023
    WHERE yt.PurchaseYear = 2025
)
SELECT
    CustomerID,
    CustomerName,
    Purchased2023,
    Purchased2024,
    Purchased2025,
    Change2025,
    CASE
        WHEN Purchased2024 > 0 THEN CAST((Change2025 / Purchased2024) * 100 AS DECIMAL(10,2))
        ELSE NULL
    END as PercentageChange
FROM Change2025
ORDER BY Change2025 DESC;
'''

# Run query and fetch results
results = pd.read_sql(query, conn)

# Calculate summary statistics first
total_2023 = results['Purchased2023'].sum()
total_2024 = results['Purchased2024'].sum()
total_2025 = results['Purchased2025'].sum()
total_change = total_2025 - total_2024
percent_change = (total_change / total_2024 * 100) if total_2024 > 0 else 0

# Create a copy of results for display
display_results = results.copy()

# Format the display values
display_results['Purchased2023'] = display_results['Purchased2023'].apply(lambda x: f'${x:,.2f}')
display_results['Purchased2024'] = display_results['Purchased2024'].apply(lambda x: f'${x:,.2f}')
display_results['Purchased2025'] = display_results['Purchased2025'].apply(lambda x: f'${x:,.2f}')
display_results['Change2025'] = display_results['Change2025'].apply(lambda x: f'${x:,.2f}')
display_results['PercentageChange'] = display_results['PercentageChange'].apply(lambda x: f'{x:.2f}%' if pd.notnull(x) else 'N/A')

# Display results in a formatted table
print('\nCustomer Purchase Change Report 2023-2025 (All Companies)')
print('Ordered by largest change from 2024 to 2025')
print('-' * 120)
print(results.to_string(index=False))
print('-' * 120)

print(f'\nSummary:')
print(f'Total Sales 2023: ${total_2023:,.2f}')
print(f'Total Sales 2024: ${total_2024:,.2f}')
print(f'Total Sales 2025: ${total_2025:,.2f}')
print(f'Total Change 2024→2025: ${total_change:,.2f}')
print(f'Percentage Change 2024→2025: {percent_change:.2f}%')

# Save to Excel for easy desktop viewing
excel_path = 'customer_change_report_2025.xlsx'
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

# Find max decline/gain for scaling
changes = [ws[f'F{row}'].value for row in range(2, ws.max_row + 1)]
max_decline = abs(min(changes)) if changes else 1
max_gain = max(changes) if changes else 1

for row in range(2, ws.max_row + 1):
    change = ws[f'F{row}'].value
    if change > 0:
        ws[f'F{row}'].fill = get_green_shade(change, max_gain)
    elif change < 0:
        ws[f'F{row}'].fill = get_red_shade(abs(change), max_decline)

wb.save(excel_path)
print('\nReport saved to customer_change_report_2025.xlsx with magnitude-based color coding.')
