-- ADO.NET Connection String for SQL Server (MSSQL)
"Server=192.168.68.54;Database=pcay;User Id=paradise;Password=paradise1;TrustServerCertificate=True;"

-- This query retrieves the names of all companies that were purchased in the years 2023, 2024, and 2025.
-- It assumes two tables:
-- 1. 'Companies' with columns 'id' and 'name'.
-- 2. 'Purchases' with columns 'company_id' (a foreign key to Companies.id) and 'purchase_date'.

SELECT DISTINCT
    c.name
FROM
    Companies c
JOIN
    Purchases p ON c.id = p.company_id
WHERE
    YEAR(p.purchase_date) IN (2023, 2024, 2025);
