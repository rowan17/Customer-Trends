-- Yearly purchase totals per customer (Top 10 only)
SELECT TOP 10
    c.ListID AS CustomerID,
    c.Name AS CustomerName,
    YEAR(il.TxnDate) AS PurchaseYear,
    SUM(il.InvoiceLineAmount) AS TotalPurchased
FROM dbo.Customer c
JOIN dbo.InvoiceLine il ON c.ListID = il.CustomerRefListID
WHERE il.TxnDate IS NOT NULL
GROUP BY c.ListID, c.Name, YEAR(il.TxnDate)
ORDER BY c.Name, PurchaseYear;


-- Bulk decline detection: Flag customers with declining yearly totals
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
YearlyWithPrev AS (
    SELECT 
        yt.CustomerID,
        yt.CustomerName,
        yt.PurchaseYear,
        yt.TotalPurchased,
        LAG(yt.TotalPurchased) OVER (PARTITION BY yt.CustomerID ORDER BY yt.PurchaseYear) AS PrevYearTotal
    FROM YearlyTotals yt
)
SELECT TOP 50
    CustomerID,
    CustomerName,
    PurchaseYear,
    TotalPurchased,
    PrevYearTotal,
    CASE WHEN PrevYearTotal IS NOT NULL AND TotalPurchased < PrevYearTotal THEN 'Decline' ELSE '' END AS DeclineFlag
FROM YearlyWithPrev
WHERE PrevYearTotal IS NOT NULL
ORDER BY CustomerName, PurchaseYear;


-- Declining customers in 2025, ordered by absolute decline
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
Decline2025 AS (
    SELECT 
        yt.CustomerID,
        yt.CustomerName,
        yt.TotalPurchased AS Purchased2025,
        prev.TotalPurchased AS Purchased2024,
        (prev.TotalPurchased - yt.TotalPurchased) AS AbsoluteDecline,
        CASE WHEN prev.TotalPurchased > 0 THEN (prev.TotalPurchased - yt.TotalPurchased) / prev.TotalPurchased ELSE NULL END AS RelativeDecline
    FROM YearlyTotals yt
    LEFT JOIN YearlyTotals prev
        ON yt.CustomerID = prev.CustomerID AND prev.PurchaseYear = 2024
    WHERE yt.PurchaseYear = 2025 AND prev.TotalPurchased IS NOT NULL AND yt.TotalPurchased < prev.TotalPurchased
)
SELECT TOP 50
    CustomerID,
    CustomerName,
    Purchased2024,
    Purchased2025,
    AbsoluteDecline,
    RelativeDecline
FROM Decline2025
ORDER BY AbsoluteDecline DESC;

