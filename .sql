-- This query calculates the total income and the specific income for a list of 5 products for each year.
-- The total income column only reflects the sum of these 5 products.
-- It assumes the product codes are stored in the 'InvoiceLineItemRefFullName' column of the 'dbo.InvoiceLine' table.

SELECT
    YEAR(TxnDate) AS sales_year,
    SUM(InvoiceLineAmount) AS total_income,
    SUM(CASE WHEN InvoiceLineItemRefFullName = 'INH:IHP_WCT25' THEN InvoiceLineAmount ELSE 0 END) AS income_ihp_wct25,
    SUM(CASE WHEN InvoiceLineItemRefFullName = 'INH:IHP_ECT25' THEN InvoiceLineAmount ELSE 0 END) AS income_ihp_ect25,
    SUM(CASE WHEN InvoiceLineItemRefFullName = 'INH:IHP_PCT25' THEN InvoiceLineAmount ELSE 0 END) AS income_ihp_pct25,
    SUM(CASE WHEN InvoiceLineItemRefFullName = 'INH:IHP_ACT25' THEN InvoiceLineAmount ELSE 0 END) AS income_ihp_act25,
    SUM(CASE WHEN InvoiceLineItemRefFullName = 'INH:IHP_TAPS25' THEN InvoiceLineAmount ELSE 0 END) AS income_ihp_taps25
FROM
    dbo.InvoiceLine
WHERE
    InvoiceLineItemRefFullName IN ('INH:IHP_WCT25', 'INH:IHP_ECT25', 'INH:IHP_PCT25', 'INH:IHP_ACT25', 'INH:IHP_TAPS25')
GROUP BY
    YEAR(TxnDate)
ORDER BY
    sales_year;
