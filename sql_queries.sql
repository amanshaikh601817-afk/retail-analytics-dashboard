-- =====================================================
-- sql_queries.sql - SQL queries for the Retail Sales Dashboard
--
-- These queries run against a small local SQLite database
-- called "retail_sales.db", which is filled from the SAME
-- cleaned data as the dashboard KPI cards:
--   * only valid sales rows are stored
--     (not cancelled, Quantity > 0, UnitPrice > 0)
--   * Revenue column = Quantity x UnitPrice
--   * all money amounts are in British pounds (£)
--
-- NOTE: the dashboard reads this file and runs each query,
-- so the SQL shown here is exactly what the "SQL Analysis"
-- section displays on screen.
-- =====================================================


-- =====================================================
-- 1) Monthly revenue
--    Adds up the revenue of every valid sale for each month.
-- =====================================================
SELECT Month          AS Month,
       ROUND(SUM(Revenue), 2) AS Revenue_GBP
FROM sales
GROUP BY Month
ORDER BY Month;


-- =====================================================
-- 2) Top 10 products by revenue
--    Groups by product description and keeps the 10 biggest earners.
-- =====================================================
SELECT Description       AS Product,
       ROUND(SUM(Revenue), 2) AS Revenue_GBP
FROM sales
GROUP BY Description
ORDER BY Revenue_GBP DESC
LIMIT 10;


-- =====================================================
-- 3) Revenue by country
--    Total revenue for every country, best earner first.
-- =====================================================
SELECT Country             AS Country,
       ROUND(SUM(Revenue), 2) AS Revenue_GBP
FROM sales
GROUP BY Country
ORDER BY Revenue_GBP DESC;


-- =====================================================
-- 4) Order count by month
--    Counts distinct invoices (one order = one InvoiceNo) per month.
-- =====================================================
SELECT Month                       AS Month,
       COUNT(DISTINCT InvoiceNo)   AS Orders
FROM sales
GROUP BY Month
ORDER BY Month;


-- =====================================================
-- 5) Repeat customers
--    Customers who shopped more than once (more than one
--    distinct InvoiceNo). Rows with a missing CustomerID
--    are skipped because we cannot identify that customer.
-- =====================================================
SELECT CustomerID                                             AS CustomerID,
       COUNT(DISTINCT InvoiceNo)                              AS Orders,
       ROUND(SUM(Revenue), 2)                                 AS Revenue_GBP
FROM sales
WHERE CustomerID IS NOT NULL
GROUP BY CustomerID
HAVING COUNT(DISTINCT InvoiceNo) > 1
ORDER BY Orders DESC, Revenue_GBP DESC;
