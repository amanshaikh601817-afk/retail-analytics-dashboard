# =====================================================
# sql_analysis.py - the SQLite layer for the dashboard
#
# What this file does (in plain words):
#   1. Creates a small local database file "retail_sales.db".
#   2. Loads the SAME cleaned sales rows the dashboard KPIs use
#      into a table called "sales".
#   3. Prevents duplicate loading: if the table already holds
#      exactly the same number of rows, nothing is re-inserted,
#      so rerunning the app never doubles up the data.
#   4. Reads "sql_queries.sql" and runs each query on request.
# =====================================================

import re
import sqlite3
import pandas as pd


DB_FILE = "retail_sales.db"      # the local database file in this folder
SQL_FILE = "sql_queries.sql"     # the file with the SQL queries to display

# The "sales" table layout. Only valid sales rows are stored.
CREATE_TABLE_SQL = """
CREATE TABLE IF NOT EXISTS sales (
    InvoiceNo   TEXT NOT NULL,
    StockCode   TEXT,
    Description TEXT,
    Quantity    INTEGER,
    UnitPrice   REAL,
    InvoiceDate TEXT,
    CustomerID  TEXT,
    Country     TEXT,
    Revenue     REAL,
    Month       TEXT
)
"""

INSERT_ROW_SQL = """
INSERT INTO sales
    (InvoiceNo, StockCode, Description, Quantity, UnitPrice,
     InvoiceDate, CustomerID, Country, Revenue, Month)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
"""


def _money_text(value):
    """Formats a number as British pounds, e.g. 1234.5 -> '£1,234.50'."""
    return f"£{value:,.2f}"


def _customer_to_text(value):
    """CustomerID like 17850.0 becomes '17850'; missing stays empty (None)."""
    if pd.isna(value):
        return None
    return str(int(value))


def load_sales_into_db(valid_sales, db_path=DB_FILE):
    """
    Fills the 'sales' table with the cleaned rows.

    valid_sales -- the DataFrame the dashboard already built
                   (rows that are not cancelled and Quantity > 0
                   and UnitPrice > 0, with a 'revenue' and 'month' column).

    Duplicate protection: it first counts the rows already in the
    table. If that count already matches the DataFrame, it does
    nothing at all - so rerunning the app cannot add copies.
    """
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(CREATE_TABLE_SQL)

        existing_rows = conn.execute("SELECT COUNT(*) FROM sales").fetchone()[0]

        if existing_rows != len(valid_sales):
            # Table is empty or holds old/wrong data -> replace it.
            conn.execute("DELETE FROM sales")

            records = []
            for row in valid_sales.itertuples(index=False):
                records.append(
                    (
                        str(row.InvoiceNo),
                        str(row.StockCode),
                        str(row.Description),
                        int(row.Quantity),
                        float(row.UnitPrice),
                        row.InvoiceDate.strftime("%Y-%m-%d %H:%M:%S"),
                        _customer_to_text(row.CustomerID),
                        str(row.Country),
                        float(row.revenue),
                        str(row.month),
                    )
                )
            conn.executemany(INSERT_ROW_SQL, records)

        conn.commit()
    finally:
        conn.close()


def connect_database(db_path=DB_FILE):
    """Opens the database and returns the connection object."""
    return sqlite3.connect(db_path)


def load_sql_queries(path=SQL_FILE):
    """
    Reads sql_queries.sql and returns a list of (title, sql) pairs.

    How: the file uses "--- 1) Title ---" style comment banners before
    each statement, and every statement ends with a semicolon.
    Only the statements are kept; comment-only text is skipped.
    """
    with open(path, "r", encoding="utf-8") as handle:
        text = handle.read()

    queries = []
    for chunk in text.split(";"):
        comment_lines = [
            line.strip().lstrip("-").strip()
            for line in chunk.splitlines()
            if line.strip().startswith("--")
        ]
        # The preferred title is a numbered banner line like "1) Monthly revenue".
        # Fallback: the last comment that is not just "-----" filler.
        numbered = [
            c
            for c in comment_lines
            if re.match(r"^\d+\)\s*\S", c)
        ]
        title = numbered[-1] if numbered else next(
            (c for c in reversed(comment_lines) if c and not set(c) <= {"=", "-"}),
            "SQL query",
        )
        executable = [
            line
            for line in chunk.splitlines()
            if line.strip() and not line.strip().startswith("--")
        ]
        if executable:
            queries.append((title, "\n".join(executable)))
    return queries


def run_query(conn, sql):
    """Runs one SELECT query and returns the result as a DataFrame."""
    return pd.read_sql_query(sql, conn)


def run_all_queries(conn):
    """Loads the .sql file and returns the queries already executed."""
    return [(title, run_query(conn, sql)) for title, sql in load_sql_queries()]
