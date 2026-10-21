# =====================================================
# data_prep.py - shared data layer for the dashboards
#
# This module holds the EXACT data-cleaning rules and metric
# definitions that the Streamlit app (app.py) uses, so the new
# Plotly Dash app can reuse them without touching app.py:
#
#   - Read Online Retail.xlsx (read-only)
#   - Cancelled invoices = InvoiceNo starting with "C"
#   - Rows with Quantity <= 0 or UnitPrice <= 0 are not real sales
#   - Excluded rows are only FLAGGED, never deleted or edited
#   - Revenue per row = Quantity * UnitPrice (GBP)
#   - Average Order Value = Total Revenue / Valid Orders
#   - SQL results come from sql_queries.sql via retail_sales.db
#
# prepare_data() computes everything ONCE and returns one object
# with all summary tables and facts the slides need.
# =====================================================

import os
from functools import lru_cache
from types import SimpleNamespace

import pandas as pd

import sql_analysis                       # SQLite helper (retail_sales.db)

REAL_DATA_FILE = "Online Retail.xlsx"

# When the local Excel file is missing (e.g. in a cloud deployment where
# the file was never committed to the repository), the loader below pulls
# the GENUINE dataset from the official UCI repository with ucimlrepo.
UCI_DATASET_ID = 352  # "Online Retail" (id 352) on archive.ics.uci.edu

_gbp = sql_analysis._money_text           # 1234.5 -> '£1,234.50'

# Performance/contrast colours shared by chart builders
ACCENT_CYAN = "#22d3ee"
ACCENT_BLUE = "#2563eb"
ACCENT_VIOLET = "#8b5cf6"
POSITIVE = "#34d399"                      # emerald - increases
NEGATIVE = "#f87171"                      # red     - decreases
NEUTRAL = "#f59e0b"                       # amber   - little/no change
CHART_TEXT = "#cfe0f2"
TICK_TEXT = "#9db8d4"


def _period_label(month_value):
    """Turns '2011-11' into 'Nov 2011' for easy reading."""
    try:
        return pd.Period(str(month_value), freq="M").strftime("%b %Y")
    except Exception:
        return str(month_value)


def _mom_pct(value, prev):
    """Month-over-month % change; None when there is nothing to compare."""
    if prev is None or pd.isna(prev):
        return None
    if prev <= 0:                          # zero revenue before -> not comparable
        return None
    return (value - prev) / prev * 100


def _no_comparison(pct):
    """True when a MoM value cannot be compared (None went through pandas
    and came back as NaN)."""
    return pct is None or pd.isna(pct)


def _mom_label(pct):
    if _no_comparison(pct):
        return "n/a"
    return f"{pct:+,.1f}%"


def _mom_direction(pct):
    if _no_comparison(pct):
        return "No comparison"
    if pct > 2:
        return "Increase"
    if pct < -2:
        return "Decrease"
    return "Little change"


def _mom_color(pct):
    # green up, red down, amber flat or not comparable
    if _no_comparison(pct):
        return NEUTRAL
    if pct > 2:
        return POSITIVE
    if pct < -2:
        return NEGATIVE
    return NEUTRAL


def _product_color(rev, avg_product_revenue, band=0.25):
    """Slide 3 colour code vs the dataset-wide average product revenue."""
    if avg_product_revenue <= 0:
        return NEUTRAL
    share = rev / avg_product_revenue
    if share - 1 > band:
        return POSITIVE
    if abs(share - 1) <= band:
        return NEUTRAL
    return NEGATIVE


def load_raw_dataframe(path=REAL_DATA_FILE):
    """
    Loads the raw transactions in two steps, so the app also works on a
    hosting service where the 23MB Excel file was never committed:

      1. If the local Excel file exists -> read it (primary source).
      2. Otherwise fetch the genuine "Online Retail" dataset from the
         official UCI repository via the `ucimlrepo` package (id 352),
         and shape it into exactly the same columns as the Excel file.

    Raises RuntimeError with a clear message when neither source is
    available -> dash_app.py shows the error instead of fake data.
    """
    if os.path.exists(path):
        return pd.read_excel(path, engine="openpyxl")

    try:
        from ucimlrepo import fetch_ucirepo
    except ImportError:
        raise RuntimeError(
            f"Local file '{path}' was not found and the ucimlrepo package is "
            "not installed. Install dependencies with "
            "'pip install -r requirements.txt' so the app can retrieve the "
            "dataset from the UCI repository."
        )

    try:
        online_retail = fetch_ucirepo(id=UCI_DATASET_ID)
        df = online_retail.data.original if online_retail.data.original is not None \
            else online_retail.data.features
        df = df.copy()
    except Exception as exc:
        raise RuntimeError(
            f"Local file '{path}' was not found and the UCI Online Retail "
            f"dataset (id {UCI_DATASET_ID}) could not be retrieved. Check the "
            f"internet connection and try again. Original error: {exc}"
        )

    # Shape the UCI data into exactly the columns the cleaning step expects.
    required = ["InvoiceNo", "StockCode", "Description", "Quantity",
                "InvoiceDate", "UnitPrice", "CustomerID", "Country"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise RuntimeError(
            f"The UCI dataset (id {UCI_DATASET_ID}) is missing the columns "
            f"{missing}; it cannot be used as the data source."
        )
    df["InvoiceNo"] = df["InvoiceNo"].astype(str).str.strip()
    df["InvoiceDate"] = pd.to_datetime(df["InvoiceDate"])
    df["Quantity"] = pd.to_numeric(df["Quantity"])
    df["UnitPrice"] = pd.to_numeric(df["UnitPrice"])
    df["StockCode"] = df["StockCode"].astype(str)
    df["Description"] = df["Description"].astype(object)
    df["CustomerID"] = pd.to_numeric(df["CustomerID"])
    df["Country"] = df["Country"].astype(str)
    return df[required]


@lru_cache(maxsize=1)
def prepare_data(path=REAL_DATA_FILE):
    """Computes every table/number the slides show. Cached per process."""

    # ---------------- 1. Load and clean (same rules as app.py) --------
    raw_df = load_raw_dataframe(path)

    total_rows = len(raw_df)
    rows_with_missing = int(raw_df.isna().any(axis=1).sum())
    missing_customer_rows = int(raw_df["CustomerID"].isna().sum())
    missing_description_rows = int(raw_df["Description"].isna().sum())

    raw = raw_df.copy()

    # Rule 1: cancelled orders = InvoiceNo starting with "C"
    raw["InvoiceNo"] = raw["InvoiceNo"].astype(str).str.strip()
    raw["is_cancelled"] = raw["InvoiceNo"].str.upper().str.startswith("C")

    # Missing descriptions are labelled "Unknown" so charts never break
    raw["Description"] = raw["Description"].fillna("Unknown")

    # Rule 2: not a real sale if Quantity <= 0 or UnitPrice <= 0
    raw["is_invalid"] = (raw["Quantity"] <= 0) | (raw["UnitPrice"] <= 0)

    excluded_mask = raw["is_cancelled"] | raw["is_invalid"]

    cancelled_rows = int(raw["is_cancelled"].sum())
    nonpositive_qty_rows = int((raw["Quantity"] <= 0).sum())
    nonpositive_price_rows = int((raw["UnitPrice"] <= 0).sum())
    excluded_rows = int(excluded_mask.sum())
    overlap_rows = int((raw["is_cancelled"] & raw["is_invalid"]).sum())

    raw["row_status"] = "Included in sales"
    raw.loc[raw["is_cancelled"], "row_status"] = "Cancelled invoice"
    raw.loc[(~raw["is_cancelled"]) & raw["is_invalid"], "row_status"] = "Non-positive quantity / price"

    # Rows kept for all sales calculations
    valid_sales = raw.loc[~excluded_mask].copy()
    valid_sales["revenue"] = valid_sales["Quantity"] * valid_sales["UnitPrice"]
    valid_sales["month"] = valid_sales["InvoiceDate"].dt.to_period("M").astype(str)
    raw["line_total"] = raw["Quantity"] * raw["UnitPrice"]

    safe_data = len(valid_sales) > 0

    # ---------------- 2. KPIs (same definitions) ----------------------
    total_revenue = float(valid_sales["revenue"].sum())
    distinct_invoices = int(valid_sales["InvoiceNo"].nunique())
    avg_order_value = total_revenue / distinct_invoices if safe_data else 0.0
    valid_customers = int(valid_sales.loc[valid_sales["CustomerID"].notna(), "CustomerID"].nunique())
    identified_customers = int(raw.loc[raw["CustomerID"].notna(), "CustomerID"].nunique())

    # Revenue formula verification (independent second pass)
    verify_total = float(
        (raw.loc[~excluded_mask, "Quantity"] * raw.loc[~excluded_mask, "UnitPrice"]).sum()
    ) if safe_data else 0.0
    revenue_difference = abs(verify_total - total_revenue)
    revenue_verified = revenue_difference < 0.005

    # ---------------- 3. Monthly table with MoM change ----------------
    monthly = (
        valid_sales.groupby("month")["revenue"].sum().reset_index()
        .sort_values("month")
        .reset_index(drop=True)
    )
    monthly["prev"] = monthly["revenue"].shift(1)
    monthly["mom_pct"] = [
        _mom_pct(v, p) for v, p in zip(monthly["revenue"], monthly["prev"])
    ]
    monthly["mom_label"] = monthly["mom_pct"].map(_mom_label)
    monthly["direction"] = monthly["mom_pct"].map(_mom_direction)
    monthly["bar_color"] = monthly["mom_pct"].map(_mom_color)

    # ---------------- 4. Product totals + slide-3 colours -------------
    product_totals = (
        valid_sales.groupby("Description")["revenue"].sum().reset_index()
        .sort_values("revenue", ascending=False)
        .reset_index(drop=True)
    )
    avg_product_revenue = float(product_totals["revenue"].mean()) if len(product_totals) else 0.0
    top_products = product_totals.head(10).copy()
    top_products["bar_color"] = top_products["revenue"].map(
        lambda r: _product_color(r, avg_product_revenue)
    )

    # ---------------- 5. Shared SQL results (duplicate-safe load) -----
    sql_analysis.load_sales_into_db(valid_sales)
    sql_conn = sql_analysis.connect_database()
    sql_monthly, sql_top, sql_country, sql_orders_count, sql_repeat = (
        df for _, df in sql_analysis.run_all_queries(sql_conn)
    )
    sql_conn.close()

    country_leading = sql_country.head(5).copy() if len(sql_country) else sql_country.copy()
    if total_revenue > 0 and len(country_leading):
        country_leading["share_pct"] = country_leading["Revenue_GBP"] / total_revenue * 100.

    repeat_ids = int(sql_repeat["CustomerID"].nunique()) if len(sql_repeat) else 0
    repeat_revenue = float(sql_repeat["Revenue_GBP"].sum()) if len(sql_repeat) else 0.0
    repeat_revenue_share = repeat_revenue / total_revenue * 100 if total_revenue > 0 else 0.0
    avg_orders_repeat = float(sql_repeat["Orders"].mean()) if len(sql_repeat) else 0.0
    single_purchase_share = (100 - (repeat_ids / valid_customers * 100)) if valid_customers > 0 else 0.0

    if len(monthly):
        best_month = monthly.loc[monthly["revenue"].idxmax()]
        weakest_month = monthly.loc[monthly["revenue"].idxmin()]
        first_month, last_month = monthly.iloc[0], monthly.iloc[-1]
        last_pct = last_month["mom_pct"]
    else:
        best_month = weakest_month = first_month = last_month = None
        last_pct = None

    uk_share = float(country_leading.iloc[0]["share_pct"]) if len(country_leading) else 0.0
    uk_revenue = float(country_leading.iloc[0]["Revenue_GBP"]) if len(country_leading) else 0.0
    top3_share = (
        float(product_totals.head(3)["revenue"].sum()) / total_revenue * 100
        if (safe_data and len(product_totals)) else 0.0
    )

    def _pct_txt(pct):
        return "n/a" if _no_comparison(pct) else f"{pct:+,.1f}%"

    # ---------------- 6. Raw data preview (same as Streamlit) ---------
    preview_df = pd.concat(
        [
            raw.loc[~excluded_mask].head(60),
            raw.loc[raw["is_cancelled"]].head(20),
            raw.loc[(~raw["is_cancelled"]) & raw["is_invalid"]].head(20),
        ]
    ).copy()
    preview_df["CustomerID"] = preview_df["CustomerID"].apply(
        lambda v: "" if pd.isna(v) else str(int(v))
    )
    preview_df = preview_df[
        ["InvoiceNo", "InvoiceDate", "Description", "Country",
         "Quantity", "UnitPrice", "line_total", "CustomerID", "row_status"]
    ].rename(columns={"line_total": "Line Total (Quantity × UnitPrice)",
                      "row_status": "Status"})

    return SimpleNamespace(
        safe_data=safe_data,
        total_rows=total_rows,
        rows_with_missing=rows_with_missing,
        missing_customer_rows=missing_customer_rows,
        missing_description_rows=missing_description_rows,
        cancelled_rows=cancelled_rows,
        nonpositive_qty_rows=nonpositive_qty_rows,
        nonpositive_price_rows=nonpositive_price_rows,
        excluded_rows=excluded_rows,
        overlap_rows=overlap_rows,
        kept_rows=len(valid_sales),
        total_revenue=total_revenue,
        distinct_invoices=distinct_invoices,
        avg_order_value=avg_order_value,
        valid_customers=valid_customers,
        identified_customers=identified_customers,
        verify_total=verify_total,
        revenue_difference=revenue_difference,
        revenue_verified=revenue_verified,
        monthly=monthly,
        product_totals=product_totals,
        avg_product_revenue=avg_product_revenue,
        top_products=top_products,
        sql_monthly=sql_monthly,
        sql_top=sql_top,
        sql_country=sql_country,
        sql_orders_count=sql_orders_count,
        sql_repeat=sql_repeat,
        country_leading=country_leading,
        repeat_ids=repeat_ids,
        repeat_revenue=repeat_revenue,
        repeat_revenue_share=repeat_revenue_share,
        avg_orders_repeat=avg_orders_repeat,
        single_purchase_share=single_purchase_share,
        best_month=best_month,
        weakest_month=weakest_month,
        first_month=first_month,
        last_month=last_month,
        last_pct=last_pct,
        uk_share=uk_share,
        uk_revenue=uk_revenue,
        top3_share=top3_share,
        preview_df=preview_df,
        _pct_txt=_pct_txt,
        _period_label=_period_label,
    )

