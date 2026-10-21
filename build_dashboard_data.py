# =====================================================
# build_dashboard_data.py - LOCAL pre-computation step
#
# Run this on your own machine, NOT on the hosting service:
#
#     python build_dashboard_data.py
#
# What it does:
#   1. Loads the GENUINE Online Retail.xlsx (full 541,909 rows).
#   2. Applies the SAME cleaning rules as data_prep.py (the one
#      source of truth - this script just calls prepare_data()).
#   3. Runs the 5 real SQL queries from sql_queries.sql against
#      the SQLite `sales` table (retail_sales.db).
#   4. Exports ONLY compact precomputed results to dashboard_data/:
#        kpis.json        every scalar KPI, data-quality count and
#                         recommendation input as plain JSON
#        monthly.csv          monthly revenue + MoM change labels
#        top_products.csv     top 10 products by revenue (+colour band)
#        sql_monthly.csv      SQL query 1 result
#        sql_top.csv          SQL query 2 result
#        sql_country.csv      SQL query 3 result (all countries)
#        orders_count.csv     SQL query 4 result (orders per month)
#        repeat_top.csv       SQL query 5 result - top 10 repeat
#                             customers (full-table stats -> kpis.json)
#        preview.csv          the ~100-row raw data preview slide
#
# The raw transactions and the SQLite database are NEVER exported:
# dashboard_data/ is what the hosted app reads, and every file in it
# is small (a few hundred KB total) so Render's 512MB free plan is
# never put under memory pressure.
#
# Every number is computed from the real source data. Nothing is
# hardcoded, and no fake/sample data is ever written. If the Excel
# file is missing, this script exits with instructions instead.
# =====================================================

import json
import os
import sys

import pandas as pd

import data_prep                      # ONE source of truth for cleaning
import sql_analysis                   # verifies the SQL layer stays intact

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dashboard_data")
SOURCE_FILE = data_prep.REAL_DATA_FILE


def _export_csv(df: pd.DataFrame, name: str) -> int:
    """Writes one compact CSV and returns its size in bytes."""
    path = os.path.join(DATA_DIR, name)
    df.to_csv(path, index=False, encoding="utf-8-sig")
    return os.path.getsize(path)


def _to_plain(value):
    """JSON-safe conversion of numpy/pandas scalars (rounding kept out)."""
    if value is None:
        return None
    if isinstance(value, (pd.Timestamp,)):
        return value.isoformat()
    if pd.isna(value):
        return None
    if hasattr(value, "item"):
        return value.item()
    return value


def main():
    print("=== build_dashboard_data: local full data run ===")

    if not os.path.exists(SOURCE_FILE):
        sys.exit(
            f"ERROR: '{SOURCE_FILE}' was not found in the project folder.\n"
            "\nThis script processes the genuine dataset locally. To run it:\n"
            f"  1. Download the UCI Online Retail dataset (id 352) "
            f"(https://archive.ics.uci.edu/dataset/352/online+retail)\n"
            f"  2. Put the Excel file in the project folder as '{SOURCE_FILE}'.\n"
            "  3. Run 'python build_dashboard_data.py' again."
        )

    print(f"[1/4] Loading + cleaning {SOURCE_FILE} (same rules as data_prep.py)...")
    D = data_prep.prepare_data()

    if not D.safe_data:
        sys.exit("ERROR: no valid sales rows remained after cleaning; "
                 "refusing to export empty dashboard data.")

    if not D.revenue_verified:
        sys.exit("ERROR: the independent revenue recompute does not match the "
                 f"KPI (difference £{D.revenue_difference:.4f}). Data pipeline "
                 "broken - not exporting.")

    os.makedirs(DATA_DIR, exist_ok=True)

    print("[2/4] Exporting compact precomputed results...")
    sizes = {}

    # ---- KPI + data-quality + recommendation inputs (all scalars) ----
    kpis = {
        "generated_at": pd.Timestamp.now(tz=None).isoformat(),
        "source": {
            "file": SOURCE_FILE,
            "cleaning_module": "data_prep.prepare_data()",
            "sql_file": sql_analysis.SQL_FILE,
        },
        "safe_data": bool(D.safe_data),
        # --- data quality counts ---
        "total_rows": int(D.total_rows),
        "rows_with_missing": int(D.rows_with_missing),
        "missing_customer_rows": int(D.missing_customer_rows),
        "missing_description_rows": int(D.missing_description_rows),
        "cancelled_rows": int(D.cancelled_rows),
        "nonpositive_qty_rows": int(D.nonpositive_qty_rows),
        "nonpositive_price_rows": int(D.nonpositive_price_rows),
        "excluded_rows": int(D.excluded_rows),
        "overlap_rows": int(D.overlap_rows),
        "kept_rows": int(D.kept_rows),
        # --- headline KPIs ---
        "total_revenue": float(D.total_revenue),
        "distinct_invoices": int(D.distinct_invoices),
        "avg_order_value": float(D.avg_order_value),
        "valid_customers": int(D.valid_customers),
        "identified_customers": int(D.identified_customers),
        "verify_total": float(D.verify_total),
        "revenue_difference": float(D.revenue_difference),
        "revenue_verified": bool(D.revenue_verified),
        # --- product / chart scalars ---
        "avg_product_revenue": float(D.avg_product_revenue),
        "n_products": int(len(D.product_totals)),
        # --- monthly landmarks ---
        "first_month": {"month": str(D.first_month["month"]),
                        "revenue": float(D.first_month["revenue"])},
        "last_month": {"month": str(D.last_month["month"]),
                       "revenue": float(D.last_month["revenue"])},
        "best_month": {"month": str(D.best_month["month"]),
                       "revenue": float(D.best_month["revenue"])},
        "weakest_month": {"month": str(D.weakest_month["month"]),
                          "revenue": float(D.weakest_month["revenue"])},
        "last_pct": _to_plain(D.last_pct),
        # --- country / repeat-customer inputs for recommendations ---
        "country_leading": [
            {"Country": str(r["Country"]),
             "Revenue_GBP": float(r["Revenue_GBP"]),
             "share_pct": float(r["share_pct"])}
            for _, r in D.country_leading.iterrows()
        ],
        "n_countries": int(len(D.sql_country)),
        "uk_share": float(D.uk_share),
        "uk_revenue": float(D.uk_revenue),
        "top3_share": float(D.top3_share),
        "repeat_ids": int(D.repeat_ids),
        "repeat_revenue": float(D.repeat_revenue),
        "repeat_revenue_share": float(D.repeat_revenue_share),
        "avg_orders_repeat": float(D.avg_orders_repeat),
        "single_purchase_share": float(D.single_purchase_share),
    }
    kpis_path = os.path.join(DATA_DIR, "kpis.json")
    with open(kpis_path, "w", encoding="utf-8") as handle:
        json.dump(kpis, handle, ensure_ascii=False, indent=2)
    sizes["kpis.json"] = os.path.getsize(kpis_path)

    # ---- slide 2: monthly revenue + MoM change ----
    monthly_export = D.monthly[["month", "revenue", "mom_pct",
                                "mom_label", "direction", "bar_color"]].copy()
    sizes["monthly.csv"] = _export_csv(monthly_export, "monthly.csv")

    # ---- slide 3: top 10 products (SQL top-10 == pandas top-10; both kept) ----
    sizes["top_products.csv"] = _export_csv(
        D.top_products[["Description", "revenue", "bar_color"]], "top_products.csv")

    # ---- slide 4 + 5: the 5 real SQL result tables ----
    sizes["sql_monthly.csv"] = _export_csv(D.sql_monthly, "sql_monthly.csv")
    sizes["sql_top.csv"] = _export_csv(D.sql_top, "sql_top.csv")
    sizes["sql_country.csv"] = _export_csv(D.sql_country, "sql_country.csv")
    sizes["orders_count.csv"] = _export_csv(D.sql_orders_count, "orders_count.csv")
    # Only the top 10 rows are displayed (slide 4 table); the full-table
    # stats (repeat_ids, repeat_revenue, ...) live in kpis.json.
    sizes["repeat_top.csv"] = _export_csv(
        D.sql_repeat.head(10)[["CustomerID", "Orders", "Revenue_GBP"]],
        "repeat_top.csv")

    # ---- slide 5: raw data preview (sample rows only, ~100 rows) ----
    sizes["preview.csv"] = _export_csv(D.preview_df, "preview.csv")

    total_bytes = sum(sizes.values())
    print("[3/4] Sanity checks against the computed originals...")
    # Internal consistency first (export size vs computed originals) - fatal.
    bad = []
    if len(monthly_export) != len(D.monthly):
        bad.append("monthly rows")
    if len(D.top_products) != 10:
        bad.append("top products rows")
    if len(D.sql_country) != kpis["n_countries"]:
        bad.append("countries row count")
    if bad:
        sys.exit("ERROR: export consistency checks failed (" + "; ".join(bad) +
                 "); investigate before use.")
    # Reference numbers printed for the eye (NOT gates - they would hardcode
    # findings, and the build must work for refreshed data too).
    print(f"      ok  monthly rows: {len(monthly_export)}")
    print(f"      ok  top products: {len(D.top_products)}")
    print(f"      ok  countries: {len(D.sql_country)}")
    print(f"      ok  kept rows: {kpis['kept_rows']:,}")
    print(f"      ok  total revenue: £{kpis['total_revenue']:,.2f}")

    print(f"[4/4] Wrote dashboard_data/ ({total_bytes / 1024:.0f} KB total):")
    for name, size in sorted(sizes.items()):
        print(f"      {name:<18} {size / 1024:8.1f} KB")

    print("\nDone. Every slide number is now precomputed from "
          f"{kpis['total_rows']:,} real rows. Commit dashboard_data/ "
          "so the hosted app can load it.")
    print("Verify with:  python test_dashboard_data.py")


if __name__ == "__main__":
    main()
