# =====================================================
# test_dashboard_data.py - verify the compact artifacts
#
# Confirms that:
#   1. dash_app.load_dashboard_data() can load every compact
#      file in dashboard_data/ (what the hosted app does).
#   2. Every KPI, table, data-quality count, recommendation
#      input and SQL result in those files is EXACTLY what the
#      full local pipeline (data_prep.prepare_data on the
#      genuine Online Retail.xlsx + sql_queries.sql) produces.
#   3. dash_app.server (the Gunicorn WSGI entry point) imports
#      cleanly with that data.
#
# Run AFTER build_dashboard_data.py:
#     python test_dashboard_data.py
# Exit code 0 = all checks pass.
# =====================================================

import sys

import pandas as pd

import data_prep
import dash_app

PASSED = []
FAILED = []


def check(name, actual, expected, tol=0.0):
    ok = False
    if isinstance(expected, float) or isinstance(expected, int):
        ok = actual == expected or abs(float(actual) - float(expected)) <= tol
    else:
        ok = actual == expected
    (PASSED if ok else FAILED).append(
        f"{name}: got {actual!r}, expected {expected!r}")
    marker = "PASS" if ok else "FAIL"
    print(f"  [{marker}] {name}: {actual!r}" + ("" if ok else f" (expected {expected!r})"))


def dataframes_equal(a, b, tol=0.005):
    if a.shape != b.shape:
        return False, f"shape {a.shape} != {b.shape}"
    for col in b.columns:
        if col not in a.columns:
            return False, f"missing column {col}"
        for left, right in zip(a[col], b[col]):
            if pd.isna(left) and pd.isna(right):
                continue
            if isinstance(right, (int, float)) and isinstance(left, (int, float)):
                if abs(left - right) > tol:
                    return False, f"{col}: {left} != {right}"
            elif str(left) != str(right):
                return False, f"{col}: {left!r} != {right!r}"
    return True, "ok"


def main():
    print("=" * 60)
    print("Test 1: hosted-style load from dashboard_data/ only")
    print("=" * 60)
    if dash_app.D is None:
        print(f"  [FAIL] dash_app could not load dashboard_data/: {dash_app.LOAD_ERROR}")
        print("         Run 'python build_dashboard_data.py' first.")
        return 1
    if dash_app.LOAD_ERROR is not None:
        print(f"  [FAIL] unexpected LOAD_ERROR: {dash_app.LOAD_ERROR}")
        return 1
    print("  [PASS] dash_app.D loaded, LOAD_ERROR is None")
    C = dash_app.D          # compact (precomputed) data
    W = dash_app.server     # WSGI entry point for Gunicorn
    print(f"  [PASS] dash_app.server imports cleanly: {W!r}")

    print()
    print("=" * 60)
    print("Test 2: recompute the full pipeline (Excel + SQL)")
    print("=" * 60)
    D = data_prep.prepare_data()
    print(f"  full pipeline: {D.total_rows:,} rows -> {D.kept_rows:,} valid rows")

    print()
    print("=" * 60)
    print("Test 3: compact artifacts == full pipeline results")
    print("=" * 60)

    # --- data quality counts ---
    for attr in ["total_rows", "rows_with_missing", "missing_customer_rows",
                 "missing_description_rows", "cancelled_rows",
                 "nonpositive_qty_rows", "nonpositive_price_rows",
                 "excluded_rows", "overlap_rows", "kept_rows"]:
        check(f"data quality.{attr}", getattr(C, attr), getattr(D, attr))

    # --- headline KPIs ---
    check("KPI.total_revenue", C.total_revenue, D.total_revenue, 1e-6)
    check("KPI.distinct_invoices", C.distinct_invoices, D.distinct_invoices)
    check("KPI.avg_order_value", C.avg_order_value, D.avg_order_value, 1e-6)
    check("KPI.valid_customers", C.valid_customers, D.valid_customers)
    check("KPI.identified_customers", C.identified_customers,
          D.identified_customers)
    check("KPI.verify_total", C.verify_total, D.verify_total, 1e-6)
    check("KPI.revenue_verified", C.revenue_verified, D.revenue_verified)

    # --- monthly table (slide 2) ---
    exported_cols = ["month", "revenue", "mom_pct", "mom_label",
                     "direction", "bar_color"]
    ok, why = dataframes_equal(
        C.monthly[exported_cols], D.monthly[exported_cols], 1e-6)
    check("slide2.monthly table match", ok, True) if ok else check(
        f"slide2.monthly mismatch: {why}", True, False)
    check("slide2.monthly revenue sum", float(C.monthly["revenue"].sum()),
          float(D.monthly["revenue"].sum()), 1e-6)
    check("slide2.MoM labels match", C.monthly["mom_label"].tolist(),
          D.monthly["mom_label"].tolist())
    check("slide2.bar colors match", C.monthly["bar_color"].tolist(),
          D.monthly["bar_color"].tolist())

    # --- products (slide 3) ---
    check("slide3.top product", str(C.top_products.iloc[0]["Description"]),
          str(D.top_products.iloc[0]["Description"]))
    check("slide3.top product revenue", float(C.top_products.iloc[0]["revenue"]),
          float(D.product_totals.iloc[0]["revenue"]), 1e-6)
    check("slide3.avg_product_revenue", C.avg_product_revenue,
          D.avg_product_revenue, 1e-6)
    check("slide3.n_products", C.n_products, len(D.product_totals))
    check("slide3.top3_share", C.top3_share, D.top3_share, 1e-6)
    ok, why = dataframes_equal(C.top_products, D.top_products, 1e-6)
    (PASSED if ok else FAILED).append(f"slide3.top_products: {why}")
    print(f"  [PASS] slide3.top_products frame matches" if ok
          else f"  [FAIL] slide3.top_products: {why}")

    # --- SQL results (slides 4-5) ---
    ok, why = dataframes_equal(C.sql_monthly, D.sql_monthly, 1e-6)
    (PASSED if ok else FAILED).append(f"sql.query1 monthly: {why}")
    print(f"  [PASS] sql.query1 monthly matches" if ok
          else f"  [FAIL] sql.query1 monthly: {why}")
    ok, why = dataframes_equal(C.sql_top, D.sql_top, 1e-6)
    (PASSED if ok else FAILED).append(f"sql.query2 top10: {why}")
    print(f"  [PASS] sql.query2 top10 matches" if ok
          else f"  [FAIL] sql.query2 top10: {why}")
    ok, why = dataframes_equal(C.sql_country, D.sql_country, 1e-6)
    (PASSED if ok else FAILED).append(f"sql.query3 country: {why}")
    print(f"  [PASS] sql.query3 country matches" if ok
          else f"  [FAIL] sql.query3 country: {why}")
    ok, why = dataframes_equal(C.sql_orders_count, D.sql_orders_count, 0)
    (PASSED if ok else FAILED).append(f"sql.query4 orders: {why}")
    print(f"  [PASS] sql.query4 orders matches" if ok
          else f"  [FAIL] sql.query4 orders: {why}")
    ok, why = dataframes_equal(C.sql_repeat, D.sql_repeat.head(10), 1e-6)
    (PASSED if ok else FAILED).append(f"sql.query5 repeat top10: {why}")
    print(f"  [PASS] sql.query5 repeat top10 matches" if ok
          else f"  [FAIL] sql.query5 repeat top10: {why}")
    check("sql.repeat_ids (full table)", C.repeat_ids, D.repeat_ids)
    check("sql.repeat_revenue (full table)", C.repeat_revenue,
          D.repeat_revenue, 1e-6)
    check("sql.repeat_revenue_share", C.repeat_revenue_share,
          D.repeat_revenue_share, 1e-6)
    check("sql.avg_orders_repeat", C.avg_orders_repeat, D.avg_orders_repeat, 1e-6)
    check("sql.single_purchase_share", C.single_purchase_share,
          D.single_purchase_share, 1e-6)

    # --- country / recommendation inputs (slides 4, 6) ---
    check("slide4.uk_share", C.uk_share, D.uk_share, 1e-6)
    check("slide6.uk_revenue", C.uk_revenue, D.uk_revenue, 1e-6)
    check("slide4.n_countries", len(C.sql_country), len(D.sql_country))
    check("slide6.country_leading top", str(C.country_leading.iloc[0]["Country"]),
          str(D.country_leading.iloc[0]["Country"]))
    check("slide6.best_month", str(C.best_month["month"]),
          str(D.best_month["month"]))
    check("slide6.best_month revenue", float(C.best_month["revenue"]),
          float(D.best_month["revenue"]), 1e-6)
    check("slide6.weakest_month", str(C.weakest_month["month"]),
          str(D.weakest_month["month"]))
    check("slide6.first_month", str(C.first_month["month"]),
          str(D.first_month["month"]))
    check("slide6.last_month", str(C.last_month["month"]),
          str(D.last_month["month"]))
    last_pct = C.last_pct
    last_pct_d = D.last_pct
    same_pct = (last_pct is None and last_pct_d is None and pd.isna(last_pct_d)) or \
        (last_pct is not None and last_pct_d is not None and
         not pd.isna(last_pct_d) and abs(float(last_pct) - float(last_pct_d)) < 1e-6)
    check("slide2.last_pct", same_pct, True)
    ok, why = dataframes_equal(C.country_leading, D.country_leading, 1e-6)
    (PASSED if ok else FAILED).append(f"slide4.country_leading: {why}")
    print(f"  [PASS] slide4.country_leading frame matches" if ok
          else f"  [FAIL] slide4.country_leading: {why}")

    # --- preview rows (slide 5) ---
    check("slide5.preview shape", C.preview_df.shape, D.preview_df.shape)
    check("slide5.preview first row invoice",
          str(C.preview_df.iloc[0]["InvoiceNo"]),
          str(D.preview_df.iloc[0]["InvoiceNo"]))

    # --- GBP formatting helper unchanged ---
    check("format._gbp", C._gbp(1234.5), data_prep._gbp(1234.5))

    print()
    print("=" * 60)
    if FAILED:
        print(f"RESULT: {len(PASSED)} checks passed, {len(FAILED)} FAILED")
        for line in FAILED:
            print("   " + line)
        return 1
    print(f"RESULT: all {len(PASSED)} checks passed - compact artifacts "
          "reproduce the full pipeline exactly.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
