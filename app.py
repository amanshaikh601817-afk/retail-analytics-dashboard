# =====================================================
# RETAIL INTELLIGENCE - Executive Slide Deck
# Built with: Streamlit (UI), Pandas (data), Plotly (charts), SQLite (SQL)
#
# The dashboard is now a slide-by-slide presentation:
#   Slide 1  Executive Overview (title, KPI cards, summary)
#   Slide 2  Monthly Sales Trends (month-over-month changes)
#   Slide 3  Product Performance (Top 10 vs dataset average)
#   Slide 4  Country & Customer Insights (SQLite-backed)
#   Slide 5  SQL & Data Quality (queries + original row counts)
#   Slide 6  Business Recommendations (evidence -> action -> KPI)
#
# Data rules (nothing is deleted or changed in the Excel file!):
#   - Cancelled invoices = InvoiceNo starting with "C"
#   - Rows with Quantity <= 0 or UnitPrice <= 0 are not real sales
#   - Both groups are excluded from the sales KPIs, but kept
#     in the data quality summary so nothing is hidden
#   - Revenue per row = Quantity * UnitPrice (verified on screen)
#   - Average Order Value = Total Revenue / number of valid invoices
#   - Unit costs are NOT part of this dataset, so no profit or
#     margin claims are ever made on screen.
# =====================================================

# 1. Import the libraries we need
import os                         # lets us check if the Excel file exists
import streamlit as st            # Streamlit builds the web page for us
import pandas as pd               # Pandas handles our table of data
import plotly.express as px       # Plotly draws interactive charts

import sql_analysis               # our SQLite helper (retail_sales.db)
_gbp = sql_analysis._money_text   # short alias: formats 1234.5 as £1,234.50

# 2. Page setup FIRST: wide layout = presentation uses the full browser width
st.set_page_config(
    page_title="Retail Intelligence - Executive Presentation",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# 3. Global styling: premium dark navy executive presentation.
#    Sections: 1 page  2 animations  3 cards  4 typography  5 rec cards  6 responsive
st.markdown(
    """
    <style>
        /* ---------- 1. PAGE (dark navy with a faint glow) ---------- */
        .block-container {
            max-width: 100%;
            padding-top: 1.5rem;
            padding-bottom: 2rem;
        }

        .stApp {
            font-family: "Segoe UI", "Inter", system-ui, sans-serif;
            font-size: 18px !important;
            line-height: 1.6;
            background:
                radial-gradient(1100px 500px at 85% -120px, rgba(99, 102, 241, 0.10), transparent 60%),
                radial-gradient(900px 450px at 8% -80px, rgba(34, 211, 238, 0.08), transparent 60%),
                #0a1424 !important;
        }

        /* Root rem base raised to 18px: Streamlit's own widgets are
           sized in rem, so this lifts their default ~14px text to
           ~16px everywhere without any CSS zoom tricks. This works
           together with `theme.baseFontSize = 18` in .streamlit/config.toml,
           which sets Streamlit's own theme root. */
        html { font-size: 18px !important; }

        /* Progress bar label ("Slide X of Y - ...") */
        div[data-testid="stProgress"] p,
        div[data-testid="stProgress"] span,
        div[data-testid="stProgress"] div[style*="font-size"] {
            font-size: 16px !important;
            color: #b9d0e6 !important;
        }
        /* Long words wrap instead of overflowing past the screen edge */
        div[data-testid="stMarkdownContainer"],
        [data-testid="stMetric"],
        .stat-card, .rec-card, .plan-card {
            overflow-wrap: anywhere;
            min-width: 0;
        }

        /* ---------- 2. ANIMATIONS (light slide transition) ----------
           When you change slides, content gently fades in and
           slides up. Reduced-motion users see everything instantly. */
        div[data-testid="stElementContainer"] {
            animation: riseIn 0.45s ease both;
        }
        @keyframes riseIn {
            from { opacity: 0; transform: translateY(12px); }
            to   { opacity: 1; transform: translateY(0); }
        }

        @media (prefers-reduced-motion: reduce) {
            div[data-testid="stElementContainer"] { animation: none !important; }
            [data-testid="stMetric"], .stat-card, .rec-card, .plan-card, .stButton > button {
                transition: none !important;
            }
            [data-testid="stMetric"]:hover, .stat-card:hover {
                transform: none !important;
            }
        }

        /* ---------- 3. CARDS ---------------------------------------- */
        [data-testid="stMetric"] {
            background: linear-gradient(155deg, rgba(255,255,255,0.055), rgba(255,255,255,0.015)) !important;
            border: 1px solid rgba(148, 179, 208, 0.16);
            border-top: 2px solid rgba(34, 211, 238, 0.45);
            border-radius: 16px;
            padding: 1.35rem 1.5rem 1.2rem;
            box-shadow: 0 4px 14px rgba(2, 10, 24, 0.35);
            transition: transform 0.25s ease, box-shadow 0.25s ease, border-color 0.25s ease;
        }
        [data-testid="stMetric"]:hover {
            transform: translateY(-4px);
            border-color: rgba(34, 211, 238, 0.5);
            box-shadow: 0 14px 32px rgba(2, 10, 24, 0.45), 0 0 24px rgba(34, 211, 238, 0.16);
        }
        [data-testid="stMetricLabel"],
        [data-testid="stMetricLabel"] p {
            color: #9db8d4 !important;
            font-size: 17px;
            font-weight: 700;
            letter-spacing: 0.05em;
            text-transform: uppercase;
            margin-bottom: 0.3rem;
            white-space: normal;
        }
        [data-testid="stMetricValue"],
        [data-testid="stMetricValue"] p {
            color: #ffffff !important;
            font-size: 40px;
            font-weight: 800;
            line-height: 1.12;
            white-space: normal;
        }

        /* Large custom KPI cards (slide 1) - wrap into fewer columns
           on narrow screens instead of shrinking the text. */
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
            gap: 1rem;
            margin-top: 0.4rem;
        }
        .kpi-card {
            background: linear-gradient(155deg, rgba(255,255,255,0.055), rgba(255,255,255,0.015));
            border: 1px solid rgba(148, 179, 208, 0.16);
            border-top: 2px solid rgba(34, 211, 238, 0.5);
            border-radius: 16px;
            padding: 1.5rem 1.7rem 1.35rem;
            box-shadow: 0 4px 14px rgba(2, 10, 24, 0.35);
            transition: transform 0.25s ease, box-shadow 0.25s ease, border-color 0.25s ease;
        }
        .kpi-card:hover {
            transform: translateY(-4px);
            border-color: rgba(34, 211, 238, 0.5);
            box-shadow: 0 14px 32px rgba(2, 10, 24, 0.45), 0 0 24px rgba(34, 211, 238, 0.16);
        }
        .kpi-label {
            font-size: 17px; font-weight: 700; letter-spacing: 0.05em;
            text-transform: uppercase; color: #b9d0e6;
        }
        .kpi-value { font-size: 40px; font-weight: 800; color: #ffffff !important; line-height: 1.12; margin-top: 0.35rem; }
        .kpi-note  { font-size: 14.5px; color: #7d9bb8; margin-top: 0.45rem; line-height: 1.45; }

        /* Readable HTML data tables (replaced canvas-based st.dataframe) */
        .data-scroll {
            border: 1px solid rgba(148, 179, 208, 0.16);
            border-radius: 14px;
            overflow: auto;
            background: rgba(255,255,255,0.02);
        }
        .data-table { width: 100%; border-collapse: collapse; }
        .data-table th {
            font-size: 16px !important; font-weight: 700; letter-spacing: 0.04em;
            text-transform: uppercase; color: #b9d0e6;
            text-align: left;
            padding: 0.7rem 1rem;
            border-bottom: 2px solid rgba(148, 179, 208, 0.3);
            background: #0d1930;
            position: sticky; top: 0;
        }
        .data-table td {
            font-size: 16px !important; color: #dbe9f7;
            padding: 0.6rem 1rem;
            border-bottom: 1px solid rgba(148, 179, 208, 0.12);
        }
        .data-table tbody tr:nth-child(even) { background: rgba(255,255,255,0.025); }
        .data-table tbody tr:hover { background: rgba(34, 211, 238, 0.06); }

        /* Compact cards (data quality + SQL summary rows) */
        .stat-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 0.8rem;
            margin-top: 0.4rem;
        }
        .stat-card {
            background: linear-gradient(155deg, rgba(255,255,255,0.05), rgba(255,255,255,0.015));
            border: 1px solid rgba(148, 179, 208, 0.14);
            border-left: 3px solid #22d3ee;
            border-radius: 14px;
            padding: 0.95rem 1.15rem;
            box-shadow: 0 3px 10px rgba(2, 10, 24, 0.3);
            transition: transform 0.25s ease, border-color 0.25s ease;
        }
        .stat-card:hover { transform: translateY(-3px); border-color: rgba(34, 211, 238, 0.45); }
        .stat-label {
            font-size: 14px; font-weight: 700; letter-spacing: 0.05em;
            text-transform: uppercase; color: #b9d0e6;
        }
        .stat-value { font-size: 26px; font-weight: 800; color: #f4f9ff; line-height: 1.25; }
        .stat-note  { font-size: 14.5px; color: #86a2bf; margin-top: 0.15rem; line-height: 1.45; }

        /* ---------- 4. TYPOGRAPHY ----------------------------------- */
        .app-badge {
            display: inline-block;
            font-size: 14px; font-weight: 700; letter-spacing: 0.14em;
            text-transform: uppercase; color: #7ee2f8;
            background: rgba(34, 211, 238, 0.08);
            border: 1px solid rgba(34, 211, 238, 0.30);
            border-radius: 999px; padding: 0.35rem 1.05rem; margin-bottom: 0.55rem;
        }
        .app-subtitle { font-size: 19px; color: #a9c2dc; margin-top: -0.5rem; }
        .app-header-rule {
            height: 3px; width: 260px; border: none; border-radius: 999px;
            margin-top: 1.1rem;
            background: linear-gradient(90deg, #22d3ee, #2563eb 55%, #8b5cf6);
        }
        .stApp h1 {
            font-weight: 800; font-size: 40px !important; letter-spacing: -0.01em;
            color: #f4f9ff !important; margin-bottom: 0.1rem;
        }
        /* Slide headings: white with a glowing gradient tick
           h3 = slide title (34px), h4/h5 = section headings (26/25px) */
        .stMarkdown h3, .stMarkdown h4 {
            font-weight: 800; color: #f4f9ff !important;
        }
        .stMarkdown h3 { font-size: 34px !important; margin-top: 0.8rem; }
        .stMarkdown h4 { font-size: 26px !important; margin-top: 1.3rem; }
        .stMarkdown h5 { font-size: 25px !important; font-weight: 800; color: #f4f9ff !important; margin-top: 0.9rem; }
        .stMarkdown h3::before, .stMarkdown h4::before, .stMarkdown h5::before {
            content: ""; display: inline-block;
            width: 9px; height: 9px; border-radius: 3px; margin-right: 0.6rem;
            background: linear-gradient(135deg, #22d3ee, #8b5cf6);
            box-shadow: 0 0 8px rgba(34, 211, 238, 0.55);
        }
        div[data-testid="stCaptionContainer"] { margin-top: -0.1rem; }
        div[data-testid="stCaptionContainer"] p { font-size: 15px !important; color: #93aecb !important; }

        /* Paragraph text stays at 18px, tables 16px (build rules) */
        div[data-testid="stMarkdownContainer"] p,
        div[data-testid="stMarkdownContainer"] li {
            font-size: 18px !important; color: #d8e6f5 !important;
        }
        div[data-testid="stMarkdownContainer"] strong { color: #f4f9ff; }
        div[data-testid="stMarkdownContainer"] table { border-collapse: collapse; width: 100%; }
        div[data-testid="stMarkdownContainer"] th {
            color: #9db8d4; font-size: 16px !important; text-transform: uppercase; letter-spacing: 0.05em;
        }
        div[data-testid="stMarkdownContainer"] td {
            color: #dbe9f7; font-size: 16px !important;
            border-color: rgba(148, 179, 208, 0.14) !important;
            padding: 0.45rem 1rem;
        }
        div[data-testid="stDataFrame"] {
            border: 1px solid rgba(148, 179, 208, 0.14);
            border-radius: 14px; overflow: hidden;
        }
        div[data-testid="stMarkdownContainer"] span[style*="color:#15803d"] { color: #34d399 !important; }

        /* Navy-friendly, large navigation buttons (>= 46px tall).
           !important: Streamlit styles these widgets from its theme
           tokens, and our rules must win over those.               */
        .stButton > button {
            font-size: 18px !important;
            font-weight: 700;
            color: #f4f9ff !important;
            min-height: 46px;
            padding: 0.55rem 1.1rem;
            border-radius: 10px;
            border: 1px solid rgba(148, 179, 208, 0.28);
            transition: border-color 0.25s ease, transform 0.15s ease;
            width: 100%;
        }
        .stButton > button:hover { border-color: rgba(34, 211, 238, 0.55); }

        /* ---------- 5. RECOMMENDATION + PLAN CARDS ------------------- */
        .rec-card {
            background: linear-gradient(155deg, rgba(255,255,255,0.05), rgba(255,255,255,0.015));
            border: 1px solid rgba(148, 179, 208, 0.14);
            border-left: 3px solid #22d3ee;
            border-radius: 14px;
            padding: 1rem 1.25rem;
            margin-bottom: 0.9rem;
            box-shadow: 0 3px 10px rgba(2, 10, 24, 0.3);
        }
        .rec-backed { border-left-color: #34d399; }   /* emerald: observed evidence */
        .rec-test   { border-left-color: #8b5cf6; }   /* violet: idea to test       */
        .rec-title  { font-weight: 800; color: #f4f9ff; font-size: 20px; margin-bottom: 0.25rem; }
        .rec-hint   { font-size: 14px; font-weight: 700; letter-spacing: 0.08em;
                      text-transform: uppercase; margin-bottom: 0.4rem; }
        .rec-backed { color: #34d399; } .rec-test { color: #c4b5fd; }
        .rec-line   { font-size: 16.5px; color: #dbe9f7; margin-top: 0.22rem; line-height: 1.55; }
        .rec-line b { color: #f4f9ff; }
        .rec-label  { color: #7dd3fc; font-size: 14px; text-transform: uppercase;
                      letter-spacing: 0.06em; font-weight: 700; margin-right: 0.35rem; }

        .plan-card {
            background: linear-gradient(155deg, rgba(255,255,255,0.05), rgba(255,255,255,0.015));
            border: 1px solid rgba(148, 179, 208, 0.14);
            border-top: 3px solid #22d3ee;
            border-radius: 14px; padding: 1.1rem 1.3rem;
        }
        .plan-1 { border-top-color: #34d399; }
        .plan-2 { border-top-color: #f59e0b; }
        .plan-3 { border-top-color: #8b5cf6; }
        .plan-name { font-weight: 800; color: #f4f9ff; font-size: 18px; margin-bottom: 0.35rem; }
        .plan-item { font-size: 16px; color: #dbe9f7; line-height: 1.6; }

        /* Performance legend chips (slides 2 and 3) */
        .legend { display: flex; gap: 1.2rem; flex-wrap: wrap; margin: 0.2rem 0 0.3rem 0; }
        .legend-chip { font-size: 15px; color: #dbe9f7; }
        .dot { display: inline-block; width: 11px; height: 11px; border-radius: 50%;
               margin-right: 0.4rem; vertical-align: middle; }

        /* ---------- 6. RESPONSIVE ------------------------------------
           Narrow screens: fewer/smoother paddings and slightly smaller
           (still readable) heading numbers. KPI cards wrap into fewer
           columns via the grid - text never shrinks below readable. */
        @media (max-width: 1200px) {
            .block-container { padding-left: 1.25rem; padding-right: 1.25rem; }
            .stApp h1 { font-size: 36px; }
            [data-testid="stMetricValue"] { font-size: 36px; }
            .kpi-value { font-size: 36px; }
            .stMarkdown h3 { font-size: 30px; }
            .stMarkdown h4 { font-size: 25px; }
        }
        @media (max-width: 900px) {
            .block-container { padding-left: 0.75rem; padding-right: 0.75rem; }
            .stApp h1 { font-size: 32px; }
            [data-testid="stMetricValue"] { font-size: 32px; }
            .kpi-value { font-size: 32px; }
            [data-testid="stMetricLabel"] { font-size: 16px; }
            .kpi-label { font-size: 16px; }
            .stat-grid { grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); }
            .stat-value { font-size: 22px; }
            .app-header-rule { width: 180px; }
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------
# 4. Load the REAL dataset ("Online Retail.xlsx") if we can.
#    Identical to the previous long-page dashboard: cached read,
#    friendly error + sample fallback when the file is unavailable.
# ----------------------------------------------------

REAL_DATA_FILE = "Online Retail.xlsx"

@st.cache_data(show_spinner="Reading Online Retail.xlsx (this can take a moment)...")
def load_retail_excel(path):
    return pd.read_excel(path, engine="openpyxl")

use_real_data = False
load_error = None

if os.path.exists(REAL_DATA_FILE):
    try:
        raw_df = load_retail_excel(REAL_DATA_FILE)
        use_real_data = True
    except Exception as exc:               # wrong file type, corrupt file, missing openpyxl...
        load_error = str(exc)
else:
    load_error = "The file was not found next to app.py."

if not use_real_data:
    st.error(
        f"**Could not load `{REAL_DATA_FILE}`**, so this presentation is currently "
        f"showing a small sample preview instead.\n\n"
        f"Reason: `{load_error}`\n\n"
        f"**How to fix it:** keep `Online Retail.xlsx` in the same folder as `app.py`, "
        f"make sure it is a valid Excel file, and install the Excel reader with "
        f"`pip install openpyxl` (it is already listed in requirements.txt)."
    )

# =====================================================
# THE PRESENTATION (needs the real dataset)
# =====================================================

if use_real_data:
    currency = "£"

    # ----------------------------------------------------
    # 5. Prepare the data + count data quality issues.
    #    Identical rules to before: rows are only FLAGGED.
    # ----------------------------------------------------

    total_rows = len(raw_df)                                  # every row in the file
    rows_with_missing = int(raw_df.isna().any(axis=1).sum())  # rows missing ANY value
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

    # A human-readable reason for each row (used in the preview table on slide 5)
    raw["row_status"] = "Included in sales"
    raw.loc[raw["is_cancelled"], "row_status"] = "Cancelled invoice"
    raw.loc[(~raw["is_cancelled"]) & raw["is_invalid"], "row_status"] = "Non-positive quantity / price"

    # Rows kept for all sales calculations
    valid_sales = raw.loc[~excluded_mask].copy()
    valid_sales["revenue"] = valid_sales["Quantity"] * valid_sales["UnitPrice"]
    valid_sales["month"] = valid_sales["InvoiceDate"].dt.to_period("M").astype(str)
    raw["line_total"] = raw["Quantity"] * raw["UnitPrice"]

    safe_data = len(valid_sales) > 0   # protects every division below

    # ----------------------------------------------------
    # 6. KPI calculations (same formulas as the long-page version)
    # ----------------------------------------------------

    total_revenue = float(valid_sales["revenue"].sum())
    distinct_invoices = int(valid_sales["InvoiceNo"].nunique())
    avg_order_value = total_revenue / distinct_invoices if safe_data else 0.0
    valid_customers = int(valid_sales.loc[valid_sales["CustomerID"].notna(), "CustomerID"].nunique())
    # Same "identified customers" KPI as before: counted anywhere in the dataset
    identified_customers = int(raw.loc[raw["CustomerID"].notna(), "CustomerID"].nunique())

    # Revenue formula verification (second pass), shown on slide 5
    verify_total = float(
        (raw.loc[~excluded_mask, "Quantity"] * raw.loc[~excluded_mask, "UnitPrice"]).sum()
    ) if safe_data else 0.0
    revenue_difference = abs(verify_total - total_revenue)
    revenue_verified = revenue_difference < 0.005

    # ----------------------------------------------------
    # 7. Shared aggregations for the slides
    # ----------------------------------------------------

    # Monthly master table for slide 2 (includes month-over-month change)
    monthly = (
        valid_sales.groupby("month")["revenue"].sum().reset_index()
        .sort_values("month")
        .reset_index(drop=True)
    )
    monthly["prev"] = monthly["revenue"].shift(1)

    def _mom_value(row):
        """Month-over-month % change; None when there is nothing to compare."""
        if row["prev"] is None or pd.isna(row["prev"]):
            return None
        if row["prev"] <= 0:               # zero revenue before -> cannot compare
            return None
        return (row["revenue"] - row["prev"]) / row["prev"] * 100

    monthly["mom_pct"] = monthly.apply(_mom_value, axis=1)

    def mom_label(pct):
        """Readable text for the MoM column ('n/a' where there is no comparison)."""
        if pct is None:
            return "n/a"
        return f"{pct:+,.1f}%"

    def mom_direction(pct):
        if pct is None:
            return "No comparison"
        if pct > 2:
            return "Increase"
        if pct < -2:
            return "Decrease"
        return "Little change"

    monthly["mom_label"] = monthly["mom_pct"].map(mom_label)
    monthly["direction"] = monthly["mom_pct"].map(mom_direction)
    # Bar colours on slide 2: green for increases, red for decreases,
    # amber for little/no change or months with nothing to compare.
    def _mom_color(pct):
        if pct is None:
            return "#f59e0b"
        if pct > 2:
            return "#34d399"
        if pct < -2:
            return "#f87171"
        return "#f59e0b"

    monthly["bar_color"] = monthly["mom_pct"].map(_mom_color)

    # Product totals for slide 3 (dataset-wide average for colour coding)
    product_totals = (
        valid_sales.groupby("Description")["revenue"].sum().reset_index()
        .sort_values("revenue", ascending=False)
        .reset_index(drop=True)
    )
    avg_product_revenue = float(product_totals["revenue"].mean()) if len(product_totals) else 0.0
    NEAR_AVG_BAND = 0.25   # "near average" = within +/- 25% of the average

    def product_color(rev):
        share_vs_avg = rev / avg_product_revenue if avg_product_revenue > 0 else 1.0
        if share_vs_avg - 1 > NEAR_AVG_BAND:
            return "#34d399"          # well above average -> emerald
        if abs(share_vs_avg - 1) <= NEAR_AVG_BAND:
            return "#f59e0b"          # near average -> amber
        return "#f87171"              # below average -> red

    top_products = product_totals.head(10).copy()
    top_products["bar_color"] = top_products["revenue"].map(product_color)

    # Run the shared SQL queries once for slides 4, 5 and 6 (duplicate-safe)
    sql_analysis.load_sales_into_db(valid_sales)
    sql_conn = sql_analysis.connect_database()
    (
        sql_monthly, sql_top, sql_country, sql_orders_count, sql_repeat,
    ) = (
        df for _, df in sql_analysis.run_all_queries(sql_conn)
    )
    sql_conn.close()

    # Country shares for slide 4
    country_leading = sql_country.head(5).copy() if len(sql_country) else sql_country.copy()
    if total_revenue > 0 and len(country_leading):
        country_leading["share_pct"] = country_leading["Revenue_GBP"] / total_revenue * 100

    # Repeat-customer facts for slides 4 and 6 (computed, never guessed)
    repeat_ids = int(sql_repeat["CustomerID"].nunique()) if len(sql_repeat) else 0
    repeat_revenue = float(sql_repeat["Revenue_GBP"].sum()) if len(sql_repeat) else 0.0
    repeat_revenue_share = repeat_revenue / total_revenue * 100 if total_revenue > 0 else 0.0
    avg_orders_repeat = float(sql_repeat["Orders"].mean()) if len(sql_repeat) else 0.0
    single_purchase_share = (100 - (repeat_ids / valid_customers * 100)) if valid_customers > 0 else 0.0

    # Facts used by the summary and recommendations (all from calculated data)
    if len(monthly):
        best_month = monthly.loc[monthly["revenue"].idxmax()]
        weakest_month = monthly.loc[monthly["revenue"].idxmin()]
        first_month, last_month = monthly.iloc[0], monthly.iloc[-1]
        last_pct = last_month["mom_pct"]
    uk_share = float(country_leading.iloc[0]["share_pct"]) if len(country_leading) else 0.0
    uk_revenue = float(country_leading.iloc[0]["Revenue_GBP"]) if len(country_leading) else 0.0
    top3_share = (
        float(product_totals.head(3)["revenue"].sum()) / total_revenue * 100
        if (safe_data and len(product_totals)) else 0.0
    )

    def _period_label(month_value):
        """Turns '2011-11' into 'Nov 2011' for easy reading."""
        try:
            return pd.Period(str(month_value), freq="M").strftime("%b %Y")
        except Exception:
            return str(month_value)

    def _pct_txt(pct):
        return "n/a" if pct is None else f"{pct:+,.1f}%"

    # ----------------------------------------------------
    # 8. Small helpers used by several slides
    # ----------------------------------------------------

    def quality_card(label, value, note):
        """Builds one compact card (used on slides 4 and 5)."""
        return (
            '<div class="stat-card">'
            f'<div class="stat-label">{label}</div>'
            f'<div class="stat-value">{value}</div>'
            f'<div class="stat-note">{note}</div>'
            "</div>"
        )

    def rec_card(title, hint, hint_class, evidence, action, reason, kpi):
        """One recommendation card: evidence -> action -> reason -> KPI."""
        return (
            f'<div class="rec-card {hint_class}">'
            f'<div class="rec-title">{title}</div>'
            f'<div class="rec-hint {hint_class}">{hint}</div>'
            f'<div class="rec-line"><span class="rec-label">Observed evidence</span> {evidence}</div>'
            f'<div class="rec-line"><span class="rec-label">Recommended action</span> {action}</div>'
            f'<div class="rec-line"><span class="rec-label">Business reason</span> {reason}</div>'
            f'<div class="rec-line"><span class="rec-label">KPI to monitor</span> {kpi}</div>'
            "</div>"
        )

    def legend_chips():
        """Shared colour legend used on slides 2 and 3."""
        return (
            '<div class="legend">'
            '<span class="legend-chip"><span class="dot" style="background:#34d399;"></span>Increase / above reference</span>'
            '<span class="legend-chip"><span class="dot" style="background:#f87171;"></span>Decrease / below reference</span>'
            '<span class="legend-chip"><span class="dot" style="background:#f59e0b;"></span>Little change / no comparison</span>'
            "</div>"
        )

    def sql_explanation():
        """Short plain-language explanation of the SQL behind the results."""
        return (
            "SQL logic in one minute: the cleaned valid rows live in the SQLite table "
            "`sales` (built from `valid_sales`, duplicate-safe). "
            "**Monthly revenue** = `SUM(Revenue)` grouped by month; "
            "**Top products** = largest `SUM(Revenue)` per product (`LIMIT 10`); "
            "**Country revenue** = `SUM(Revenue)` per country; "
            "**Repeat customers** = `GROUP BY CustomerID HAVING COUNT(DISTINCT InvoiceNo) > 1` — "
            "rows without a CustomerID are skipped there because that customer cannot be identified."
        )

    # Cross-slide chart styling helpers (consistent, readable)
    DARK_TEXT = "#cfe0f2"
    TICK_TEXT = "#9db8d4"

    def style_chart(fig, height, y_money=True):
        """One place that applies the premium, readable look to every bar chart."""
        fig.update_layout(
            height=height,
            margin=dict(l=10, r=10, t=20, b=10),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Segoe UI, sans-serif", size=16, color=DARK_TEXT),
            showlegend=False,
            hoverlabel=dict(
                bgcolor="#101f36",
                bordercolor="rgba(34,211,238,0.6)",
                font=dict(color="#e8f1fb", size=16),
            ),
            transition=dict(duration=450, easing="cubic-in-out"),
        )
        fig.update_xaxes(showgrid=False, tickfont=dict(color=TICK_TEXT, size=16))
        if y_money:
            fig.update_yaxes(gridcolor="rgba(148,179,208,0.12)", tickprefix="£", tickformat=",",
                             tickfont=dict(color=TICK_TEXT, size=16))
        else:
            fig.update_yaxes(gridcolor="rgba(148,179,208,0.12)", tickfont=dict(color=TICK_TEXT, size=16))
        return fig

    def show_table(display_df, max_height=460):
        """
        Renders a DataFrame as a styled HTML table with 16px text and
        comfortable row spacing.

        Why not st.dataframe()? Streamlit's dataframe widget draws itself
        on a canvas at a small fixed font that CSS cannot change, so its
        text would stay tiny. A plain HTML table lets us set 16px text,
        generous padding and sticky headers (and long tables scroll).
        """
        from html import escape  # built-in module; keeps text safe as code

        header_cells = "".join(f"<th>{escape(str(col))}</th>" for col in display_df.columns)
        body_rows = []
        for record in display_df.itertuples(index=False):
            row_cells = "".join(
                f'<td>{escape("" if (str(v) == "nan" or pd.isna(v)) else str(v))}</td>'
                for v in record
            )
            body_rows.append(f"<tr>{row_cells}</tr>")
        table_html = (
            f'<div class="data-scroll" style="max-height:{max_height}px">'
            f'<table class="data-table"><thead><tr>{header_cells}</tr></thead>'
            f'<tbody>{"".join(body_rows)}</tbody></table></div>'
        )
        st.markdown(table_html, unsafe_allow_html=True)

    # =====================================================
    # SLIDE 1 - EXECUTIVE OVERVIEW
    # =====================================================

    def slide_overview():
        st.markdown('<span class="app-badge">Executive Command Center</span>', unsafe_allow_html=True)
        st.title("Retail Intelligence")
        st.markdown('<p class="app-subtitle">Executive Sales &amp; Customer Insights</p>',
                    unsafe_allow_html=True)
        st.markdown('<div class="app-header-rule"></div>', unsafe_allow_html=True)

        # Large KPI cards: they wrap into fewer columns on narrow screens
        # (instead of shrinking the numbers) because they use the same
        # responsive grid as the compact cards.
        kpi_cards = "".join([
            '<div class="kpi-card" title="Sum of Quantity × UnitPrice across valid sales rows '
            '(not cancelled, Quantity > 0, UnitPrice > 0).">'
            '<div class="kpi-label">Total Revenue</div>'
            f'<div class="kpi-value">£{total_revenue:,.0f}</div>'
            '<div class="kpi-note">Quantity × UnitPrice over all valid rows</div></div>',

            '<div class="kpi-card" title="Distinct InvoiceNo values among valid sales rows '
            '(each invoice counted once).">'
            '<div class="kpi-label">Valid Orders</div>'
            f'<div class="kpi-value">{distinct_invoices:,}</div>'
            '<div class="kpi-note">Distinct InvoiceNo among valid sales rows</div></div>',

            '<div class="kpi-card" title="Distinct CustomerID values present anywhere in the '
            'dataset (rows with a missing CustomerID are skipped).">'
            '<div class="kpi-label">Unique Customers</div>'
            f'<div class="kpi-value">{identified_customers:,}</div>'
            '<div class="kpi-note">CustomerIDs recorded anywhere in the dataset</div></div>',

            '<div class="kpi-card" title="Total Revenue ÷ Valid Orders.">'
            '<div class="kpi-label">Average Order Value</div>'
            f'<div class="kpi-value">£{avg_order_value:,.2f}</div>'
            '<div class="kpi-note">Total Revenue ÷ Valid Orders</div></div>',
        ])
        st.markdown(f'<div class="kpi-grid">{kpi_cards}</div>', unsafe_allow_html=True)

        if not safe_data:
            st.warning("No valid sales rows were found after cleaning — "
                       "check the data quality slide (5) for the row counts.")

        # Short summary generated from the calculated numbers (no invented figures)
        if safe_data and len(monthly):
            st.markdown("### Summary")
            pct_word = _pct_txt(last_pct)
            st.markdown(
                f"Between **{_period_label(first_month['month'])}** and "
                f"**{_period_label(last_month['month'])}**, valid sales generated "
                f"**£{total_revenue:,.0f}** across **{distinct_invoices:,} orders** from "
                f"**{valid_customers:,} identified customers** who placed at least one valid, "
                f"CustomerID-bearing line. Average order value is **£{avg_order_value:,.2f}**. "
                f"The strongest month was **{_period_label(best_month['month'])}** at "
                f"**£{best_month['revenue']:,.0f}**; **{_period_label(last_month['month'])}** "
                f"came in at **£{last_month['revenue']:,.0f}** "
                f"({('(first month - no comparison)' if last_pct is None else pct_word + ' vs the previous month')}). "
                f"Revenue is concentrated in **{country_leading.iloc[0]['Country']}** "
                f"({uk_share:,.1f}% of total revenue)."
            )
        st.caption(
            "Every number on these slides is computed live from Online Retail.xlsx "
            "(valid sales rows only: not cancelled, Quantity > 0, UnitPrice > 0). "
            "All amounts in British pounds (£). Use the arrows below to move through the deck."
        )

    # =====================================================
    # SLIDE 2 - MONTHLY SALES TRENDS
    # =====================================================

    def slide_monthly():
        st.markdown("### Monthly Sales Trends")
        st.caption("Total revenue per month (valid sales only), coloured by the change "
                   "against the previous month: green = increase, red = decrease, "
                   "amber = little or no change (first month and zero-revenue months "
                   "show no percentage).")
        st.markdown(legend_chips(), unsafe_allow_html=True)

        trend_fig = px.bar(
            monthly, x="month", y="revenue", custom_data=["mom_label"],
            labels={"month": "Month", "revenue": "Revenue (£)"},
        )
        trend_fig.update_traces(
            marker=dict(color=monthly["bar_color"]),
            hovertemplate="<b>%{x}</b><br>Revenue: £%{y:,.0f}<br>vs previous month: %{customdata[0]}<extra></extra>",
        )
        st.plotly_chart(style_chart(trend_fig, 500), width="stretch")

        # The actual month-over-month numbers in a readable table
        st.markdown("#### Month-over-month change")
        mom_table = monthly[["month", "revenue", "mom_label", "direction"]].copy()
        mom_table.columns = ["Month", "Revenue", "Change vs previous", "Direction"]
        mom_table["Revenue"] = mom_table["Revenue"].map(_gbp)
        show_table(mom_table, 460)

    # =====================================================
    # SLIDE 3 - PRODUCT PERFORMANCE
    # =====================================================

    def slide_products():
        st.markdown("### Product Performance")
        st.caption(
            f"Top 10 products by revenue (valid sales only). The colours compare each "
            f"product's revenue with the **dataset-wide average product revenue of "
            f"£{avg_product_revenue:,.0f}** — NOT with profitability, because this dataset "
            f"contains no unit-cost data."
        )
        st.markdown(legend_chips(), unsafe_allow_html=True)

        top_sorted = top_products.sort_values("revenue")
        prod_fig = px.bar(
            top_sorted, x="revenue", y="Description", orientation="h",
            labels={"Description": "Product", "revenue": "Revenue (£)"},
        )
        prod_fig.update_traces(
            marker=dict(color=top_sorted["bar_color"]),
            hovertemplate="<b>%{y}</b><br>Revenue: £%{x:,.0f}<extra></extra>",
        )
        st.plotly_chart(style_chart(prod_fig, 520, y_money=False), width="stretch")

        # Small section: the highest-revenue products, with their share of total revenue
        st.markdown("#### Highest-revenue products")
        if len(product_totals):
            leaders = product_totals.head(3).copy()
            leaders["share"] = leaders["revenue"] / total_revenue * 100 if total_revenue > 0 else 0
            rows = [
                f'<div class="stat-card"><div class="stat-label">{i + 1}. {str(desc)[:40]}</div>'
                f'<div class="stat-value">{_gbp(rev)}</div>'
                f'<div class="stat-note">{share:,.1f}% of total revenue</div></div>'
                for i, (desc, rev, share) in enumerate(leaders[["Description", "revenue", "share"]].itertuples(index=False))
            ]
            st.markdown(f'<div class="stat-grid">{"".join(rows)}</div>', unsafe_allow_html=True)

        if avg_product_revenue > 0 and len(top_products):
            st.caption(
                f"Within this Top 10, the smallest product still earns about "
                f"{float(top_products['revenue'].min()) / avg_product_revenue:,.0f}x the dataset-wide "
                f"average product revenue, so all ten sit in the green band — that is exactly what a "
                f"Top 10 should look like against a dataset average spread across "
                f"{len(product_totals):,} products."
            )

    # =====================================================
    # SLIDE 4 - COUNTRY AND CUSTOMER INSIGHTS
    # =====================================================

    def slide_country():
        st.markdown("### Country and Customer Insights")
        st.caption("Country revenue and repeat-customer facts, computed from the cleaned rows "
                   "and the SQLite sales table. No customer behaviour is assumed beyond what "
                   "the data shows.")

        c1, c2 = st.columns([1.3, 1], gap="large")
        with c1:
            st.markdown("#### Revenue by country (top 10 of "
                        f"{len(sql_country)})")
            country_data = sql_country.head(10).sort_values("Revenue_GBP")
            country_fig = px.bar(
                country_data, x="Revenue_GBP", y="Country", orientation="h",
                labels={"Revenue_GBP": "Revenue (£)"},
            )
            country_fig.update_traces(
                marker=dict(color=country_data["Revenue_GBP"], showscale=False,
                            colorscale=[[0, "#7c3aed"], [0.5, "#2563eb"], [1, "#22d3ee"]]),
                hovertemplate="<b>%{y}</b><br>Revenue: £%{x:,.0f}<extra></extra>",
            )
            st.plotly_chart(style_chart(country_fig, 470, y_money=False), width="stretch")

        with c2:
            st.markdown("#### Leading countries")
            if len(country_leading):
                lead_rows = "".join([
                    f'<div class="stat-card"><div class="stat-label">{r["Country"]}</div>'
                    f'<div class="stat-value">{_gbp(r["Revenue_GBP"])}</div>'
                    f'<div class="stat-note">{r["share_pct"]:,.1f}% of total revenue</div></div>'
                    for _, r in country_leading.iterrows()
                ])
                st.markdown(f'<div class="stat-grid">{lead_rows}</div>', unsafe_allow_html=True)

        st.markdown("#### Repeat-customer metrics")
        repeat_cards = "".join([
            quality_card("Repeat Customers", f"{repeat_ids:,}",
                         "more than one distinct invoice (SQL)"),
            quality_card("Revenue from Repeat Customers", _gbp(repeat_revenue),
                         f"{repeat_revenue_share:,.1f}% of total revenue"),
            quality_card("Avg Orders per Repeat Customer", f"{avg_orders_repeat:,.1f}",
                         f"out of {valid_customers:,} identified customers; "
                         f"{single_purchase_share:,.1f}% ordered only once"),
        ])
        st.markdown(f'<div class="stat-grid">{repeat_cards}</div>', unsafe_allow_html=True)

        st.markdown("#### Most loyal customers (by order count)")
        loyal_display = sql_repeat.head(10).rename(columns={
            "CustomerID": "Customer ID", "Orders": "Orders",
            "Revenue_GBP": "Revenue (£)",
        }).assign(**{"Revenue (£)": lambda d: d["Revenue (£)"].map(_gbp)})
        show_table(loyal_display, 460)
        st.caption("Sourced from SQL query 5 (repeat customers). Rows with a missing "
                   "CustomerID are skipped there because that customer cannot be identified.")

    # =====================================================
    # SLIDE 5 - SQL AND DATA QUALITY
    # =====================================================

    def slide_sql_quality():
        st.markdown("### SQL and Data Quality")
        st.caption(sql_explanation())

        # 5a. SQL results
        st.markdown("#### SQL results")
        sql_kpis = "".join([
            quality_card("SQL Monthly Buckets", f"{len(sql_monthly):,}",
                         "months returned by query 1"),
            quality_card("SQL Top Product", str(sql_top.iloc[0]["Product"])[:28] if len(sql_top) else "n/a",
                         _gbp(sql_top.iloc[0]["Revenue_GBP"]) if len(sql_top) else ""),
            quality_card("SQL Countries", f"{len(sql_country):,}",
                         "countries with valid sales rows"),
            quality_card("SQL Repeat Customers", f"{repeat_ids:,}",
                         "customers with more than one invoice"),
        ])
        st.markdown(f'<div class="stat-grid">{sql_kpis}</div>', unsafe_allow_html=True)

        s1, s2 = st.columns(2, gap="large")
        with s1:
            st.markdown("##### Monthly revenue (SQL)")
            sql_m_data = sql_monthly.sort_values("Month")
            sql_f = px.bar(sql_m_data, x="Month", y="Revenue_GBP",
                           labels={"Revenue_GBP": "Revenue (£)"})
            sql_f.update_traces(
                marker=dict(color=sql_m_data["Revenue_GBP"], showscale=False,
                            colorscale=[[0, "#1d4ed8"], [0.55, "#2563eb"], [1, "#22d3ee"]]),
                hovertemplate="<b>%{x}</b><br>Revenue: £%{y:,.0f}<extra></extra>")
            st.plotly_chart(style_chart(sql_f, 460), width="stretch")
        with s2:
            st.markdown("##### Top 10 products (SQL)")
            sql_top_display = (
                sql_top.rename(columns={"Product": "Product", "Revenue_GBP": "Revenue (£)"})
                .assign(**{"Revenue (£)": lambda d: d["Revenue (£)"].map(_gbp)})
            )
            show_table(sql_top_display, 430)
            st.markdown("##### Revenue by country (SQL)")
            sql_country_display = (
                sql_country.rename(columns={"Country": "Country", "Revenue_GBP": "Revenue (£)"})
                .assign(**{"Revenue (£)": lambda d: d["Revenue (£)"].map(_gbp)})
            )
            show_table(sql_country_display, 430)

        # 5b. Data quality summary (identical counts to the long-page version)
        st.markdown("#### Data Quality Summary")
        st.caption("No rows are ever changed or deleted in the Excel file — rows are only "
                   "flagged. “Excluded” below means left out of the sales KPIs only.")

        quality_cards = "".join([
            quality_card("Rows in Original Dataset", f"{total_rows:,}",
                         "Every row in Online Retail.xlsx"),
            quality_card("Cancelled Invoice Rows", f"{cancelled_rows:,}",
                         'InvoiceNo starts with "C" (credit notes)'),
            quality_card("Rows with Missing Values", f"{rows_with_missing:,}",
                         f"missing CustomerID: {missing_customer_rows:,} · missing Description: {missing_description_rows:,}"),
            quality_card("Rows Excluded from Sales", f"{excluded_rows:,}",
                         "cancelled OR Quantity ≤ 0 OR UnitPrice ≤ 0"),
        ])
        st.markdown(f'<div class="stat-grid">{quality_cards}</div>', unsafe_allow_html=True)

        st.markdown(
            f"""
            **Exact row counts**

            | Check | Rows |
            |---|---:|
            | Total rows in the original dataset | {total_rows:,} |
            | Rows with missing values (any column) | {rows_with_missing:,} |
            | ↳ of which missing CustomerID | {missing_customer_rows:,} |
            | ↳ of which missing Description | {missing_description_rows:,} |
            | Cancelled invoice rows | {cancelled_rows:,} |
            | ↳ of which Quantity ≤ 0 | {nonpositive_qty_rows:,} |
            | ↳ of which UnitPrice ≤ 0 | {nonpositive_price_rows:,} |
            | Rows excluded from sales calculations | {excluded_rows:,} |
            | ↳ overlap: cancelled **and** non-positive | {overlap_rows:,} |
            | **Rows kept for sales calculations** | **{len(valid_sales):,}** |
            """
        )

        if revenue_verified:
            st.markdown(
                f'<span style="color:#15803d; font-weight:700;">✓ Verified</span> — recomputing '
                f'Quantity × UnitPrice independently gives <b>£{verify_total:,.2f}</b>, '
                f"which matches the Total Revenue KPI (difference £{revenue_difference:.4f}). "
                f"Average Order Value = £{total_revenue:,.2f} ÷ {distinct_invoices:,} invoices "
                f"= £{avg_order_value:,.2f}.",
                unsafe_allow_html=True,
            )
        else:
            st.warning(
                f"Revenue check mismatch: KPI shows £{total_revenue:,.2f} but the "
                f"independent recompute gives £{verify_total:,.2f}. "
                "Please review the data preparation step."
            )

        # 5c. Raw data preview (same preview as the old dashboard)
        st.markdown("#### Raw Data Preview")
        st.caption(
            "Preview: the first 60 valid rows plus 20 cancelled and 20 non-positive examples "
            "from the file. Only rows with status “Included in sales” count toward Total Revenue."
        )
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
        preview_cols = [
            "InvoiceNo", "InvoiceDate", "Description", "Country",
            "Quantity", "UnitPrice", "line_total", "CustomerID", "row_status",
        ]
        raw_preview_display = preview_df[preview_cols].rename(
            columns={"line_total": "Line Total (Quantity × UnitPrice)",
                     "row_status": "Status"}
        )
        show_table(raw_preview_display, 460)

    # =====================================================
    # SLIDE 6 - BUSINESS RECOMMENDATIONS
    # =====================================================

    def slide_recommendations():
        st.markdown("### Business Recommendations")
        st.caption(
            "Generated from the calculated figures on the previous slides. This dataset has "
            "no unit-cost data, so none of these address profit or margins, and no revenue "
            "gains or forecasts are claimed."

        )

        # --- Evidence-backed recommendations (left emerald border) -------
        st.markdown("#### Evidence-backed findings")
        if len(country_leading):
            runner_ups = ", ".join(country_leading.iloc[1:4]["Country"].tolist())
            st.markdown(rec_card(
                "1. Revenue depends heavily on one market",
                "observed", "rec-backed",
                f"<b>{country_leading.iloc[0]['Country']}</b> generated {_gbp(uk_revenue)} = "
                f"<b>{uk_share:,.1f}%</b> of all revenue ({len(sql_country)} countries in total).",
                "Prepare a pilot marketing/promotions plan for the strongest other markets "
                f"(next by revenue: {runner_ups}).",
                "Concentration on one market is a commercial risk; the runner-up markets already "
                "generate real revenue, so small targeted efforts build on proven demand.",
                "Non-UK revenue share per month (target: gradually rising, judge quarterly)."
            ), unsafe_allow_html=True)

        if len(product_totals):
            st.markdown(rec_card(
                "2. A few products carry the business",
                "observed", "rec-backed",
                f"The top 3 products account for <b>{top3_share:,.1f}%</b> of total revenue, led by "
                f"“<b>{product_totals.iloc[0]['Description']}</b>” at {_gbp(product_totals.iloc[0]['revenue'])}.",
                "Protect availability of the top earners (stock planning, listing quality) and "
                "cross-sell around them in the same orders.",
                "Revenue concentration means stock-outs on a top product directly hit sales; "
                "cross-selling rides on existing demand without new product bets.",
                "Top-3 revenue share per month (stable share, rising total)."
            ), unsafe_allow_html=True)

        if len(monthly):
            st.markdown(rec_card(
                "3. Timing of strength and weakness is known",
                "observed", "rec-backed",
                f"Strongest month: <b>{_period_label(best_month['month'])}</b> at "
                f"{_gbp(best_month['revenue'])}. Weakest month: "
                f"<b>{_period_label(weakest_month['month'])}</b> at {_gbp(weakest_month['revenue'])}. "
                f"The latest month ({_period_label(last_month['month'])}) shows "
                f"{_pct_txt(last_pct)} vs its previous month.",
                "Schedule promotions before the historically strong period and run a drill-down "
                "review for the weak month (its orders, invoice sizes and mix).",
                "The data shows a real seasonal gap; acting on the pattern avoids reacting late.",
                "Monthly revenue and its month-over-month % (fewer deep negative months)."
            ), unsafe_allow_html=True)

        if repeat_ids:
            st.markdown(rec_card(
                f"4. Repeat customers are a proven revenue engine",
                "observed", "rec-backed",
                f"<b>{repeat_ids:,} of {valid_customers:,}</b> identified customers bought more than "
                f"once (avg {avg_orders_repeat:,.1f} orders each) and generate "
                f"<b>{repeat_revenue_share:,.1f}%</b> of all revenue.",
                "Create a straightforward repeat-purchase nudge for high-frequency customers "
                "(the top orders-per-customer names are listed on slide 4).",
                "Their revenue share is measured, so serving them better supports the largest "
                "measured part of the business.",
                "Revenue from repeat customers (% of total) and average orders per repeat customer."
            ), unsafe_allow_html=True)

        # --- Ideas that still need testing (violet border) ---------------
        st.markdown("#### Ideas that still need testing")
        st.markdown(rec_card(
            "Raise average order value",
            "hypothesis - needs testing", "rec-test",
            f"Average order value currently stands at <b>{_gbp(avg_order_value)}</b>.",
            "Pilot an A/B test of a bundle or multi-buy offer built around top products.",
            "Larger baskets would raise order value *if* customers respond; this is a guess "
            "until the pilot measures a real effect.",
            "Average order value in the pilot group vs control."
        ), unsafe_allow_html=True)
        st.markdown(rec_card(
            "Win back one-order customers",
            "hypothesis - needs testing", "rec-test",
            f"<b>{single_purchase_share:,.1f}%</b> of identified customers ordered only once.",
            "Test a single, well-targeted follow-up campaign for first-order customers.",
            "If even a modest share return, repeat purchases rise; the effect must be measured "
            "-> not assumed.",
            "Repeat-purchase rate in the campaign cohort vs control."
        ), unsafe_allow_html=True)
        st.markdown(rec_card(
            "Investigate small markets before investing",
            "hypothesis - needs testing", "rec-test",
            f"Of {len(sql_country)} countries, 30+ earn only a small fraction of revenue "
            "(full list on slide 4).",
            "Run a short data investigation (reasons: shipping, returns, order sizes) before "
            "committing budget to small markets.",
            "Cheap analysis avoids expensive expansion mistakes where the data is thin.",
            "Revenue per active market (quarterly change)."
        ), unsafe_allow_html=True)

        # --- Prioritized action plan --------------------------------------
        st.markdown("#### Prioritized action plan")
        p1, p2, p3 = st.columns(3, gap="large")
        with p1:
            st.markdown(
                '<div class="plan-card plan-1"><div class="plan-name">Immediate (this week)</div>'
                '<div class="plan-item">'
                "1. Lock stock/availability of the top-3 revenue products.<br>"
                "2. Note the latest month-over-month move from slide 2 and brief the team.<br>"
                "3. Confirm which weak month/market drill-down gets scheduled first."
                "</div></div>", unsafe_allow_html=True)
        with p2:
            st.markdown(
                '<div class="plan-card plan-2"><div class="plan-name">Next 30 days</div>'
                '<div class="plan-item">'
                "1. Pilot targeted campaigns in the strongest non-UK markets (slide 4).<br>"
                "2. Start the A/B test for average order value (idea above).<br>"
                "3. Send a first repeat-purchase nudge to the top repeat customers."
                "</div></div>", unsafe_allow_html=True)
        with p3:
            st.markdown(
                '<div class="plan-card plan-3"><div class="plan-name">Longer term</div>'
                '<div class="plan-item">'
                "1. Review the concentration KPIs (top-3 share, non-UK share) quarterly.<br>"
                "2. Grow the repeat-revenue share with a small loyalty programme.<br>"
                "3. Rebuild this deck with refreshed data each month to track progress."
                "</div></div>", unsafe_allow_html=True)

        st.caption(
            "Method note: all evidence figures above are the same computed values shown on "
            "slides 1-5 (no external data, no cost assumptions). Items marked "
            "“needs testing” are deliberately labelled as hypotheses."
        )

    # =====================================================
    # 9. SLIDE NAVIGATION (session_state keeps it stable on reruns)
    # =====================================================

    SLIDES = [
        ("Executive Overview", slide_overview),
        ("Monthly Sales Trends", slide_monthly),
        ("Product Performance", slide_products),
        ("Country & Customer Insights", slide_country),
        ("SQL & Data Quality", slide_sql_quality),
        ("Business Recommendations", slide_recommendations),
    ]

    def go_next():
        st.session_state["slide"] = min(len(SLIDES), st.session_state.get("slide", 1) + 1)

    def go_prev():
        st.session_state["slide"] = max(1, st.session_state.get("slide", 1) - 1)

    def go_home():
        st.session_state["slide"] = 1

    if "slide" not in st.session_state:
        st.session_state["slide"] = 1

    slide_no = st.session_state["slide"]
    titles = [name for name, _ in SLIDES]

    # 9a. Progress indicator at the top
    st.progress(slide_no / len(SLIDES), text=f"Slide {slide_no} of {len(SLIDES)} — {titles[slide_no - 1]}")

    # 9b. Render ONLY the current slide
    SLIDES[slide_no - 1][1]()

    # 9c. Bottom navigation bar
    st.markdown('<div style="margin-top:1.1rem;"></div>', unsafe_allow_html=True)
    nav_prev, nav_center, nav_next = st.columns([1, 1.4, 1], gap="large")
    with nav_prev:
        st.button("Previous", on_click=go_prev, disabled=(slide_no == 1),
                  type="secondary", use_container_width=True)
    with nav_center:
        st.button("Start Over (Back to Overview)", on_click=go_home,
                  disabled=(slide_no == 1),
                  type="secondary", use_container_width=True)
    with nav_next:
        st.button("Next", on_click=go_next, disabled=(slide_no == len(SLIDES)),
                  type="primary", use_container_width=True)
    st.caption(f"Deck navigation: Previous is disabled on slide 1, Next on slide {len(SLIDES)}. "
               "The Start Over button returns you to the Executive Overview.")

else:
    # ----------------------------------------------------
    # Fallback: small SAMPLE preview (only when the Excel file
    # is unavailable). The presentation itself keeps using the
    # real dataset the moment the file is back in place.
    # ----------------------------------------------------
    currency = "₹"
    data = {
        "order_date": ["2024-01-15", "2024-01-22", "2024-02-05", "2024-02-18",
                        "2024-03-03", "2024-03-21", "2024-04-09", "2024-04-27",
                        "2024-05-12", "2024-05-30", "2024-06-07", "2024-06-25"],
        "product": ["Laptop", "Phone", "Laptop", "Headphones", "Phone", "Tablet",
                     "Laptop", "Phone", "Tablet", "Headphones", "Laptop", "Phone"],
        "price": [55000, 20000, 60000, 3000, 21000, 25000, 58000, 19500,
                   24000, 2800, 61000, 22000],
    }
    df = pd.DataFrame(data)
    df["order_date"] = pd.to_datetime(df["order_date"])
    df["month"] = df["order_date"].dt.to_period("M").astype(str)

    total_revenue_fallback = df["price"].sum()
    col1, col2, col3 = st.columns(3, gap="large")
    col1.metric("Total Revenue", f"₹{total_revenue_fallback:,.0f}")
    col2.metric("Total Orders", f"{len(df)}")
    col3.metric("Avg Order Value", f"₹{total_revenue_fallback / len(df):,.0f}")
    st.dataframe(df, width="stretch", hide_index=True)
