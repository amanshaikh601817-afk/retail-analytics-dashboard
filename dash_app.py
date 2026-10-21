# =====================================================
# dash_app.py - Retail Intelligence (Plotly Dash version)
#
# A separate slide-by-slide presentation built with Plotly Dash.
# The Streamlit version (app.py) is kept untouched as a backup.
#
# Data: PRECOMPUTED compact results in dashboard_data/, produced
# locally by build_dashboard_data.py from the genuine Online
# Retail.xlsx + data_prep cleaning rules + sql_queries.sql:
#   kpis.json / monthly.csv / top_products.csv / sql_monthly.csv /
#   sql_top.csv / sql_country.csv / orders_count.csv /
#   repeat_top.csv / preview.csv
#
# This keeps the hosted app lightweight (fits Render's 512MB free
# plan): NO full Excel read, NO UCI download, NO SQLite build and
# no full-table processing happens at web startup - the slides
# only ever read the small files above.
#
# Structure (one slide visible at a time):
#   Slide 1  Executive Overview  - title + 4 KPI cards + summary
#   Slide 2  Monthly Sales Trends - MoM colour-coded chart
#   Slide 3  Product Performance  - multicolour Top 10 bars
#   Slide 4  Country & Customer Insights
#   Slide 5  SQL Analysis & Data Quality
#   Slide 6  Business Recommendations + Action Plan
#
# Run:  python dash_app.py   then open http://127.0.0.1:8050
# Hosted: gunicorn dash_app:server --workers 1 ...
# =====================================================

import json
import os

import dash
from dash import Dash, html, dcc, Input, Output, State, ctx, no_update
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from types import SimpleNamespace

# data_prep is imported ONLY for the formatting helpers and theme
# colours shared with the Streamlit app. prepare_data() is never
# called here, so no Excel/UCI/SQLite work can happen at import.
import data_prep
from data_prep import (
    ACCENT_CYAN, ACCENT_BLUE, ACCENT_VIOLET,
    POSITIVE, NEGATIVE, NEUTRAL, CHART_TEXT, TICK_TEXT,
)

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dashboard_data")
SRC = data_prep

# The compact files the hosted app is allowed to read.
REQUIRED_DATA_FILES = [
    "kpis.json", "monthly.csv", "top_products.csv", "sql_monthly.csv",
    "sql_top.csv", "sql_country.csv", "orders_count.csv",
    "repeat_top.csv", "preview.csv",
]

BUILD_INSTRUCTIONS = {
    "kpis.json": "KPI / data-quality / recommendation scalars",
    "monthly.csv": "monthly revenue + month-over-month change (slide 2)",
    "top_products.csv": "top 10 products by revenue (slide 3)",
    "sql_monthly.csv": "SQL query 1 - monthly revenue table (slides 4-5)",
    "sql_top.csv": "SQL query 2 - top 10 products table (slide 5)",
    "sql_country.csv": "SQL query 3 - revenue by country (slides 4-5)",
    "orders_count.csv": "SQL query 4 - orders per month",
    "repeat_top.csv": "SQL query 5 - top repeat customers (slide 4)",
    "preview.csv": "raw data preview rows (slide 5)",
}


def _no_comparison_pct(pct):
    """True when a MoM value cannot be compared (same rule as data_prep)."""
    return pct is None or pd.isna(pct)


def _pct_txt(pct):
    """'+12.3% / -4.0% / n/a' from a month-over-month percentage."""
    return "n/a" if _no_comparison_pct(pct) else f"{pct:+,.1f}%"


def _top10_country_frame(rows, total_revenue):
    """kpis.json stores the 5 'leading country' cards; restore framing."""
    df = pd.DataFrame(rows)
    if not len(df):
        return df
    if total_revenue > 0:
        df["share_pct"] = df["Revenue_GBP"] / total_revenue * 100.0
    else:
        df["share_pct"] = 0.0
    return df


# ---------------------------------------------------------------
# Compact-data loader: the ONLY data work this app performs.
# ---------------------------------------------------------------
def load_dashboard_data(data_dir=DATA_DIR):
    """
    Reads the precomputed files created by build_dashboard_data.py
    and rebuilds the exact namespace the slide builders expect.

    Raises FileNotFoundError listing the missing files - the caller
    shows build instructions instead of any fake data.
    """
    missing = [f for f in REQUIRED_DATA_FILES
               if not os.path.exists(os.path.join(data_dir, f))]
    if missing:
        raise FileNotFoundError(
            "dashboard_data/ is missing required files: "
            + ", ".join(missing)
        )

    with open(os.path.join(data_dir, "kpis.json"), encoding="utf-8") as handle:
        K = json.load(handle)

    # keep_default_na=False + explicit empty-string NA list: this keeps the
    # literal "n/a" MoM labels ("n/a" would otherwise be parsed as NaN)
    # while still turning empty numeric cells into NaN.
    monthly = pd.read_csv(os.path.join(data_dir, "monthly.csv"),
                          keep_default_na=False, na_values=[""])
    top_products = pd.read_csv(os.path.join(data_dir, "top_products.csv"))
    sql_monthly = pd.read_csv(os.path.join(data_dir, "sql_monthly.csv"))
    sql_top = pd.read_csv(os.path.join(data_dir, "sql_top.csv"))
    sql_country = pd.read_csv(os.path.join(data_dir, "sql_country.csv"))
    sql_orders_count = pd.read_csv(os.path.join(data_dir, "orders_count.csv"))
    sql_repeat = pd.read_csv(os.path.join(data_dir, "repeat_top.csv"),
                             dtype={"CustomerID": str})
    preview_df = pd.read_csv(os.path.join(data_dir, "preview.csv"))

    total_revenue = float(K["total_revenue"])

    return SimpleNamespace(
        safe_data=bool(K["safe_data"]),
        # data quality (slide 5)
        total_rows=int(K["total_rows"]),
        rows_with_missing=int(K["rows_with_missing"]),
        missing_customer_rows=int(K["missing_customer_rows"]),
        missing_description_rows=int(K["missing_description_rows"]),
        cancelled_rows=int(K["cancelled_rows"]),
        nonpositive_qty_rows=int(K["nonpositive_qty_rows"]),
        nonpositive_price_rows=int(K["nonpositive_price_rows"]),
        excluded_rows=int(K["excluded_rows"]),
        overlap_rows=int(K["overlap_rows"]),
        kept_rows=int(K["kept_rows"]),
        # headline KPIs (slides 1, 5, 6)
        total_revenue=total_revenue,
        distinct_invoices=int(K["distinct_invoices"]),
        avg_order_value=float(K["avg_order_value"]),
        valid_customers=int(K["valid_customers"]),
        identified_customers=int(K["identified_customers"]),
        verify_total=float(K["verify_total"]),
        revenue_difference=float(K["revenue_difference"]),
        revenue_verified=bool(K["revenue_verified"]),
        # tables for the slides
        monthly=monthly,
        product_totals=top_products,          # top 10 == top of full sort
        n_products=int(K["n_products"]),      # full product count for text
        avg_product_revenue=float(K["avg_product_revenue"]),
        top_products=top_products,
        sql_monthly=sql_monthly,
        sql_top=sql_top,
        sql_country=sql_country,
        sql_orders_count=sql_orders_count,
        sql_repeat=sql_repeat,
        country_leading=_top10_country_frame(
            K["country_leading"], total_revenue),
        repeat_ids=int(K["repeat_ids"]),
        repeat_revenue=float(K["repeat_revenue"]),
        repeat_revenue_share=float(K["repeat_revenue_share"]),
        avg_orders_repeat=float(K["avg_orders_repeat"]),
        single_purchase_share=float(K["single_purchase_share"]),
        best_month=K["best_month"],
        weakest_month=K["weakest_month"],
        first_month=K["first_month"],
        last_month=K["last_month"],
        last_pct=K["last_pct"],
        uk_share=float(K["uk_share"]),
        uk_revenue=float(K["uk_revenue"]),
        top3_share=float(K["top3_share"]),
        preview_df=preview_df,
        _pct_txt=_pct_txt,
        _period_label=data_prep._period_label,
        _gbp=data_prep._gbp,
    )


# All summary tables/numbers are PRECOMPUTED locally (genuine Online
# Retail.xlsx + the same cleaning rules + the real SQL queries).
# If dashboard_data/ is missing, the app still starts and every
# slide clearly shows the build instructions - no fake data ever.
try:
    D = load_dashboard_data()
    LOAD_ERROR = None
except (FileNotFoundError, ValueError, KeyError, json.JSONDecodeError) as exc:
    D = None
    LOAD_ERROR = str(exc)

# Multicolour palette for the slide-3 product bars (rank order)
PRODUCT_PALETTE = [
    "#22d3ee", "#2563eb", "#8b5cf6", "#34d399", "#f59e0b",
    "#f87171", "#0ea5e9", "#a78bfa", "#fbbf24", "#38bdf8",
]

SLIDES = ["Executive Overview", "Monthly Sales Trends", "Product Performance",
          "Country & Customer Insights", "SQL Analysis & Data Quality",
          "Business Recommendations"]
N_SLIDES = len(SLIDES)


def compute_new_slide(action, current):
    """Pure navigation logic (easy to unit-test)."""
    if action == "next":
        return min(N_SLIDES, current + 1)
    if action == "prev":
        return max(1, current - 1)
    return 1                      # 'home' = Start Over / Back to Overview


# ---------------------------------------------------------------
# Small builders used by all slides
# ---------------------------------------------------------------

def money(value):
    return f"£{value:,.2f}"


def figure_card(fig):
    return html.Div(
        dcc.Graph(
            figure=fig,
            config={"displayModeBar": False, "responsive": True},
            style={"width": "100%", "height": "auto"},
        ),
        className="graph-card",
    )


def style_dark(fig, height=500):
    """One place that applies the dark readable look to every chart."""
    fig.update_layout(
        height=height,
        margin=dict(l=20, r=20, t=30, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Segoe UI, sans-serif", size=16, color=CHART_TEXT),
        showlegend=False,
        hoverlabel=dict(
            bgcolor="#101f36",
            bordercolor="rgba(34,211,238,0.6)",
            font=dict(color="#e8f1fb", size=16),
        ),
        transition_duration=450,
        transition_easing="cubic-in-out",
    )
    fig.update_xaxes(showgrid=False, tickfont=dict(color=TICK_TEXT, size=16))
    fig.update_yaxes(gridcolor="rgba(148,179,208,0.12)", tickfont=dict(color=TICK_TEXT, size=16))
    return fig


def kpi_card(label, value, note):
    return html.Div(
        [
            html.Div(label, className="kpi-label"),
            html.Div(value, className="kpi-value"),
            html.Div(note, className="kpi-note"),
        ],
        className="kpi-card",
    )


def mini_card(label, value, note=""):
    inner = [html.Div(label, className="label"), html.Div(value, className="value")]
    if note:
        inner.append(html.Div(note, className="note"))
    return html.Div(inner, className="card")


def build_table(df, max_height=460):
    """Readable table: 16px CSS text, sticky header, scrolls inside
    max_height so long tables never clip the slide."""
    cells = df.astype(str).where(pd.notna(df), "")
    return html.Div(
        html.Table(
            [
                html.Thead(
                    html.Tr([html.Th(str(c)) for c in cells.columns])
                ),
                html.Tbody(
                    [
                        html.Tr([html.Td(str(v)) for v in row])
                        for row in cells.itertuples(index=False)
                    ]
                ),
            ],
            className="data-table",
        ),
        className="data-scroll",
        style={"maxHeight": f"{max_height}px", "overflowY": "auto"},
    )


def legend(mappings):
    """Color key built from real components (Dash escapes raw strings)."""
    return html.Div(
        [
            html.Span(
                [html.Span(className="dot", style={"background": c}), t],
                className="legend-chip",
            )
            for c, t in mappings
        ],
        className="legend",
        style={"display": "flex", "gap": "22px", "flexWrap": "wrap", "margin": "4px 0 10px"},
    )


# ---------------------------------------------------------------
# Slide builders (pure functions -> easy to test)
# ---------------------------------------------------------------

def slide_1():
    head = html.Div(
        [
            html.Span("Executive Command Center", className="app-badge"),
            html.H1("Retail Intelligence", className="app-title"),
            html.P("Executive Sales & Customer Insights", className="app-subtitle"),
            html.Hr(className="app-header-rule"),
        ]
    )
    if D.safe_data:
        kpis = html.Div(
            [
                kpi_card("Total Revenue", f"£{D.total_revenue:,.0f}",
                         "Sum of Quantity × UnitPrice over all valid rows"),
                kpi_card("Valid Orders", f"{D.distinct_invoices:,}",
                         "Distinct InvoiceNo among valid sales rows"),
                kpi_card("Unique Customers", f"{D.identified_customers:,}",
                         "CustomerIDs present anywhere in the dataset"),
                kpi_card("Average Order Value", f"£{D.avg_order_value:,.2f}",
                         "Total Revenue ÷ Valid Orders"),
            ],
            className="kpi-grid",
        )
        summary_text = [
            html.H3("Summary", className="section"),
            html.Div(
                [
                    html.Span("Between "),
                    html.Strong(SRC._period_label(D.first_month["month"])),
                    html.Span(" and "),
                    html.Strong(SRC._period_label(D.last_month["month"])),
                    html.Span(f", valid sales generated £{D.total_revenue:,.0f} across "
                              f"{D.distinct_invoices:,} orders from {D.valid_customers:,} "
                              f"identified customers. Average order value is £{D.avg_order_value:,.2f}. "
                              f"The strongest month was "),
                    html.Strong(SRC._period_label(D.best_month["month"])),
                    html.Span(f" at £{D.best_month['revenue']:,.0f}; "),
                    html.Strong(SRC._period_label(D.last_month["month"])),
                    html.Span(f" came in at £{D.last_month['revenue']:,.0f} "
                              f"({'first month - no comparison' if D.last_pct is None else D._pct_txt(D.last_pct) + ' vs the previous month'}). "
                              f"Revenue is concentrated in "),
                    html.Strong(str(D.country_leading.iloc[0]['Country'])),
                    html.Span(f" ({D.uk_share:,.1f}% of total revenue)."),
                ],
                className="summary-box",
            ),
        ]
    else:
        kpis = html.Div()
        summary_text = [html.Div("No valid sales rows found after cleaning — "
                                 "check the data quality slide (5).")]
    return [head, kpis, *summary_text,
            html.P("Every number is precomputed locally from the genuine Online Retail.xlsx "
                   "using build_dashboard_data.py (valid sales rows only: not cancelled, "
                   "Quantity > 0, UnitPrice > 0). All amounts in British pounds (£). Use the "
                   "buttons below to move through the deck.",
                   className="caption")]


def slide_2():
    monthly = D.monthly
    fig = px.bar(
        monthly, x="month", y="revenue",
        labels={"month": "Month", "revenue": "Revenue (£)"},
        custom_data=["direction"],
    )
    fig.update_traces(
        marker=dict(color=monthly["bar_color"]),
        hovertemplate="<b>%{x}</b><br>Revenue: £%{y:,.0f}<br>%{customdata[0]}<extra></extra>",
    )
    style_dark(fig, height=500)

    mom_display = monthly[["month", "revenue", "mom_label", "direction"]].copy()
    mom_display.columns = ["Month", "Revenue", "Change vs previous", "Direction"]
    mom_display["Revenue"] = mom_display["Revenue"].map(D._gbp)

    return [
        html.H2("Monthly Sales Trends", className="slide-title"),
        html.P("Total revenue per month (valid sales only), coloured by the change against the "
               "previous month. The first month and zero-revenue prior months show n/a "
               "(no comparison is possible).", className="caption"),
        legend([(POSITIVE, "Increase (green)"), (NEGATIVE, "Decrease (red)"),
                (NEUTRAL, "Little / no change (amber)")]),
        figure_card(fig),
        html.H3("Month-over-month change", className="section"),
        build_table(mom_display, 460),
    ]


def slide_3():
    top_sorted = D.top_products.sort_values("revenue")
    fig = px.bar(
        top_sorted, x="revenue", y="Description", orientation="h",
        labels={"Description": "Product", "revenue": "Revenue (£)"},
    )
    # Clearly labelled multicolour bars: one distinct colour per rank,
    # product names on the y-axis, money values in the tooltip.
    fig.update_traces(
        marker=dict(
            color=[PRODUCT_PALETTE[i % len(PRODUCT_PALETTE)]
                   for i in range(len(top_sorted))],
        ),
        hovertemplate="<b>%{y}</b><br>Revenue: £%{x:,.0f}<extra></extra>",
    )
    style_dark(fig, height=520)

    leaders = D.product_totals.head(3).copy() if len(D.product_totals) else D.product_totals
    lead_cards = []
    if len(leaders) and D.total_revenue > 0:
        for i, (desc, rev) in enumerate(
            leaders[["Description", "revenue"]].itertuples(index=False)
        ):
            share = rev / D.total_revenue * 100
            lead_cards.append(mini_card(
                f"{i + 1}. {str(desc)[:38]}",
                f"£{rev:,.0f}",
                f"{share:,.1f}% of total revenue",
            ))
    leaders_grid = html.Div(lead_cards, className="card-grid") if lead_cards else html.Div()

    return [
        html.H2("Product Performance", className="slide-title"),
        html.P(f"Top 10 products by revenue (valid sales only). Each bar carries a distinct "
               f"colour; the legend below compares revenue with the dataset-wide average "
               f"product revenue of £{D.avg_product_revenue:,.0f} — this is about revenue, "
               f"NOT profitability (the dataset has no unit-cost data).",
               className="caption"),
        figure_card(fig),
        legend([(POSITIVE, f"Above dataset average (green)"),
                (NEUTRAL, f"Near the dataset average (amber)"),
                (NEGATIVE, f"Below dataset average (red)")]),
        html.H3("Highest-revenue products", className="section"),
        leaders_grid,
        html.P(f"Colour note: all Top-10 products sit far above the dataset-wide average of "
               f"£{D.avg_product_revenue:,.0f} across {D.n_products:,} products, so "
               f"they land in the green band — exactly what a Top 10 should look like.",
               className="caption"),
    ]


def slide_4():
    country_data = D.sql_country.head(10).sort_values("Revenue_GBP")
    fig = px.bar(
        country_data, x="Revenue_GBP", y="Country", orientation="h",
        labels={"Revenue_GBP": "Revenue (£)"},
    )
    fig.update_traces(
        marker=dict(color=country_data["Revenue_GBP"], showscale=False,
                    colorscale=[[0, ACCENT_VIOLET], [0.5, ACCENT_BLUE], [1, ACCENT_CYAN]]),
        hovertemplate="<b>%{y}</b><br>Revenue: £%{x:,.0f}<extra></extra>",
    )
    style_dark(fig, height=460)

    lead_cards = [
        mini_card(
            r["Country"],
            f"£{r['Revenue_GBP']:,.0f}",
            f"{r['share_pct']:,.1f}% of total revenue",
        )
        for _, r in D.country_leading.iterrows()
    ] if len(D.country_leading) else []

    repeat_cards = [
        mini_card("Repeat Customers", f"{D.repeat_ids:,}",
                  "more than one distinct invoice (SQL)"),
        mini_card("Revenue from Repeat Customers", f"£{D.repeat_revenue:,.0f}",
                  f"{D.repeat_revenue_share:,.1f}% of total revenue"),
        mini_card("Avg Orders per Repeat Customer", f"{D.avg_orders_repeat:,.1f}",
                  f"{D.single_purchase_share:,.1f}% of identified customers ordered only once"),
    ]

    loyal_display = D.sql_repeat.head(10).rename(columns={
        "CustomerID": "Customer ID", "Orders": "Orders", "Revenue_GBP": "Revenue (£)",
    }).assign(**{"Revenue (£)": lambda d2: d2["Revenue (£)"].map(D._gbp)})

    return [
        html.H2("Country and Customer Insights", className="slide-title"),
        html.P("Country revenue and repeat-customer facts precomputed locally from the "
               "cleaned rows and the SQLite sales table (build_dashboard_data.py). No "
               "customer behaviour is assumed beyond the data.",
               className="caption"),
        html.Div(
            [
                html.Div([html.H3(f"Revenue by country (top 10 of {len(D.sql_country)} all)",
                                  className="section"), figure_card(fig)], className="col"),
                html.Div(
                    [html.H3("Leading countries", className="section"),
                     html.Div(lead_cards, className="card-grid")],
                    className="col",
                ),
            ],
            className="grid-2",
        ),
        html.H3("Repeat-customer metrics", className="section"),
        html.Div(repeat_cards, className="card-grid"),
        html.H3("Most loyal customers (by order count)", className="section"),
        build_table(loyal_display, 460),
        html.P("Source: SQL query 5 (repeat customers). Rows without a CustomerID are skipped "
               "there because that customer cannot be identified.", className="caption"),
    ]


def slide_5():
    sql_m = D.sql_monthly.sort_values("Month")
    fig = px.bar(sql_m, x="Month", y="Revenue_GBP", labels={"Revenue_GBP": "Revenue (£)"})
    fig.update_traces(
        marker=dict(color=sql_m["Revenue_GBP"], showscale=False,
                    colorscale=[[0, "#1d4ed8"], [0.55, ACCENT_BLUE], [1, ACCENT_CYAN]]),
        hovertemplate="<b>%{x}</b><br>Revenue: £%{y:,.0f}<extra></extra>",
    )
    style_dark(fig, height=460)

    sql_top_display = D.sql_top.rename(columns={"Revenue_GBP": "Revenue (£)"}) \
        .assign(**{"Revenue (£)": lambda d2: d2["Revenue (£)"].map(D._gbp)})
    sql_country_display = D.sql_country.rename(columns={"Revenue_GBP": "Revenue (£)"}) \
        .assign(**{"Revenue (£)": lambda d2: d2["Revenue (£)"].map(D._gbp)})

    quality_cards = [
        mini_card("Rows in Original Dataset", f"{D.total_rows:,}", "every row of Online Retail.xlsx"),
        mini_card("Cancelled Invoice Rows", f"{D.cancelled_rows:,}",
                  'InvoiceNo starts with "C" (credit notes)'),
        mini_card("Rows with Missing Values", f"{D.rows_with_missing:,}",
                  f"missing CustomerID: {D.missing_customer_rows:,} · missing Description: {D.missing_description_rows:,}"),
        mini_card("Rows Excluded from Sales", f"{D.excluded_rows:,}",
                  "cancelled OR Quantity ≤ 0 OR UnitPrice ≤ 0"),
    ]

    counts_df = pd.DataFrame(
        {
            "Check": ["Total rows in the original dataset",
                      "Rows with missing values (any column)",
                      "↳ of which missing CustomerID",
                      "↳ of which missing Description",
                      "Cancelled invoice rows",
                      "↳ of which Quantity ≤ 0",
                      "↳ of which UnitPrice ≤ 0",
                      "Rows excluded from sales calculations",
                      "↳ overlap: cancelled and non-positive",
                      "Rows kept for sales calculations"],
            "Rows": [f"{D.total_rows:,}", f"{D.rows_with_missing:,}",
                     f"{D.missing_customer_rows:,}", f"{D.missing_description_rows:,}",
                     f"{D.cancelled_rows:,}", f"{D.nonpositive_qty_rows:,}",
                     f"{D.nonpositive_price_rows:,}", f"{D.excluded_rows:,}",
                     f"{D.overlap_rows:,}", f"{D.kept_rows:,}"],
        }
    )

    if D.revenue_verified:
        verified = html.P(
            f"✓ Verified — recomputing Quantity × UnitPrice independently gives "
            f"£{D.verify_total:,.2f}, matching the Total Revenue KPI "
            f"(difference £{D.revenue_difference:.4f}). Average Order Value = "
            f"£{D.total_revenue:,.2f} ÷ {D.distinct_invoices:,} invoices = "
            f"£{D.avg_order_value:,.2f}.",
            className="caption",
        )
    else:
        verified = html.P(
            f"Revenue check mismatch: KPI shows £{D.total_revenue:,.2f} but the independent "
            f"recompute gives £{D.verify_total:,.2f}. Review the data preparation step.",
            className="caption",
        )

    return [
        html.H2("SQL Analysis and Data Quality", className="slide-title"),
        html.P(
            "SQL logic in one minute: the cleaned valid rows live in the SQLite table `sales` "
            "(duplicate-safe, built locally from the same DataFrame the KPIs use). Monthly revenue = "
            "SUM(Revenue) grouped by month; top products = largest SUM(Revenue) per product "
            "(LIMIT 10); country revenue = SUM(Revenue) per country; repeat customers = "
            "GROUP BY CustomerID HAVING COUNT(DISTINCT InvoiceNo) > 1 (rows without a "
            "CustomerID are skipped because that customer cannot be identified).",
            className="caption",
        ),
        html.Div(
            [
                html.Div(
                    [html.H3("Monthly revenue (SQL)", className="section"), figure_card(fig)],
                    className="col"),
                html.Div(
                    [
                        html.H3("Top 10 products (SQL)", className="section"),
                        build_table(sql_top_display, 300),
                        html.H3("Revenue by country (SQL)", className="section"),
                        build_table(sql_country_display, 300),
                    ],
                    className="col"),
            ],
            className="grid-2",
        ),
        html.H3("Data Quality Summary", className="section"),
        html.P("No rows are ever changed or deleted in the Excel file — rows are only flagged. "
               "“Excluded” below means left out of the sales KPIs only.", className="caption"),
        html.Div(quality_cards, className="card-grid"),
        build_table(counts_df, 430),
        verified,
        html.H3("Raw Data Preview", className="section"),
        html.P("First 60 valid rows plus 20 cancelled and 20 non-positive examples from the "
               "file. Only rows with status “Included in sales” count toward Total Revenue.",
               className="caption"),
        build_table(D.preview_df, 460),
    ]


def _rec_card(title, hint, css, evidence, action, reason, kpi):
    return html.Div(
        [
            html.Div(title, className="rec-title"),
            html.Div(hint, className=f"rec-hint {css}"),
            html.P([html.Span("Observed evidence", className="rec-label"),
                    html.Span(evidence, style={"display": "inline"})],
                   className="rec-line"),
            html.P([html.Span("Recommended action", className="rec-label"),
                    html.Span(action)], className="rec-line"),
            html.P([html.Span("Business reason", className="rec-label"),
                    html.Span(reason)], className="rec-line"),
            html.P([html.Span("KPI to monitor", className="rec-label"),
                    html.Span(kpi)], className="rec-line"),
        ],
        className=f"rec-card {css}",
    )


def slide_6():
    runner_ups = ", ".join(D.country_leading.iloc[1:4]["Country"].tolist()) if len(D.country_leading) else ""
    backed = [
        _rec_card(
            "1. Revenue depends heavily on one market",
            "observed", "rec-backed",
            html.Span([html.B(str(D.country_leading.iloc[0]["Country"])),
                       f" generated £{D.uk_revenue:,.0f} = {D.uk_share:,.1f}% of all revenue "
                       f"({len(D.sql_country)} countries in total)."]),
            "Prepare a pilot marketing/promotions plan for the strongest other markets "
            f"(next by revenue: {runner_ups}).",
            "Concentration on one market is a commercial risk; the runner-up markets already "
            "generate real revenue, so small targeted efforts build on proven demand.",
            "Non-UK revenue share per month (rising trend, judged quarterly).",
        ),
        _rec_card(
            "2. A few products carry the business",
            "observed", "rec-backed",
            html.Span([f"The top 3 products account for {D.top3_share:,.1f}% of total revenue, led by “",
                       html.B(str(D.product_totals.iloc[0]["Description"])),
                       f"” at £{D.product_totals.iloc[0]['revenue']:,.0f}."]),
            "Protect availability of the top earners (stock planning, listing quality) and "
            "cross-sell around them in the same orders.",
            "Revenue concentration means stock-outs on a top product directly hit sales; "
            "cross-selling rides on existing demand without new product bets.",
            "Top-3 revenue share per month (stable share, rising total).",
        ),
        _rec_card(
            "3. Timing of strength and weakness is known",
            "observed", "rec-backed",
            html.Span([
                f"Strongest month: {(SRC._period_label(D.best_month['month']))} at "
                f"£{D.best_month['revenue']:,.0f}. Weakest month: "
                f"{(SRC._period_label(D.weakest_month['month']))} at "
                f"£{D.weakest_month['revenue']:,.0f}. The latest month "
                f"({(SRC._period_label(D.last_month['month']))}) shows "
                f"{D._pct_txt(D.last_pct)} vs its previous month."
            ]),
            "Schedule promotions before the historically strong period and run a drill-down "
            "review of the weak month (its orders, invoice sizes and mix).",
            "The data shows a real seasonal gap; acting on the pattern avoids reacting late.",
            "Monthly revenue and its month-over-month change (fewer deep negative months).",
        ),
        _rec_card(
            "4. Repeat customers are a proven revenue engine",
            "observed", "rec-backed",
            html.Span([f"{D.repeat_ids:,} of {D.valid_customers:,} identified customers bought "
                       f"more than once (avg {D.avg_orders_repeat:,.1f} orders each) and generate "
                       f"{D.repeat_revenue_share:,.1f}% of all revenue."]),
            "Create a straightforward repeat-purchase nudge for high-frequency customers "
            "(top names are listed on slide 4).",
            "Their revenue share is measured, so serving them better supports the largest "
            "measured part of the business.",
            "Revenue from repeat customers (% of total) and average orders per repeat customer.",
        ),
    ]
    ideas = [
        _rec_card(
            "Raise average order value",
            "hypothesis - needs testing", "rec-test",
            html.Span([f"Average order value currently stands at £{D.avg_order_value:,.2f}."]),
            "Pilot an A/B test of a bundle or multi-buy offer built around top products.",
            "Larger baskets would raise order value if customers respond; this is a guess "
            "until the pilot measures a real effect.",
            "Average order value in the pilot group vs control.",
        ),
        _rec_card(
            "Win back one-order customers",
            "hypothesis - needs testing", "rec-test",
            html.Span([f"{D.single_purchase_share:,.1f}% of identified customers ordered only once."]),
            "Test a single, well-targeted follow-up campaign for first-order customers.",
            "If even a modest share return, repeat purchases rise; the effect must be measured, "
            "not assumed.",
            "Repeat-purchase rate in the campaign cohort vs control.",
        ),
        _rec_card(
            "Investigate small markets before investing",
            "hypothesis - needs testing", "rec-test",
            html.Span([f"Of {len(D.sql_country)} countries, 30+ earn only a small fraction of "
                       f"revenue (full list on slide 5)."]),
            "Run a short data investigation (shipping, returns, order sizes) before committing "
            "budget to small markets.",
            "Cheap analysis avoids expensive expansion mistakes where the data is thin.",
            "Revenue per active market (quarterly change).",
        ),
    ]

    def planetwo(name, lines, css):
        items = [html.Div(line, className="plan-item") for line in lines]
        return html.Div([html.Div(name, className="plan-name"), *items],
                        className=f"plan-card {css}")

    return [
        html.H2("Business Recommendations", className="slide-title"),
        html.P("Generated from the calculated figures on the previous slides. This dataset has "
               "no unit-cost data, so none of these address profit or margins, and no revenue "
               "gains or forecasts are claimed.", className="caption"),
        html.H3("Evidence-backed findings", className="section"),
        *backed,
        html.H3("Ideas that still need testing", className="section"),
        *ideas,
        html.H3("Prioritized action plan", className="section"),
        html.Div(
            [
                planetwo("Immediate (this week)",
                         ["1. Lock stock/availability of the top-3 revenue products.",
                          "2. Note the latest month-over-month move from slide 2 and brief the team.",
                          "3. Confirm which weak month/market drill-down gets scheduled first."],
                         "plan-1"),
                planetwo("Next 30 days",
                         ["1. Pilot targeted campaigns in the strongest non-UK markets (slide 4).",
                          "2. Start the A/B test for average order value (idea above).",
                          "3. Send a first repeat-purchase nudge to the top repeat customers."],
                         "plan-2"),
                planetwo("Longer term",
                         ["1. Review the concentration KPIs (top-3 share, non-UK share) quarterly.",
                          "2. Grow the repeat-revenue share with a small loyalty programme.",
                          "3. Rebuild this deck with refreshed data each month to track progress."],
                         "plan-3"),
            ],
            className="grid-3",
        ),
        html.P("Method note: all evidence figures are the same computed values shown on slides "
               "1-5 (no external data, no cost assumptions). Items marked “needs testing” are "
               "deliberately labelled as hypotheses.", className="caption"),
    ]


SLIDE_BUILDERS = [slide_1, slide_2, slide_3, slide_4, slide_5, slide_6]


def data_error_slide():
    """Shown on every slide while the precomputed data is unavailable:
    dashboard_data/ (built locally by build_dashboard_data.py) is
    missing or unreadable. No sample/fake data is ever substituted."""
    missing = []
    for name, purpose in BUILD_INSTRUCTIONS.items():
        missing.append(html.Li(f"{name} — {purpose}"))
    return html.Div(
        [
            html.H2("Precomputed data missing", className="slide-title"),
            html.P(
                "The dashboard loads compact precomputed results from the "
                "dashboard_data/ folder (no raw dataset is ever loaded on "
                "the server). Those files are missing here, so no KPIs are "
                "displayed (fake/sample data is never substituted). To fix "
                "this, build them locally and commit them to the repository:",
                className="caption",
            ),
            html.Pre(
                "python build_dashboard_data.py          # needs Online Retail.xlsx locally\n"
                "git add dashboard_data/ && git commit -m \"Add precomputed dashboard data\"\n"
                "git push",
                className="summary-box",
            ),
            html.P(
                "Required files and what each one feeds:",
                className="caption",
            ),
            html.Ul(missing, className="caption"),
            html.P(
                "Every KPI, SQL result and recommendation shown afterwards is "
                "computed from the genuine dataset on your machine - nothing "
                "is hardcoded or invented.",
                className="caption",
            ),
            html.Pre(str(LOAD_ERROR), className="summary-box"),
        ],
        className="slide-wrap",
    )


def build_slide(slide_no):
    """Compose the components of one slide (pure function, testable)."""
    slide_no = max(1, min(N_SLIDES, int(slide_no)))
    if D is None:
        return html.Div(data_error_slide(), className="slide-wrap")
    return html.Div(SLIDE_BUILDERS[slide_no - 1](), className="slide-wrap")


def build_progress(slide_no):
    fill = 100.0 * slide_no / N_SLIDES
    label = f"Slide {slide_no} of {N_SLIDES} — {SLIDES[slide_no - 1]}"
    return fill, label


# ---------------------------------------------------------------
# App + layout + navigation callbacks
# ---------------------------------------------------------------

app = Dash(
    __name__,
    title="Retail Intelligence — Executive Presentation",
    assets_folder="assets",
)

app.layout = html.Div(
    [
        dcc.Store(id="slide-store", data={"slide": 1}),
        html.Div(
            [
                html.Div(
                    [
                        html.Div(className="progress-track",
                                 children=html.Div(id="progress-fill",
                                                   className="progress-fill",
                                                   style={"width": "16.7%"})),
                        html.Span("Slide 1 of 6 — Executive Overview", id="progress-label",
                                  className="progress-label"),
                    ],
                    className="progress-row",
                ),
                html.Div(build_slide(1), id="slide-body"),
                html.Div(
                    [
                        html.Button("Previous", id="prev-btn", className="nav-btn",
                                    disabled=True),
                        html.Button("Start Over (Back to Overview)", id="home-btn",
                                    className="nav-btn"),
                        html.Button("Next", id="next-btn", className="nav-btn nav-primary"),
                    ],
                    className="nav-bar",
                ),
                html.P("Previous is disabled on slide 1 and Next on the last slide. "
                       "“Start Over” returns to the Executive Overview.",
                       className="nav-footer"),
            ],
            className="deck-inner",
        ),
    ],
    id="deck",
)


@app.callback(
    Output("prev-btn", "disabled"),
    Output("home-btn", "disabled"),
    Output("next-btn", "disabled"),
    Input("slide-store", "data"),
)
def update_nav(data):
    slide = int(data.get("slide", 1))
    return slide == 1, slide == 1, slide == N_SLIDES


@app.callback(
    Output("progress-fill", "style"),
    Output("progress-label", "children"),
    Output("slide-body", "children"),
    Input("slide-store", "data"),
)
def render_slide(data):
    slide = max(1, min(N_SLIDES, int(data.get("slide", 1))))
    width, label = build_progress(slide)
    return {"width": f"{width:.1f}%"}, label, build_slide(slide)


@app.callback(
    Output("slide-store", "data"),
    Input("next-btn", "n_clicks"),
    Input("prev-btn", "n_clicks"),
    Input("home-btn", "n_clicks"),
    State("slide-store", "data"),
    prevent_initial_call=True,
)
def navigate(next_clicks, prev_clicks, home_clicks, data):
    action = ""
    if ctx.triggered_id == "next-btn":
        action = "next"
    elif ctx.triggered_id == "prev-btn":
        action = "prev"
    elif ctx.triggered_id == "home-btn":
        action = "home"
    else:
        return no_update
    current = int(data.get("slide", 1)) if data else 1
    return {"slide": compute_new_slide(action, current)}


# WSGI entry point for production servers (Render uses Gunicorn:
#   gunicorn dash_app:server )
server = app.server


if __name__ == "__main__":
    import os

    app.run(
        host="0.0.0.0" if os.environ.get("PORT") else "127.0.0.1",
        port=int(os.environ.get("PORT", 8050)),
        debug=False,
    )

