# Retail Intelligence — Retail Sales Analytics Dashboard

An end-to-end retail analytics project: I cleaned a real 541,909-row transaction dataset, built KPIs, wrote SQL queries against a SQLite database, and turned the results into a six-slide executive dashboard built with Plotly Dash.

- **Data:** Online Retail.xlsx — 541,909 rows of an online giftware store (Dec 2010 – Dec 2011)
- **Stack:** Python, Pandas, SQLite + SQL, Plotly (Express + Graph Objects), Plotly Dash
- **Result:** £10,666,684.54 in valid sales revenue across 19,960 orders — presented as an interactive slide deck every stakeholder can click through
- **Everything in this README is verified from the actual dataset.** No number here was invented or estimated.

---

## Table of Contents

1. [The Business Problem](#1-the-business-problem)
2. [About the Dataset](#2-about-the-dataset)
3. [Tools & How I Used Them](#3-tools--how-i-used-them)
4. [Project Structure](#4-project-structure)
5. [Installation & Run Instructions (Windows)](#5-installation--run-instructions-windows)
6. [My Workflow, Step by Step](#6-my-workflow-step-by-step)
7. [Data Cleaning Decisions](#7-data-cleaning-decisions)
8. [KPI Definitions](#8-kpi-definitions)
9. [The SQL Queries Explained](#9-the-sql-queries-explained)
10. [Dashboard Design](#10-dashboard-design)
11. [Five Verified Insights](#11-five-verified-insights)
12. [Actual Findings vs. Recommendations That Need Testing](#12-actual-findings-vs-recommendations-that-need-testing)
13. [15 Interview Questions & Simple Answers](#13-15-interview-questions--simple-answers)
14. [Honest Limitations](#14-honest-limitations)

---

## 1. The Business Problem

An online giftware retailer (UK-based, selling to consumers and small wholesalers in ~38 countries) has one year of raw transaction data sitting in an Excel file. Leadership cannot answer basic questions from the raw file:

- How is revenue trending month by month, and where are the strong and weak periods?
- Which products actually earn the money?
- Which countries matter, and how dependent are we on one market?
- Do customers come back and buy again — and how much revenue do repeat customers drive?
- Is the raw data even trustworthy? (Cancellations, negative quantities, missing customer IDs…)

My job: turn that messy Excel file into a small set of **trustworthy KPIs and true stories**, and present them as a dashboard that an executive can scroll through in five minutes.

**What this project deliberately does NOT do** (so I never overstate it):

- It does **not** calculate profit or margin — the dataset has no unit-cost column, so I claim nothing about profitability.
- It does **not** forecast revenue — there is no forecasting code.
- It does **not** invent new data. Every row that is excluded from KPIs is *flagged*, counted and shown, never silently deleted.
- Recommendations about future actions are clearly labelled as things that **need testing**, separate from measured facts.

---

## 2. About the Dataset

`Online Retail.xlsx` — a well-known public dataset of transactions from a UK-based, non-store online retailer selling unique all-occasion giftware. I use it **read-only**; the app never writes back to the Excel file.

**Source & attribution:** the dataset is the *UCI "Online Retail" dataset* (ID 352), hosted on the official UCI Machine Learning Repository:
https://archive.ics.uci.edu/dataset/352/online+retail
I downloaded by Chen, Sain, Sain, Guo, Ford & Huang (Daqing Chen, 2012). *Online Retail*. UCI Machine Learning Repository. DOI: 10.24432/C5BW33.

**How the app finds the data:**

1. **If `Online Retail.xlsx` exists in the project folder** (a normal local run) it is read directly — nothing ever leaves your laptop.
2. **If the file is missing** (e.g. the Excel file was never committed to GitHub), the app fetches the *genuine* dataset from the official UCI repository with the `ucimlrepo` package (`fetch_ucirepo(id=352)`) and shapes it into exactly the same columns as the Excel file.
3. **If both fail** (no local file and no network), the dashboard shows a clear "Data unavailable" error on every slide. No fake/sample data is ever substituted, and the app never starts with invented numbers.

| Fact | Value |
|---|---|
| Rows | 541,909 |
| Date range | 1 Dec 2010 → 9 Dec 2011 |
| Columns | InvoiceNo, StockCode, Description, Quantity, InvoiceDate, UnitPrice, CustomerID, Country |
| Missing CustomerID | 135,080 rows (≈ 25% of the file) |
| Missing Description | 1,454 rows |
| Cancelled invoice rows (InvoiceNo starts with "C") | 9,288 |
| Valid sales rows kept for KPIs | 530,104 |
| Distinct products | 4,026 |
| Countries | 38 |
| Currency | British pounds (£) |

**What one row means:** one line on an invoice — e.g. invoice 536365 bought 6 units of "WHITE HANGING HEART T-LIGHT HOLDER" at £3.39 each. One invoice (`InvoiceNo`) = one order; a customer can appear on many invoices.

---

## 3. Tools & How I Used Them

| Tool | Where it lives in this project | What I actually did with it |
|---|---|---|
| **Python 3** | all files | Glue language: reads the Excel file, orchestrates cleaning, KPI math, database loading and the web app. |
| **Pandas** | `data_prep.py` | `read_excel()` to load 541,909 rows; boolean masks to *flag* cancelled/invalid rows; `groupby()` to build monthly revenue, product totals and country tables; month-over-month % changes; built the cleaned `valid_sales` DataFrame that feeds both the KPIs and the database. |
| **SQLite (sqlite3 library)** | `sql_analysis.py`, `retail_sales.db` | A real local database file. Cleaned rows are inserted into a `sales` table with **duplicate protection** (a row-count check prevents re-loading on every app refresh). |
| **SQL** | `sql_queries.sql` | 5 hand-written, commented queries (monthly revenue, top-10 products, revenue by country, orders per month, repeat customers) run via `pd.read_sql_query()` so the dashboard shows exactly the SQL it executes. |
| **Plotly Express / Graph Objects** | `dash_app.py` | Interactive bar charts: monthly revenue coloured by direction of change, multicolour Top-10 product bars, gradient country bars, custom hover tooltips with actual values. |
| **Plotly Dash** | `dash_app.py` | The presentation layer: a six-slide deck with one slide visible at a time, Previous/Next buttons, progress bar, `dcc.Store` for slide state, and callbacks that rebuild only the visible slide. |
| **Streamlit** | `app.py` (kept as backup) | The original scrolling version of the same analysis. It is **preserved untouched** as a second, earlier deliverable. |
| **openpyxl** | via `pandas.read_excel` | The engine that lets Pandas read `.xlsx` files. |

**One design decision worth noting:** all cleaning rules and KPI definitions live in exactly one file (`data_prep.py`), computed once and cached with `functools.lru_cache`. Both dashboards import that one module, so both always show identical numbers.

---

## 4. Project Structure

```
retail-analytics-dashboard/
├── Online Retail.xlsx        # source data, 541,909 rows (read-only, never modified; optional — see "How the app finds the data")
├── retail_sales.db           # SQLite database (sales table, rebuilt from loaded data on every startup)
├── data_prep.py              # ONE source of truth: data loader (Excel → UCI fallback) + cleaning rules + all KPIs (cached)
├── sql_analysis.py           # SQLite helper: safe load + query runner
├── sql_queries.sql           # the 5 commented SQL queries the dashboard runs
├── dash_app.py               # the Plotly Dash slide deck (START HERE) — exposes `server = app.server` for Gunicorn
├── app.py                    # original Streamlit version (backup, untouched)
├── app_backup_*.py           # older snapshots of app.py
├── requirements.txt          # python dependencies
├── render.yaml               # Render deployment config (build + start commands)
├── backups_pre_render/       # snapshot of key files taken before the deployment preparation
├── assets/style.css          # dark executive theme for the dashboards
└── .streamlit/config.toml    # Streamlit dark theme
```

---

## 5. Installation & Run Instructions (Windows)

You need Python 3.10+ installed (tested on Python 3.14). Everything runs locally — no internet required after installing the packages.

**Step 1 — open a terminal in the project folder** (PowerShell or Command Prompt):

```powershell
cd "C:\Users\<you>\Documents\retail-analytics-dashboard"
```

**Step 2 — create and activate a virtual environment** (keeps this project's packages separate from your system):

```powershell
python -m venv .venv
.venv\Scripts\activate
```

Your prompt now starts with `(.venv)`.

**Step 3 — install the dependencies:**

```powershell
pip install -r requirements.txt
```

(Installs: pandas, openpyxl, plotly, dash, ucimlrepo, gunicorn, streamlit.)

**Step 4 — run the dashboard:**

```powershell
python dash_app.py
```

**Step 5 — open it in your browser:**

```
http://127.0.0.1:8050
```

Use the **Previous / Next / Start Over** buttons to move through the 6 slides (Previous is disabled on slide 1, Next on the last slide; the progress bar at the top shows where you are).

**Good to know:**

- The first run reads the full 541,909-row Excel file, so allow roughly **30–60 seconds** before the browser page becomes responsive. After that, results are cached in memory for the session.
- `retail_sales.db` **does not need to be present at all**: on every startup the app rebuilds the `sales` table from the loaded data (it only inserts when the row count doesn't match), so it can never double-load. This is what makes deployment work — hosting services do not keep files built at runtime.
- Stop the app with `Ctrl+C` in the terminal.
- To view the older Streamlit version instead: `streamlit run app.py` (opens automatically in the browser on port 8501).

---

## 5b. Deployment on Render

### Files that prepare this project for Render

| File | Purpose |
|---|---|
| `render.yaml` | Render service config: `runtime: python`, `buildCommand: pip install -r requirements.txt`, `startCommand: gunicorn dash_app:server --workers 1 --threads 4 --bind 0.0.0.0:$PORT` |
| `dash_app.py` | exposes **`server = app.server`** — Gunicorn's entry point (`gunicorn dash_app:server` loads `dash_app.py` and treats `server` as the WSGI app) |
| `requirements.txt` | full dependency list, including `gunicorn`, `ucimlrepo` and `openpyxl` |
| `data_prep.py` | data-loader fallback: local Excel → genuine UCI dataset (id 352) via `ucimlrepo`; SQLite `sales` table rebuilt from that data on every startup |

### Step-by-step

1. Push the repository to GitHub (the Excel file is intentionally **not** required to be committed — `ucimlrepo` retrieves the genuine dataset on the server).
2. On [render.com](https://render.com): **New → Blueprint**, select your repository. Render reads `render.yaml` automatically.
   - If you prefer **New → Web Service** manually: Runtime **Python**, Build command `pip install -r requirements.txt`, Start command `gunicorn dash_app:server --workers 1 --threads 4 --bind 0.0.0.0:$PORT`.
3. Choose the **Free** plan and create the service. The first deploy takes a few minutes (dependencies + a one-time download of the UCI dataset ≈ 60–90 s of extra startup time).
4. Open the service URL → the six-slide deck should appear exactly as it does locally.

### Deployment notes & honest warnings

- **Gunicorn is Linux-only** — `pip install gunicorn` also works on Windows but `gunicorn <…>` cannot *start* there (Win32API). That is expected: render.yaml's start command only ever runs on Render's Linux instances. Locally you keep using `python dash_app.py`; the WSGI entry point it feeds is verified with `from dash_app import server` (returns the Flask app).
- **Memory:** the dataset has 541,909 rows and the cleaned pipeline needs roughly 1–2 GB RAM. The Render **free** plan (512 MB) may hit its memory limit while loading; **Starter (512 MB) behaves the same, so if the worker gets killed, choose a plan with at least 2 GB RAM (Standard)**. `--workers 1` (explicit in the start command) matters because each extra worker would re-load another ~2 GB.

---

## 6. My Workflow, Step by Step

```
Online Retail.xlsx (read-only)
        │
        ▼
Pandas cleaning + flagging           data_prep.py
        │
        ├──► valid_sales DataFrame (530,104 rows)  ──► KPIs, monthly/product/country tables
        │
        ▼
SQLite sales table (row-count-protected)     sql_analysis.py
        │
        ▼
5 commented SQL queries                      sql_queries.sql
        │
        ▼
Pandas DataFrames from pd.read_sql_query
        │
        ▼
6-slide Plotly Dash deck                     dash_app.py + assets/style.css
```

1. **Load once:** `prepare_data()` in `data_prep.py` reads the Excel file, applies the cleaning rules, computes every KPI, summary table and the SQL results, and returns a single object. `@lru_cache` means it happens once per run, not on every slide click.
2. **Flag, don't delete:** invalid rows get a status label ("Cancelled invoice", "Non-positive quantity / price") and appear on the Data Quality slide. Only their contribution to *sales KPIs* is excluded.
3. **Switch to SQL for the interview-grade questions:** cleaned rows go into SQLite, and the five standard SQL patterns (`SUM`, `GROUP BY`, `COUNT(DISTINCT)`, `HAVING`, `LIMIT`) run from `sql_queries.sql`.
4. **Verify:** revenue is independently recomputed and compared with the KPI — the difference must be < £0.005. On this dataset it matches to the penny (£0.0000 difference), and slide 5 shows this check on screen.
5. **Present:** Dash renders each slide as a function of the slide number held in `dcc.Store`; navigation callbacks touch *only presentation state*, so numbers can never be changed by clicking.

---

## 7. Data Cleaning Decisions

The dataset is real and messy. Here are the exact decisions (with exact counts) and *why*:

| # | Decision | Rule | How many rows | Why |
|---|---|---|---|---|
| 1 | Treat `C`-prefixed invoices as cancellations | `InvoiceNo` starts with "C" | 9,288 | In this dataset a "C" invoice is a credit note (an order being reversed). Each cancelled row also carries Quantity ≤ 0 — counting it as "sales" would double the books. |
| 2 | Drop negative/zero economics from *sales KPIs only* | `Quantity ≤ 0` OR `UnitPrice ≤ 0` | 10,624 with Quantity ≤ 0; 2,517 with UnitPrice ≤ 0 (ranks overlap with cancellations) | A sale with a non-positive quantity or price is a return, an adjustment or a fee — not revenue. |
| 3 | Flag, never delete | rows get `row_status` labels | 0 rows deleted | Total excluded from KPIs: **11,805** of 541,909 (that's the union of rules 1+2; the 9,288 cancelled rows overlap with the non-positive group). 530,104 rows remain. Nothing is hidden — slide 5 shows every count. |
| 4 | Missing `CustomerID` kept | 135,080 rows | kept in valid sales | These are still real sales (likely walk-in / guest checkouts). Excluding them would distort revenue; excluding them only from *repeat-customer analysis*, where identity is unknowable, is the safe choice. |
| 5 | Missing `Description` → "Unknown" | `fillna("Unknown")` | 1,454 rows | Keeps product grouping charts working; the StockCode and £ amounts are still valid. |
| 6 | One invoice ≠ one row per customer | `COUNT(DISTINCT InvoiceNo)` | — | Sizes and repeats use distinct invoices, so a multi-line order is never counted as multiple orders. |
| 7 | Revenue per row re-verified | `Quantity × UnitPrice` recomputed a second time | diff = £0.0000 | A cheap independent check that the KPI math is correct — shown live on slide 5. |

**The interview sound-bite version:** *"I never delete bad rows — I flag them so every KPI is auditable. 11,805 of 541,909 rows are excluded from sales KPIs and separately counted, so anyone can argue with my cleaning and recompute."*

---

## 8. KPI Definitions

All defined in `data_prep.py`, used identically by both dashboards. Amounts in GBP (£).

| KPI | Formula | Value on this data |
|---|---|---|
| **Total Revenue** | SUM of `Quantity × UnitPrice` over valid rows (not cancelled, Quantity > 0, UnitPrice > 0) | £10,666,684.54 |
| **Valid Orders** | COUNT of distinct `InvoiceNo` within valid rows | 19,960 |
| **Average Order Value (AOV)** | Total Revenue ÷ Valid Orders | £534.40 |
| **Unique Customers** (KPI card) | distinct `CustomerID` anywhere in the dataset | 4,372 |
| **Identified valid customers** (used in the summary text) | distinct `CustomerID` within valid rows | 4,338 |
| **MoM (Month-over-Month change)** | (this month − last month) ÷ last month × 100. "n/a" when there is no previous month to compare | colour-coded per month |
| **Repeat customer** | a `CustomerID` with > 1 distinct valid invoice (SQL `HAVING COUNT(DISTINCT InvoiceNo) > 1`) | 2,845 customers |
| **Repeat revenue share** | revenue from repeat customers ÷ Total Revenue | 77.8% |
| **Single-purchase share** | identified valid customers minus repeat customers, as % of identified valid customers | 34.4% |
| **Country revenue share** | country revenue ÷ Total Revenue | UK = 84.6% |

Two honest details I can defend if challenged:

- **Why the two "customers" numbers differ:** 4,372 CustomerIDs exist in the whole file; 4,338 of them appear on valid sales rows. Both are shown with their exact definitions, so nobody thinks it's an error.
- **Why AOV is a plain ratio, not "median":** leadership asked "what does the average order look like", so Total ÷ Orders is the answer; the monthly Orders table is right next to it for context.

---

## 9. The SQL Queries Explained

The cleaned valid-sales rows live in a SQLite table named `sales` (columns: `InvoiceNo, StockCode, Description, Quantity, UnitPrice, InvoiceDate, CustomerID, Country, Revenue, Month`). `sql_queries.sql` contains five queries, each with an English comment banner. The app reads the file and runs each query, so **the SQL shown on screen is the SQL actually executed.**

**1) Monthly revenue**
```sql
SELECT Month, ROUND(SUM(Revenue), 2) AS Revenue_GBP
FROM sales GROUP BY Month ORDER BY Month;
```
*Every valid row is summed into its `YYYY-MM` bucket. This feeds the trend chart and MoM logic.*

**2) Top 10 products by revenue**
```sql
SELECT Description, ROUND(SUM(Revenue), 2) AS Revenue_GBP
FROM sales GROUP BY Description ORDER BY Revenue_GBP DESC LIMIT 10;
```
*One row per product, summed, `LIMIT 10` keeps the podium.*

**3) Revenue by country**
```sql
SELECT Country, ROUND(SUM(Revenue), 2) AS Revenue_GBP
FROM sales GROUP BY Country ORDER BY Revenue_GBP DESC;
```
*One row per country, biggest first (38 rows).*

**4) Order count by month**
```sql
SELECT Month, COUNT(DISTINCT InvoiceNo) AS Orders
FROM sales GROUP BY Month ORDER BY Month;
```
*`COUNT(DISTINCT)` is the key point: an order with 5 product lines and 5 rows must count once.*

**5) Repeat customers**
```sql
SELECT CustomerID, COUNT(DISTINCT InvoiceNo) AS Orders,
       ROUND(SUM(Revenue), 2) AS Revenue_GBP
FROM sales
WHERE CustomerID IS NOT NULL
GROUP BY CustomerID
HAVING COUNT(DISTINCT InvoiceNo) > 1
ORDER BY Orders DESC, Revenue_GBP DESC;
```
*`WHERE` filters before grouping; `HAVING` filters after — it keeps only customers with more than one invoice. Rows with no CustomerID are skipped because that customer cannot be identified.*

---

## 10. Dashboard Design

The dashboard is a **slide deck, not a scroll page**: one slide fills the screen, buttons move between slides. I chose this format because executive reviews happen live and verbally — a deck keeps attention on one story at a time.

| Slide | Content |
|---|---|
| 1 — Executive Overview | Title + the 4 headline KPI cards + a 5-sentence summary written from the actual numbers |
| 2 — Monthly Sales Trends | Monthly revenue bars, colour-coded by MoM direction (green = increase, red = decrease, amber = little/no change), plus a full month-over-month table |
| 3 — Product Performance | Top-10 revenue bars (one colour per rank) + a comparison against the dataset-wide average product revenue (£2,649) — explicitly labelled *revenue, not profitability* |
| 4 — Country & Customer Insights | Country revenue (top 10 of 38), leading-country cards, repeat-customer cards, top repeat-customer table |
| 5 — SQL Analysis & Data Quality | The three leading SQL results, the full data-quality audit trail (row counts), the revenue verification check, and a raw-data preview showing included / cancelled / invalid rows side by side |
| 6 — Business Recommendations | Cards that each carry **observed evidence → recommended action → business reason → KPI to monitor**; items not proven are labelled "hypothesis — needs testing" |

Design decisions worth talking about:

- **Dark navy theme, 18px root font, cyan accents** (`assets/style.css` + `.streamlit/config.toml`): presentation legibility — readable from a projector across a meeting room.
- **One shared data layer:** `data_prep.py` computes everything once; slides are pure render functions. Navigation only changes a number in `dcc.Store`, so no button click can alter a metric.
- **Colour means something:** the same green/amber/red semantics are used on slide 2 (MoM direction) and slide 3 (vs average), so the audience learns the colour language once.
- **Fail-safe month math:** the first month and any month after a zero-revenue month show "n/a" instead of a meaningless `+∞%`, because division by zero is not a valid business number.
- **Duplicates impossible by design:** `load_sales_into_db()` only inserts when the table's row count doesn't match the DataFrame; refreshing the app a hundred times cannot double the data.

---

## 11. Five Verified Insights

Every number below was computed by running this project's own code against `Online Retail.xlsx` (the same code the dashboard shows). For each insight: what I measured, why it matters to the business, and what I would recommend. **Important:** the *finding* is the measured part; the *recommendation* is a proposed course of action based on it and should be piloted before being called a success.

---

### Insight 1 — One market carries the business: the UK is 84.6% of all revenue

**Finding (measured).** Of 38 countries in the dataset, the United Kingdom generated **£9,025,222.08** — **84.6%** of the £10,666,684.54 total. The next four countries together are a fraction of that: Netherlands £285,446.34 (2.7%), EIRE £283,453.96 (2.7%), Germany £228,867.14 (2.1%), France £209,715.11 (2.0%).

**Business impact.** If the UK market hits trouble (economy, shipping costs, a competitor, Season disruption to domestic logistics), roughly **six out of every seven pounds of revenue is exposed**. That is a concentration risk the leadership team should see on page one, and international growth cannot be judged honestly while the base market dwarfs everything else.

**Recommendation (needs a pilot, not a promise).** The runner-up markets already generate real revenue with no special effort — that is proven demand. Run a small, measured promotional pilot in the two strongest non-UK markets (Netherlands and EIRE by revenue) and judge results after one quarter using monthly non-UK revenue share as the KPI. A 2.7% market doubling is worth more than it sounds in absolute pounds and de-risks the concentration.

---

### Insight 2 — Demand is seasonal: Sep–Nov brings 34.9% of the year's revenue, February is the trough

**Finding (measured).** The three months Sep–Nov 2011 produced **£3,723,065.80** — **34.9%** of total revenue in just 3 of 13 months. November 2011 is the strongest month (**£1,509,496.33**, +30.7% vs October), and the rally starts in August (+5.6%) and continues +39.4% in September and +9.1% in October. February 2011 is the weakest month (**£523,631.89**).

**Business impact.** Stock, staffing and marketing decisions made without this calendar misfire: February plans built on November levels cause overbuying and cash tied up in inventory, while the autumn build-up catches under-resourced operations flat-footed. The pattern is measured, not guessed.

**Recommendation (needs testing).** Front-load inventory and promo planning for the Sep–Nov build-up starting in summer, and mount a focused drill-down of February (which products, which customers, invoice sizes) before assuming any fixes. One caveat I always state: **December 2011 is a partial month** (the file ends 9 Dec 2011), so the −57.7% December figure on the trend chart is a data artefact, not evidence of a demand collapse — I label it that way rather than let a leader misread it.

---

### Insight 3 — Repeat customers generate 77.8% of all revenue

**Finding (measured).** **2,845** customers (of 4,338 identified customers on valid rows — 65.6%) purchased on more than one invoice. They average **6.0 orders** each and their spend is **£8,295,096.17** — **77.8% of total revenue**. (SQL query 5; rows without a CustomerID are excluded there because identity is unknowable.)

**Business impact.** The business is fundamentally a *repeat-purchase and wholesale-replenishment* business. That changes where money should go: the biggest lever is serving existing customers well (availability, service, sensible re-ordering), not just chasing new ones. It also frames risk: losing even a modest slice of repeat buyers costs far more than the same slice of one-time buyers.

**Recommendation (needs testing).** Make repeat buying frictionless for the proven base: keep top products in stock (slide 4 lists the top repeat customers by order count), and pilot a light, respectful re-order reminder for high-frequency customers over one quarter. Measure repeat-revenue share and average orders per repeat customer against a control before/after — the pilot exists to prove the effect, not to assume it.

---

### Insight 4 — About a third of identified customers never came back: 34.4% ordered only once

**Finding (measured).** **1,493** of 4,338 identified customers (**34.4%**) had exactly one valid invoice; the rest are accounted for by the 2,845 repeat customers of Insight 3 (and repeat customers hold 77.8% of revenue, so the remaining ~22% sits with everyone else).

**Business impact.** This is the mirror image of Insight 3 and the clearest growth gap in the data: a large pool of customers who tried us once. Even a small improvement in their return rate has an outsized effect — winning them back costs far less than acquiring strangers, because they already exist in the data with purchase history.

**Recommendation (needs testing).** Run a single, well-targeted follow-up test aimed at first-order customers (not a blanket discount — a measured experiment with a control group). KPI: repeat-purchase rate of the campaign cohort vs control after one quarter. I deliberately label this a **hypothesis**: the data says these customers exist and don't return, it does not say why, so the test has to include learning *why* (checkout friction, product fit, price) and not just push a discount.

---

### Insight 5 — The revenue leader isn't even a product: the "top product" is a postage line, and revenue is spread across a long tail of 4,026 SKUs

**Finding (measured).** The #1 line by revenue is **"DOTCOM POSTAGE"** (£206,248.77) — the shipping fee charged on orders, not a stock item — and **"Manual"** (£78,112.82, rank 8) is booking-side adjustments, also not merchandise. The top genuine merchandise item is **"REGENCY CAKESTAND 3 TIER"** (£174,484.74), followed by "PAPER CRAFT , LITTLE BIRDIE" (£168,469.60) and "WHITE HANGING HEART T-LIGHT HOLDER" (£106,292.77). Overall spread: Top-3 lines = **5.15%** of revenue, Top-10 lines = **10.8%**, across **4,026** distinct product descriptions.

**Business impact.** Two things leaders would otherwise miss. (1) "Top product by revenue" lists made carelessly are contaminated by service lines — a product decision like "discontinue the bottom 90% of SKUs" or "spotlight the No.1 product" would be made on the wrong item. (2) The long tail is the *real* product strategy: no single SKU dominates, so assortment breadth and availability of the strongest décor lines matter more than one hero product.

**Recommendation (needs testing).** Treat top merchandise lines (e.g. REGENCY CAKESTAND, WHITE HANGING HEART) as availability-critical and plan their stock like A-items; review how postage is charged/passed through and what "Manual" adjustments represent before making any judgments from them; and test whether fixtures/bundles built around top décor SKUs lift order value in a controlled pilot before rolling out.

---

## 12. Actual Findings vs. Recommendations That Need Testing

This distinction is the most important one in any analyst's review — I keep it everywhere, including on slide 6.

**✅ Actual findings (measured, recomputable, shown on the dashboard):**

- 84.6% UK revenue concentration and the country table behind it
- The monthly revenue curve, the Sep–Nov 34.9% seasonal share, Nov 2011 peak, Feb 2011 trough
- Repeat customers = 2,845 people, 6.0 avg orders, 77.8% of revenue
- 34.4% of identified customers ordered exactly once
- The full data-quality audit: 541,909 rows → 11,805 excluded → 530,104 kept, revenue verified to £0.0000 difference
- The "top product" nuance (postage + manual lines outrank real merchandise)

**🔬 Recommendations / hypotheses (proposals built on the findings — not yet proven):**

- Promotional pilots in Netherlands/EIRE to reduce UK dependence
- Front-loading autumn inventory *before* the wishlist that might change this year
- A repeat-purchase nudge for high-frequency customers and a win-back test for one-time buyers
- Bundling top décor SKUs to raise average order value
- Stock-covering top merchandise SKUs like A-items

**Why the split matters:** findings are facts that survive audit; recommendations are arguments about the future that need a measured experiment (with a control group and a named KPI) before they can be declared successful. The dashboard itself carries this on slide 6 — evidence-backed cards vs "hypothesis — needs testing" cards — and slide 6 explicitly states that the dataset has no cost data, so no profit or margin claims are made anywhere.

---

## 13. 15 Interview Questions & Simple Answers

**1) What is this project, in one minute?**
I took a real retail Excel file with 541,909 rows of sales, cleaned it with Pandas, computed the key sales KPIs, loaded the clean data into a SQLite database and wrote five SQL queries, then presented everything as a six-slide executive dashboard built with Plotly Dash. The result explains where revenue comes from (84.6% UK), when it comes (autumn), who it comes from (repeat customers, 77.8% of revenue), and the data-quality story behind every number.

**2) Walk me through how you built it end to end.**
First I explored the Excel file and found the quality problems (cancelled invoices, negative quantities, ~25% missing customer IDs). Then I wrote cleaning rules in Pandas that *flag* bad rows instead of deleting them. Then I computed KPIs like Total Revenue (£10.67M) and Average Order Value (£534.40). Then I loaded only the valid sales rows into SQLite and wrote five commented SQL queries. Finally I built the Dash slide deck on top, with one shared data module so both dashboards show identical numbers.

**3) How exactly did you clean the data?**
Three flags: (1) invoices starting with "C" are cancelled credit notes — 9,288 rows; (2) rows with Quantity ≤ 0 or UnitPrice ≤ 0 are not real sales; (3) missing Descriptions became "Unknown" and missing CustomerIDs were kept for revenue but excluded from repeat-customer analysis. Nothing was deleted: 11,805 rows were excluded from KPIs and counted publicly on the data-quality slide, leaving 530,104 valid rows.

**4) Why not just delete the bad rows?**
Because "bad" is a business judgment, and deleting hides it. Flagging keeps the audit trail — anyone can see exactly how much data was excluded and why. If a stakeholder disagrees with my rule, they can recompute with their own rule and compare. A silent `DROP` makes that impossible.

**5) How did you calculate Total Revenue and Average Order Value?**
Revenue per row is Quantity × UnitPrice, summed over valid rows = £10,666,684.54. Valid Orders = COUNT(DISTINCT InvoiceNo) = 19,960. AOV = Revenue ÷ Orders = £534.40. Distinct matters: one order with five product lines must count once, not five times. I also re-verified revenue a second, independent way — the two match to £0.0000, and that check is shown on screen.

**6) Why did you use SQL and SQLite at all? Pandas could do the same.**
Two reasons. First, SQL is the language of analytics interviews and most companies' warehouses — demonstrating `GROUP BY`, `COUNT(DISTINCT)`, `HAVING` and `LIMIT` on a real database is the point. Second, SQLite makes the project realistic: the cleaned data sits in an actual database file (`retail_sales.db`), pandas and SQL produce the same numbers side by side, and the dashboard literally runs the SQL file it shows on screen.

**7) Explain one of your SQL queries and why it's written that way.**
The repeat-customers query. `WHERE CustomerID IS NOT NULL` filters before grouping (unidentifiable customers can't be customers); `GROUP BY CustomerID` makes one row per customer; `COUNT(DISTINCT InvoiceNo) > 1` in `HAVING` keeps only those with more than one order — that's the definition of "repeat" in billing data. This one found 2,845 repeat customers generating 77.8% of revenue.

**8) What's the difference between WHERE and HAVING?**
`WHERE` filters individual rows before grouping; `HAVING` filters groups after aggregation. In query 5, `WHERE` drops rows with no CustomerID; `HAVING` keeps only customers with more than one invoice.

**9) What does `@lru_cache` do in your project and why did you add it?**
It makes `prepare_data()` run only once per process. Without it, every slide navigation would re-read a 23MB Excel file and re-run the SQL — the app would feel broken. With it, the first load takes a minute and every click afterwards is instant. The trade-off is stale data within a session, which is acceptable because the session is a presentation.

**10) Why Dash? How is it different from Streamlit (your first version)?**
Streamlit reruns the whole script top-to-bottom on every interaction — great for quick dashboards. Dash is component-and-callback based: my app holds its slide number in `dcc.Store`, and clicking Next only rebuilds the visible slide. That state control plus a literal slide-deck layout (progress bar, Previous/Next/disable logic) fits a presentation better. I kept the original Streamlit app untouched as backup.

**11) How do the callbacks in your Dash app work?**
Three callbacks: one reads the slide number from `dcc.Store` and renders the slide body plus progress bar; one enables/disables the nav buttons (Previous disabled on slide 1, Next disabled on slide 6); one listens to button clicks and writes the new slide number into the store. Navigation is presentation-only state, so no click can ever change a metric.

**12) What was the biggest data-quality surprise?**
That "top product by revenue" was contaminated by non-products: the #1 line, "DOTCOM POSTAGE" (£206,248.77), is a shipping fee, and "Manual" (£78,112.82) is an adjustments label. Also, 135,080 rows (~25%) have no CustomerID. Both changed how I present results — I name postage/manual explicitly and restrict repeat-customer claims to identified customers.

**13) Which insight are you most proud of, and how would you validate the follow-up action?**
That repeat customers generate 77.8% of revenue from 65.6% of identified customers. It reframes the business as repeat-purchase-driven. To validate the nudge idea I'd run a control-grouped pilot for one quarter and measure repeat-revenue share and orders per repeat customer against control — proving effect, not assuming it.

**14) What are the limitations of this dataset and analysis?**
No cost data, so no profit or margin claims anywhere. December 2011 is partial (file ends 9 Dec), so the last MoM number is an artefact I label as such. A quarter of rows lack customer identity, so customer analysis excludes them (silently dropping them would inflate revenue analysis instead). And recommendations are hypotheses to be tested — the data shows who exists, not why they buy.

**15) If revenue suddenly dropped 20% next month, how would you investigate?**
Same structure every time: first verify the drop is real (data quality — nothing loaded twice, nothing missing); then decompose it in the KPI's own terms — by month-over-month bridge, product mix, country mix, and repeat vs one-time customers; compare against the seasonal baseline (February, not November). I'd use exactly the queries on slides 2–5, then take the leading hypothesis to a controlled check rather than an opinion.

---

## 14. Honest Limitations

Stating limitations is a strength in interviews — here are mine, plainly:

1. **No profit numbers, by choice.** The dataset has no unit-cost column. Every chart shows revenue; slide 3's caption explicitly says "revenue, NOT profitability." I never say "top profit maker."
2. **December 2011 is a partial month.** The data ends 9 Dec 2011, so the last bar's −57.7% is partly calendar, not demand. I call it out on the README and slide 6 evidence rather than let it mislead.
3. **~25% of rows lack CustomerID** (135,080). Revenue includes them; customer-level KPIs correctly exclude them. Both numbers are labelled where they appear.
4. **Cancelled orders are excluded from "sales" KPIs** — correct for revenue analysis, but a full audit-log view would also track returns as their own storyline (a natural next version).
5. **No forecasting, segmentation or ML.** This project is cleaning + KPIs + SQL + presentation. If asked, the next extensions I'd propose are: a completed-December adjustment note, RFM-style segmentation on identified customers, and margin analysis once cost data exists.
6. **Two backup files and one legacy app exist** (`app.py`, `app_backup_*.py`) — intentional: the Streamlit original is preserved untouched, with `dash_app.py` as the polished deliverable.

---

*Numbers in this README were produced by running this repository's own code (`data_prep.prepare_data()` + `sql_queries.sql` against `retail_sales.db`) on the local `Online Retail.xlsx`. Recompute anything: `python -c "import data_prep as dp; D = dp.prepare_data(); print(D.total_revenue, D.avg_order_value, D.uk_share)"`.*
