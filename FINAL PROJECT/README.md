<div align="center">

# -- ! Data Intelligence Dashboard ! --
### *8-Sheet Excel Workbook Covering Customer, Product, Region & Monthly Analysis, Regression, Pivot Tables, Interactive Dashboard & Scenario Planning*

[![Excel](https://img.shields.io/badge/Microsoft%20Excel-365%20%2F%202021-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Sheets](https://img.shields.io/badge/Worksheets-8-FF6F00?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Analysis](https://img.shields.io/badge/Analysis-Customer%20%7C%20Product%20%7C%20Region-4CAF50?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Dashboard](https://img.shields.io/badge/Dashboard-KPIs%20%7C%20Slicers%20%7C%20Charts-9C27B0?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)

<br/>

> *"Data tells a story only when someone takes the time to ask it the right questions."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Project Workflow](#-project-workflow)
- [📑 Sheet-by-Sheet Guide](#-sheet-by-sheet-guide)
  - [📥 Raw Data](#-1-raw-data--raw-data)
  - [📊 Analysis](#-2-analysis--analysis)
  - [📈 Regression](#-3-regression--regression)
  - [🎨 Visualizations](#-4-visualizations--visualizations)
  - [🧭 Dashboard](#-5-dashboard--dashboard)
  - [🧪 What-If](#-6-what-if--what-if)
  - [🎲 Scenario Summary 2](#-7-scenario-summary-2--scenario-summary-2)
  - [📝 Final Report](#-8-final-report--final_report)
- [🛠️ Tech Stack](#️-tech-stack)
- [📈 Sample Results](#-sample-results)
- [💡 Key Business Insights](#-key-business-insights)
- [⚠️ Known Issues & Notes](#️-known-issues--notes)
- [🏆 Advantages](#-advantages)
- [🚀 How to Use](#-how-to-use)
- [📄 License](#-license)
- [👤 Author](#-author)
- [🙏 Acknowledgements](#-acknowledgements)

---

## 📌 Overview

The **Data Intelligence Dashboard** is an intermediate-level workbook (`Data_Intelligence_Dashboard.xlsx`) made of **8 worksheets** that analyse **250 customer transactions** from **50 customers**, across **5 regions**, **3 product categories** and **10 products**, between **11 April 2024 and 11 February 2025**.

Starting from one clean `SalesData` table, the workbook builds customer segmentation, product rankings, regional and monthly performance, a regression model, three pivot tables with a Region slicer and Date timeline, an interactive dashboard, and a revenue-target scenario model.

This project is designed to:
- Organise raw transactions inside a structured **Excel Table** (`SalesData`)
- Calculate headline KPIs: **Total Revenue, Transactions, Units Sold, Unique Customers**
- Classify every customer as **High Value** or **Regular** using an `IF` rule
- Rank **products** and **regions** using `SUMIF`, `COUNTIF` and `TEXTJOIN + FILTER`
- Track monthly revenue and **growth %** using `EOMONTH`, `SUMIF` and `IFERROR`
- Measure the Quantity → Total Amount relationship using **Linear Regression**
- Summarise data with **Pivot Tables**, a **Region Slicer** and a **Date Timeline**
- Present KPIs and charts in a single **Dashboard** with a target-achievement bar
- Plan growth with a **What-If model** and **Scenario Manager** (Conservative / Expected / Optimistic)

---

## 🎯 Problem Statement

> **Objective:** Analyse customer transaction data to find out *how much we earn, who our most valuable customers are, which products and regions lead, how revenue moves month to month,* and *how much growth is needed to reach a 300,000 revenue target.*

You are creating a complete transaction-analysis report for a business manager. Each worksheet must answer one business question using a specific Excel technique, and the final dashboard and report must bring the key results together.

| 📂 Category | 📄 Sheets | 🔍 Description |
|-------------|-----------|----------------|
| Raw Data | Raw Data | 250 transactions stored as `SalesData` with tenure, month-end and product-code helper columns |
| Core Analysis | Analysis | KPIs, customer value, product, region and monthly tables |
| Predictive Analysis | Regression | Does a higher Quantity lead to a higher Total Amount? |
| Visual Summary | Visualizations, Dashboard | Pivot tables, slicer, timeline, charts and KPI cards |
| Scenario Planning | What-If, Scenario Summary 2 | Revenue projection and three growth scenarios |
| Reporting | Final_Report | Objective, key metrics, insights and recommendations |

The goal is to demonstrate **end-to-end business data analysis skills** — from raw data to decision-ready insights — using only Excel.

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 🗄️ **Structured Table** | `SalesData` makes every formula auto-readable, e.g. `SalesData[Total_Amount]` |
| 📊 **KPI Cards** | Total Revenue **229,192.47**, **250** transactions, **753** units, **50** customers |
| 🏅 **Customer Segmentation** | Customers with total purchases ≥ **5,000** are flagged *High Value* with icon-set formatting |
| 🔍 **Dynamic Arrays** | `SORT`, `UNIQUE`, `FILTER`, `XLOOKUP`, `XMATCH` and `TEXTJOIN` find top customer, top product and top region automatically |
| 📅 **Monthly Trend** | `EOMONTH` + `SUMIF` build month-wise revenue; icon sets show growth or decline |
| 📈 **Regression Model** | R² = **0.2098**, slope = **353.79**, with a scatter plot and fitted line |
| 🧮 **Pivot Tables** | Revenue by Category, Month and Region, with a **Region Slicer** and **Date Timeline** |
| 🎯 **Target Tracking** | Data-bar shows **76.40%** achievement against a 300,000 target |
| 🧪 **What-If Model** | Finds the growth rate (**30.89%**) required to reach the revenue target |
| 🎲 **Scenario Manager** | Conservative (5%), Expected (10%) and Optimistic (20%) growth in one summary |

---

## 🏗️ Project Structure

```
📦 data-intelligence-dashboard/
│
├── 📄 Data_Intelligence_Dashboard.xlsx   ← Main workbook (8 worksheets)
│     ├── Dashboard                       ← KPI cards, 3 charts, target bar
│     ├── Final_Report                    ← Objective, metrics, insights, recommendations
│     ├── Analysis                        ← KPIs, customer / product / region / month tables
│     ├── Visualizations                  ← 3 pivot tables, 2 charts, slicer & timeline
│     ├── What-If                         ← Revenue growth & target model
│     ├── Scenario Summary 2              ← Scenario Manager report
│     ├── Raw Data                        ← Source data (SalesData table)
│     └── Regression                      ← Quantity vs Total_Amount regression
│
└── 📄 README.md                          ← Project documentation
```

---

## 🔄 Project Workflow

```
Open Workbook
      │
      ▼
┌─────────────────────────────┐
│   Raw Data                  │  ← Raw Data  (SalesData, 250 rows)
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Calculate & Segment       │  ← Analysis (KPIs, High Value Customers,
│                             │     Products, Regions, Monthly Growth)
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Find Relationships        │  ← Regression (Quantity vs Total_Amount)
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Summarise & Visualise     │  ← Visualizations (Pivots, Slicer, Timeline)
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Report Results            │  ← Dashboard + Final_Report
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Plan Scenarios            │  ← What-If + Scenario Summary 2
└────────────┬────────────────┘
             │
             ▼
      Analysis Complete ✅
```

---

## 📑 Sheet-by-Sheet Guide

### 📥 1. Raw Data — `Raw Data`

> The foundation of the project: **250 transactions** stored as an Excel Table named **`SalesData`** (columns `A` to `Q`).

| Column | Header | Description |
|--------|--------|-------------|
| A | Transaction_ID | Unique transaction code (`TRX0001` – `TRX0250`) |
| B | Date | Order date (11-Apr-2024 → 11-Feb-2025) |
| C | Customer_ID | Customer code (`CUST001` – `CUST050`) |
| D | Customer_Name | Name of the customer |
| E | Product_ID | Product code (`P001` – `P010`) |
| F | Product_Name | Blender / Bookshelf / Coffee Maker / Desk / Headphones / Keyboard / Laptop / Monitor / Office Chair / Smartphone |
| G | Category | Appliances / Electronics / Furniture |
| H | Quantity | Units sold in the transaction |
| I | Unit_Price | Price per unit (59.99 – 899.99) |
| J | Payment_Method | Cash / Credit Card / Debit Card / PayPal |
| K | Region | Central / East / North / South / West |
| L | Customer_Segment | Basic / Standard / Premium |
| M | Customer_Since | Date the customer joined |
| N | Total_Amount | Transaction value (`Quantity × Unit_Price`) |
| O | Customer_Tenure_Months | Months since the customer joined |
| P | Month_End | Last day of the transaction month |
| Q | Product_Abbreviation | First 3 letters of the product in capitals (e.g. `LAP`) |

**Calculated columns:**
```excel
Customer_Tenure_Months  =DATEDIF(M2,TODAY(),"m")
Month_End               =EOMONTH(B2,0)
Product_Abbreviation    =UPPER(LEFT(F2,3))
```

---

### 📊 2. Analysis — `Analysis`

> The calculation engine of the workbook. Every KPI, table and ranking used by the Dashboard and Final_Report is built here.

**A. Core KPIs (`A1:B13`)**

| Cell | Formula | Result |
|------|---------|--------|
| `B1` Total Revenue | `=SUM(SalesData[Total_Amount])` | **229,192.47** |
| `B2` Total Transactions | `=COUNTA(SalesData[Transaction_ID])` | **250** |
| `B3` Total Units Sold | `=SUM(SalesData[Quantity])` | **753** |
| `B4` Average Transaction Value | `=AVERAGE(SalesData[Total_Amount])` | **916.77** |
| `B5` Unique Customers | `=COUNTA(UNIQUE(SalesData[Customer_ID]))` | **50** |
| `B7:B9` Date helpers | `=TODAY()` · `=NOW()` · `=EOMONTH(B7,0)` | Live date values |
| `B11` High Value Customers | `=COUNTIF(F2:F51,"High Value")` | **20** |
| `B12` High Value Revenue | `=SUMIF(F2:F51,"High Value",E2:E51)` | **152,555.88** |
| `B13` High Value Revenue Share | `=B12/B1` | **66.56%** |

**B. Customer value table (`D1:F51`)**

| Column | Formula | Purpose |
|--------|---------|---------|
| `E` Customer_Total_Revenue | `=SUMIF(SalesData[Customer_ID],D2,SalesData[Total_Amount])` | Total purchase per customer |
| `F` Customer_Value | `=IF(E2>=5000,"High Value","Regular")` | Segment label (threshold = 5,000) |
| `H2` High Value list | `=FILTER(D2:E51,E2:E51>=5000,"No High Value Customers")` | Spills all high-value customers |
| `E53` Top Customer ID | `=TEXTJOIN(", ",TRUE,FILTER(D2:D51,E2:E51=MAX(E2:E51)))` | **CUST025** |
| `E54` Highest Customer Revenue | `=MAX(E2:E51)` | **15,659.65** |
| `E55` Top Customer Name | `=XLOOKUP(E53,SalesData[Customer_ID],SalesData[Customer_Name],"ID not found")` | **Mark Carter** |

**C. Product table (`K1:N11`)**

| Product | Purchase Count | Units Sold | Product Revenue |
|---------|----------------|------------|-----------------|
| Blender | 29 | 87 | 5,219.13 |
| **Bookshelf** | **35** | **102** | 15,298.98 |
| Coffee Maker | 14 | 39 | 3,119.61 |
| Desk | 26 | 78 | 23,399.22 |
| Headphones | 16 | 47 | 7,049.53 |
| Keyboard | 34 | 89 | 8,009.11 |
| **Laptop** | 24 | 75 | **67,499.25** |
| Monitor | 26 | 88 | 21,999.12 |
| Office Chair | 15 | 52 | 10,399.48 |
| Smartphone | 31 | 96 | 67,199.04 |

```excel
=COUNTIF(SalesData[Product_Name],K2)                    ← Purchase_Count
=SUMIF(SalesData[Product_Name],K2,SalesData[Quantity])  ← Units_Sold
=SUMIF(SalesData[Product_Name],K2,SalesData[Total_Amount]) ← Product_Revenue
```
> Product highlights: **Most Frequently Purchased = Bookshelf (35)**, **Most Sold = Bookshelf (102 units)**, **Top Revenue = Laptop (67,499.25)**. Data-bar conditional formatting is applied to the product summary.

**D. Region table (`P1:Q9`)**

| Region | Regional Revenue |
|--------|------------------|
| Central | 41,288.34 |
| **East** | **59,288.39** |
| North | 50,808.31 |
| South | 36,398.75 |
| West | 41,408.68 |

> `Q8` returns the **Top Sales Region** (East) with `TEXTJOIN + FILTER`, and `Q9 =MAX(Q2:Q6)` returns its revenue.

**E. Monthly table (`S1:U14`)**

| Column | Formula | Purpose |
|--------|---------|---------|
| `S` Month_End | `=SORT(UNIQUE(SalesData[Month_End]))` | Spills each unique month-end |
| `T` Monthly_Revenue | `=SUMIF(SalesData[Month_End],S2,SalesData[Total_Amount])` | Revenue in that month |
| `U` Monthly_Growth | `=IFERROR((T3-T2)/T2,0)` | Change vs. previous month |

| Month | Revenue | Growth % |
|-------|---------|----------|
| Apr 2024 | 20,919.33 | — |
| May 2024 | 17,039.35 | -18.5% |
| Jun 2024 | 19,299.45 | +13.3% |
| Jul 2024 | 18,479.40 | -4.2% |
| Aug 2024 | 15,759.43 | -14.7% |
| Sep 2024 | 12,989.51 | -17.6% |
| **Oct 2024** | **27,009.13** | **+107.9%** |
| Nov 2024 | 28,049.14 | +3.9% |
| Dec 2024 | 29,009.00 | +3.4% |
| **Jan 2025** | **32,259.04** | **+11.2%** |
| Feb 2025 | 8,379.69 | -74.0% |

> An **icon-set conditional format** on `U3:U14` shows arrows for rising and falling months. `T15` holds the **Latest Transaction Date** (`=MAX(SalesData[Date])`) and the **Best Sales Month** is **Jan 2025** (32,259.04).

**F. Regression summary, target and status (`A15:B24`)**

| Cell | Content | Value |
|------|---------|-------|
| `B17` | Regression R Square | 0.2098 |
| `B18` | Interpretation | *Quantity explains approximately 20.98% of the variation in Total Amount.* |
| `B20` | Report Status | Completed |
| `B21` | Report Timestamp | `=IF(B20="","",IF(ISNUMBER(B21),IF(B21>0,B21,NOW()),NOW()))` — a self-referencing formula that stamps the time once the status is filled |
| `B23` | Revenue Target | 300,000 |
| `B24` | Target Achievement | `=B1/B23` → **76.40%** |

**G. Name matching (`W1:Y30`)**

> Compares **Customer List 1** (all 50 names via `SORT(UNIQUE())`) with **Customer List 2** (customers who bought in the *East* region) using `FILTER` and `XMATCH`, returning the **Matching Names** — 29 unique customers have East-region purchases.

---

### 📈 3. Regression — `Regression`

> Simple linear regression: **Total_Amount (Y)** predicted from **Quantity (X)** across 250 observations, with a *Quantity Line Fit Plot* scatter chart and a full residual table.

| Regression Statistic | Value |
|----------------------|-------|
| Multiple R | 0.4581 |
| R Square | 0.2098 |
| Adjusted R Square | 0.2067 |
| Standard Error | 919.61 |
| Observations | 250 |
| F-statistic | 65.86 |
| Significance F | 2.26 × 10⁻¹⁴ |

| Term | Coefficient | P-value | 95% Confidence Interval |
|------|-------------|---------|-------------------------|
| Intercept | -148.85 | 0.3010 | -431.71 to 134.00 |
| Quantity | **353.79** | 2.26 × 10⁻¹⁴ | 267.93 to 439.66 |

**Regression equation:**
```
Predicted Total_Amount = -148.85 + (353.79 × Quantity)
```

**Interpretation:**
- Quantity explains approximately **20.98%** of the variation in Total Amount.
- The relationship is **statistically significant** (Significance F ≈ 0), but the model is weak.
- Each **extra unit** is associated with about **353.79** more in transaction value.
- Quantity alone cannot predict revenue accurately because **unit prices vary widely** (59.99 to 899.99).
- The sheet also includes the **Residual Output** table (Observation, Predicted Total_Amount, Residuals) for all 250 rows.

---

### 🎨 4. Visualizations — `Visualizations`

> Three **Pivot Tables** (Sum of Total_Amount) that feed two charts, controlled by a **Region Slicer** and a **Date Timeline**.

| Pivot | Rows | Result |
|-------|------|--------|
| Pivot 1 | Category | Appliances 8,338.74 · Electronics 171,756.05 · Furniture 49,097.68 |
| Pivot 2 | Month | 11 months, Apr 2024 → Feb 2025 (Grand Total 229,192.47) |
| Pivot 3 | Region | Central 41,288.34 · East 59,288.39 · North 50,808.31 · South 36,398.75 · West 41,408.68 |

| # | Chart | Type | Shows |
|---|-------|------|-------|
| 1 | Sales by Category | Bar | Revenue per product category |
| 2 | Sales by Region | Pie | Regional share of revenue |

> The **Region slicer** and **Date timeline** let the viewer filter the pivots interactively.

---

### 🧭 5. Dashboard — `Dashboard`

> A one-page executive summary: KPI cards, three charts and a target-achievement bar.

**KPI cards:**

| KPI | Formula | Value |
|-----|---------|-------|
| Total Revenue | `=Analysis!B1` | **229,192.47** |
| Total Transactions | `=Analysis!B2` | **250** |
| Total Units Sold | `=Analysis!B3` | **753** |
| Unique Customers | `=Analysis!B5` | **50** |
| Revenue Target Achievement | `=Analysis!B1/300000` | **76.40%** (shown as a data bar) |

**Charts on the dashboard:**

| # | Chart | Type | Shows |
|---|-------|------|-------|
| 1 | Sales by Category | Bar | Category comparison |
| 2 | Monthly Sales Trend | Line | Revenue over 11 months |
| 3 | Sales by Region | Pie | Regional contribution |

---

### 🧪 6. What-If — `What-If`

> **Sales Scenario Analysis:** a simple growth model that answers *"how much must revenue grow to hit the target?"*

| Cell | Label | Formula / Value |
|------|-------|-----------------|
| `B3` | Base Revenue | `=Analysis!B1` → 229,192.47 |
| `B4` | Expected Growth Rate | `0.3089` *(input cell)* |
| `B5` | Projected Revenue | `=B3*(1+B4)` → **300,000.00** |
| `B8` | Target Revenue | `300000` |
| `B9` | Required Growth Rate | `=B8/B3-1` → **30.89%** |

> Revenue must grow by about **30.89%** (≈ **70,807.53** more) to reach the 300,000 target.

---

### 🎲 7. Scenario Summary 2 — `Scenario Summary 2`

> Output of Excel's **Scenario Manager** — three growth scenarios applied to the changing cell `What-If!B4`, with the result cell `What-If!B5` (Projected Revenue).

| Scenario | Growth Rate | Projected Revenue | Gain vs. Base |
|----------|-------------|-------------------|---------------|
| Conservative Growth | 5% | 240,652.09 | +11,459.62 |
| Expected Growth | 10% | 252,111.72 | +22,919.25 |
| Optimistic Growth | 20% | 275,030.96 | +45,838.49 |
| *Required for target* | *30.89%* | *300,000.00* | *+70,807.53* |

> Even the **Optimistic** scenario (20%) falls short of the 300,000 target by about **24,969**.

---

### 📝 8. Final Report — `Final_Report`

> A written management report that links to live workbook values.

| Section | Content |
|---------|---------|
| **Project Objective** | Analyse customer transactions to find sales trends, high-value customers, leading products and regional performance |
| **Key Metrics** | Total Revenue 229,192.47 · 250 transactions · 753 units · 50 customers (linked to `Analysis`) |
| **Key Insights** | High-value customers, regional and monthly peaks, top customer, regression result, scenario results |
| **Recommendations** | Retain high-value customers, keep **Bookshelf** stocked, promote **Laptop** sales, and investigate East's strength and January's peak |

---

## 🛠️ Tech Stack

| Tool | Version | Purpose |
|------|---------|---------|
| 📗 **Microsoft Excel** | 365 / 2021 or later | Workbook and formula engine |
| 🗄️ **Excel Tables** | Built-in | Structured references (`SalesData[Total_Amount]`) |
| 📊 **SUM / COUNTA / AVERAGE** | Built-in | Headline KPIs |
| 🎯 **SUMIF / COUNTIF** | Built-in | Customer, product, region and month totals |
| 🔀 **IF / IFERROR** | Built-in | Customer segmentation and safe growth % |
| 📅 **DATEDIF / EOMONTH / TODAY / NOW** | Built-in | Tenure, month-end and live timestamps |
| ✨ **SORT / UNIQUE / FILTER** | Excel 365 / 2021 | Dynamic array outputs |
| 🔍 **XLOOKUP / XMATCH / TEXTJOIN** | Excel 365 / 2021 | Top customer, top product and name matching |
| 🧮 **PivotTable, Slicer & Timeline** | Built-in | Category / Month / Region summaries with filters |
| 📐 **Data Analysis ToolPak** | Add-in | Regression output |
| 🚦 **Conditional Formatting** | Built-in | Icon sets and data bars |
| 🧪 **What-If Analysis (Scenario Manager)** | Built-in | Growth scenarios |
| 📈 **Charts** | Built-in | Bar, Line, Pie, Scatter |

---

## 📈 Sample Results

Key outputs captured from the workbook:

```
Raw Data                Records                       →  250 transactions | 50 customers
Analysis                Total Revenue                 →  229,192.47
Analysis                Average Transaction Value     →  916.77
Analysis                High Value Customers          →  20 of 50 (66.56% of revenue)
Analysis                Top Customer                  →  CUST025 Mark Carter : 15,659.65
Analysis                Top Revenue Product           →  Laptop : 67,499.25
Analysis                Most Frequently Purchased     →  Bookshelf : 35 purchases, 102 units
Analysis                Top Region                    →  East : 59,288.39
Analysis                Lowest Region                 →  South : 36,398.75
Analysis                Best Month                    →  Jan 2025 : 32,259.04
Analysis                Target Achievement            →  76.40% of 300,000
Regression              R Square                      →  0.2098 (20.98%)
Regression              Equation                      →  Total_Amount = -148.85 + 353.79 × Quantity
What-If                 Required Growth for Target    →  30.89%
Scenario Summary 2      5% / 10% / 20% growth         →  240,652.09 / 252,111.72 / 275,030.96
```

---

## 💡 Key Business Insights

| # | Insight | Evidence |
|---|---------|----------|
| 1 | **Revenue is 76% of the way to target** | 229,192.47 earned against a 300,000 target; 70,807.53 still needed |
| 2 | **Electronics dominates revenue** | 171,756.05, about **74.9%** of all sales, versus Furniture 21.4% and Appliances 3.6% |
| 3 | **Two products drive almost 60% of revenue** | Laptop (67,499.25) and Smartphone (67,199.04) together contribute **58.8%** |
| 4 | **Bookshelf wins on volume, not revenue** | Highest purchase count (35) and units (102), but only 15,298.98 in revenue |
| 5 | **East is the strongest region** | 59,288.39 (25.9%), followed by North (50,808.31); South is lowest at 36,398.75 |
| 6 | **A few customers carry the business** | 20 of 50 customers (40%) produce **66.56%** of revenue |
| 7 | **Top customer is a retention priority** | Mark Carter (CUST025) spends 15,659.65, about 6.8% of all revenue |
| 8 | **Revenue peaks in winter** | Oct 2024 jumped +107.9%, and revenue stayed high through Jan 2025 (32,259.04) |
| 9 | **Quantity alone is a weak predictor** | R² = 0.21 — price mix matters more than units sold |
| 10 | **Target needs ambitious growth** | 30.89% growth is required; even the 20% optimistic scenario reaches only 275,030.96 |

---

## ⚠️ Known Issues & Notes

| # | Note | Detail |
|---|------|--------|
| 1 | **Excel version** | `SORT`, `UNIQUE`, `FILTER`, `XLOOKUP`, `XMATCH` and `TEXTJOIN` need Excel 365 or 2021+; older versions show `#NAME?` |
| 2 | **Partial months** | April 2024 starts on the 11th and February 2025 ends on the 11th, so the **-74.0%** growth in Feb 2025 reflects an incomplete month, not a real collapse |
| 3 | **Extra month rows** | The monthly table's formula range extends past Feb 2025; the empty rows show `0` revenue and `-100%` growth and can be ignored or trimmed |
| 4 | **Volatile functions** | `TODAY()`, `NOW()` and `DATEDIF` (Customer_Tenure_Months) recalculate every time the file opens |
| 5 | **Hardcoded target** | The Dashboard uses `=Analysis!B1/300000`; changing the target in `Analysis!B23` does not update the Dashboard bar unless the formula is re-linked |
| 6 | **Fixed ranges** | Customer formulas use `D2:E51` / `F2:F51` (50 customers); adding more customers requires extending these ranges |
| 7 | **Growth rate is typed in** | `What-If!B4` holds `0.3089` as a value, so re-run **Goal Seek** if the base revenue changes |
| 8 | **Scenario snapshot** | `Scenario Summary 2` is a static report; rerun **Scenario Manager** after changing base revenue |
| 9 | **Self-referencing timestamp** | `Analysis!B21` refers to itself, so Excel's *iterative calculation* must be enabled for it to work as intended |
| 10 | **Sheet name spacing** | The sheet is named `What-If ` with a trailing space, so use quotes when referencing it in formulas |
| 11 | **Static pivots** | Pivot tables must be refreshed (**Data → Refresh All**) after the source data changes |
| 12 | **Spill ranges** | Dynamic array formulas need empty cells below them, otherwise they return `#SPILL!` |
| 13 | **Weak regression** | R² = 0.21 means Quantity explains only part of revenue; adding *Unit_Price* would improve the model |
| 14 | **Currency** | The data has no currency symbol, so values are shown in the dataset's own units |

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🎓 **Complete Workflow** | Covers data → analysis → prediction → reporting → scenario planning |
| 🗄️ **Table-Driven** | `SalesData` keeps formulas readable and auto-expanding for new rows |
| 📊 **Visual Storytelling** | Charts for category, region, monthly trend and regression fit |
| 🔍 **Easy to Trace** | Every Dashboard and Final_Report figure links back to a visible formula in `Analysis` |
| 🎛️ **Interactive** | Region slicer and Date timeline filter the pivot tables instantly |
| 🎯 **Target-Focused** | Clearly shows the gap between current revenue and the 300,000 goal |
| 🧪 **Decision Support** | Scenario Manager compares conservative, expected and optimistic growth |
| 💻 **No Macros** | Pure formulas, pivot tables and charts — no VBA required |
| ⚡ **Lightweight** | A single `.xlsx` file that opens instantly |

---

## 🚀 How to Use

1. Download or clone the repository.
2. Open `Data_Intelligence_Dashboard.xlsx` in **Microsoft Excel 365 / 2021 or later**.
3. Start with the **Dashboard** for the big picture, then read **Final_Report** for the written summary.
4. Open **Analysis** and click any formula cell to read it in the **formula bar**.
5. On **Visualizations**, click a region in the **Slicer** or drag the **Timeline** to filter the pivots.
6. Change the **High Value threshold** (`5000`) inside the `IF` formula in `Analysis!F2:F51` to re-segment customers.
7. Change the **Expected Growth Rate** in `What-If!B4`, or open **Data → What-If Analysis → Scenario Manager** to switch scenarios.
8. After editing source data, use **Data → Refresh All** to update the pivot tables and charts.

---

## 📄 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for full details.

```
MIT License — Free to use, modify, and distribute with attribution.
```

---

## 👤 Author

<div align="center">

### Aastha Chodvadiya

[![GitHub](https://img.shields.io/badge/GitHub-yourhandle-181717?style=for-the-badge&logo=github&logoColor=white)](https://github.com/24se02ml135-boop/python-project)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/chodvadiya-aastha-8b79a4418/)
> *"Every formula you master turns hours of manual work into a single cell."*

**🎓 Role:** Python Learner | Data & Spreadsheet Enthusiast \
**📍 Location:** India \
**🛠️ Skills:** Excel Formulas · Data Analysis · Regression · Dashboards · Python

</div>

---

## 🙏 Acknowledgements

Special thanks to the following resources and communities that made this project possible:

- 📗 [Microsoft Excel Support](https://support.microsoft.com/excel) — Official function reference
- 🧮 [PivotTable Guide](https://support.microsoft.com/en-us/office/create-a-pivottable-to-analyze-worksheet-data-a9a84538-bfe9-40a9-a8e9-f99134456576) — Pivot table reference
- 🧪 [FILTER Function Guide](https://support.microsoft.com/en-us/office/filter-function-f4f7cb66-82eb-4767-8f7c-4877ad80c759) — Dynamic array reference
- 🎓 [ExcelJet](https://exceljet.net/) — Clear formula explanations and examples
- 🖥️ [W3Schools Excel](https://www.w3schools.com/excel/) — Beginner Excel reference
- 💬 [Stack Overflow Community](https://stackoverflow.com/) — Problem-solving support

---

<div align="center">

---

*Made with ❤️ and ☕ — Last updated: October 2026*

</div>
