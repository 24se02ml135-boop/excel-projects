<div align="center">

# -- ! Sales Analysis Project ! --
### *10-Sheet Excel Workbook Covering Descriptive Statistics, Regression, Pivot Tables, Dashboards & What-If Analysis*

[![Excel](https://img.shields.io/badge/Microsoft%20Excel-365%20%2F%202021-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Sheets](https://img.shields.io/badge/Worksheets-10-FF6F00?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Statistics](https://img.shields.io/badge/Statistics-Descriptive%20%7C%20Regression-4CAF50?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Dashboard](https://img.shields.io/badge/Dashboard-KPIs%20%7C%20Charts%20%7C%20Insights-9C27B0?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)

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
  - [🗄️ Sales Data](#️-1-sales-data--sales_data)
  - [📐 Descriptive Statistics](#-2-descriptive-statistics--descriptive_statistics)
  - [📈 Regression Analysis](#-3-regression-analysis--regression_analysis)
  - [📅 Monthly Sales](#-4-monthly-sales--monthly_sales)
  - [🧮 Region × Product Pivot](#-5-region--product-pivot--region_product_pivot)
  - [🥧 Category Sales](#-6-category-sales--category_sales)
  - [📊 Sales Histogram](#-7-sales-histogram--sales_histogram)
  - [👥 Customer Analysis](#-8-customer-analysis--customer_analysis)
  - [🖥️ Dashboard](#️-9-dashboard--dashboard)
  - [🧪 What-If Analysis](#-10-what-if-analysis--what_if_analysis)
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

The **Sales Analysis Project** is an intermediate-level workbook (`Sales_Analysis_Project.xlsx`) made of **10 worksheets** that analyse **200 sales transactions** from **30 customers**, across **5 regions** and **5 product categories**, between **April 2024 and April 2025**.

Starting from one clean `SalesTable`, the workbook builds descriptive statistics, a regression model, monthly trends, a pivot table, customer rankings, an executive dashboard and a what-if profit simulation.

This project is designed to:
- Organise raw sales data inside a structured **Excel Table** (`SalesTable`)
- Summarise numeric columns with **Descriptive Statistics** (mean, median, SD, skewness, kurtosis)
- Measure the Sales → Profit relationship using **Linear Regression**
- Track monthly performance and **growth %** with `SUMIFS`, `EDATE` and `IFERROR`
- Cross-tabulate **Region × Product Category** with a **Pivot Table**
- Rank customers using **`SORT`**, **`UNIQUE`**, **`FILTER`** and **Top 10 conditional formatting**
- Present KPIs and charts in a single **Dashboard** with written insights
- Test profit sensitivity using a **What-If Data Table**

---

## 🎯 Problem Statement

> **Objective:** Analyse company sales data to find out *how much we sell, where we sell it, who buys it, how profitable it is,* and *what happens to profit if extra discounts are given.*

You are creating a complete sales-analysis report for a business manager. Each worksheet must answer one business question using a specific Excel technique, and the final dashboard must bring the key results together.

| 📂 Category | 📄 Sheets | 🔍 Description |
|-------------|-----------|----------------|
| Raw Data | Sales_Data | 200 transactions stored as `SalesTable` with helper profit columns |
| Statistics | Descriptive_Statistics, Sales_Histogram | Summary measures and distribution of Sales, Quantity, Discount, Profit |
| Predictive Analysis | Regression_Analysis | Does higher Sales lead to higher Profit? |
| Time Analysis | Monthly_Sales | Month-wise total sales and growth % |
| Segment Analysis | Region_Product_Pivot, Category_Sales | Sales by region and by product category |
| Customer Analysis | Customer_Analysis | Total purchase per customer, top customer, high-value filter |
| Reporting | Dashboard | KPIs, four charts and six key insights |
| Scenario Planning | What_If_Analysis | Profit impact of 0% – 20% extra discount |

The goal is to demonstrate **end-to-end business data analysis skills** — from raw data to decision-ready insights — using only Excel.

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 🗄️ **Structured Table** | `SalesTable` makes every formula auto-readable, e.g. `SalesTable[Sales]` |
| 📐 **Full Statistical Summary** | 15 measures for each of Sales, Quantity, Discount and Profit |
| 📈 **Regression Model** | R² = **0.598**, slope = **0.3509**, with a scatter plot and fitted line |
| 📅 **Dynamic Month Table** | `EDATE` builds 13 months automatically; `SUMIFS` totals each month |
| 🚦 **Growth Icons** | Icon-set conditional formatting shows month-on-month rise or fall |
| 🧮 **Pivot Table + Chart** | Region × Product Category cross-tab with a clustered bar chart |
| 🔍 **Spill Formulas** | `SORT(UNIQUE())` lists customers; `FILTER` returns those above ₹8,000 |
| 🏅 **Top-10 Highlighting** | Conditional formatting colours the top 10 customers automatically |
| 🖥️ **KPI Dashboard** | Total Sales, Total Profit, Transactions, Profit Margin and live timestamp |
| 🧪 **What-If Data Table** | Shows total profit for 0%, 5%, 10%, 15%, 20% extra discount in one block |

---

## 🏗️ Project Structure

```
📦 sales-analysis-project/
│
├── 📄 Sales_Analysis_Project.xlsx      ← Main workbook (10 worksheets)
│     ├── Regression_Analysis
│     ├── Descriptive_Statistics
│     ├── Monthly_Sales
│     ├── Region_Product_Pivot
│     ├── Category_Sales
│     ├── Sales_Histogram               ← Chart sheet
│     ├── Sales_Data                    ← Source data (SalesTable)
│     ├── Customer_Analysis
│     ├── Dashboard
│     └── What_If_Analysis
│
└── 📄 README.md                        ← Project documentation
```

---

## 🔄 Project Workflow

```
Open Workbook
      │
      ▼
┌─────────────────────────────┐
│   Raw Data                  │  ← Sales_Data  (SalesTable, 200 rows)
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Understand the Data       │  ← Descriptive_Statistics | Sales_Histogram
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Find Relationships        │  ← Regression_Analysis (Sales vs Profit)
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Summarise by Segment      │  ← Monthly_Sales | Region_Product_Pivot | Category_Sales
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Analyse Customers         │  ← Customer_Analysis (Top 10, High-Value Filter)
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Report Results            │  ← Dashboard (KPIs + Charts + Insights)
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Plan Scenarios            │  ← What-If Data Table (Extra Discount)
└────────────┬────────────────┘
             │
             ▼
      Analysis Complete ✅
```

---

## 📑 Sheet-by-Sheet Guide

### 🗄️ 1. Sales Data — `Sales_Data`

> The foundation of the project: **200 transactions** stored as an Excel Table named **`SalesTable`** (range `A1:J201`).

| Column | Header | Description |
|--------|--------|-------------|
| A | Customer_ID | Customer code (`CUST001` – `CUST030`) |
| B | Customer_Name | Name of the customer |
| C | Region | Central / East / North / South / West |
| D | Product_Category | Books / Clothing / Electronics / Furniture / Office Supplies |
| E | Sales | Sales amount of the order |
| F | Quantity | Units sold (1 – 19) |
| G | Discount | Discount given (0% – 20%) |
| H | Order_Date | Order date (10-Apr-2024 → 07-Apr-2025) |
| I | Profit | Profit earned on the order |
| J | Adjusted_Profit | Profit after deducting an extra discount (used by What-If) |

**Calculated column (J):**
```excel
=SalesTable[[#This Row],[Profit]]-SalesTable[[#This Row],[Sales]]*$K$2
```
> `$K$2` holds the **extra discount %** (currently `0`), which is the input cell for the What-If Data Table.

---

### 📐 2. Descriptive Statistics — `Descriptive_Statistics`

> Excel's **Data Analysis ToolPak → Descriptive Statistics** output for four numeric columns.

| Measure | Sales | Quantity | Discount | Profit |
|---------|-------|----------|----------|--------|
| Mean | 976.09 | 9.995 | 0.0973 | 341.44 |
| Median | 976.72 | 10 | 0.10 | 332.26 |
| Mode | #N/A | 8 | 0.20 | #N/A |
| Standard Deviation | 299.27 | 5.65 | 0.0723 | 135.77 |
| Skewness | 0.147 | 0.061 | 0.097 | 0.573 |
| Kurtosis | -0.089 | -1.209 | -1.326 | 0.265 |
| Minimum | 258.51 | 1 | 0 | 64.63 |
| Maximum | 1,923.66 | 19 | 0.20 | 769.46 |
| Sum | 195,217.76 | 1,999 | 19.45 | 68,287.01 |
| Count | 200 | 200 | 200 | 200 |

**How to read it:**

| Observation | Meaning |
|-------------|---------|
| Mean ≈ Median for Sales | Sales are almost **symmetrically distributed** |
| Profit skewness = 0.57 | Profit is **slightly right-skewed** — a few orders earn much more than average |
| Mode = `#N/A` | Every Sales and Profit value is unique, so no value repeats |

---

### 📈 3. Regression Analysis — `Regression_Analysis`

> Simple linear regression: **Profit (Y)** predicted from **Sales (X)** across 200 observations, with a *Sales Line Fit Plot* scatter chart.

| Regression Statistic | Value |
|----------------------|-------|
| Multiple R | 0.7734 |
| R Square | 0.5982 |
| Adjusted R Square | 0.5961 |
| Standard Error | 86.28 |
| Observations | 200 |
| F-statistic | 294.72 |
| Significance F | 4.63 × 10⁻⁴¹ |

| Term | Coefficient | P-value | 95% Confidence Interval |
|------|-------------|---------|-------------------------|
| Intercept | -1.0316 | 0.9606 | -42.17 to 40.11 |
| Sales | **0.3509** | 4.63 × 10⁻⁴¹ | 0.3106 to 0.3912 |

**Regression equation:**
```
Predicted Profit = -1.03155 + (0.350856 × Sales)
```

**Interpretation written in the sheet:**
- Sales explains approximately **59.82%** of the variation in Profit.
- Higher Sales are associated with higher Profit.
- For every **₹1 increase in Sales**, Profit rises by about **₹0.35**.
- The model also includes the full **Residual Output** table (Observation, Predicted Profit, Residuals) for all 200 rows.

---

### 📅 4. Monthly Sales — `Monthly_Sales`

> Builds a 13-month sales table automatically and calculates month-on-month growth.

| Column | Header | Formula (row 3) | Purpose |
|--------|--------|-----------------|---------|
| A | Month | `=EDATE(A2,1)` | Next month start date (first month is `=DATE(2024,4,1)`) |
| B | Total Sales | `=SUMIFS(SalesTable[Sales],SalesTable[Order_Date],">="&A3,SalesTable[Order_Date],"<"&EDATE(A3,1))` | Sum of sales inside that month |
| C | Growth % | `=IFERROR((B3-B2)/B2,0)` | Change vs. previous month |

| Month | Total Sales | Growth % |
|-------|-------------|----------|
| Apr 2024 | 19,029.96 | — |
| May 2024 | 18,507.34 | -2.7% |
| Jun 2024 | 20,187.51 | +9.1% |
| Jul 2024 | 11,609.04 | -42.5% |
| Aug 2024 | 10,091.70 | -13.1% |
| **Sep 2024** | **20,533.50** | **+103.5%** |
| Oct 2024 | 14,215.24 | -30.8% |
| Nov 2024 | 20,135.03 | +41.6% |
| Dec 2024 | 11,120.04 | -44.8% |
| Jan 2025 | 14,914.65 | +34.1% |
| Feb 2025 | 14,006.36 | -6.1% |
| Mar 2025 | 17,655.22 | +26.1% |
| Apr 2025 | 3,212.17 | -81.8% |

> A **line chart** (*Monthly Sales Trend*) plots the series, and an **icon-set conditional format** on `C3:C14` shows arrows for rising and falling months.

---

### 🧮 5. Region × Product Pivot — `Region_Product_Pivot`

> A **Pivot Table** (Sum of Sales) with *Region* in rows and *Product Category* in columns, plus a clustered bar chart.

| Region | Books | Clothing | Electronics | Furniture | Office Supplies | **Total** |
|--------|-------|----------|-------------|-----------|-----------------|-----------|
| Central | 11,297.15 | 12,200.26 | 6,621.24 | 4,876.96 | 8,472.72 | **43,468.33** |
| East | 11,389.09 | 2,401.45 | 2,869.54 | 7,592.34 | 8,189.22 | **32,441.64** |
| North | 8,037.01 | 8,422.82 | 8,191.71 | 6,978.08 | 3,345.80 | **34,975.42** |
| South | 9,588.04 | 8,393.63 | 11,989.07 | 7,982.73 | 2,295.32 | **40,248.79** |
| West | 6,691.17 | 10,509.43 | 11,372.44 | 9,182.78 | 6,327.76 | **44,083.58** |
| **Grand Total** | **47,002.46** | **41,927.59** | **41,044.00** | **36,612.89** | **28,630.82** | **195,217.76** |

**Region-wise profit margin (calculated from `Sales_Data`):**

| Region | Sales | Profit | Margin |
|--------|-------|--------|--------|
| Central | 43,468.33 | 16,068.40 | 37.0% |
| East | 32,441.64 | 11,638.59 | 35.9% |
| North | 34,975.42 | 11,902.28 | 34.0% |
| South | 40,248.79 | 13,691.91 | 34.0% |
| West | 44,083.58 | 14,985.83 | 34.0% |

---

### 🥧 6. Category Sales — `Category_Sales`

> Total sales per product category using `SUMIF` on the table, shown as a **pie chart** (share) and a **bar chart** (comparison).

```excel
=SUMIF(SalesTable[Product_Category],A2,SalesTable[Sales])
```

| Category | Total Sales | Share | Profit | Margin |
|----------|-------------|-------|--------|--------|
| Books | 47,002.46 | 24.1% | 16,901.63 | 36.0% |
| Clothing | 41,927.59 | 21.5% | 18,784.49 | **44.8%** |
| Electronics | 41,044.00 | 21.0% | 9,182.15 | **22.4%** |
| Furniture | 36,612.89 | 18.8% | 11,579.14 | 31.6% |
| Office Supplies | 28,630.82 | 14.7% | 11,839.60 | 41.4% |

---

### 📊 7. Sales Histogram — `Sales_Histogram`

> A dedicated **chart sheet** titled **"Sales Amount Distribution"**.
> - X-axis: *Sales Amount Range*
> - Y-axis: *Number of Transactions*

It shows how the 200 order values are spread between **258.51** and **1,923.66**, and confirms the near-symmetric (bell-shaped) distribution seen in the descriptive statistics.

---

### 👥 8. Customer Analysis — `Customer_Analysis`

> Ranks customers by total purchase value and extracts high-value customers.

| Cell / Range | Formula | Purpose |
|--------------|---------|---------|
| `B3` | `=SORT(UNIQUE(SalesTable[Customer_ID]))` | Spills the unique, sorted customer list |
| `C3:C32` | `=SUMIF(SalesTable[Customer_ID],B3,SalesTable[Sales])` | Total purchase per customer |
| `F2` | `=MAX(C3:C202)` | Highest total purchase |
| `F3` | `=INDEX(B3:B202,MATCH(F2,C3:C202,0))` | Customer ID of the top buyer |
| `F6` | `8000` *(input)* | Minimum total purchase threshold |
| `E8` | `=FILTER(B3:C202,(B3:B202<>"")*(C3:C202>=$F$6),"No customers found")` | Customers at or above the threshold |

**Top 10 Customers:**

| Rank | Customer ID | Total Purchase |
|------|-------------|----------------|
| 🥇 1 | CUST015 | 11,611.06 |
| 🥈 2 | CUST028 | 10,293.87 |
| 🥉 3 | CUST007 | 9,941.88 |
| 4 | CUST008 | 9,760.80 |
| 5 | CUST026 | 9,371.81 |
| 6 | CUST021 | 8,740.18 |
| 7 | CUST003 | 8,490.18 |
| 8 | CUST012 | 8,339.00 |
| 9 | CUST030 | 7,928.71 |
| 10 | CUST004 | 7,813.59 |

> **8 customers** have total purchases of **₹8,000 or more**. Change `F6` and the `FILTER` list updates instantly. **Top-10 conditional formatting** is applied on `C3:C32` and on the filtered list.

---

### 🖥️ 9. Dashboard — `Dashboard`

> A one-page executive summary: KPI cards, four charts and written insights.

**KPI cards:**

| KPI | Formula | Value |
|-----|---------|-------|
| Total Sales | `=SUM(SalesTable[Sales])` | **195,217.76** |
| Total Profit | `=SUM(SalesTable[Profit])` | **68,287.01** |
| Total Transactions | `=ROWS(SalesTable[Sales])` | **200** |
| Profit Margin | `=E3/B3` | **34.98%** |
| Updated On | `=NOW()` | Live date & time |

**Charts on the dashboard:**

| # | Chart | Type | Shows |
|---|-------|------|-------|
| 1 | Monthly Sales Trend | Line | Sales over 13 months |
| 2 | Sales Share by Product Category | Pie | Category contribution |
| 3 | Sales by Region and Product Category | Clustered Bar | Region × category comparison |
| 4 | Total Sales by Product Category | Bar | Category ranking |

---

### 🧪 10. What-If Analysis — `What_If_Analysis`

> Uses a **one-variable Data Table** to see how **Total Profit** changes when an **extra discount** is applied to every order.

**Setup (located in `Sales_Data`, columns K–N):**

| Cell | Content | Role |
|------|---------|------|
| `K2` | `0` | **Input cell** — extra discount % |
| `J2:J201` | `=Profit - Sales × $K$2` | Adjusted profit per order |
| `N2` | `=SUM(SalesTable[Adjusted_Profit])` | Total profit (data table formula) |
| `M3:M7` | `0, 5%, 10%, 15%, 20%` | Discount scenarios |
| `N3:N7` | `{=TABLE(,K2)}` | Data table results |

**Result:**

| Extra Discount | Total Profit | Change vs. Base |
|----------------|--------------|-----------------|
| 0% | 68,287.01 | — |
| 5% | 58,526.12 | -9,760.89 |
| 10% | 48,765.23 | -19,521.78 |
| 15% | 39,004.35 | -29,282.66 |
| 20% | 29,243.46 | -39,043.55 |

> Every extra **5%** discount reduces total profit by about **₹9,761** (≈ 14.3%). At a **20%** extra discount, profit falls by more than **57%**.

---

## 🛠️ Tech Stack

| Tool | Version | Purpose |
|------|---------|---------|
| 📗 **Microsoft Excel** | 365 / 2021 or later | Workbook and formula engine |
| 🗄️ **Excel Tables** | Built-in | Structured references (`SalesTable[Sales]`) |
| 📊 **SUMIF / SUMIFS** | Built-in | Conditional totals |
| 📅 **DATE / EDATE** | Built-in | Month series generation |
| 🛡️ **IFERROR** | Built-in | Safe growth % calculation |
| 🔍 **INDEX / MATCH / MAX** | Built-in | Top customer lookup |
| ✨ **SORT / UNIQUE / FILTER** | Excel 365 / 2021 | Dynamic array outputs |
| 📐 **Data Analysis ToolPak** | Add-in | Descriptive Statistics, Regression, Histogram |
| 🧮 **PivotTable & PivotChart** | Built-in | Region × Category summary |
| 🚦 **Conditional Formatting** | Built-in | Icon sets and Top-10 rules |
| 🧪 **Data Table (What-If)** | Built-in | Scenario simulation |
| 📈 **Charts** | Built-in | Line, Pie, Bar, Scatter, Histogram |

---

## 📈 Sample Results

Key outputs captured from the workbook:

```
Sales_Data              Records                       →  200 transactions | 30 customers
Descriptive_Statistics  Mean Sales                    →  976.09
Descriptive_Statistics  Mean Profit                   →  341.44
Regression_Analysis     R Square                      →  0.5982 (59.82%)
Regression_Analysis     Equation                      →  Profit = -1.03155 + 0.350856 × Sales
Monthly_Sales           Best month                    →  Sep 2024 : 20,533.50
Monthly_Sales           Lowest full month             →  Aug 2024 : 10,091.70
Region_Product_Pivot    Top region                    →  West : 44,083.58
Region_Product_Pivot    Lowest region                 →  East : 32,441.64
Category_Sales          Top category                  →  Books : 47,002.46
Category_Sales          Lowest category               →  Office Supplies : 28,630.82
Customer_Analysis       Top customer                  →  CUST015 : 11,611.06
Customer_Analysis       Customers ≥ 8,000             →  8 customers
Dashboard               Total Sales                   →  195,217.76
Dashboard               Total Profit                  →  68,287.01
Dashboard               Profit Margin                 →  34.98%
What_If_Analysis        20% extra discount            →  Profit falls to 29,243.46
```

---

## 💡 Key Business Insights

| # | Insight | Evidence |
|---|---------|----------|
| 1 | **Overall performance is healthy** | ₹195,217.76 sales with a ≈ **35% profit margin** |
| 2 | **West is the strongest region** | Highest sales at ₹44,083.58, followed closely by Central (₹43,468.33) |
| 3 | **East needs attention** | Lowest regional sales; Clothing (₹2,401) and Electronics (₹2,870) are very weak there |
| 4 | **Books lead in sales, Clothing leads in profit** | Books = ₹47,002 sales; Clothing = ₹18,784 profit at a 44.8% margin |
| 5 | **Electronics is the least profitable** | Only 22.4% margin despite ₹41,044 in sales |
| 6 | **Sales drive profit** | Regression R² of 0.60 — a ₹1 rise in sales adds ≈ ₹0.35 profit |
| 7 | **Sales are seasonal and volatile** | Peaks in Apr–Jun, Sep, Nov; dips in Jul–Aug and Dec |
| 8 | **Customer base is concentrated** | CUST015 alone buys ₹11,611, about 6% of all sales |
| 9 | **Discounts are expensive** | Each extra 5% discount removes ≈ ₹9,761 of profit |

---

## ⚠️ Known Issues & Notes

| # | Note | Detail |
|---|------|--------|
| 1 | **Excel version** | `SORT`, `UNIQUE` and `FILTER` need Excel 365 or 2021+; older versions show `#NAME?` |
| 2 | **Partial months** | April 2024 starts on the 10th and April 2025 ends on the 7th, so the **-81.8%** growth in Apr 2025 reflects an incomplete month, not a real collapse |
| 3 | **Volatile `NOW()`** | The Dashboard timestamp recalculates on every change |
| 4 | **Mode shows `#N/A`** | Sales and Profit have no repeating values, which is expected behaviour |
| 5 | **Spill range** | `SORT(UNIQUE())` in `B3` and `FILTER` in `E8` need empty cells below them, otherwise they return `#SPILL!` |
| 6 | **Stray header** | Cell `K1` carries the default header **"Column1"** because it was auto-included in the table area; it can be renamed or cleared |
| 7 | **What-If location** | The Data Table lives in `Sales_Data!M1:N7`; the `What_If_Analysis` sheet itself is the reserved space for scenario outputs |
| 8 | **Data Table recalculation** | If Excel is set to *"Automatic except Data Tables"*, press **F9** to refresh the scenario results |
| 9 | **Hardcoded ranges** | `Customer_Analysis` uses fixed ranges (`C3:C202`); adding more than 200 customers needs the ranges extended |
| 10 | **Name repeats** | 30 Customer IDs map to only 27 unique names, so always analyse by **Customer_ID**, not by name |
| 11 | **Static pivot** | The PivotTable must be refreshed (**Data → Refresh All**) after the source data changes |

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🎓 **Complete Workflow** | Covers data → statistics → prediction → reporting → scenario planning |
| 🗄️ **Table-Driven** | `SalesTable` keeps formulas readable and auto-expanding for new rows |
| 📊 **Visual Storytelling** | Charts for trend, share, comparison, distribution and regression |
| 🔍 **Easy to Trace** | Every figure on the dashboard comes from a visible formula |
| 🧪 **Decision Support** | What-If table quantifies the cost of discounting |
| 🔄 **Interactive** | Change the discount cell `K2` or the customer threshold `F6` and results update |
| 💻 **No Macros** | Pure formulas, pivot tables and charts — no VBA required |
| ⚡ **Lightweight** | A single `.xlsx` file that opens instantly |

---

## 🚀 How to Use

1. Download or clone the repository.
2. Open `Sales_Analysis_Project.xlsx` in **Microsoft Excel 365 / 2021 or later**.
3. Start with the **Dashboard** for the big picture, then explore each sheet.
4. Click any formula cell and read it in the **formula bar**.
5. Change the **minimum purchase** in `Customer_Analysis!F6` and watch the `FILTER` list update.
6. Change the **extra discount** in `Sales_Data!K2` and watch `Adjusted_Profit` and `Total_Profit` change.
7. After editing source data, use **Data → Refresh All** to update the PivotTable and charts.
8. Press **F9** to recalculate the What-If Data Table if needed.

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
