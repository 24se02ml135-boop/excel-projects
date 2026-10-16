<div align="center">

# -- ! Retail Sales Performance Dashboard ! --
### *10-Sheet Excel Project: Data Cleaning, Formulas, Summaries, Charts & Interactive Dashboard*

[![Excel](https://img.shields.io/badge/Microsoft%20Excel-365%20%2F%202021-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Sheets](https://img.shields.io/badge/Worksheets-10-FF6F00?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Records](https://img.shields.io/badge/Cleaned%20Orders-360-4CAF50?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Charts](https://img.shields.io/badge/Charts-9-9C27B0?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)

<br/>

> *"Raw data tells you what happened. A dashboard tells you what to do next."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Project Workflow](#-project-workflow)
- [📑 Sheet-by-Sheet Guide](#-sheet-by-sheet-guide)
  - [📥 Raw Dataset](#-1-raw-dataset)
  - [🧹 Cleaned Data](#-2-cleaned-data)
  - [🔍 Lookup & Advanced Formulas](#-3-lookup--advanced-formulas)
  - [📊 Pivot Summaries](#-4-pivot-summaries)
  - [📈 Data Visualization](#-5-data-visualization)
  - [🖥️ Interactive Dashboard](#️-6-interactive-dashboard)
  - [📝 Instructions, Tasks & Documentation](#-7-instructions-tasks--documentation)
- [🛠️ Tech Stack](#️-tech-stack)
- [📈 Sample Results](#-sample-results)
- [⚠️ Known Issues & Notes](#️-known-issues--notes)
- [🏆 Advantages](#-advantages)
- [🚀 How to Use](#-how-to-use)
- [📄 License](#-license)
- [👤 Author](#-author)
- [🙏 Acknowledgements](#-acknowledgements)

---

## 📌 Overview

The **Retail Sales Performance Dashboard** is an Excel workbook (`Aastha_ExcelDashboard.xlsx`) made of **10 worksheets** that take a messy retail transaction dataset all the way to a filterable, chart-driven dashboard. It uses a **synthetic dataset of 362 raw order records** (2 duplicates, a few missing values), cleans it down to **360 unique orders**, and analyses it by region, category, month, sales channel and salesperson.

This project is designed to:
- **Import and clean** raw data (duplicates, blanks, missing values)
- Derive new columns such as **Gross Sales, Net Sales, Month, Quarter and Order Size**
- Apply **`VLOOKUP`**, **`INDEX + MATCH`** and **nested `IF`** formulas
- Build **formula-based summary tables** with `SUMIF` and `COUNTIF`
- Visualise results with **bar, pie and line charts**
- Build an **interactive dashboard** with dropdown filters and live KPIs
- **Document** the whole workbook so another person can follow it

---

## 🎯 Problem Statement

> **Objective:** Complete a 3-hour Excel dashboard assignment: choose a dataset, clean it, analyse it with formulas, visualise it and present it as an interactive dashboard.

The assignment is split into **7 weighted tasks**, completed in order:

| # | 📂 Task | 🔍 Skills Tested | ⏱️ Time (min) | ⚖️ Weight |
|---|---------|------------------|---------------|-----------|
| 1 | Dataset Selection & Import | Data import, initial cleaning | 15 | 5% |
| 2 | Data Cleaning & Preparation | Remove duplicates, handle missing values, format data types | 25 | 10% |
| 3 | Create Pivot Tables | Data analysis, summarisation | 30 | 15% |
| 4 | Advanced Formulas | VLOOKUP, INDEX-MATCH, IF conditions, nested functions | 35 | 15% |
| 5 | Data Visualization | Charts, conditional formatting | 40 | 20% |
| 6 | Interactive Dashboard | Slicers, timeline, dynamic ranges | 45 | 25% |
| 7 | Documentation | Sheet organisation, naming conventions, instructions | 20 | 10% |

The goal is to demonstrate an **end-to-end analyst workflow** inside a single Excel file.

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 🧹 **Data Cleaning** | Removes 2 duplicate Order IDs, fills blank Region with `Unknown`, recovers missing Category from the Lookup sheet, sets missing Discount to `0` |
| 💰 **Derived Metrics** | `Net Sales = Quantity × Unit Price × (1 − Discount)` calculated for every order |
| 🧠 **Order Size Banding** | Nested `IF` labels each order **High / Medium / Low** |
| 🔍 **Two Lookup Styles** | `VLOOKUP` returns category, `INDEX + MATCH` returns base price |
| 📊 **Live Summary Tables** | `SUMIF` / `COUNTIF` tables for Region, Category, Month and Channel |
| 📈 **9 Charts** | Bar, pie and line charts across the Dashboard and Data Visualization sheets |
| 🎛️ **Dropdown Filters** | Region, Category and Channel filters drive the KPI cards and tables |
| 🛡️ **Error Handling** | `IFERROR` returns `0` when a filter combination has no data |

---

## 🏗️ Project Structure

```
📦 excel-dashboard-project/
│
├── 📄 Aastha_ExcelDashboard.xlsx      ← Main workbook (10 worksheets)
│     ├── Instructions                 ← Assignment rules
│     ├── Tasks                        ← 7 weighted tasks
│     ├── Raw Dataset                  ← 362 original records
│     ├── Cleaned Data                 ← 360 unique orders + derived columns
│     ├── Lookup                       ← Product → Category → Base Price
│     ├── Pivot Summaries              ← Region / Category / Month / Channel tables
│     ├── Advanced Formulas            ← VLOOKUP, INDEX-MATCH, nested IF
│     ├── Dashboard                    ← Filters, KPI cards, filtered tables & charts
│     ├── Data Visualization           ← Static charts for the full dataset
│     └── Documentation                ← Method, usage and limitations
│
└── 📄 README.md                       ← Project documentation
```

---

## 🔄 Project Workflow

```
Open Workbook
      │
      ▼
┌─────────────────────────────┐
│   Import                    │  ← Raw Dataset (362 rows, 10 columns)
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Clean & Prepare           │  ← Cleaned Data (360 rows, 15 columns)
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Enrich with Formulas      │  ← Lookup | Advanced Formulas
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Summarise                 │  ← Pivot Summaries (SUMIF / COUNTIF)
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Visualise                 │  ← Data Visualization (4 charts)
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Interactive Dashboard     │  ← Dropdown filters + KPIs + 5 charts
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Document                  │  ← Documentation sheet
└────────────┬────────────────┘
             │
             ▼
      Workbook Complete ✅
```

---

## 📑 Sheet-by-Sheet Guide

### 📥 1. Raw Dataset

> The untouched source data: **362 records, 10 columns**, covering orders from **2 Jan 2025 to 28 Dec 2025**.

| Column | Description |
|--------|-------------|
| Order ID | Unique order code, e.g. `ORD-0001` |
| Order Date | Date of the order |
| Region | North / South / East / West |
| Product | 10 products (Laptop, Phone, Desk, Chair...) |
| Category | Electronics / Furniture / Stationery / Accessories |
| Quantity, Unit Price | Order size and price per unit |
| Discount | Fraction (e.g. `0.1` = 10%) |
| Sales Channel | Online / Retail / Wholesale |
| Salesperson | Aarav, Diya, Kabir, Meera, Riya, Vivaan |

**Intentional data problems:** 2 duplicate Order IDs, 1 blank Region, 1 blank Category and 1 blank Discount.

---

### 🧹 2. Cleaned Data

> **360 unique orders** with 5 derived columns added to the original 10.

| Column | Header | Formula (row 2) | Purpose |
|--------|--------|-----------------|---------|
| K | Gross Sales | `=F2*G2` | Quantity × Unit Price |
| L | Net Sales | `=K2*(1-H2)` | Sales after discount |
| M | Month | `=TEXT(B2,"mmm")` | Month name from date |
| N | Quarter | `="Q"&ROUNDUP(MONTH(B2)/3,0)` | Q1 – Q4 |
| O | Order Size | `=IF(L2>=100000,"High",IF(L2>=30000,"Medium","Low"))` | Value band |

**Cleaning rules applied:**

| Problem | Fix |
|---------|-----|
| Duplicate Order ID | Removed (362 → 360 rows) |
| Blank Region | Replaced with `Unknown` |
| Missing Category | Looked up from the `Lookup` sheet using Product |
| Missing Discount | Set to `0` |

**Order Size bands:**

| Net Sales | Band |
|-----------|------|
| 100,000 and above | High |
| 30,000 – 99,999 | Medium |
| Below 30,000 | Low |

---

### 🔍 3. Lookup & Advanced Formulas

**`Lookup` sheet** holds the product master: 10 products with Category and Base Price.

**`Advanced Formulas` sheet** demonstrates three formula styles on the first 20 orders:

| Column | Formula (row 2) | Result |
|--------|-----------------|--------|
| VLOOKUP Category | `=VLOOKUP(B2,Lookup!$A$2:$C$11,2,FALSE)` | Phone → **Electronics** |
| INDEX-MATCH Price | `=INDEX(Lookup!$C$2:$C$11,MATCH(B2,Lookup!$A$2:$A$11,0))` | Phone → **28000** |
| Nested IF Sales Band | `=IF('Cleaned Data'!L2>=100000,"High",IF('Cleaned Data'!L2>=30000,"Medium","Low"))` | ORD-0001 → **Low** |

**Key Concepts Used:**

| Concept | Detail |
|---------|--------|
| `FALSE` / `0` match | Forces an exact match in `VLOOKUP` and `MATCH` |
| `$` anchors | Keeps the Lookup range fixed while copying formulas down |
| Nested `IF` | Three outcomes from two conditions |

---

### 📊 4. Pivot Summaries

> Four formula-based summary tables (built with `SUMIF` / `COUNTIF` instead of native PivotTables).

```excel
Total Sales  =SUMIF('Cleaned Data'!$C$2:$C$361, A2, 'Cleaned Data'!$L$2:$L$361)
Orders       =COUNTIF('Cleaned Data'!$C$2:$C$361, A2)
```

| Table | Grouped By | Measures |
|-------|-----------|----------|
| 1 | Region | Total Sales, Orders |
| 2 | Category | Total Sales, Units Sold |
| 3 | Month | Total Sales, Orders |
| 4 | Channel | Total Sales, Orders |

---

### 📈 5. Data Visualization

> Four charts built from the full cleaned dataset.

| Chart | Type | Shows |
|-------|------|-------|
| Region | Bar | Net Sales by region |
| Category | Pie | Share of Net Sales by category |
| Month | Line | Monthly sales trend |
| Channel | Bar | Net Sales by channel |

---

### 🖥️ 6. Interactive Dashboard

> Three dropdown filters control the KPI cards and summary tables.

| Filter | Cell | Options |
|--------|------|---------|
| Region | `B4` | All, North, South, East, West, Unknown |
| Category | `B5` | All, Electronics, Furniture, Stationery, Accessories |
| Channel | `B6` | All, Online, Retail, Wholesale |

**KPI cards:**

| KPI | Formula idea |
|-----|--------------|
| Net Sales | `SUMIFS` on Net Sales with all three filters |
| Order Count | `COUNTIFS` with all three filters |
| Units Sold | `SUMIFS` on Quantity with all three filters |
| Avg Order Value | `Net Sales ÷ Order Count` |

**Filter logic (wildcard trick):**
```excel
=IFERROR(SUMIFS('Cleaned Data'!$L$2:$L$361,
   'Cleaned Data'!$C$2:$C$361, IF($B$4="All","*",$B$4),
   'Cleaned Data'!$E$2:$E$361, IF($B$5="All","*",$B$5),
   'Cleaned Data'!$I$2:$I$361, IF($B$6="All","*",$B$6)),0)
```
> When a filter is set to **All**, it is replaced by the wildcard `*`, which matches every value. `IFERROR` returns `0` if nothing matches.

The sheet also shows **filtered Region, Category and Month tables** and **5 charts** (Sales by Region, Sales by Category, Monthly Sales Trend and two summary charts).

---

### 📝 7. Instructions, Tasks & Documentation

| Sheet | Purpose |
|-------|---------|
| Instructions | 6 assignment rules (3-hour limit, choose a dataset, save every 15 min...) |
| Tasks | 7 tasks with skills, time estimate and weightage |
| Documentation | 20-row reference: project, dataset, cleaning method, key metric, usage, refresh steps, limitations |

---

## 🛠️ Tech Stack

| Tool | Version | Purpose |
|------|---------|---------|
| 📗 **Microsoft Excel** | 365 / 2021 or later | Workbook, formulas and charts |
| 🧹 **Data Cleaning** | Built-in | Duplicate removal, missing-value handling |
| 🔍 **VLOOKUP / INDEX / MATCH** | Built-in | Lookups |
| 🧠 **IF (nested)** | Built-in | Order size and sales bands |
| 📊 **SUMIF / SUMIFS / COUNTIF / COUNTIFS** | Built-in | Summaries and filtered KPIs |
| 📅 **TEXT / MONTH / ROUNDUP** | Built-in | Month and quarter extraction |
| 🛡️ **IFERROR** | Built-in | Error prevention |
| 📈 **Bar / Pie / Line Charts** | Built-in | Visualisation |
| 🎛️ **Data Validation** | Built-in | Dropdown filters |

---

## 📈 Sample Results

Key outputs from the cleaned dataset (360 orders, all filters set to **All**):

```
Total Net Sales      →  22,614,134
Total Gross Sales    →  25,141,790
Order Count          →  360
Units Sold           →  2,258
Avg Order Value      →  62,817

Net Sales by Region
  South    7,226,252  (83 orders)
  East     7,001,700  (107 orders)
  North    4,230,562  (77 orders)
  West     4,151,140  (92 orders)
  Unknown      4,480  (1 order)

Net Sales by Category
  Electronics  16,088,550
  Furniture     5,814,750
  Accessories     581,360
  Stationery      129,474

Net Sales by Channel
  Wholesale  8,329,634  (129 orders)
  Online     7,669,874  (120 orders)
  Retail     6,614,626  (111 orders)

Monthly Peak / Low
  Highest: Dec 3,634,122   |   Lowest: Feb 1,022,370

Top Product    Laptop  9,763,000  (27 orders)
Top Salesperson  Kabir  4,472,958
Order Size     High 55 | Medium 90 | Low 215
```

---

## ⚠️ Known Issues & Notes

| # | Note | Detail |
|---|------|--------|
| 1 | **Synthetic data** | The dataset was generated for practice; it is not from Kaggle or UCI, so results are illustrative |
| 2 | **Dropdowns, not slicers** | Filters use Data Validation dropdowns; native slicers and timelines are not used |
| 3 | **No native PivotTables** | Summary tables are `SUMIF` / `COUNTIF` formulas, so they do not refresh via *Refresh All* |
| 4 | **Channel chart labels** | The channel table on `Data Visualization` lists *Online, In-Store, Other*, but the data only has *Online, Retail, Wholesale*. Only Online has a value there; `Pivot Summaries` shows the correct channels |
| 5 | **Order counts ignore filters** | In the Dashboard Region and Month tables, the *Orders* column uses `COUNTIF` on one column only, so it does not change with the dropdowns (revenue columns do) |
| 6 | **Hardcoded ranges** | Formulas use fixed ranges ending at row 361; new records need extended ranges, or the data should be converted to an Excel Table |
| 7 | **Partial formula demo** | `Advanced Formulas` covers only the first 20 orders |
| 8 | **Unknown region** | One order has no region and is grouped as `Unknown` |
| 9 | **Calculated columns** | Columns K–O should not be edited by hand |

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🔄 **End-to-End** | Covers import, cleaning, analysis, visuals and dashboard in one file |
| 🎛️ **Interactive** | Change a dropdown and every KPI updates instantly |
| 🔍 **Traceable** | Every number comes from a visible formula, with no hidden macros |
| 🛡️ **Safe** | `IFERROR` prevents ugly error values on empty filter results |
| 📝 **Documented** | Built-in Documentation sheet explains method, usage and limits |
| ⚡ **Lightweight** | A single `.xlsx` file that needs no add-ins or VBA |
| 🧪 **Extensible** | Easy to upgrade with PivotTables, slicers, Tables and `TODAY()` |

---

## 🚀 How to Use

1. Download or clone the repository.
2. Open `Aastha_ExcelDashboard.xlsx` in **Microsoft Excel 365 / 2021 or later**.
3. Make sure **Formulas → Calculation Options → Automatic** is on.
4. Go to the **Dashboard** sheet.
5. Pick values in **Region (B4)**, **Category (B5)** and **Channel (B6)**.
6. Watch the KPI cards, tables and charts update.
7. Review `Raw Dataset` → `Cleaned Data` to see how the data was prepared.
8. To add new records, extend the ranges past row 361 in `Cleaned Data` and update the summary and chart references.

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
> *"Every chart you build turns a table of numbers into a decision."*

**🎓 Role:** Python Learner | Data & Spreadsheet Enthusiast \
**📍 Location:** India \
**🛠️ Skills:** Excel Dashboards · Data Cleaning · Lookups · Conditional Logic · Data Analysis · Python

</div>

---

## 🙏 Acknowledgements

Special thanks to the following resources and communities that made this project possible:

- 📗 [Microsoft Excel Support](https://support.microsoft.com/excel) — Official function reference
- 📊 [Create a chart in Excel](https://support.microsoft.com/en-us/office/create-a-chart-from-start-to-finish-0baf399e-dd61-4e18-8a73-b3fd5d5680c2) — Charting guide
- 🎛️ [Data Validation in Excel](https://support.microsoft.com/en-us/office/apply-data-validation-to-cells-29fecbcc-d1b9-42c1-9d76-eff3ce5f7249) — Dropdown filter reference
- 🎓 [ExcelJet](https://exceljet.net/) — Clear formula explanations and examples
- 🖥️ [W3Schools Excel](https://www.w3schools.com/excel/) — Beginner Excel reference
- 💬 [Stack Overflow Community](https://stackoverflow.com/) — Problem-solving support

---

<div align="center">

---

*Made with ❤️ and ☕ — Last updated: October 2026*

</div>
