<div align="center">

# -- ! Excel Formulas Practice Project ! --
### *16-Sheet Workbook Covering Cell References, Logic, Lookups, Text, Dates, Math & Dynamic Arrays*

[![Excel](https://img.shields.io/badge/Microsoft%20Excel-365%20%2F%202021-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Sheets](https://img.shields.io/badge/Worksheets-16-FF6F00?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Lookups](https://img.shields.io/badge/Lookups-VLOOKUP%20%7C%20XLOOKUP%20%7C%20INDEX--MATCH-4CAF50?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Dynamic](https://img.shields.io/badge/Dynamic%20Arrays-FILTER%20%7C%20XMATCH-9C27B0?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)

<br/>

> *"A formula is just a question you teach the spreadsheet to answer for you."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Project Workflow](#-project-workflow)
- [📑 Sheet-by-Sheet Guide](#-sheet-by-sheet-guide)
  - [🔗 Cell References](#-1-cell-references--01_relative_absolute)
  - [🧠 Logical Functions](#-2-logical-functions--02-to-03)
  - [📊 Conditional Aggregation](#-3-conditional-aggregation--04_countifs_sumifs)
  - [🔍 Lookup Functions](#-4-lookup-functions--05-to-08-10-to-11)
  - [🔤 Text Functions](#-5-text-functions--09_text)
  - [🧭 Dynamic References](#-6-dynamic-references--12-to-13)
  - [📅 Date & Math](#-7-date--math-functions--14-to-15)
  - [🧪 Dynamic Arrays](#-8-dynamic-arrays--16_filter)
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

The **Excel Formulas Practice Project** is a beginner-to-intermediate workbook (`Project_1_excel.xlsx`) made of **16 worksheets**, each dedicated to one formula concept. Every sheet contains a small, realistic dataset (students, products, sales, employees) and live formulas that demonstrate how the function works.

This project is designed to:
- Understand **relative vs absolute** cell references (`B2` vs `$C$2`)
- Build decision logic with **nested `IF`**, **`AND`** and **`OR`**
- Summarise data using **`COUNTIFS`**, **`SUMIFS`** and **`AVERAGEIFS`**
- Compare classic and modern **lookups** (`VLOOKUP`, `INDEX + MATCH`, `XLOOKUP`, `XMATCH`)
- Clean and transform text with **`LEFT`**, **`FIND`**, **`UPPER`**, **`LOWER`**
- Use **dynamic references** with **`INDIRECT`** and **`OFFSET`**
- Work with **dates** (`DATEDIF`, date subtraction) and **rounding** (`ROUND`, `CEILING`, `FLOOR`)
- Return multiple rows at once with the **`FILTER`** dynamic array function

---

## 🎯 Problem Statement

> **Objective:** Build a multi-sheet Excel workbook that demonstrates the most commonly used Excel formulas with practical datasets.

You are creating a hands-on formula reference for students learning Excel. Each worksheet must hold sample data, apply one formula family, and show the result clearly so the logic can be traced cell by cell.

| 📂 Category | 📄 Sheets | 🔍 Description |
|-------------|-----------|----------------|
| Cell References | 01 | Tax and total calculation using a fixed tax-rate cell |
| Logical Functions | 02, 03 | Grades, discounts, eligibility and pass checks |
| Conditional Aggregation | 04 | Count, sum and average with one or more conditions |
| Lookup Functions | 05, 06, 07, 08, 10, 11 | Fetch values by ID, code, or multiple criteria |
| Text Functions | 09 | Extract first names, change letter case |
| Dynamic References | 12, 13 | Read cells by text address and shifting ranges |
| Date & Math | 14, 15 | Age, day gaps, rounding variations |
| Dynamic Arrays | 16 | Filter all students scoring above 80 |

The goal is to demonstrate **practical spreadsheet problem-solving skills** through small, focused, easy-to-follow examples.

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 🔗 **Absolute References** | `$C$2` locks the tax rate while formulas are copied down 20 rows |
| 🧠 **Nested Logic** | Three-level `IF` assigns grades A / B / C / Fail |
| 🎯 **Multi-Criteria Checks** | `AND` / `OR` decide eligibility and top-performer status |
| 📊 **Conditional Totals** | `SUMIFS` finds North-region Laptop sales in one formula |
| 🔍 **6 Lookup Styles** | `VLOOKUP`, `INDEX+MATCH`, `XLOOKUP`, `XMATCH` side by side |
| 🔑 **Helper Key Column** | `Salesperson|Month` key enables two-condition lookup |
| 🔤 **Text Cleaning** | Splits first names and converts case automatically |
| 🧭 **Dynamic Addressing** | `INDIRECT` and `OFFSET` build references on the fly |
| 📅 **Date Intelligence** | Calculates age and service years with `DATEDIF` |
| 🧪 **Spill Output** | `FILTER` returns all qualifying rows in one formula |

---

## 🏗️ Project Structure

```
📦 excel-formulas-project/
│
├── 📄 Project_1_excel.xlsx      ← Main workbook (16 worksheets)
│     ├── 01_Relative_Absolute
│     ├── 02_Nested_IF
│     ├── 03_IF_AND_OR
│     ├── 04_COUNTIFS_SUMIFS
│     ├── 05_VLOOKUP_Student
│     ├── 06_VLOOKUP_Product
│     ├── 07_INDEX_MATCH
│     ├── 08_Employee
│     ├── 09_TEXT
│     ├── 10_XLOOKUP
│     ├── 11_XMATCH
│     ├── 12_INDIRECT
│     ├── 13_OFFSET
│     ├── 14_Date_Time
│     ├── 15_Math
│     └── 16_FILTER
│
└── 📄 README.md                 ← Project documentation
```

---

## 🔄 Project Workflow

```
Open Workbook
      │
      ▼
┌─────────────────────────────┐
│   Basics                    │  ← 01 Relative / Absolute references
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Decision Logic            │  ← 02 Nested IF  |  03 IF + AND / OR
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Summarise Data            │  ← 04 COUNTIFS / SUMIFS / AVERAGEIFS
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Find Data (Lookups)       │  ← 05 06 VLOOKUP | 07 INDEX-MATCH | 08 10 XLOOKUP | 11 XMATCH
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Transform Data            │  ← 09 TEXT | 12 INDIRECT | 13 OFFSET
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Dates & Numbers           │  ← 14 Date/Time | 15 ROUND / CEILING / FLOOR
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Dynamic Arrays            │  ← 16 FILTER
└────────────┬────────────────┘
             │
             ▼
      Workbook Complete ✅
```

---

## 📑 Sheet-by-Sheet Guide

### 🔗 1. Cell References — `01_Relative_Absolute`

> Calculates tax, final price and total amount for 20 customers using one fixed tax-rate cell.

| Column | Header | Formula (row 2) | Purpose |
|--------|--------|-----------------|---------|
| A–D | Name, Price, Tax Rate, Quantity | — | Input data |
| E | Tax Amount | `=B2*$C$2` | **Absolute** reference locks the tax rate |
| F | final price | `=B2+E2` | **Relative** reference moves with each row |
| G | Total Amount | `=B2*D2` | Price × Quantity |

**Key Concepts Used:**

| Concept | Detail |
|---------|--------|
| Relative `B2` | Changes automatically when the formula is copied down |
| Absolute `$C$2` | Stays fixed no matter where the formula is copied |

---

### 🧠 2. Logical Functions — `02_Nested_IF` & `03_IF_AND_OR`

**Sheet 02 – Grade assignment (Math marks)**
```excel
=IF(C2>=80,"A",IF(C2>=60,"B",IF(C2>=40,"C","Fail")))
```

| Marks | Grade |
|-------|-------|
| 80 and above | A |
| 60 – 79 | B |
| 40 – 59 | C |
| Below 40 | Fail |

**Sheet 02 – Discount slabs (second table)**
```excel
=IF(C25>=5000,C25*20%,IF(C25>=3000,C25*10%,0))
```

| Price | Discount |
|-------|----------|
| 5000 and above | 20% |
| 3000 – 4999 | 10% |
| Below 3000 | None |

**Sheet 03 – OR condition (Eligibility)**
```excel
=IF(OR(C2>=5000,D2>=10),"Eligible","Not Eligible")
```
> Eligible if the **price ≥ 5000** *or* the **quantity ≥ 10**.

**Sheet 03 – AND condition (Top performer)**
```excel
=IF(AND(C27>80,D27>80),"Yes","No")
```
> "Yes" only when **both** Math and Science are above 80.

---

### 📊 3. Conditional Aggregation — `04_COUNTIFS_SUMIFS`

| Result Label | Formula | Output |
|--------------|---------|--------|
| Count Math > 50 | `=COUNTIFS(C26:C45,">50")` | **17** |
| North Laptop Sales | `=SUMIFS(E2:E21,C2:C21,"North",D2:D21,"Laptop")` | **265000** |
| Average Math > 60 | `=AVERAGEIFS(C26:C45,C26:C45,">60")` | **79.07** |

The sheet holds two tables: a **sales table** (Region, Product, Sales) and a **marks table** (Math).

---

### 🔍 4. Lookup Functions — `05` to `08`, `10` to `11`

| Sheet | Function | Formula | Result |
|-------|----------|---------|--------|
| 05_VLOOKUP_Student | `VLOOKUP` | `=VLOOKUP(D14,A2:B21,2,FALSE)` | ID **102** → **Nikita** |
| 06_VLOOKUP_Product | `VLOOKUP` | `=VLOOKUP(G14,A2:C21,3,FALSE)` | Code **p111** → **5500** |
| 07_INDEX_MATCH | `INDEX + MATCH` | `=INDEX($C$2:$C$61,MATCH(F2&"\|"&G2,$D$2:$D$61,0))` | Rahul + Feb → **52000** |
| 08_Employee | `XLOOKUP` | `=XLOOKUP(J1,A2:A11,D2:D11,"Not Found")` | ID **205** → **48000** |
| 10_XLOOKUP | `XLOOKUP` | `=XLOOKUP(J12,A2:A11,C2:C11,"Not Found")` | ID **210** → **43000** |
| 10_XLOOKUP | `XLOOKUP` (2 criteria) | `=XLOOKUP(1,(E2:E11=J15)*(F2:F11=J16),G2:G11,"Not Found")` | rahul + Feb → **52000** |
| 11_XMATCH | `XMATCH` | `=XMATCH(E1,A2:A21,0)` | Tablet → position **4** |

**Two-condition lookup with a helper key (Sheet 07):**
```excel
=A2&"|"&B2          ← Key column D, e.g. "Rahul|Jan"
=INDEX($C$2:$C$61, MATCH(F2&"|"&G2, $D$2:$D$61, 0))
```

**Key Concepts Used:**

| Concept | Detail |
|---------|--------|
| `FALSE` (exact match) | Forces `VLOOKUP` to return only exact matches |
| Helper key | Joins two columns so `MATCH` can search on both at once |
| `XLOOKUP` not-found text | Built-in fallback message, no `IFERROR` needed |
| Boolean multiplication | `(cond1)*(cond2)` makes `XLOOKUP` act like `AND` |

---

### 🔤 5. Text Functions — `09_TEXT`

| Column | Formula | Example (`Aastha Patel`) |
|--------|---------|--------------------------|
| First Name | `=LEFT(A2,FIND(" ",A2)-1)` | `Aastha` |
| UPPER | `=UPPER(A2)` | `AASTHA PATEL` |
| LOWER | `=LOWER(A2)` | `aastha patel` |

> `FIND(" ",A2)` locates the space; `LEFT` then takes every character before it.

---

### 🧭 6. Dynamic References — `12_INDIRECT` & `13_OFFSET`

| Sheet | Formula | Input | Result |
|-------|---------|-------|--------|
| 12_INDIRECT | `=INDIRECT(E1)` | `E1 = "B5"` | **21000** |
| 13_OFFSET | `=OFFSET(B2,4,0)` | Start `B2`, move 4 rows down | **25000** |
| 13_OFFSET | `=SUM(OFFSET(B2,0,0,5,1))` | 5-row range from `B2` | **91000** |

**Key Concepts Used:**

| Function | Purpose |
|----------|---------|
| `INDIRECT` | Converts a text address (`"B5"`) into a real cell reference |
| `OFFSET` | Moves from a start cell by rows / columns, optionally sizing a range |

---

### 📅 7. Date & Math Functions — `14_Date_Time` & `15_Math`

**Sheet 14 – Age and day difference**
```excel
Age             =DATEDIF(B2,DATE(2026,9,30),"Y")
Days Difference =D2-C2
```

| Name | Age | Days Difference |
|------|-----|-----------------|
| Priya | 22 | 272 |
| Nikita | 23 | 241 |

**Sheet 08 – Years of service**
```excel
=DATEDIF(E2,DATE(2026,9,30),"Y")
```

**Sheet 15 – Rounding comparison (`125.678`)**

| Function | Formula | Result |
|----------|---------|--------|
| ROUND | `=ROUND(B2,2)` | 125.68 |
| CEILING | `=CEILING(B2,1)` | 126 |
| FLOOR | `=FLOOR(B2,1)` | 125 |

---

### 🧪 8. Dynamic Arrays — `16_FILTER`

```excel
=FILTER(A2:C21, C2:C21>80, "No students")
```

> A **single formula** in `F1` spills across `F1:H8` and returns every student scoring above 80.

| Student ID | Name | Score |
|-----------|------|-------|
| 101 | Aastha | 85 |
| 105 | Amit | 91 |
| 109 | Pooja | 88 |
| 111 | Meera | 83 |
| 114 | Anjali | 94 |
| 116 | Isha | 87 |
| 118 | Komal | 81 |
| 120 | Hetal | 90 |

---

## 🛠️ Tech Stack

| Tool | Version | Purpose |
|------|---------|---------|
| 📗 **Microsoft Excel** | 365 / 2021 or later | Workbook and formula engine |
| 🔗 **Cell References** | Built-in | Relative and absolute addressing |
| 🧠 **IF / AND / OR** | Built-in | Decision logic |
| 📊 **COUNTIFS / SUMIFS / AVERAGEIFS** | Built-in | Conditional summaries |
| 🔍 **VLOOKUP / INDEX / MATCH** | Built-in | Classic lookups |
| ✨ **XLOOKUP / XMATCH** | Excel 365 / 2021 | Modern lookups |
| 🔤 **LEFT / FIND / UPPER / LOWER** | Built-in | Text handling |
| 🧭 **INDIRECT / OFFSET** | Built-in | Dynamic references |
| 📅 **DATEDIF / DATE** | Built-in | Date calculations |
| 🔢 **ROUND / CEILING / FLOOR** | Built-in | Number rounding |
| 🧪 **FILTER** | Excel 365 / 2021 | Dynamic array output |

---

## 📈 Sample Results

Key outputs captured from the workbook:

```
01_Relative_Absolute   Aastha : Price 7000, Qty 12  →  Tax 1260 | Final 8260 | Total 84000
02_Nested_IF           Aastha : Math 75             →  Grade B
02_Nested_IF           Aastha : Price 6000          →  Discount 1200
04_COUNTIFS_SUMIFS     Count Math > 50              →  17
04_COUNTIFS_SUMIFS     North Laptop Sales           →  265000
04_COUNTIFS_SUMIFS     Average Math > 60            →  79.07
05_VLOOKUP_Student     ID 102                       →  Nikita
06_VLOOKUP_Product     Code p111                    →  5500
07_INDEX_MATCH         Rahul | Feb                  →  52000
08_Employee            ID 205                       →  Salary 48000 | Joined 30-Jun-2020
10_XLOOKUP             ID 210                       →  43000
11_XMATCH              Tablet                       →  Position 4
12_INDIRECT            B5                           →  21000
13_OFFSET              OFFSET(B2,4,0)               →  25000
13_OFFSET              SUM of 5 rows                →  91000
14_Date_Time           Priya                        →  Age 22 | 272 days
16_FILTER              Score > 80                   →  8 students returned
```

---

## ⚠️ Known Issues & Notes

| # | Note | Detail |
|---|------|--------|
| 1 | **Excel version** | `XLOOKUP`, `XMATCH` and `FILTER` need Excel 365 or 2021+; older versions show `#NAME?` |
| 2 | **Fixed reference date** | `DATEDIF` uses `DATE(2026,9,30)` instead of `TODAY()`, so ages do not update automatically |
| 3 | **Volatile functions** | `INDIRECT` and `OFFSET` recalculate on every change and can slow very large workbooks |
| 4 | **Spill range** | `FILTER` output in `F1:H8` needs empty cells around it, otherwise it returns `#SPILL!` |
| 5 | **Case sensitivity** | Lookups are not case-sensitive, so `p111` finds `P111` and `rahul` finds `Rahul` |
| 6 | **Trailing space** | The first name in sheet 01 is stored as `"Aastha "` (with a space), which may affect exact-match lookups |
| 7 | **Hardcoded ranges** | Formulas use fixed ranges (e.g. `A2:A21`) and will not expand automatically when rows are added; Excel Tables would solve this |
| 8 | **Row-wise copy** | Sheet 07 uses `$` anchors on ranges, while sheets 05–06 do not, so copying those formulas down will shift the range |

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🎓 **Beginner Friendly** | One concept per sheet with small, readable datasets |
| 📚 **Wide Coverage** | 16 topics from basic references to dynamic arrays |
| 🔄 **Classic vs Modern** | Shows `VLOOKUP` and `INDEX+MATCH` next to `XLOOKUP` and `XMATCH` |
| 🔍 **Easy to Trace** | Formulas sit beside their data, so every result can be checked by hand |
| 🧪 **Extensible** | Easily upgraded with `IFERROR`, `TODAY()`, Tables and Pivot Tables |
| 💻 **No Macros** | Pure formulas, no VBA or add-ins required |
| ⚡ **Lightweight** | A single `.xlsx` file that opens instantly |

---

## 🚀 How to Use

1. Download or clone the repository.
2. Open `Project_1_excel.xlsx` in **Microsoft Excel 365 / 2021 or later**.
3. Move through sheets `01` to `16` in order.
4. Click any formula cell and read it in the **formula bar**.
5. Change an input (e.g. the search ID in `05_VLOOKUP_Student!D14`) and watch the result update.
6. Use **Formulas → Evaluate Formula** to step through nested logic.

---

## 📄 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for full details.

```
MIT License — Free to use, modify, and distribute with attribution.
```

---

## 👤 Author

<div align="center">

### Priya Vadsak

[![GitHub](https://img.shields.io/badge/GitHub-23se02ml213--design-181717?style=for-the-badge&logo=github&logoColor=white)](https://github.com/23se02ml213-design)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/priya-vadsak-4347294300)

> *"Every formula you master turns hours of manual work into a single cell."*

**🎓 Role:** Python Learner | Data & Spreadsheet Enthusiast \
**📍 Location:** India \
**🛠️ Skills:** Excel Formulas · Lookups · Conditional Logic · Data Analysis · Python

</div>

---

## 🙏 Acknowledgements

Special thanks to the following resources and communities that made this project possible:

- 📗 [Microsoft Excel Support](https://support.microsoft.com/excel) — Official function reference
- 🔍 [XLOOKUP Documentation](https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929) — Modern lookup guide
- 🧪 [FILTER Function Guide](https://support.microsoft.com/en-us/office/filter-function-f4f7cb66-82eb-4767-8f7c-4877ad80c759) — Dynamic array reference
- 🎓 [ExcelJet](https://exceljet.net/) — Clear formula explanations and examples
- 🖥️ [W3Schools Excel](https://www.w3schools.com/excel/) — Beginner Excel reference
- 💬 [Stack Overflow Community](https://stackoverflow.com/) — Problem-solving support

---

<div align="center">

---

*Made with ❤️ and ☕ — Last updated: October 2026*

</div>
