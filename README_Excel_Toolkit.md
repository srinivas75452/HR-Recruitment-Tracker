# 📊 Advanced Excel Toolkit — XLOOKUP · INDEX-MATCH · Power Query · TAT Tracker

> A practical reference workbook demonstrating **advanced Excel skills** used daily in MIS, HR Analytics, and Operations reporting.  
> Every formula is live, editable, and annotated with plain-English explanations.

---

## 📌 Why This Exists

VLOOKUP is the most searched Excel function — but it's also the most limited.  
This toolkit shows the **modern Excel stack** that powers real operations dashboards:
- `XLOOKUP` replacing VLOOKUP/HLOOKUP
- `INDEX-MATCH` for multi-condition and left-lookups
- **Power Query** for repeatable data cleaning
- **TAT Templates** for SLA breach monitoring (used in warehouse operations across 23 states)

---

## 📂 Sheets Overview

| Sheet | Focus | Skill |
|-------|-------|-------|
| `Index` | Contents + skill map | — |
| `1_XLOOKUP_Advanced` | 6 XLOOKUP patterns with live source data | Advanced Excel |
| `2_INDEX_MATCH` | 5 INDEX-MATCH patterns vs VLOOKUP | Advanced Excel |
| `3_Power_Query_Clean` | Dirty vs Clean data side-by-side + M code steps | Power Query |
| `5_MIS_TAT_Template` | Delivery TAT tracker with SLA alerts | MIS Reporting |

---

## 🔍 XLOOKUP Patterns Covered

| Pattern | Formula Highlight | Use Case |
|---------|------------------|----------|
| Basic lookup | `=XLOOKUP(id, A:A, B:B, "Not Found")` | Replace VLOOKUP |
| Wildcard match | `match_mode=2` with `"Priya*"` | Partial name search |
| Reverse / last match | `search_mode=-1` | Find most recent record |
| Nested (2D) | `XLOOKUP` inside `XLOOKUP` | Row + column lookup |
| With aggregation | `MAX(XLOOKUP(...))` | Highest salary by grade |
| Array formula | `AVERAGE(IF(...))` | Conditional average |

---

## 🔍 INDEX-MATCH Patterns Covered

| Pattern | Use Case |
|---------|----------|
| Basic INDEX-MATCH | Left-side lookup (impossible with VLOOKUP) |
| Two-condition match | Brand + Category simultaneously |
| 2D dynamic lookup | Change column number to pull any field |
| MIN/MAX item finder | Which product has lowest stock? |
| Approximate match banding | Tier data into Low / Medium / High |

---

## 🧹 Power Query — Dirty → Clean Demo

**Before (dirty data problems):**
- Extra spaces in IDs and names
- Inconsistent case: `HYDERABAD`, `hyderabad`, `Hyderabad`
- Salary as text with commas: `"65,000"`
- Boolean as mixed: `1`, `TRUE`, `yes`, `FALSE`, `no`
- Missing values shown as `N/A` instead of null
- Date formats inconsistent: `01-jan-2022`, `2022/02/15`, `04/04/2022`

**Power Query M Code Steps documented** (8 transformation steps shown in sheet)

---

## 📦 TAT Template — SLA Breach Monitor

Based on real MIS work managing 35,000+ SKUs across 1,500+ brands at KDL:

- Calculates actual vs expected delivery TAT
- Flags **SLA Breach** (Yes/No) automatically
- Classifies delay severity: `On Time` / `Delayed` / `Critical Delay`
- Priority flag: `HIGH` / `MEDIUM` / `LOW` based on delay days
- Summary block: On-Time %, Avg TAT, Critical breach count

---

## 🛠️ Skills Demonstrated

```
Excel Functions:  XLOOKUP, INDEX, MATCH, IFERROR, COUNTIF, AVERAGEIF,
                  MAX, MIN, DAYS, IF, TEXT, COUNTA, AVERAGE
Advanced:         Array formulas (Ctrl+Shift+Enter), Wildcard matching,
                  Approximate match banding, 2D lookup
Power Query:      Text.Trim, Text.Proper, Number.From, Table.ReplaceValue,
                  Table.RemoveRowsWithErrors, Table.SelectRows (M Language)
MIS:              TAT calculation, SLA breach detection, priority flagging,
                  summary KPI blocks
Design:           Color-coded tables, alternating rows, professional headers,
                  data validation dropdowns
```

---

## 🚀 How to Use

1. Download `Excel_Advanced_Toolkit.xlsx`
2. Open in Excel (no macros needed — pure formulas)
3. Start at the `Index` sheet to navigate
4. **Change the lookup values** in yellow cells to see formulas react live
5. Use the Power Query sheet as a reference for your own data cleaning projects

---

## 👤 Author

**Srinivas G** — Data & Business Operations Analyst  
Hyderabad, India | [srinivas75452@gmail.com](mailto:srinivas75452@gmail.com)  
Skills: Excel · Power BI · VBA · Power Query · MIS Reporting · HR Analytics

---

*All data is fictional/sample. Formulas are fully functional and editable.*
