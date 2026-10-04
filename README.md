# PIVC Bypass Analysis — Power BI Dashboard Solution

Welcome to the **PIVC Bypass Analysis Dashboard** project. This repository contains the complete production-grade toolkit to build, test, and present an executive Power BI dashboard for tracking Physical/Video KYC (PIVC) bypass exceptions.

---

## 🚀 Quick Launch

- **Double-click [index.html](file:///c:/Users/dhair/OneDrive/Desktop/Dashboard%20for%20PIVC%20bypass/index.html)** in any browser (Chrome, Edge) to launch the **Live Interactive Executive Dashboard**.
- Test all 4 slicers, real-time KPI updates, 8 interactive visuals (Pie Chart Studio, 100% Stacked Bar Charts, Multi-Metric Line Graphs, Daily Timeline), and paginated search drill-downs!

---

## 📦 Project Contents

1. [**`index.html`**](file:///c:/Users/dhair/OneDrive/Desktop/Dashboard%20for%20PIVC%20bypass/index.html) — Interactive Executive Web Dashboard simulator matching the Power BI layout.
2. [**`Executive_PowerBI_Dashboard_Guide.md`**](file:///c:/Users/dhair/OneDrive/Desktop/Dashboard%20for%20PIVC%20bypass/Executive_PowerBI_Dashboard_Guide.md) — Comprehensive, visual-by-visual Power BI Desktop building guide.
3. [**`DAX_Measures_and_Calculations.dax`**](file:///c:/Users/dhair/OneDrive/Desktop/Dashboard%20for%20PIVC%20bypass/DAX_Measures_and_Calculations.dax) — Production-ready DAX for calculated columns, date table, and required measures.
4. [**`PowerQuery_M_Code.m`**](file:///c:/Users/dhair/OneDrive/Desktop/Dashboard%20for%20PIVC%20bypass/PowerQuery_M_Code.m) — Upstream ETL script for Power Query Advanced Editor.
5. [**`PIVC_bypass_500_records.csv`**](file:///c:/Users/dhair/OneDrive/Desktop/Dashboard%20for%20PIVC%20bypass/PIVC_bypass_500_records.csv) — 501 real records with `APPROVAL_SUBSTATUS` (`Exceptional` vs `DPIVC DONE`), `REMARKS`, `USER_ID`, `APPLICATION_NO`, and `REQUEST_NO`.
6. [**`pivc_theme.json`**](file:///c:/Users/dhair/OneDrive/Desktop/Dashboard%20for%20PIVC%20bypass/pivc_theme.json) — Executive color palette and visual styling file for Power BI Desktop (*View > Themes > Browse for themes*).

---

## 📊 Core DAX Formulas Summary

### 1. Calculated Column: Bypass Flag (Dual-Condition)
```dax
Bypass Flag = 
IF (
    CONTAINSSTRING ( 'PIVC_Data'[APPROVAL_SUBSTATUS], "Except" )
        || CONTAINSSTRING ( 'PIVC_Data'[REMARKS], "Bypass" ),
    "Bypass",
    "Regular"
)
```
*Logic: A record is a **Bypass** if `APPROVAL_SUBSTATUS` contains `"Except"` (matching `"Exceptional"` or `"Exception"`) OR if `REMARKS` contains `"Bypass"` (matching `"Bypass request-AS-..."`). When `APPROVAL_SUBSTATUS` is `"DPIVC DONE"` or blank, it evaluates cleanly to `"Regular"` (System Done).*

### 2. Calculated Columns: Month-Year & Chronological Sorting
```dax
Month-Year = 
FORMAT ( 'PIVC_Data'[APPROVAL_DATE], "mmm yyyy" )

Month-Year Sort = 
YEAR ( 'PIVC_Data'[APPROVAL_DATE] ) * 100 + MONTH ( 'PIVC_Data'[APPROVAL_DATE] )
```
*In Power BI Desktop: Select `Month-Year` -> Column Tools -> **Sort by Column** -> choose `Month-Year Sort`.*

### 3. Measures
```dax
Total Applications = 
DISTINCTCOUNT ( 'PIVC_Data'[APPLICATION_NO] )

Total Bypass Applications = 
CALCULATE (
    DISTINCTCOUNT ( 'PIVC_Data'[APPLICATION_NO] ),
    'PIVC_Data'[Bypass Flag] = "Bypass"
)

Bypass Percentage = 
DIVIDE (
    [Total Bypass Applications],
    [Total Applications],
    0
)
```
