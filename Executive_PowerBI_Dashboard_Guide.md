# PIVC Bypass Analysis — Professional Power BI Dashboard Guide

This comprehensive guide walks you through building an **Executive Management Power BI Dashboard** for tracking **PIVC (Physical / Video KYC) Bypass Applications**, analyzing monthly trends, monitoring user performance, and evaluating bypass percentages.

---

## 📁 Project Deliverables in Your Workspace

All necessary files and ready-to-use assets have been created in your project folder:

| File Name | Purpose |
| :--- | :--- |
| [**`index.html`**](file:///c:/Users/dhair/OneDrive/Desktop/Dashboard%20for%20PIVC%20bypass/index.html) | **Interactive Executive Dashboard Web App**: Open in Chrome/Edge to test all visuals, slicers, and drill-downs instantly! |
| [**`pivc_bypass_dataset.csv`**](file:///c:/Users/dhair/OneDrive/Desktop/Dashboard%20for%20PIVC%20bypass/pivc_bypass_dataset.csv) | Realistic sample dataset (150+ records across 9 months) matching your exact 6 columns. |
| [**`DAX_Measures_and_Calculations.dax`**](file:///c:/Users/dhair/OneDrive/Desktop/Dashboard%20for%20PIVC%20bypass/DAX_Measures_and_Calculations.dax) | Clean, copy-paste DAX formulas for calculated columns, date dimension, and core measures. |
| [**`PowerQuery_M_Code.m`**](file:///c:/Users/dhair/OneDrive/Desktop/Dashboard%20for%20PIVC%20bypass/PowerQuery_M_Code.m) | Power Query transformation script for automated ETL and high-compression ingestion. |
| [**`pivc_theme.json`**](file:///c:/Users/dhair/OneDrive/Desktop/Dashboard%20for%20PIVC%20bypass/pivc_theme.json) | Executive Power BI Theme palette (Navy `#0F2042`, Coral Bypass `#D83B01`, Cobalt `#0078D4`). |

---

## 1. Data Schema & Business Rule

Your input dataset contains the following 6 fields:
- `APPLICATION_NO`: Unique application identifier
- `APPROVAL_DATE`: Date of application approval/review
- `APPROVAL_STATUS`: Status (`Approved`, `Rejected`, `Pending`, etc.)
- `USER_ID`: Identifier of the agent/user submitting or handling the application
- `ASSIGNED_TO`: Compliance / Operations analyst assigned
- `REMARKS`: Case notes and remarks

### 📌 Core Business Rule:
> **Any record where the `REMARKS` column contains the text `"Bypass request"` is classified as a PIVC Bypass case.**

---

## 2. Calculated Columns & Date Hierarchy

### A. Calculated Column: `Bypass Flag` (Dual-Condition Identification)
In Power BI Desktop, navigate to **Data View** or right-click your table (`PIVC_Data`) and select **New Column**:

```dax
Bypass Flag = 
IF (
    CONTAINSSTRING ( 'PIVC_Data'[APPROVAL_SUBSTATUS], "Except" )
        || CONTAINSSTRING ( 'PIVC_Data'[REMARKS], "Bypass" ),
    "Bypass",
    "Regular"
)
```
*Business Rule Notes for `PIVC_bypass_500_records.csv`:*
- **Exception Substatus**: When a case is bypassed, `APPROVAL_SUBSTATUS` is `"Exceptional"` (or `"Exception"`).
- **System Automated KYC**: When PIVC is successfully done through the automated system, `APPROVAL_SUBSTATUS` is `"DPIVC DONE"` (or blank/null), which safely evaluates to `"Regular"`.
- **Remarks Match**: Catches bypass remarks such as `"Bypass request-AS-715441"`.
- *Note:* If your column in Power BI is named `APPROVAL_STATUS` rather than `APPROVAL_SUBSTATUS`, simply use `'PIVC_Data'[APPROVAL_STATUS]`.

---

### B. Calculated Column: `Month-Year` and Chronological Sorting
To ensure months appear chronologically (e.g. `Jan 2024 -> Feb 2024 -> Mar 2024`) instead of alphabetical order (`Apr, Aug, Dec, Feb...`), create **two columns**:

#### 1. Display Column:
```dax
Month-Year = 
FORMAT ( 'PIVC_Data'[APPROVAL_DATE], "mmm yyyy" )
```

#### 2. Sort Key Column:
```dax
Month-Year Sort = 
YEAR ( 'PIVC_Data'[APPROVAL_DATE] ) * 100 + MONTH ( 'PIVC_Data'[APPROVAL_DATE] )
```

#### 3. Crucial Configuration Step in Power BI Desktop:
1. In the **Fields / Data pane**, click on **`Month-Year`**.
2. Click the **Column tools** tab on the top ribbon.
3. Click **Sort by column** in the ribbon.
4. Select **`Month-Year Sort`** from the dropdown.
   *(Now all charts, tables, and slicers using `Month-Year` will sort January through December chronologically!)*

---

### C. (Recommended) Standalone Enterprise Date Dimension: `Dim_Date`
For enterprise-grade time-intelligence without circular dependencies, go to **Modeling > New Table**:

```dax
Dim_Date = 
VAR MinDate = MIN ( 'PIVC_Data'[APPROVAL_DATE] )
VAR MaxDate = MAX ( 'PIVC_Data'[APPROVAL_DATE] )
RETURN
    ADDCOLUMNS (
        CALENDAR ( DATE ( YEAR ( MinDate ), 1, 1 ), DATE ( YEAR ( MaxDate ), 12, 31 ) ),
        "Year", YEAR ( [Date] ),
        "Month Number", MONTH ( [Date] ),
        "Month Name", FORMAT ( [Date], "mmmm" ),
        "Month Short", FORMAT ( [Date], "mmm" ),
        "Month-Year", FORMAT ( [Date], "mmm yyyy" ),
        "Month-Year Sort", YEAR ( [Date] ) * 100 + MONTH ( [Date] )
    )
```
*Relationship: In Model View, drag `Dim_Date[Date]` (1) to `'PIVC_Data'[APPROVAL_DATE]` (\*).*

---

## 3. Required DAX Measures

Create a dedicated Measures Table (e.g., `_Measures`) or add to `'PIVC_Data'`:

### Measure 1: Total Applications
```dax
Total Applications = 
DISTINCTCOUNT ( 'PIVC_Data'[APPLICATION_NO] )
```

### Measure 2: Total Bypass Applications
```dax
Total Bypass Applications = 
CALCULATE (
    DISTINCTCOUNT ( 'PIVC_Data'[APPLICATION_NO] ),
    'PIVC_Data'[Bypass Flag] = "Bypass"
)
```

### Measure 3: Bypass Percentage
```dax
Bypass Percentage = 
DIVIDE (
    [Total Bypass Applications],
    [Total Applications],
    0
)
```
*Formatting Tip: Select the measure, go to the **Measure tools** ribbon, and set format to **Percentage** with **1 decimal place** (`0.0%`).*

### Measure 4: Total Regular Applications (Supporting Metric)
```dax
Total Regular Applications = 
CALCULATE (
    DISTINCTCOUNT ( 'PIVC_Data'[APPLICATION_NO] ),
    'PIVC_Data'[Bypass Flag] = "Regular"
)
```

---

## 4. Visual Layout Blueprint

Follow this layout grid on a standard **16:9 canvas** (1280x720 or 1920x1080):

```
+----------------------------------------------------------------------------------------------------+
|  [Logo] PIVC Bypass Analysis Dashboard | Executive Risk & Operations Monitor   [Date: Today]       |
+----------------------------------------------------------------------------------------------------+
|  SLICERS: [ Month-Year v ]   [ Approval Status v ]   [ User ID v ]   [ Assigned To v ]             |
+----------------------------------------------------------------------------------------------------+
|  +-------------------+  +-------------------+  +-------------------+  +-------------------------+  |
|  | TOTAL APPS        |  | BYPASS APPS       |  | BYPASS %          |  | REGULAR APPS            |  |
|  |      152          |  |       52          |  |     34.2%         |  |         100             |  |
|  +-------------------+  +-------------------+  +-------------------+  +-------------------------+  |
+----------------------------------------------------------------------------------------------------+
|  +---------------------------------------+  +---------------------------------------------------+  |
|  | Visual 1: Bypass Count by Month       |  | Visual 2: Monthly Bypass Trend                    |  |
|  | (Clustered Column Chart)              |  | (Line & Area Chart)                               |  |
|  +---------------------------------------+  +---------------------------------------------------+  |
+----------------------------------------------------------------------------------------------------+
|  +---------------------------------------+  +---------------------------------------------------+  |
|  | Visual 3: User-wise Bypass Count      |  | Visual 4: Assignee-wise Bypass Count              |  |
|  | (Horizontal Clustered Bar Chart)      |  | (Horizontal Clustered Bar Chart)                  |  |
|  +---------------------------------------+  +---------------------------------------------------+  |
+----------------------------------------------------------------------------------------------------+
|  +---------------------------------------+  +---------------------------------------------------+  |
|  | Visual 5: Bypass vs Non-Bypass        |  | Visual 6: Monthly Executive Summary Matrix        |  |
|  | (Donut Chart)                         |  | (Matrix / Table Visual with Data Bars)            |  |
|  +---------------------------------------+  +---------------------------------------------------+  |
+----------------------------------------------------------------------------------------------------+
```

---

## 5. Visual-by-Visual Step-by-Step Configuration

### 1. Top KPI Cards
- **Visual Type**: **Card** (or **New Card** visual in modern Power BI)
- **Card 1**: Field = `[Total Applications]`
  - Callout Value Font: Segoe UI Bold, size 28, Color: Navy `#0F2042`
  - Category Label: "Total Applications"
- **Card 2**: Field = `[Total Bypass Applications]`
  - Callout Value Font: Segoe UI Bold, size 28, Color: Coral `#D83B01`
  - Category Label: "Total Bypass Applications"
- **Card 3**: Field = `[Bypass Percentage]`
  - Callout Value Font: Segoe UI Bold, size 28, Color: Orange/Amber `#D83B01`
  - Category Label: "Bypass Percentage"

---

### 2. Visual 1: Column Chart — Bypass Count by Month
- **Visual Type**: **Clustered Column Chart**
- **X-Axis**: `'PIVC_Data'[Month-Year]`
- **Y-Axis**: `[Total Bypass Applications]`
- **Formatting**:
  - Columns Color: `#D83B01` (Coral Accent)
  - Data Labels: **On** (Position: Outside end)
  - Sort Axis: Sort by `Month-Year` -> **Sort ascending** (uses the `Month-Year Sort` key automatically!)

---

### 3. Visual 2: Line Chart — Monthly Bypass Trend
- **Visual Type**: **Line Chart** (or Line and Clustered Column Chart)
- **X-Axis**: `'PIVC_Data'[Month-Year]`
- **Y-Axis**: `[Bypass Percentage]`
- **Formatting**:
  - Line Color: `#0078D4` (Cobalt Blue), Stroke width: 3px
  - Markers: **On** (Circle, 5px, White fill with blue stroke)
  - Data Labels: **On** (Formatted as `0.0%`)
  - Y-Axis bounds: Minimum = 0, Maximum = 0.5 (or auto)

---

### 4. Visual 3: Bar Chart — User-wise Bypass Count
- **Visual Type**: **Clustered Bar Chart** (Horizontal)
- **Y-Axis**: `'PIVC_Data'[USER_ID]`
- **X-Axis**: `[Total Bypass Applications]`
- **Formatting**:
  - Sort Axis: Sort by `Total Bypass Applications` -> **Sort descending**
  - Bar Color: `#0078D4`
  - Data Labels: **On**

---

### 5. Visual 4: Bar Chart — Assignee-wise Bypass Count
- **Visual Type**: **Clustered Bar Chart** (Horizontal)
- **Y-Axis**: `'PIVC_Data'[ASSIGNED_TO]`
- **X-Axis**: `[Total Bypass Applications]`
- **Formatting**:
  - Sort Axis: Sort by `Total Bypass Applications` -> **Sort descending**
  - Bar Color: `#D83B01`
  - Data Labels: **On**

---

### 6. Visual 5: Donut Chart — Bypass vs Non-Bypass Applications
- **Visual Type**: **Donut Chart**
- **Legend**: `'PIVC_Data'[Bypass Flag]`
- **Values**: `[Total Applications]` (or Distinct Count of `APPLICATION_NO`)
- **Colors**:
  - `"Bypass"`: `#D83B01` (Coral)
  - `"Regular"`: `#0078D4` (Blue)
- **Detail Labels**: Category + Percent of Total

---

### 7. Visual 6: Matrix / Table Visual
- **Visual Type**: **Matrix**
- **Rows**: `'PIVC_Data'[Month-Year]`
- **Values**:
  1. `[Total Applications]`
  2. `[Total Bypass Applications]`
  3. `[Bypass Percentage]`
- **Formatting**:
  - Style: **Bold header**
  - Background color: Alternating rows
  - Conditional Formatting: Under `[Bypass Percentage]`, enable **Data bars** (positive bar color: Coral `#D83B01`).

---

### 8. Slicers Configuration
Add 4 slicers along the top banner:
1. **Month-Year**: Slicer visual -> Field: `'PIVC_Data'[Month-Year]` (Style: Dropdown or Horizontal Tiles).
2. **Approval Status**: Slicer visual -> Field: `'PIVC_Data'[APPROVAL_STATUS]` (Style: Dropdown).
3. **User ID**: Slicer visual -> Field: `'PIVC_Data'[USER_ID]` (Style: Dropdown with search enabled).
4. **Assigned To**: Slicer visual -> Field: `'PIVC_Data'[ASSIGNED_TO]` (Style: Dropdown with search enabled).

---

## 6. How to Import the Executive Theme into Power BI Desktop

1. Open Power BI Desktop.
2. In the top ribbon, click the **View** tab.
3. Click the dropdown arrow next to the **Themes** gallery.
4. Click **Browse for themes**.
5. Select [`pivc_theme.json`](file:///c:/Users/dhair/OneDrive/Desktop/Dashboard%20for%20PIVC%20bypass/pivc_theme.json).
6. Power BI Desktop will automatically apply the corporate palette, rounded visual corners, and font hierarchy!

---

## 7. Immediate Testing via the Interactive Simulator

To preview how your executive dashboard will look and behave with full data filtering:
1. Double-click [**`index.html`**](file:///c:/Users/dhair/OneDrive/Desktop/Dashboard%20for%20PIVC%20bypass/index.html) in your workspace folder.
2. It opens directly in Google Chrome or Microsoft Edge.
3. You can interactively test:
   - Selecting any month from the **Month-Year** dropdown.
   - Filtering by **Approval Status**, **User ID**, or **Assigned To**.
   - Reviewing the instant recalculation of all KPIs and charts.
   - Clicking **"Power BI DAX Code"** for instant copy-pasting.
   - Clicking **"Export CSV"** to extract filtered audit records.
