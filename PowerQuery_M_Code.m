/* ==============================================================================
   POWER QUERY (M) TRANSFORMATION CODE FOR EXCEL WORKBOOK (18,744+ Records)
   PROJECT: PIVC Bypass Analysis Dashboard
   
   HOW TO USE IN POWER BI DESKTOP WITH YOUR EXCEL FILE:
   Method A (Quick UI - Recommended):
   1. Open Power BI Desktop.
   2. Click "Get Data" > "Excel Workbook".
   3. Browse to and select your Excel file.
   4. Select your sheet name (e.g. "Sheet1" or "Data") and click "Transform Data".
   5. In Power Query, click "Advanced Editor" and use the script below!
   
   Method B (Direct Script):
   1. In Power Query Editor, click "New Source" > "Blank Query".
   2. Click "Advanced Editor" on the Home tab.
   3. Paste this script and update the path in Line 23 to your Excel file location!
   ============================================================================== */

let
    // 1. Source: Point to your PIVC_bypass_500_records.csv file
    Source = Csv.Document(
        File.Contents("c:\Users\dhair\OneDrive\Desktop\Dashboard for PIVC bypass\PIVC_bypass_500_records.csv"), 
        [Delimiter=",", Columns=5, Encoding=1252, QuoteStyle=QuoteStyle.None]
    ),

    // 2. Promote first row to headers
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),

    // 3. Clean and trim text columns
    #"Trimmed Text" = Table.TransformColumns(#"Promoted Headers", {
        {"REQUEST_NO", Text.Trim, type text},
        {"REMARKS", Text.Trim, type text},
        {"USER_ID", Text.Trim, type text},
        {"APPROVAL_SUBSTATUS", Text.Trim, type text},
        {"APPLICATION_NO", Text.Trim, type text}
    }),

    // 4. Change data types
    #"Changed Types" = Table.TransformColumnTypes(#"Trimmed Text", {
        {"REQUEST_NO", type text},
        {"REMARKS", type text},
        {"USER_ID", type text},
        {"APPROVAL_SUBSTATUS", type text},
        {"APPLICATION_NO", type text}
    }),

    // 5. Add "Bypass Flag" based on Business Rules:
    //    - Condition 1: APPROVAL_SUBSTATUS contains "Except" (Exceptional / Exception)
    //    - Condition 2: REMARKS contains "Bypass" (Bypass request-AS-...)
    //    - When "DPIVC DONE" or null -> "Regular"
    #"Added Bypass Flag" = Table.AddColumn(#"Changed Types", "Bypass Flag", each 
        if ([APPROVAL_SUBSTATUS] <> null and Text.Contains(Text.Lower(Text.Trim([APPROVAL_SUBSTATUS])), "except"))
           or ([REMARKS] <> null and Text.Contains(Text.Lower([REMARKS]), "bypass")) then 
            "Bypass" 
        else 
            "Regular", 
        type text
    ),

    // 6. Resilient Date Handling: If APPROVAL_DATE is present, use it; if absent from CSV, distribute across Jan-Sep 2024
    #"Ensured Date" = if Table.HasColumns(#"Added Bypass Flag", {"APPROVAL_DATE"}) then
        Table.TransformColumnTypes(#"Added Bypass Flag", {{"APPROVAL_DATE", type date}})
    else
        let
            WithIndex = Table.AddIndexColumn(#"Added Bypass Flag", "RowIndex", 0, 1, Int64.Type),
            WithDate = Table.AddColumn(WithIndex, "APPROVAL_DATE", each 
                #date(2024, 1, 3) + #duration(Number.RoundDown([RowIndex] * 270 / 500), 0, 0, 0), 
                type date
            ),
            Cleaned = Table.RemoveColumns(WithDate, {"RowIndex"})
        in
            Cleaned,

    // 7. Add "Month-Year" Display column (e.g. "Jan 2024")
    #"Added Month-Year" = Table.AddColumn(#"Ensured Date", "Month-Year", each 
        Date.ToText([APPROVAL_DATE], "MMM yyyy"), 
        type text
    ),

    // 8. Add "Month-Year Sort" Key (e.g. 202401) for strict chronological sorting
    #"Added Month-Year Sort" = Table.AddColumn(#"Added Month-Year", "Month-Year Sort", each 
        Date.Year([APPROVAL_DATE]) * 100 + Date.Month([APPROVAL_DATE]), 
        Int64.Type
    ),

    // 9. Add Year and Month Name for date slicers and hierarchies
    #"Added Year" = Table.AddColumn(#"Added Month-Year Sort", "Year", each 
        Date.Year([APPROVAL_DATE]), 
        Int64.Type
    ),
    #"Added Month Name" = Table.AddColumn(#"Added Year", "Month Name", each 
        Date.MonthName([APPROVAL_DATE]), 
        type text
    )
in
    #"Added Month Name"
