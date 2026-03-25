# Quotalyze: Features and Constraints

## Core Features

1. **Tender & Project Management**
   - Session-based workspace setup allowing multiple distinct projects (tenders) to be processed cleanly without files overlapping.
   - Simple UI to clear sessions and start over.

2. **Automated Quotation Alignment**
   - Users can upload one blank "Reference" quotation and multiple quoted "Vendor" Excel files.
   - The application automatically aligns each vendor's `Rate` and `Amount` to the corresponding items in the reference document.

3. **Intelligent Amount Verification & Math Correction**
   - The system automatically verifies the math for each line item (`Rate * Qty = Amount`).
   - If a vendor provides an incorrect amount, the system automatically corrects it using the correct math and logs the discrepancy in an error list displayed on the frontend.
   - Detects and strictly ignores non-mathematical or text-based quantities (e.g. "L.S.").

4. **Missing Totals Backfilling**
   - Automatically detects "Total" or "Grand Total" rows at the bottom of the BOQ. 
   - If a vendor leaves the grand total blank, the system automatically sums everything above it and plugs in the calculated total.

5. **Competitor Analytics & Visual Highlighting**
   - Analyzes rates item-by-item to determine the Minimum (Lowest) and Maximum (Highest) prices.
   - Expand analytics to Subtotals and Grand Totals by tracking mathematically lowest/highest **Amount** cells.
   - Populates two dedicated columns in the report: `Lowest_Vendor` and `Highest_Vendor`.
   - Modifies the output template to visually highlight the lowest values in **Light Green** and the highest values in **Light Coral (Red)**.

6. **Incomplete Quote Detection**
   - Validates whether any vendor has skipped pricing for necessary items.
   - Throws a direct "Partial Quotations" alert summarizing exactly which vendors failed to quote all required items.

7. **Native Template Overlaying & Export**
   - Avoids generating generic, ugly spreadsheets. Instead, it overlays the final analysis data directly onto a clone of the original Reference Excel file. 
   - Perfectly preserves the original client headers, footers, logos, and custom row formatting.

8. **Automated XML Scrubbing (Anti-Corruption Engine)**
   - Safely intercepts and scrubs uploaded `.xlsx` files to remove corrupted or hidden `<definedNames>` blocks inside the workbook XML.
   - This actively prevents the fatal `Invalid XML` crashes common when files are converted via "Save As" from ancient `.xls` formats.

9. **Estimated Rate Comparison System**
   - Allows users to optionally upload an "Estimated Rate" file.
   - Automatically compares each vendor's rates mathematically against the internal estimate.
   - Populates a `Closest_to_Estimate` column and actively highlights the closest vendor's cell in bright **Yellow** (which overrides Red/Green coloring).

---

## Application Constraints & Strict Rules

1. **File Format Restriction (`.xlsx` Only)**
   - The system exclusively supports the modern Office Open XML format (`.xlsx`). 
   - Uploading older Excel 97-2003 `.xls` files is prohibited, and users must use Excel's "Save As" function to convert them before uploading.

2. **No Formulas or External References**
   - Reference and Vendor files **must contain values only**. 
   - If there are formulas, macros, or references linking to external workbooks or hidden internal sheets, the application's XML parsers will fail to read the file correctly.

3. **Strict Table Header Requirements**
   - The application's alignment engine relies directly on detecting predefined column headers.
   - The main table within the reference file must strictly contain standard columns like `S.No`, `Description`, `Qty`, and `Unit`. A severe spelling error or completely deviated naming structure will cause the analysis to crash.

4. **Vendor Naming Convention**
   - The application does not parse inside the file to find out who the vendor is.
   - The **name of the uploaded Excel file** itself dictates the vendor's identity. (e.g., A file named `Aster Plumbing.xlsx` will cause "Aster Plumbing" to appear as the header on the final report).

5. **Partial Analysis Limitations**
   - Items are considered valid for pricing verification only if a valid numeric Quantity exists. If a Quantity is textual but not standard, the app will skip math verification for that specific row.
   - The application cannot compare items that are completely absent from the Reference file but added arbitrarily by the Vendor.

6. **Blank Column Prohibition**
   - The reference file should **not have any blank columns** before the `S.No` (Serial Number) column, as this will break the alignment detection framework.
