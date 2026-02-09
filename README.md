# MCD Bill to Tenant Manager Automator ⚡💧🔥

A comprehensive Google Apps Script solution designed to automate data fetching, formula injection, variance analysis, and multi-layered summation for Utility management (Electricity, Water, and LPG). 

This tool synchronizes data from external sources and handles complex accounting calculations across specific property sectors.

## 🚀 Key Features

*   **Dynamic Data Syncing**: Fetches data from external spreadsheets based on URL links provided in cell `A1` of each utility tab.
*   **Intelligent Formula Injection**: Automatically applies accounting formulas to columns while skipping rows with existing "Fix Rates" or manual entries.
*   **Automatic Sub-Totalling & Grand Totals**: A recursive algorithm that scans Column A for "Sub Total" labels, sums the preceding block, and aggregates them into a final "Total" row.
*   **Mathematical Variance Correction**: Instead of summing percentages (which results in errors), the script recalculates variance percentages at the sub-total level for mathematical accuracy.
*   **Smart Variance Scanner**: Scans data for anomalies (variances outside ±30%) and generates an automated issue log in a separate `IssueLogs` tab.
*   **Audit Logging**: Submits active records to a centralized master database with user and timestamp tracking.

---

## 🛠️ Configuration & Setup

To use this script, ensure the following constants in the code match your spreadsheet environment:

```javascript
const SOURCE_DB_URL = "YOUR_SOURCE_LINK";
const PBTT_DB_ID    = "YOUR_DATABASE_ID";

const CONFIG = {
  headerRow: 12,      // Where the column labels live
  dataStartRow: 13,   // Where the utility entries start
  minCols: 34         // Minimum column width for the logic
};
```

### Required Columns in Sheet
The script relies on a specific layout starting from **Row 12 (Headers)**:
*   **Column A**: Identifier (e.g., Shop Name, "Sub Total", or "Total").
*   **Columns J - AK**: Contain specific values like Rates, Consumption (J, K, L), Previous Readings (AF), and calculated Variances (AG, AH, AJ, AK).

---

## 📈 Operational Workflow

### 1. Fetch Data
*   The script reads the URL from cell **A1**.
*   It clears existing data below Row 12 and imports fresh rows.
*   It performs **dynamic mapping**: translating source columns to specific target columns (e.g., Source L to Target AF).

### 2. Run Formulas
*   Injects calculation logic into every row.
*   **Conditional Skipping**: If "fix rate" is found in columns O, P, or Z, the script preserves manual inputs.
*   **Theoretical Adjustments**: Specifically handles "Theoretical" entries in Column K to prevent formula errors.

### 3. Automated Summation Engine
This is the core logic that handles your sub-totals:
1.  **Search Phase**: Locates "Sub Total" and "Total" labels in Column A.
2.  **Partitioning**: Divides the sheet into blocks between Sub Totals.
3.  **Summation**: Injects `=SUM(Start:End)` formulas for all numerical value columns.
4.  **Recalculation**: For Variance %, it applies the division formula: `Amount Variance / Previous Total` rather than summing percentages.
5.  **Grand Total**: Aggregates only the "Sub Total" cells into the final "Total" row to prevent double-counting.

### 4. Scan Tab
*   Reviews Columns **AH** and **AK**.
*   If any variance is **> 30%** or **< -30%**, the cell is highlighted Red.
*   The issue is automatically logged into the **IssueLogs** sheet with a timestamp and specific cell reference.

---

## 🧮 Calculation Logic

| Column | Name | Logic |
| :--- | :--- | :--- |
| **L** | Total Consumption | `(Current - Previous) * Multiplier` |
| **P** | Calculated Amount | `Consumption * Rate` (unless "fix rate") |
| **Q** | Amount + VAT | `Amount * 1.12` |
| **AH** | Cons. Var % | `(Current Cons - Historical Cons) / Historical` |
| **AK** | Amount Var % | `(Current Amount - Historical Amount) / Historical` |

---

## 💻 Technical Details

*   **Language**: Google Apps Script (JavaScript-based).
*   **Permissions**: Requires `SpreadsheetApp` and `DriveApp` scope for cross-workbook operations.
*   **Trigger**: Custom Menu in Google Sheets UI (`Utility Manager`).

---

## 🛠 Developer Usage
To contribute or modify:
1.  Open the Google Sheet.
2.  Go to `Extensions` > `Apps Script`.
3.  Paste the contents of `CODE.gs` into the editor.
4.  Save and refresh the Spreadsheet.

---

### Author
[Your Name/Organization]
*Maintained for Property Management and Utility Tracking workflows.*
