Below is a professionally structured `README.md` for your GitHub repository. It includes sections for project overview, setup instructions, and a technical breakdown of how the script functions.

***

# Utility Manager: Automation for Electricity, Water, & LPG

This repository contains a Google Apps Script designed to automate the management, calculation, and reporting of utility data (Electricity, Water, and LPG) within Google Sheets. It bridges the gap between raw data sources and financial reporting by automating complex calculations and data validation.

## 🚀 Features

*   **Multi-Utility Support:** Dedicated workflows for Electricity (`Elec`), Water, and LPG.
*   **Dynamic Data Fetching:** Imports data from external source spreadsheets based on links provided directly in the sheet.
*   **Smart Calculations:** 
    *   Applies conditional formulas for Fixed Rates, Special Rates, and Theoretical readings.
    *   Protects manual overrides (the script identifies and skips cells that shouldn't be overwritten).
*   **Data Integrity Scanning:** Highlights missing required data and logs variance issues.
*   **Master DB Submission:** Exports active PBTT records to a centralized master database with a single click.
*   **Custom UI:** Adds a "Utility Manager" menu to the Google Sheets interface for ease of use.

---

## 🛠️ Setup Instructions

### 1. Prepare your Google Sheet
Ensure your Google Sheet has three tabs named exactly:
*   `Elec`
*   `Water`
*   `LPG`

Each tab must have:
*   **Cell A5:** The URL of the Source Spreadsheet where raw data is located.
*   **Cell L5:** The Base Rate/Price Reference.
*   **Cell L6:** The Secondary Reference (e.g., Discount or specific Multiplier).
*   **Headers on Row 12** and **Data starting on Row 13**.

### 2. Install the Script
1.  In your Google Sheet, go to **Extensions** > **Apps Script**.
2.  Delete any existing code in the editor (`Code.gs`).
3.  Copy the code from the script file in this repository and paste it into the editor.
4.  Update the **Configuration Constants** at the top of the script:
    *   `SOURCE_DB_URL`: Default source spreadsheet link.
    *   `PBTT_DB_ID`: The File ID of the master spreadsheet where "Submit" data will be sent.
5.  Click the **Save** (💾) icon and rename the project to "Utility Manager".

### 3. Authorization
1.  Refresh your Google Sheet.
2.  A new menu named **Utility Manager** will appear.
3.  Click any function (e.g., *Utility Manager > Electricity > Scan Tab*).
4.  Google will prompt for permissions. Select your account, click "Advanced," and click "Go to Utility Manager (unsafe)" to grant access.

---

## 📖 How it Works

### 1. The Menu (onOpen)
The `onOpen` function automatically builds a nested menu in your toolbar. This allows users to trigger specific utility workflows without touching the script.

### 2. Fetch Data (`fetchDataOnly`)
This function looks at **Cell A5** in the active tab, opens that external spreadsheet, and pulls the relevant utility data.
*   **Protection Logic:** It automatically skips specific columns (K, L, O, P, Z) to ensure it doesn't overwrite manual inputs or specific formulas existing in the destination sheet.
*   **Auto-Limit:** The script stops importing once it detects a row starting with the word "TOTAL."

### 3. Run Formulas (`applyFormulasToSheet`)
This is the core logic engine. Based on whether you are on the Electricity, Water, or LPG tab, the script injects appropriate formulas starting from row 13.
*   **Conditions:** It checks column values (like "fix rate" or "special rate") to decide which formulas to apply to which columns.
*   **Overlap Safety:** It includes "Overwrite Protection," ensuring that if a user has manually entered data in columns `O` (Rate) or `Z`, the automation will not overwrite it.

### 4. Scan and Validate
*   **Scan Tab:** Checks required columns (5, 11, 31, 34) for empty values and highlights them in yellow.
*   **Issue Logging:** Monitors variances. If utility usage exceeds or drops below 30% compared to previous months, it logs an entry into an "IssueLogs" tab.

### 5. Submit PBTT (`recordActivePBTT`)
Takes the summary values from cells `E1`, `E2`, `E4`, and `E5` of your current tab and appends them as a new row in a centralized master database (`PBTT_DB_ID`). It also logs the user who submitted it and a timestamp.

---

## ⚙️ Configuration Reference

| Variable | Description | Default/Example |
| :--- | :--- | :--- |
| `headerRow` | Row number where your headers live. | `12` |
| `dataStartRow` | Row where the data actually begins. | `13` |
| `SOURCE_DB_URL` | Primary URL for source fetching. | *Your Sheet URL* |
| `PBTT_DB_ID` | File ID for the master destination DB. | *Alphanumeric ID from URL* |

---

## ⚠️ Requirements
*   **Editor Permissions:** You must have edit access to both the Utility sheet and the Source/Database sheets.
*   **Structure:** Do not change tab names (`Elec`, `Water`, `LPG`), or the calculation logic may fail.
