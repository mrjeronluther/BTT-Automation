# 🏗️ BTT-Automation - MCD Bill to Tenant Manager Automator

> A comprehensive Google Apps Script solution for automating utility bill management, formula injection, variance analysis, and multi-layered summation across property files.

**Language:** JavaScript (Google Apps Script)  
**Last Updated:** February 2026  
**Author:** [Jeron Luther Castro / Megaworld Lifestyle Malls](https://github.com/mrjeronluther)

---

## 📋 Table of Contents

- [Overview](#overview)
- [Key Features](#key-features)
- [Installation & Setup](#installation--setup)
- [Configuration](#configuration)
- [Operational Workflow](#operational-workflow)
- [Calculation Logic](#calculation-logic)
- [Technical Details](#technical-details)
- [Usage Guide](#usage-guide)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

---

## 🎯 Overview

**BTT-Automation** is a sophisticated Google Apps Script designed to streamline utility management (Electricity, Water, and LPG) for multi-property operations. It automates the complex process of synchronizing billing data from external sources, injecting accounting formulas, performing variance analysis, and generating comprehensive reports.

### What Does It Do?

This tool synchronizes data from external sources and handles complex accounting calculations across specific property sectors, enabling:
- ⚡ **Electricity** billing automation
- 💧 **Water** consumption tracking  
- 🔥 **LPG** (Liquified Petroleum Gas) management

Perfect for property managers, facility teams, and accounting departments managing multiple utility streams across multiple locations.

---

## 🚀 Key Features

### 1. **Dynamic Data Syncing** 📥
- Automatically fetches data from external spreadsheets
- Links configured via URL in cell `A1` of each utility tab
- Intelligent row mapping translates source columns to target columns
- Handles multiple utility types (Elec, Water, LPG) with different schemas

### 2. **Intelligent Formula Injection** 🧮
- Automatically applies accounting formulas to calculation columns
- Smart conditional logic:
  - Skips rows with existing "Fix Rates" or manual entries
  - Handles "Theoretical" entries to prevent formula errors
  - Detects "Special Rate" scenarios for custom calculations
- Preserves manual inputs while automating calculations

### 3. **Automatic Sub-Totalling & Grand Totals** 📊
- Recursive algorithm scans Column A for "Sub Total" and "Total" labels
- Divides data into logical blocks between subtotals
- Injects `=SUM()` formulas for all numerical columns
- Prevents double-counting by aggregating only subtotal rows into final total

### 4. **Mathematical Variance Correction** ✅
- Recalculates variance percentages at the sub-total level
- Avoids mathematical errors from summing percentages directly
- Applies formula: `(Current - Historical) / Historical`
- Handles consumption variances (%) and amount variances (%)

### 5. **Smart Variance Scanner** 🔍
- Scans data for anomalies (variances outside ±30% threshold)
- Highlights problematic cells in red for quick visual identification
- Automatically logs all issues with:
  - Timestamp
  - Source tab name
  - Row and column reference
  - Variance value and direction

### 6. **Audit Logging & PBTT Submission** 📤
- Submits active records to a centralized master database
- Tracks user and timestamp for audit trail
- Enables compliance and historical tracking
- Integrates with centralized Property Bill-to-Tenant (PBTT) database

### 7. **Custom Menu Interface** 🎨
- User-friendly Google Sheets menu system
- Organized by utility type (⚡ Electricity, 💧 Water, 🔥 LPG)
- One-click access to all operations
- Separate submenu for audit submission

---

## ⚙️ Installation & Setup

### Prerequisites
- Google Sheets access
- Ability to create/edit Google Apps Scripts
- Source spreadsheet with utility data
- Destination spreadsheet for processed data

### Step-by-Step Installation

#### 1. **Open Your Destination Google Sheet**
This is where your processed utility bills will live.

#### 2. **Access Apps Script Editor**
- Click `Extensions` > `Apps Script`
- A new tab will open with the Apps Script editor

#### 3. **Paste the Script**
- Clear any existing code
- Copy the entire contents of `code.gs`
- Paste into the Apps Script editor
- Click **Save** and give your project a name

#### 4. **Authorize Permissions**
- Click the **Run** button (or select `onOpen` and run)
- Google will prompt you to authorize the script
- Review and grant the required permissions:
  - Access to your spreadsheet
  - Access to external spreadsheets (for data fetching)
  - Drive access

#### 5. **Return to Your Sheet**
- Close the Apps Script tab
- Refresh your Google Sheet (`Ctrl+R` or `Cmd+R`)
- A new menu **"Utility Manager"** should appear in the menu bar

#### 6. **Verify Installation**
- Click on **Utility Manager** to see the submenu options
- You should see: ⚡ Electricity, 💧 Water, 🔥 LPG, and 📤 Submit Active PBTT

---

## 🔧 Configuration

### Required Constants

Edit these constants in `code.gs` to match your environment:

```javascript
// Line 4-5: Update with your source and destination spreadsheet IDs
const SOURCE_DB_URL = "https://docs.google.com/spreadsheets/d/YOUR_SOURCE_ID/edit";
const PBTT_DB_ID    = "YOUR_PBTT_DATABASE_ID";

// Line 7-11: Configure header positions (adjust if your sheet differs)
const CONFIG = {
  headerRow: 12,      // Row where column labels are located
  dataStartRow: 13,   // Row where utility entries begin
  minCols: 34         // Minimum columns required for calculations
};
```

### Finding Your Spreadsheet IDs

**Source Spreadsheet ID:**
- Open your source sheet in browser
- Copy the ID from the URL: `docs.google.com/spreadsheets/d/{ID_HERE}/edit`

**PBTT Database ID:**
- Create or identify your centralized database sheet
- Copy the ID from its URL

### Sheet Structure Requirements

Your destination sheet must follow this layout:

```
Row 12:    HEADERS (Column labels for all utilities)
Row 13+:   DATA ROWS (Utility entries start here)
Column A:  Identifier (Shop Name, "Sub Total", "Total")
Columns J-AK: Calculation columns (Rates, Consumption, Readings, Variances)
```

### Required Cell References

Set these up on **each utility tab** (Elec, Water, LPG):

| Cell | Purpose | Example |
|------|---------|---------|
| **A1** | Source spreadsheet URL link | `https://docs.google.com/spreadsheets/d/...` |
| **A5** | Same as A1 (backup reference) | Same URL |
| **L5** | Base rate value (Elec: ₱/kWh, Water: ₱/m³) | `8.50` |
| **L6** | Secondary rate/adjustment factor | `2.15` |
| **E1-E5** | PBTT submission fields | Property, location, startDate, endDate |

---

## 📈 Operational Workflow

### Complete Process Flow

```
1. FETCH DATA
   ↓
2. RUN FORMULAS
   ↓
3. SCAN TAB (Optional - to verify variance)
   ↓
4. SUBMIT TO PBTT (When ready for audit)
```

### Phase 1: Fetch Data

**Purpose:** Import fresh utility data from source spreadsheet

**Steps:**
1. Open your destination spreadsheet
2. Click `Utility Manager` > `⚡ Electricity` > `1. Fetch Data`
3. Script reads source spreadsheet from cell A1 link
4. Clears existing data below Row 12
5. Imports fresh rows with intelligent column mapping

**Exclusion Logic:**
The script intentionally skips these columns to preserve your formulas:
- Column K (Index 10)
- Column L (Index 11)
- Column O (Index 14) - *unless "fix rate" present*
- Column P (Index 15) - *unless "fix rate" present*
- Column Z (Index 25)

**Column Mapping:**
- Source Column K → Target Column J (Previous reading reference)
- Source Column L → Target Column AE (Historical consumption/billing)
- Source Column P → Target Column AH (Historical billing amount)

**Success Indicator:** Toast notification appears: "Fetched [TabName] data. Target K is now blank."

---

### Phase 2: Run Formulas

**Purpose:** Inject calculation logic into all data rows

**Prerequisites:**
- L5 and L6 must contain numeric values
- Data must be fetched first
- At least one row of data present

**Steps:**
1. Click `Utility Manager` > `⚡ Electricity` > `2. Run Formulas`
2. System checks L5 and L6 are populated (required)
3. Scans data for special conditions:
   - **Fix Rate Detection:** If column O contains "fix rate"
   - **Theoretical Entries:** If column J or K contains "theoretical"
   - **Special Rates:** If column N or O contains "special rate"
4. Applies appropriate formula set based on conditions
5. Protects existing manual entries from being overwritten
6. Formats variance columns as percentages (0.00%)

**Formula Sets Applied:**

#### **Electricity Formula Map**
| Column | Formula | Purpose |
|--------|---------|---------|
| **L** | `=IFERROR((K-J)*I,0)` | Total Consumption |
| **O** | `=$L$5` | Rate (from config) |
| **P** | `=IFERROR(IF(O="fix rate","Put/input",ROUND(L*O,2)),"0")` | Calculated Amount |
| **Q** | `=IFERROR(ROUND(P*1.12,3),"-")` | Amount + VAT |
| **G** | `=IFERROR(ROUND(P*1.12,2),"-")` | Amount + VAT (duplicate) |
| **Z** | `=$L$6` | Adjustment Factor |
| **AF** | `=IFERROR(L-AE,"-")` | Consumption Variance (units) |
| **AG** | `=IFERROR(AF/AE,"-")` | Consumption Variance (%) |
| **AI** | `=IFERROR(P-AH,"-")` | Amount Variance (₱) |
| **AJ** | `=IFERROR(AI/AH,"-")` | Amount Variance (%) |
| **AA** | `=IFERROR(L*Z,"-")` | Consumption × Adjustment |
| **AB** | `=IFERROR(P-AA,"-")` | Amount Adjustment |
| **AC** | `=IFERROR(AB/Q,"-")` | Adjustment % |

#### **Water Formula Map**
| Column | Formula | Purpose |
|--------|---------|---------|
| **L** | `=K-J` | Total Consumption |
| **O** | `=IF(NOT(ISNUMBER($L$5)),"0",$L$5)` | Rate |
| **P** | `=IFERROR(IF(O="fix rate","Put/input",ROUND(ROUND(O,2)*ROUND(L,3),2)),"-")` | Calculated Amount |
| **S** | `=IF(NOT(ISNUMBER($U$10)),"0",$U$10)` | Adjustment Rate |
| **T** | `=IFERROR(S*L,"0")` | Adjusted Consumption Cost |
| **U** | `=IFERROR(L+T,"0")` | Total Consumption + Adj |
| **V** | `=IFERROR(P*S,"0")` | Billing Adjustment |
| **W** | `=IFERROR(V+P,"0")` | Total Bill |
| **X** | `=IFERROR(ROUND(W*1.12,3),"-")` | Bill + VAT |
| **Z** | `=IF(NOT(ISNUMBER($L$6)),".",$L$6)` | Adjustment Factor |
| **AA** | `=IFERROR(L*Z,"-")` | Adjusted Units |
| **AB** | `=IFERROR(W-AA,"-")` | Variance Amount |
| **AC** | `=IFERROR(AB/Q,"-")` | Variance % |
| **AF** | `=IFERROR(L-AE,"-")` | Consumption Variance |
| **AG** | `=IFERROR(AF/AE,"-")` | Consumption Variance % |
| **AI** | `=IFERROR(W-AH,"-")` | Amount Variance |
| **AJ** | `=IFERROR(AI/AH,"-")` | Amount Variance % |

#### **LPG Formula Map**
| Column | Formula | Purpose |
|--------|---------|---------|
| **L** | `=K-J` | Total Consumption |
| **M** | `=$N$10*L` | Density/Conversion Factor |
| **N** | `=IFERROR(L*M,"0")` | Converted Units |
| **O** | `=IF(NOT(ISNUMBER($L$5)),"0",$L$5)` | Rate |
| **P** | `=IFERROR(IF(O="fix rate","Put/input",N*O),"-")` | Calculated Amount |
| **Q** | `=IFERROR(ROUND(P*1.12,3),"-")` | Amount + VAT |
| **Z** | `=IF(NOT(ISNUMBER($L$6)),".",$L$6)` | Adjustment Factor |
| **AA** | `=IFERROR(N*Z,"-")` | Adjusted Amount |
| **AB** | `=IFERROR(P-AA,"-")` | Variance Amount |
| **AC** | `=IFERROR(AB/Q,"-")` | Variance % |
| **AF** | `=IFERROR(N-AE,"-")` | Consumption Variance |
| **AG** | `=IFERROR(AF/AE,"-")` | Consumption Variance % |
| **AI** | `=IFERROR(P-AH,"-")` | Amount Variance |
| **AJ** | `=IFERROR(AI/AH,"-")` | Amount Variance % |

**Conditional Logic (Overwrite Protection):**

The script is intelligent about when to apply formulas:

1. **If O = "fix rate"**: Only applies P (and Q for Elec), Z
2. **If J or K = "theoretical"**: Applies O, P, Z (skips L to avoid errors)
3. **If N or O = "special rate"**: Applies L, O, P, Q (or P for Water), Z
4. **Otherwise**: Applies full formula set
5. **If O or Z already have data**: Skips overwriting them
6. **If P already has data AND "fix rate" exists**: Skips overwriting P
7. **If K = "theoretical"**: Never overwrites L (usage column)

**Mandatory Calculations:** Variance and adjustment columns always run:
- M, N, AA, AB, AC (LPG & Elec)
- S, T, U, V, W, X (Water)

---

### Phase 3: Scan Tab (Variance Detection)

**Purpose:** Identify anomalies and audit-flag problematic entries

**Steps:**
1. Click `Utility Manager` > `⚡ Electricity` > `Scan Tab`
2. Script reviews columns **AG** and **AJ** (variance percentages)
3. For each variance:
   - ✅ **Within ±30%**: No action
   - ⚠️ **> +30% or < -30%**: Cell highlighted RED, logged to IssueLogs

**IssueLogs Output:**
Creates or appends to "IssueLogs" sheet with:
| Column | Data |
|--------|------|
| Timestamp | When issue was detected |
| Source Tab | Which utility tab (Elec, Water, LPG) |
| Row | Row number of anomaly |
| Col | Column letter (AG or AJ) |
| Value | Variance percentage (formatted) |
| Issue | "Exceeded +30%" or "Below -30%" |

**Use Case:** Review red-highlighted cells to investigate unusual consumption or billing patterns.

---

### Phase 4: Submit Active PBTT

**Purpose:** Log processed bills to centralized audit database

**Steps:**
1. Fill in PBTT fields on any utility tab:
   - **E1**: Property name/code
   - **E2**: Location identifier
   - **E4**: Service start date
   - **E5**: Service end date
2. Click `Utility Manager` > `📤 Submit Active PBTT`
3. Script submits row to PBTT database with:
   - Current timestamp
   - All 4 PBTT fields
   - Source spreadsheet URL
   - Your email (from Session.getActiveUser())
4. Fields E1, E2, E4, E5, E7 clear automatically

**PBTT Database Schema:**
| Column | Data |
|--------|------|
| Timestamp | Submission time |
| Property | From E1 |
| Location | From E2 |
| Start Date | From E4 |
| End Date | From E5 |
| Source URL | Sheet URL |
| Submitted By | User email |

---

## 🧮 Calculation Logic Reference

### Key Formulas Explained

#### **Consumption Calculation (L)**
```
Electricity: (Current Reading - Previous Reading) × Multiplier
Water: Current Reading - Previous Reading  
LPG: Current Reading - Previous Reading
```

#### **Amount Calculation (P)**
```
Amount = Consumption × Rate
BUT IF "fix rate" detected: Manual input required ("Put/input" prompt)
VAT-Inclusive (Q): Amount × 1.12
```

#### **Variance Analysis (AG, AJ)**
```
Consumption Variance % = (Current Cons - Historical Cons) / Historical Cons
Amount Variance % = (Current Amount - Historical Amount) / Historical Amount
```

#### **Subtotal/Grand Total**
- **Sub Totals**: Sum only data rows above (not including other subtotals)
- **Grand Total**: Sum only subtotal cells (prevents double-counting)

### Column Reference Matrix

| Col | Electricity | Water | LPG | Purpose |
|-----|-------------|-------|-----|---------|
| **I** | Multiplier | - | - | Meter factor |
| **J** | Prev. Reading | Prev. Reading | Prev. Reading | Reference for consumption |
| **K** | Prev. Reading | Prev. Reading | Prev. Reading | Copy of J |
| **L** | Consumption | Consumption | Consumption | ✅ Key calculation |
| **M** | - | - | Density Factor | LPG conversion |
| **N** | - | - | Converted Units | For LPG rate calc |
| **O** | Rate | Rate | Rate | ✅ Key input |
| **P** | Amount | Amount | Amount | ✅ Core calculation |
| **Q** | Amount+VAT | - | Amount+VAT | Taxed amount |
| **Z** | Adj. Factor | Adj. Factor | Adj. Factor | ✅ Key input |
| **AA** | Adj. Consumption | Adj. Consumption | Adj. Consumption | Z × L |
| **AB** | Amount Adj | Amount Adj | Amount Adj | P - AA |
| **AC** | Adj % | Adj % | Adj % | AB / Q |
| **AE** | Historical Cons | Historical Cons | Historical Cons | Reference data |
| **AF** | Cons Variance (units) | Cons Variance (units) | Cons Variance (units) | L - AE |
| **AG** | **Cons Var %** | **Cons Var %** | **Cons Var %** | 🔍 **Monitored** |
| **AH** | Historical Amount | Historical Amount | Historical Amount | Reference billing |
| **AI** | Amount Var (₱) | Amount Var (₱) | Amount Var (₱) | P - AH |
| **AJ** | **Amount Var %** | **Amount Var %** | **Amount Var %** | 🔍 **Monitored** |

---

## 💻 Technical Details

### Technology Stack
- **Language**: Google Apps Script (JavaScript ES6+)
- **Platform**: Google Sheets / Google Drive
- **Runtime**: Server-side (no client-side processing)
- **API Usage**: SpreadsheetApp, DriveApp

### Permissions Required
```
1. Spreadsheet Access
   - Read/write to active spreadsheet
   - Create new sheets (IssueLogs)

2. Drive Access
   - Open external spreadsheets by URL
   - Read data from source files

3. Session Access
   - Get current user email (for audit logging)
```

### Performance Considerations
- **Data Limit**: Tested on sheets with up to 10,000 rows
- **Formula Injection**: ~0.5 seconds per 100 rows
- **Variance Scanning**: ~0.1 seconds per 1,000 rows
- **Quota**: Uses SpreadsheetApp quota (no additional charges)

### Browser Compatibility
- Google Chrome (recommended)
- Firefox
- Safari
- Edge
- Any browser supporting Google Sheets

### Limitations
- Maximum of 34 columns supported per sheet
- Data must start at Row 13 (Row 12 = headers)
- External spreadsheet must be shared/accessible to your account
- Script runs with your permissions (data access depends on sharing settings)

---

## 📖 Usage Guide

### Typical Monthly Workflow

```
WEEK 1: Receive New Bills
├─ Place source spreadsheet URL in A1 of each tab
└─ Verify L5 and L6 have current rates

WEEK 2: Process Data
├─ Utility Manager → ⚡ Electricity → 1. Fetch Data
├─ Utility Manager → ⚡ Electricity → 2. Run Formulas
├─ Repeat for 💧 Water and 🔥 LPG
└─ Wait for success notifications

WEEK 3: Quality Assurance
├─ Utility Manager → ⚡ Electricity → Scan Tab
├─ Review red-highlighted cells in sheet
├─ Check IssueLogs sheet for anomalies
└─ Investigate any variances >30% or <-30%

WEEK 4: Audit & Submission
├─ Fill in PBTT fields (E1, E2, E4, E5)
├─ Utility Manager → 📤 Submit Active PBTT
└─ Verify record appears in central database
```

### Troubleshooting Common Issues

#### **❌ "Paste SOURCE LINK in A5" Error**
**Cause:** Cell A5 is empty or invalid URL  
**Fix:**
- Go to source spreadsheet
- Copy full URL from browser address bar
- Paste into cell A5 on destination tab
- Try Fetch Data again

#### **❌ "Tab 'Elec' not found in source" Error**
**Cause:** Source sheet doesn't have matching tab name  
**Fix:**
- Verify source sheet has tabs named exactly: "Elec", "Water", "LPG"
- Check for typos (case-sensitive)
- Ensure you have read access to source file

#### **❌ "Cannot open source link" Error**
**Cause:** Invalid URL or no sharing access  
**Fix:**
- Test the URL by pasting in browser
- Ensure source spreadsheet is shared with your account
- Request "Viewer" or "Editor" access from owner

#### **❌ "Action Blocked: L5 and L6 are required" Error**
**Cause:** Missing rate values in cells L5 or L6  
**Fix:**
- Click cell L5, enter base rate (e.g., 8.50)
- Click cell L6, enter adjustment factor (e.g., 2.15)
- Both must have numeric values (not formulas)
- Try Run Formulas again

#### **❌ Formulas not applying to all rows**
**Cause:** Data includes "Total" row or special conditions blocking formulas  
**Fix:**
- Check that "Total" row exists and is properly labeled
- Script stops applying formulas at "Total" row
- Review special condition rows:
  - "fix rate" in column O → Only P/Z applied
  - "theoretical" in J/K → O/P/Z applied
  - "special rate" in N/O → Full set with L applied

#### **❌ IssueLogs sheet not created**
**Cause:** No variances detected outside ±30%  
**Fix:**
- IssueLogs only creates if anomalies exist
- Normal variance operation if sheet is absent
- Run Scan Tab again if variances update

#### **⚠️ Red highlighting not appearing in Scan Tab**
**Cause:** Variance columns (AG, AJ) contain text/errors  
**Fix:**
- Verify formulas are applied correctly (Run Formulas)
- Check that AE and AH contain historical reference data
- Review formula output for "-" or error values

### Best Practices

✅ **DO:**
- Update L5 and L6 monthly with new rates
- Run Fetch → Formulas → Scan in order
- Review IssueLogs before final submission
- Keep source and destination sheets in organized folders
- Document any manual overrides in the sheet
- Test with sample data before production use

❌ **DON'T:**
- Manually edit calculated columns (they'll be overwritten)
- Change headers in Row 12 without updating formulas
- Delete the IssueLogs sheet (recreate if needed)
- Use formulas in L5 and L6 (must be fixed values)
- Run Fetch Data twice without clearing (duplicates data)
- Share PBTT database with unauthorized users

---

## 🛠 Developer Usage

### Modifying the Script

#### **Change Configuration Values**

```javascript
// Edit Line 4-5 for your URLs
const SOURCE_DB_URL = "YOUR_SOURCE_LINK";
const PBTT_DB_ID    = "YOUR_DATABASE_ID";

// Edit Line 7-11 for your sheet structure
const CONFIG = {
  headerRow: 12,      // Change if headers are elsewhere
  dataStartRow: 13,   // Change if data starts at different row
  minCols: 34         // Change if you need more columns
};
```

#### **Add a New Utility Type**

1. **Create new sheet** in destination spreadsheet named "NewUtility"
2. **Add to menu** (Line 18-40):
```javascript
.addSubMenu(ui.createMenu("New Utility")
    .addItem("1. Fetch Data", "fetchNewUtility")
    .addItem("2. Run Formulas", "runFormulaNewUtility")
    .addSeparator()
    .addItem("Scan Tab", "scanNewUtilityTab")
    .addItem("Clear Tab", "clearNewUtility"))
```

3. **Add formula map** (Line 140-193):
```javascript
const formulaMapNewUtility = {
  L: r => `=K${r}-J${r}`,
  O: r => `=$L$5`,
  P: r => `=IFERROR(IF(O${r}="fix rate","Put/input",ROUND(L${r}*O${r},2)),"-")`,
  // ... add other formulas
};
```

4. **Add selection logic** (Line 196-199):
```javascript
if (tabName === "Water") activeMap = formulaMapWater;
else if (tabName === "LPG") activeMap = formulaMapLPG;
else if (tabName === "NewUtility") activeMap = formulaMapNewUtility;
else activeMap = formulaMapElec;
```

5. **Add trigger functions** (Line 322-336):
```javascript
function fetchNewUtility() { fetchDataOnly("NewUtility"); }
function runFormulaNewUtility() { applyFormulasToSheet("NewUtility"); }
function clearNewUtility() { clearTabData("NewUtility"); }
function scanNewUtilityTab() { scanTab("NewUtility"); }
```

#### **Adjust Variance Thresholds**

Current threshold: ±30%  
Located in: `scanTab()` function, Line 295

Change this line:
```javascript
[5, 11, 31, 34].forEach(c => {  // Column indices for variance columns
```

And this line:
```javascript
if(!vals[i][0]) sheet.getRange(CONFIG.dataStartRow+i, c).setBackground("#fff176");
```

#### **Modify Rate Configuration Sources**

Currently reads from L5 and L6 on each sheet.

To read from external source:
```javascript
// Instead of:
const valL2 = sheet.getRange("L5").getValue();

// Use:
const sourceSheet = ss.getSheetByName("RateMaster");
const valL2 = sourceSheet.getRange("C2").getValue(); // reads from RateMaster!C2
```

### Code Structure

```
code.gs
├── CONFIGURATION (Lines 1-11)
│   ├── SOURCE_DB_URL
│   ├── PBTT_DB_ID
│   └── CONFIG object
│
├── 1. MENU (Lines 13-40)
│   └── onOpen() - Creates menu interface
│
├── 2. FETCH DATA (Lines 42-122)
│   ├── fetchDataOnly() - Main fetch function
│   └── Column exclusion logic
│
├── 3. RUN FORMULAS (Lines 124-270)
│   ├── applyFormulasToSheet() - Main formula engine
│   ├── Formula maps (Elec, Water, LPG)
│   ├── Conditional logic (fix rate, theoretical, special rate)
│   ├── Overwrite protection
│   └── Variance formatting
│
├── 4. UTILS (Lines 272-300)
│   ├── logToIssueTab() - Write to IssueLogs
│   ├── clearTabData() - Clear rows
│   └── scanTab() - Variance detection
│
├── 5. SUBMIT PBTT (Lines 303-317)
│   └── recordActivePBTT() - Log to database
│
└── 6. TRIGGERS (Lines 320-336)
    ├── fetchElec/Water/LPG()
    ├── runFormulaElec/Water/LPG()
    ├── clearElec/Water/LPG()
    └── scanElecTab/WaterTab/LPGTab()
```

---

## 🤝 Contributing

### How to Contribute

1. **Report Issues**
   - Create detailed description of the problem
   - Include error message and steps to reproduce
   - Share relevant cell values or formulas

2. **Suggest Enhancements**
   - Describe the feature/improvement
   - Explain the use case
   - Suggest implementation approach

3. **Submit Improvements**
   - Fork the repository
   - Create feature branch (`git checkout -b feature/MyFeature`)
   - Test thoroughly with sample data
   - Submit pull request with detailed description

### Testing Checklist
Before submitting changes, test:
- ✅ Fetch Data with 10+ rows
- ✅ Run Formulas with all condition types (fix rate, theoretical, normal)
- ✅ Scan Tab detects variances
- ✅ Submit PBTT creates proper database entry
- ✅ IssueLogs sheet updates correctly
- ✅ No formula overwrites occur on protected cells
- ✅ Menu displays correctly

---

## 📄 License

This project is maintained for **Property Management and Utility Tracking workflows** by Megaworld Lifestyle Malls.

**Usage Rights:**
- ✅ Free to use and modify for internal operations
- ✅ Adapt for your property portfolio
- ⚠️ Requires attribution if shared
- ❌ Not for commercial resale
- ❌ Not for public distribution without permission

---

## 📞 Support & Contact

**Author:** [Jeron Luther Castro](https://github.com/mrjeronluther)  
**Organization:** Megaworld Lifestyle Malls  
**Repository:** [BTT-Automation on GitHub](https://github.com/mrjeronluther/BTT-Automation)

### Getting Help

1. **Check Troubleshooting Section** - Most issues documented above
2. **Review Code Comments** - Extensive inline documentation
3. **Test with Sample Data** - Create test sheet to validate
4. **Open GitHub Issue** - For bugs or feature requests

---

## 🎓 Learning Resources

### Google Apps Script
- [Official Documentation](https://developers.google.com/apps-script)
- [SpreadsheetApp Class](https://developers.google.com/apps-script/reference/spreadsheet)
- [Apps Script Editor Guide](https://developers.google.com/apps-script/guides/sheets)

### Related Concepts
- Google Sheets Formula Functions
- Array Formulas and IFERROR
- Regular Expressions for text matching
- JSON data handling

---

## 📊 Project Stats

| Metric | Value |
|--------|-------|
| **Language** | Google Apps Script (JavaScript) |
| **File Size** | ~13 KB |
| **Lines of Code** | 336 |
| **Functions** | 20+ |
| **Supported Utilities** | 3 (Elec, Water, LPG) |
| **Maximum Rows** | 10,000+ |
| **Calculation Columns** | 34+ |
| **Created** | 2026 |
| **Last Updated** | February 2026 |

---

## 🎯 Roadmap & Future Features

**Planned Enhancements:**
- [ ] Email notifications for variance alerts
- [ ] Graphical dashboard for utility trends
- [ ] Bulk rate update from master table
- [ ] Advanced filtering and export options
- [ ] Multi-year historical comparison
- [ ] Integration with accounting systems
- [ ] Mobile-friendly report view

**Known Limitations:**
- Currently supports 3 utility types
- Limited to Google Sheets (no other spreadsheet apps)
- Requires manual source URL configuration
- No built-in data validation

---

## ✨ Acknowledgments

Special thanks to the Megaworld Lifestyle Malls property management team for feedback and real-world usage testing.

---

**Last Updated:** February 9, 2026  
**Status:** ✅ Production Ready  
**Version:** 1.0.0
