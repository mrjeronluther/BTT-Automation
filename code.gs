/* =================================
CONFIGURATION & MAPPING
================================= */
const PBTT_DB_ID    = "1hMMUd4ho50HP63dc2fRAo--iK-m7YotamkKtsDGT_Us";
const BACKUP_REGISTRY_ID = "10-ywOh509BNRMd0C-Mb8b5gibbu62D_K8U8cWYcV59U"; 
const BACKUP_FOLDER_ID = "1aokNFrCuVdLWs4AylG7LNekCtfQ5B1-p";
const CELL_LIMIT_MAX = 10000000; // Google's absolute limit (10M)
const CELL_ROTATION_LIMIT = 8000000; // Accurate threshold to trigger rotation (80%)



const CONFIG = {
  headerRow: 12,
  dataStartRow: 13,
  minCols: 34
};

// EASILY ADJUST SOURCE -> TARGET MAPPING HERE
const FETCH_MAPS = {
  "Elec": {
    "J": "K",   // target col : source col
    "AF": "L",  
    "AI": "P"   
  },
  "Water": {
    "J": "K",
    "AF": "L", 
    "AI": "W"  
  },
  "LPG": {
    "J": "K",
    "AF": "N",
    "AI": "P"
  }
};

// COLUMNS TO BE LEFT BLANK DURING FETCH (To be filled by Run Formula)
const EXCLUSIONS = {
  "Elec": ["K", "L", "O", "P", "Q", "Z", "AA", "AB", "AC", "AG", "AH", "AJ", "AK"],
  "Water": ["K", "L", "O", "P", "Q", "S","T", "U", "V", "W", "X", "Z", "AA", "AB", "AC", "AG", "AH", "AJ", "AK"],
  "LPG": ["K", "L", "O", "P", "M","N","Q", "Z", "AA", "AB", "AC", "AG", "AH", "AJ", "AK"]
};

/* =================================
1. MENU
================================= */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Utility Manager")

    .addSubMenu(ui.createMenu("⚡ Electricity")
        .addItem("1. Fetch Data", "masterFetchElec") // Points to wrapper
        .addItem("2. Run Formulas", "runFormulaElec"))

    .addSubMenu(ui.createMenu("💧 Water")
        .addItem("1. Fetch Data", "masterFetchWater") // Points to wrapper
        .addItem("2. Run Formulas", "runFormulaWater"))

    .addSubMenu(ui.createMenu("🔥 LPG")
        .addItem("1. Fetch Data", "masterFetchLPG") // Points to wrapper
        .addItem("2. Run Formulas", "runFormulaLPG"))

    .addSeparator()
    .addItem("🔍 Scan All Tabs (Elec, Water, LPG)", "scanAllTabs")
    .addItem("📤 Submit Active PBTT", "recordActivePBTT") 
    .addItem("🛠️ Run this if Ref/dvGen/dvPeriod error occur", "INSTALL_SYSTEM")
    .addToUi();
}


// Wrapper for Electricity
function masterFetchElec() {
  INITIALIZE_SYSTEM_BUTTON(); // Function 1: Sync & Ref#
  fetchElec();                // Function 2: The actual fetch
}
// Wrapper for Water
function masterFetchWater() {
  fetchWater();               // Function 2
}
// Wrapper for LPG
function masterFetchLPG() {
  fetchLPG();                // Function 2
}

/* =================================
   2. SCAN WRAPPER & LOG CLEANING
================================= */
function scanAllTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabsToScan = ["Elec", "Water", "LPG"];
  
  // --- 1. REQUIREMENT CHECKER (New Fix) ---
  const requirements = {
    "Elec":  { config: ["L5", "L6"], cols: ["L","O","P","Q","Z","AA","AB","AC","AG","AH","AJ","AK"] },
    "Water": { config: ["L5", "L6", "U10"], cols: ["L","O","P","S","T","U","V","W","X","Z","AA","AB","AC","AG","AH","AJ","AK"] },
    "LPG":   { config: ["L5", "L6", "N10"], cols: ["L","M","N","O","P","Q","Z","AA","AB","AC","AG","AH","AJ","AK"] }
  };

  let stopErrors = [];
  tabsToScan.forEach(tabName => {
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) return;
    let missing = [];
    const req = requirements[tabName];
    // Check Cells
    req.config.forEach(c => { if(sheet.getRange(c).getValue() === "") missing.push(`Cell ${c}`); });
    // Check Cols (Checking first data row)
    req.cols.forEach(col => { if(sheet.getRange(col + CONFIG.dataStartRow).getValue() === "") missing.push(`Col ${col}`); });
    
    if (missing.length > 0) stopErrors.push(`[${tabName}]: ${missing.join(", ")}`);
  });

  if (stopErrors.length > 0) {
    SpreadsheetApp.getUi().alert("🚫 SCAN CANCELLED - DATA MISSING\n\n" + stopErrors.join("\n\n"));
    return;
  }

  // --- 2. CLEAR LOGS ---
  const logSheetNames = ["Basic Anomalies", "Client Rate Anomalies"];
  logSheetNames.forEach(name => {
    let s = ss.getSheetByName(name) || ss.insertSheet(name);
    if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 6).clearContent();
    if (s.getLastRow() === 0) s.appendRow(["Timestamp", "Tab", "Cell", "Column Label", "Error Message", "Remarks"]);
  });

  // --- 3. RUN SCANS SILENTLY ---
  tabsToScan.forEach(tabName => {
    scanTab(tabName, false); // Make sure your scanTab has the alerts REMOVED as shown below
  });

  // --- 4. SHOW FINAL SUMMARY MODAL ---
  const stdTotal = Math.max(0, ss.getSheetByName("Basic Anomalies").getLastRow() - 1);
  const kaTotal = Math.max(0, ss.getSheetByName("Client Rate Anomalies").getLastRow() - 1);
  showScanSuccessModal(tabsToScan, stdTotal, kaTotal);
}

// Function for the Modal Pop up
function showScanSuccessModal(scannedTabs, totalStd, totalKA) {
  const htmlContent = `
    <html>
      <head>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/css/materialize.min.css">
        <style>body { padding: 25px; } .header { font-size: 1.3em; font-weight: bold; color: #2e7d32; border-bottom: 2px solid #eee; margin-bottom: 20px; }</style>
      </head>
      <body>
        <div class="header">📋 Global Scan Summary</div>
        <p>Sheets Analyzed: <b>${scannedTabs.join(", ")}</b></p>
        <div style="padding:15px; background:#f5f5f5; border-radius:10px;">
          <p>Standard Anomalies: <span style="font-weight:bold; color:red; float:right;">${totalStd}</span></p>
          <p>KA Identification Issues: <span style="font-weight:bold; color:red; float:right;">${totalKA}</span></p>
        </div>
        <p><small>Review full logs in 'Basic Anomalies' and 'Client Rate Anomalies' sheets.</small></p>
        <div style="text-align:center; margin-top:20px;">
          <button class="btn green darken-2" onclick="google.script.host.close()">Understood</button>
        </div>
      </body>
    </html>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(400).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "System Update");
}


/* =================================
2. HELPER: COL LETTER TO INDEX
================================= */
function colToIdx(letter) {
  let column = 0;
  for (let i = 0; i < letter.length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, letter.length - i - 1);
  }
  return column - 1; 
}



/* =================================
3. FETCH DATA (DYNAMIC MAPPING)
================================= */
function fetchDataOnly(tabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return;

  const dataStartRow = 13; 

  // ==========================================
  // CHECK FOR OVERWRITE (Col E13 Downwards)
  // ==========================================
  const lastSheetRow = sheet.getLastRow();
  let hasExistingData = false;

  if (lastSheetRow >= dataStartRow) {
    const colE_values = sheet.getRange(dataStartRow, 5, lastSheetRow - dataStartRow + 1, 1).getValues();
    hasExistingData = colE_values.some(row => {
      const val = row[0];
      if (val === "" || val === null || val === undefined || val === false) return false;
      return String(val).trim() !== ""; 
    });
  }


  const sourceLink = sheet.getRange("A1").getValue();
  if (!sourceLink) { 
    SpreadsheetApp.getUi().alert("Paste SOURCE LINK in cell C14 in Instructions Tab.");
    return; 
  }

  let sourceSS;
  try { 
    sourceSS = SpreadsheetApp.openByUrl(sourceLink); 
  } catch (e) { 
    SpreadsheetApp.getUi().alert("Cannot open source link."); 
    return; 
  }

  if (sourceSS.getId() === ss.getId()) {
    SpreadsheetApp.getUi().alert("FETCH CANCELLED: You are using the current spreadsheet's URL. Please paste an external source link in C14 in Instruction Tab.");
    return;
  }
  
  const sourceSheet = sourceSS.getSheetByName(tabName);
  if (!sourceSheet) { 
    SpreadsheetApp.getUi().alert(`Tab "${tabName}" not found in source.`); 
    return; 
  }

  const lastSourceRow = sourceSheet.getLastRow();
  const lastSourceCol = Math.max(sourceSheet.getLastColumn(), 29); 
  if (lastSourceRow < dataStartRow) return;

  const rawData = sourceSheet.getRange(dataStartRow, 1, lastSourceRow - dataStartRow + 1, lastSourceCol).getValues();

  // --- STAGE 1: VALIDATION PASS (Data Integrity only) ---
  for (let i = 0; i < rawData.length; i++) {
    let row = rawData[i];
    let vA_source = String(row[0] || "").trim();
    let vA_low = vA_source.toLowerCase();
    
    // Create logic for what represents a subtotal text line
    let isSubTotal = vA_low.includes("subtotal") || vA_low.includes("sub-total") || vA_low.includes("sub total");
    
    // Ignore footer text but allow subtotal text safely
    if (vA_low.includes("total") && !isSubTotal) break; 

    // Skip Data Validation IF row is a recognized subtotal
    if (!isSubTotal) {
      let vB = String(row[1] || "").trim();
      let vE = String(row[4] || "").trim();

      if (vE !== "" && vB === "") {
        SpreadsheetApp.getUi().alert(
          `FETCH CANCELLED: Data Integrity Error\n\nRow ${i + dataStartRow} in the Source Tab "${tabName}" is incomplete. Column E has a value but Column B is missing.`
        );
        return; 
      }
    }
  }

  // --- STAGE 2: PROCESSING ---
  if (sheet.getMaxColumns() < 29) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 29 - sheet.getMaxColumns());
  }

  const destWidth = sheet.getMaxColumns(); 
  const pasteArray = [];
  const skipIndices = (EXCLUSIONS[tabName] || []).map(letter => colToIdx(letter));
  const mapping = FETCH_MAPS[tabName] || {};

  let totalFound = false;
  let rowCounter = 0; 

  const addFooterToPasteArray = (labelA) => {
    if (pasteArray.length > 0) {
      let lastRow = pasteArray[pasteArray.length - 1];
      if (!lastRow.every(cell => cell === "")) pasteArray.push(new Array(destWidth).fill(""));
    }
    let totalRow = new Array(destWidth).fill("");
    totalRow[0] = labelA; 
    pasteArray.push(totalRow);
    for (let s = 0; s < 3; s++) pasteArray.push(new Array(destWidth).fill(""));
    let sigRow = new Array(destWidth).fill("");
    sigRow[0] = "Prepared By:"; sigRow[7] = "Checked By:"; sigRow[28] = "Noted By:";
    pasteArray.push(sigRow);
  };

  for (let i = 0; i < rawData.length; i++) {
    let sourceRow = rawData[i];
    let valA_source = String(sourceRow[0] || "").trim();
    let valA_lower = valA_source.toLowerCase();
    
    // Redefining isSubTotal again to verify current line
    let isSubTotal = valA_lower.includes("subtotal") || valA_lower.includes("sub-total") || valA_lower.includes("sub total");

    if (valA_lower.includes("total") && !isSubTotal) {
      totalFound = true;
      addFooterToPasteArray(valA_source); 
      break; 
    }

    let vB = String(sourceRow[1] || "").trim();
    let vE = String(sourceRow[4] || "").trim();

    // EXEMPTION APPLIED HERE: If it's normal valid data (B+E) *OR* a subtotal... PULL THE ROW!
    if ((vB !== "" && vE !== "") || isSubTotal) {
      let destRow = new Array(destWidth).fill(""); 
      
      // Determine what to place in Column A based on what the row actually is
      if (!isSubTotal) {
        rowCounter++;
        destRow[0] = rowCounter; // auto-numbering only counts real items
      } else {
        destRow[0] = valA_source; // writes the subtotal word over properly!
      }

      for (let c = 1; c < sourceRow.length; c++) {
        if (skipIndices.includes(c)) continue;
        if (c < destWidth) destRow[c] = sourceRow[c];
      }

      Object.keys(mapping).forEach(targetCol => {
        let sIdx = colToIdx(mapping[targetCol]);
        let tIdx = colToIdx(targetCol);
        if (sourceRow[sIdx] !== undefined) destRow[tIdx] = sourceRow[sIdx];
      });
      pasteArray.push(destRow);
    }
  }

  // If no Total row was ever found, add a generic one at the end
  if (!totalFound) addFooterToPasteArray("TOTAL"); 

  // --- FINAL STEP: AUTO-CLEAR ROW 13+, PASTE, AND ALIGN ---
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  const rowsToClear = maxRows - dataStartRow + 1;

  if (rowsToClear > 0) {
    sheet.getRange(dataStartRow, 1, rowsToClear, maxCols).clearContent();
  }
  
  if (pasteArray.length > 0) {
    const destinationRange = sheet.getRange(dataStartRow, 1, pasteArray.length, destWidth);
    destinationRange.setValues(pasteArray);
    destinationRange.setHorizontalAlignment("center");
    destinationRange.setVerticalAlignment("middle");
  }
  
  SpreadsheetApp.getActive().toast(`Fetch complete. Subtotals retained correctly. Data on/after TOTAL row ignored.`, "Success");
}

/* =================================
4. RUN FORMULAS (FULL UPDATED VERSION)
================================= */
function applyFormulasToSheet(tabName) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return;

  // 1. Mandatory Validations
  const valL5 = sheet.getRange("L5").getValue();
  const valL6 = sheet.getRange("L6").getValue();

  if (valL5 === "" || valL6 === "" || isNaN(valL5) || isNaN(valL6)) {
    SpreadsheetApp.getUi().alert("❌ Action Blocked: L5 and L6 must contain numeric values.");
    return;
  }

  if (tabName === "LPG") {
    const valN10 = sheet.getRange("N10").getValue();
    if (valN10 === "" || isNaN(valN10)) {
      SpreadsheetApp.getUi().alert("❌ Action Blocked: N10 must contain a numeric value for LPG formulas.");
      return;
    }
  }

  if (tabName === "Water") {
    const valU10 = sheet.getRange("U10").getValue();
    if (valU10 === "" || isNaN(valU10)) {
      SpreadsheetApp.getUi().alert("❌ Action Blocked: U10 must contain a numeric value for Water formulas.");
      return;
    }
  }

  if (Number(valL5) <= Number(valL6)) {
    SpreadsheetApp.getUi().alert("❌ Action Blocked: L5 must be greater than L6.");
    return;
  }

  // --- FORMULA DEFINITIONS ---
  const formulaMapElec = {
    L: (r) => `=IFERROR((K${r}-J${r})*I${r},"-")`,
    O: (r) => `=IF(NOT(ISNUMBER($L$5)),"-",$L$5)`,
    P: (r) => `=IFERROR(IF(O${r}="fix rate","Put/input",ROUND(L${r}*O${r}, 2)),"-")`,
    Q: (r) => `=IFERROR(ROUND(P${r}*1.12, 2), "-")`,
    Z: (r) => `=$L$6`,
    AA: (r) => `=IFERROR(L${r}*Z${r},"-")`,
    AB: (r) => `=IFERROR(P${r}-AA${r},"-")`,
    AC: (r) => `=IFERROR((O${r}-Z${r})/Z${r}, "-")`,
    AG: (r) => `=IFERROR(L${r}-AF${r},"-")`,
    AH: (r) => `=IFERROR(AG${r}/AF${r},"-")`,
    AJ: (r) => `=IFERROR(P${r}-AI${r},"-")`,
    AK: (r) => `=IFERROR(AI${r}/AJ${r},"-")`,
  };

  const formulaMapWater = {
    L: (r) => `=IFERROR(K${r}-J${r}, "-")`,
    O: (r) => `=IF(NOT(ISNUMBER($L$5)),"-",$L$5)`,
    P: (r) => `=IFERROR(IF(O${r}="fix rate", "Put/input", ROUND(ROUND(O${r}, 2) * ROUND(L${r}, 2), 2)),"-")`,
    S: (r) => `=IF(NOT(ISNUMBER($U$10)),"-",$U$10)`,
    T: (r) => `=IFERROR(S${r}*L${r},"-")`,
    U: (r) => `=IFERROR(L${r}+T${r},"-")`,
    V: (r) => `=IF(OR(J${r}="Fix Rate", O${r}="Fix Rate"), "-", IF(AND(ISNUMBER(P${r}), ISNUMBER(S${r})), P${r}*S${r}, "-"))`,
    W: (r) => `=IFERROR(V${r}+P${r},"-")`,
    X: (r) => `=IFERROR(IF(J${r}="fix rate", ROUND(P${r}*1.12, 2), ROUND(W${r}*1.12, 2)), "-")`,
    Z: (r) => `=IF(NOT(ISNUMBER($L$6)),"-",$L$6)`,
    AA: (r) => `=IFERROR(L${r}*Z${r},"-")`,
    AB: (r) => `=IFERROR(W${r}-AA${r},"-")`,
    AC: (r) => `=IFERROR((O${r}-Z${r})/Z${r}, "-")`,
    AG: (r) => `=IFERROR(L${r}-AF${r},"-")`,
    AH: (r) => `=IFERROR(AG${r}/AF${r},"-")`,
    AJ: (r) => `=IFERROR(W${r}-AI${r},"-")`,
    AK: (r) => `=IFERROR(AI${r}/AJ${r},"-")`,
  };

  const formulaMapLPG = {
    L: (r) => `=IFERROR(K${r}-J${r}, "-")`,
    M: (r) => `=if(not(isnumber($N$10)),".",$N$10)`,
    N: (r) => `=iferror(L${r}*M${r},"-")`,
    O: (r) => `=IF(NOT(ISNUMBER($L$5)),"-",$L$5)`,
    P: (r) => `=IFERROR(IF(O${r}="fix rate","Put/input", ROUND(N${r}*O${r}, 2)),"-")`,
    Q: (r) => `=IFERROR(ROUND(P${r}*1.12, 2), "-")`,
    Z: (r) => `=IF(NOT(ISNUMBER($L$6)),"-",$L$6)`,
    AA: (r) => `=IFERROR(N${r}*Z${r},"-")`,
    AB: (r) => `=IFERROR(P${r}-AA${r},"-")`,
    AC: (r) => `=IFERROR((O${r}-Z${r})/Z${r}, "-")`,
    AG: (r) => `=IFERROR(L${r}-AF${r},"-")`,
    AH: (r) => `=IFERROR(AG${r}/AF${r},"-")`,
    AJ: (r) => `=IFERROR(P${r}-AI${r},"-")`,
    AK: (r) => `=IFERROR(AI${r}/AJ${r},"-")`,
  };

  const activeMap = (tabName === "Water") ? formulaMapWater : (tabName === "LPG" ? formulaMapLPG : formulaMapElec);
  const lastRow = sheet.getLastRow();
  const fullDataA = sheet.getRange(1, 1, lastRow, 1).getValues();
  const fullDataE = sheet.getRange(1, colToIdx("E") + 1, lastRow, 1).getValues();

  let stopRow = lastRow;
  for (let i = CONFIG.dataStartRow - 1; i < lastRow; i++) {
    if (String(fullDataA[i][0]).toLowerCase().trim() === "total") {
      stopRow = i + 1;
      break;
    }
  }

  for (let i = CONFIG.dataStartRow - 1; i < stopRow; i++) {
    const r = i + 1;
    const labelA = String(fullDataA[i][0]).toLowerCase().trim();
    const valE = String(fullDataE[i][0]).trim();

    if (labelA.includes("total")) continue;

    if (valE === "") {
      sheet.getRange(r, 1, 1, sheet.getLastColumn()).clearContent();
      continue;
    }

    const rowData = sheet.getRange(r, 1, 1, 35).getValues()[0];
    const valO = String(rowData[14] || "").trim();
    const valP = String(rowData[15] || "").trim();
    const valZ = String(rowData[25] || "").trim();
    const valJ = String(rowData[9] || "").toLowerCase();
    const valK = String(rowData[10] || "").toLowerCase();

    let targetCols = Object.keys(activeMap);

    if (valO !== "") targetCols = targetCols.filter(c => c !== "O");
    if (valP !== "") targetCols = targetCols.filter(c => c !== "P");
    if (valZ !== "") targetCols = targetCols.filter(c => c !== "Z");
    if (valJ.includes("theoretical")) targetCols = targetCols.filter(c => c !== "L");
    if (valK.includes("theoretical")) targetCols = targetCols.filter(c => c !== "L");

    targetCols.forEach(colKey => {
      sheet.getRange(`${colKey}${r}`).setFormula(activeMap[colKey](r));
    });
  }

  const sumCols = ["L", "N", "P", "Q", "AA", "AB", "AF", "AG", "AI", "AJ"];
  let sectionStartRow = CONFIG.dataStartRow;
  let subTotalRowsFound = [];

  for (let i = CONFIG.dataStartRow - 1; i < stopRow; i++) {
    const rowLabel = String(fullDataA[i][0]).toLowerCase().trim();
    const r = i + 1;
    const normalizedLabel = rowLabel.replace(/[^a-z]/g, "");

    if (normalizedLabel.includes("subtotal")) {
      const rangeEnd = r - 1;
      sumCols.forEach(col => {
        sheet.getRange(`${col}${r}`).setFormula(`=SUM(${col}${sectionStartRow}:${col}${rangeEnd})`);
      });
      subTotalRowsFound.push(r);
      sectionStartRow = r + 1;
    }

    if (normalizedLabel === "total") {
      sumCols.forEach(col => {
        let formula = "";
        if (subTotalRowsFound.length > 0) {
          let refs = subTotalRowsFound.map(subR => `${col}${subR}`).join(",");
          formula = `=SUM(${refs})`;
        } else {
          const rangeEnd = r - 1;
          formula = `=SUM(${col}${CONFIG.dataStartRow}:${col}${rangeEnd})`;
        }
        sheet.getRange(`${col}${r}`).setFormula(formula);
      });
    }
  }

  const lastSheetRow = sheet.getLastRow();
  if (lastSheetRow > stopRow) {
    const footerRange = sheet.getRange(stopRow + 1, 1, lastSheetRow - stopRow, sheet.getLastColumn());
    const footerValues = footerRange.getValues();
    const cleanedFooter = footerValues.map(row => row.map(cell => (typeof cell === 'number' && cell !== "") ? "" : cell));
    footerRange.setValues(cleanedFooter);
  }

  // --- FINAL FORMATTING ---
  SpreadsheetApp.flush();

  // 1. Standard numeric format for calculation columns
  ["P", "Q", "J", "L", "AF", "AI", "AJ", "AG"].forEach(c => {
    sheet.getRange(`${c}${CONFIG.dataStartRow}:${c}${stopRow}`).setNumberFormat("#,##0.00");
  });

  // 2. Percentage format for AH and AK (Percentage with 2 decimal places)
  ["AH", "AK"].forEach(c => {
    sheet.getRange(`${c}${CONFIG.dataStartRow}:${c}${stopRow}`).setNumberFormat("0.00%");
  });

  SpreadsheetApp.getActive().toast(`Formula logic and percentage formatting complete for ${tabName}.`, "Success");
}
/* =================================
5. TRIGGER WRAPPERS
================================= */
function fetchElec() { if (confirmFetchOverwrite("Elec")) fetchDataOnly("Elec"); }
function runFormulaElec() { applyFormulasToSheet("Elec"); }
function clearElec() { clearTabData("Elec"); }
function scanElecTab() { scanTab("Elec"); }

function fetchWater() { if (confirmFetchOverwrite("Water")) fetchDataOnly("Water"); }
function runFormulaWater() { applyFormulasToSheet("Water"); }
function clearWater() { clearTabData("Water"); }
function scanWaterTab() { scanTab("Water"); }

function fetchLPG() { if (confirmFetchOverwrite("LPG")) fetchDataOnly("LPG"); }
function runFormulaLPG() { applyFormulasToSheet("LPG"); }
function clearLPG() { clearTabData("LPG"); }
function scanLPGTab() { scanTab("LPG"); }

/* =================================
6. UTILITIES (CLEANED UP)
================================= */
function clearTabData(tabName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
  if (sheet && sheet.getLastRow() >= CONFIG.dataStartRow) {
    sheet.getRange(CONFIG.dataStartRow, 1, sheet.getLastRow() - CONFIG.dataStartRow + 1, sheet.getMaxColumns()).clearContent();
  }
}



/* =================================
6. FINAL SCAN TAB (Standardized Conditions)
================================= */

function scanTab(tabName, shouldClearLogs = true) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.dataStartRow) return;

  const setupLogSheet = (name) => {
    let s = ss.getSheetByName(name) || ss.insertSheet(name);
    if (shouldClearLogs && s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow(), 6).clearContent();
    if (s.getLastRow() === 0) s.appendRow(["Timestamp", "Tab", "Cell", "Column Label", "Error Message", "Remarks"]);
    return s;
  };

  const standardLogSheet = setupLogSheet("Basic Anomalies");
  const kaLogSheet = setupLogSheet("Client Rate Anomalies");

  const dataRange = sheet.getRange(CONFIG.dataStartRow, 1, lastRow - CONFIG.dataStartRow + 1, 37);
  const dataValues = dataRange.getValues();
  const headers = sheet.getRange(CONFIG.headerRow, 1, 1, 37).getValues()[0];
  const valL5 = sheet.getRange("L5").getValue();
  const valL6 = sheet.getRange("L6").getValue();
  const rawE4 = sheet.getRange("E4").getValue();
  
  const kaRefMap = getKAData(); 
  const issueLogs = [];
  const kaLogs = [];

  const logHelper = (rowArr, rNum, colLet, msg, internalReason = "", logArray = issueLogs) => {
    const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MMM d, yyyy");
    const index29Value = rowArr[29] || "";
    const finalRemarks = internalReason ? `${internalReason} | Remarks: ${index29Value}` : index29Value;
    logArray.push([timestamp, tabName, `${colLet}${rNum}`, headers[colToIdx(colLet)] || colLet, msg, finalRemarks]);
  };

  for (let i = 0; i < dataValues.length; i++) {
    const rowNum = CONFIG.dataStartRow + i;
    const row = dataValues[i];

    // --- STEP 1: RUN CHECKLIST FIRST (This triggers the Col A/E hard-stop) ---
    runCommonChecklist(row, rowNum, (r, c, m, res, arr) => logHelper(row, r, c, m, res, arr), valL5, valL6);

    const labelA = String(row[0]).trim();
    const valE = String(row[colToIdx("E")] || "").trim();

    // --- STEP 2: SKIP ROW LOGIC ---
    // Only skip if the row is actually empty (No A and No E) or if it's a Total
    const normalizedLabelA = labelA.toLowerCase().replace(/[^a-z]/g, "");
    if (normalizedLabelA.includes("total")) continue;
    if (labelA === "" && valE === "") continue;

    // --- STEP 3: KA VALIDATION ---
    if (kaRefMap) {
      const valF = String(row[colToIdx("F")] || "").trim().toUpperCase();
      const valG = String(row[colToIdx("G")] || "").trim().toUpperCase();
      const hasKA = (valF === "KA" || valG === "KA");

      // Check Database for Match
      const matchedKey = findReferenceKey(valE, kaRefMap);
      const validCategories = matchedKey ? kaRefMap[matchedKey] : [];
      const headerE4 = superClean(rawE4); 
      let isMatch = false;

      // Determine if Site Identity (E4) matches Category assigned to Tenant
      if (validCategories.length > 0) {
        for (let k = 0; k < validCategories.length; k++) {
          let keyword = superClean(validCategories[k]);
          if (keyword !== "" && (headerE4.includes(keyword) || keyword.includes(headerE4))) {
            isMatch = true;
            break;
          }
        }
      }

      // Logic check: Calculation result vs Manual "KA" flag
      if (isMatch) {
        if (!hasKA) logHelper(row, rowNum, "F", 'user need to put "KA"', `DB match: [${valE}]`, kaLogs);
      } else {
        if (hasKA) logHelper(row, rowNum, "F", 'user need to remove "KA"', `No DB entry found for [${valE}] in Site [${headerE4}]`, kaLogs);
      }
    }

    // --- STEP 4: TAB SPECIFIC CALCULATIONS ---
    switch(tabName) {
      case "Elec":
        if (!(typeof row[colToIdx("Q")] === 'number' && row[colToIdx("Q")] > 0)) logHelper(row, rowNum, "Q", "Amount should be a number > 0");
        break;
      case "Water":
        ["S", "T", "U", "V", "W"].forEach(c => { if (String(row[colToIdx(c)]).trim() === "") logHelper(row, rowNum, c, "Formula output missing"); });
        if (!(typeof row[colToIdx("X")] === 'number' && row[colToIdx("X")] > 0)) logHelper(row, rowNum, "X", "VAT amount missing");
        break;
      case "LPG":
        const vL = row[colToIdx("L")];
        if (typeof vL === 'number') {
          if (!(typeof row[colToIdx("M")] === 'number' && row[colToIdx("M")] > 0)) logHelper(row, rowNum, "M", "Multiplier missing");
          if (!(typeof row[colToIdx("N")] === 'number' && row[colToIdx("N")] > 0)) logHelper(row, rowNum, "N", "Consumption amount error");
        }
        break;
    }
  }

  // --- WRITE TO LOGS ---
  if (issueLogs.length > 0) {
    const sIdx = standardLogSheet.getLastRow() + 1;
    standardLogSheet.getRange(sIdx, 1, issueLogs.length, 6).setValues(issueLogs);
  }
  if (kaLogs.length > 0) {
    const kIdx = kaLogSheet.getLastRow() + 1;
    kaLogSheet.getRange(kIdx, 1, kaLogs.length, 6).setValues(kaLogs);
  }

  console.log(`Scan Tab ${tabName} completed.`);
}

function findReferenceKey(cellValue, kaRefMap) {
  if (!cellValue) return null;
  const searchStr = superClean(cellValue);
  if (kaRefMap[searchStr]) return searchStr;

  const refKeys = Object.keys(kaRefMap);
  for (let i = 0; i < refKeys.length; i++) {
    const key = refKeys[i];
    if (key !== "" && (searchStr.includes(key) || key.includes(searchStr))) return key;
  }
  return null;
}

/* =================================
   OTHER HELPERS (KEEP EXISTING)
================================= */
function superClean(val) {
  if (!val) return "";
  let str = String(val).toLowerCase();
  str = str.replace(/[^a-z0-9\s]/g, ' ').replace(/[\s\u00A0]+/g, ' ').trim();
  return str;
}


/**
 * Maps Column B to an array of valid Column C values.
 * Allows one property to have multiple valid categories.
 */
/**
 * Maps Column B AND Column E (iterations) to Column C values.
 */
function getKAData() {
  const KA_REF_URL = "https://docs.google.com/spreadsheets/d/1jY-9FMha3x972o4Gz1d6DVD36d3ppjHW_WM1DHJz6ag/edit";
  try {
    const ss = SpreadsheetApp.openByUrl(KA_REF_URL);
    const sheet = ss.getSheetByName("Data");
    if (!sheet) throw new Error("Master sheet 'Data' not found.");

    const lastR = sheet.getLastRow();
    if (lastR < 2) return {};

    // Get 5 columns: A (0), B (1), C (2), D (3), E (4)
    const rawData = sheet.getRange(2, 1, lastR - 1, 5).getValues(); 
    const propertyMap = {};

    for (let i = 0; i < rawData.length; i++) {
      const row = rawData[i];
      
      const valA = String(row[0] || "").trim();      // ID (Col A)
      const valB = String(row[1] || "").trim();      // Main Name (Col B)
      const valE = String(row[4] || "").trim();      // Iterations (Col E)
      const category = superClean(row[2]);           // Category (Col C)
      
      const currentRowNum = i + 2;

      // --- ADDED CHECKER PER REQUEST ---
      // Specifically checks if Col E has data but Col A does not
      if (valE !== "" && valA === "") {
        const specificMsg = `🛑 MASTER DATABASE ERROR (Row ${currentRowNum})\n\nColumn E contains values, but Column A is blank. You must input a number first in Column A of the Master File to proceed.`;
        SpreadsheetApp.getUi().alert(specificMsg);
        throw new Error("Aborted: Missing number in Master Column A.");
      }

      // Maintain general safety for Col B as well
      if (valB !== "" && valA === "") {
        const errorMsg = `🛑 MASTER DATABASE ERROR\n\nRow ${currentRowNum} has a Main Name (Col B) but is missing an Identifier in Column A.\n\nPlease fix the Master File to proceed.`;
        SpreadsheetApp.getUi().alert(errorMsg);
        throw new Error("Master Data Violation: Missing Column A.");
      }
      // ---------------------------------

      const addKey = (name) => {
        let cleanedName = superClean(name);
        if (!cleanedName) return;
        if (!propertyMap[cleanedName]) propertyMap[cleanedName] = [];
        if (!propertyMap[cleanedName].includes(category)) propertyMap[cleanedName].push(category);
      };

      addKey(valB);
      if (valE) valE.split(",").forEach(part => addKey(part));
    }
    return propertyMap;

  } catch (e) {
    // Re-throw if it's one of our validation errors to ensure the whole scan stops
    if (e.message.includes("Aborted") || e.message.includes("Violation")) throw e;
    
    console.error("KA Ref Error: " + e.message);
    return null;
  }
}


/* =================================
REFACTORED: THE "COMMON" CHECKLIST (ALL TABS)
================================= */
function runCommonChecklist(row, rNum, log, L5, L6) {
  // Helper to fetch value by Column Letter
  const get = (colLetter) => row[colToIdx(colLetter)];
  
  // Clean values for Column A and Column E
  const valA = String(get("A") || "").trim();
  const valE = String(get("E") || "").trim();
  
  // --- MANDATORY IDENTIFIER CHECK (HARD STOP) ---
  // If Column E (Tenant) is populated, Column A (No. or Area) MUST have a value.
  if (valE !== "" && valA === "") {
    const errorMsg = `CRITICAL DATA ERROR\n\nRow ${rNum} has a Tenant Name in Column E ("${valE}") but the identifier in Column A is blank.\n\nPROCESS HALTED: Every tenant must have an Row Number in Column A to continue.`;
    
    SpreadsheetApp.getUi().alert(errorMsg);
    throw new Error(`Execution stopped at row ${rNum}: Missing Col A with populated Col E.`);
  }


  // --- STANDARD CALCULATION CHECKS ---
  const valL = get("L");
  const L_isHyphen = (String(valL).trim() === "-");
  
  // Run logic ONLY if Col E has an entry
  if (valE !== "") {

    // J, K, L Conditions: Ensure mandatory reading/results are present
    ["J", "K", "L"].forEach(c => {
      if (String(get(c)).trim() === "") log(rNum, c, "Should not be blank if E has entry");
    });

    // Column O Conditions: Billing Rate
    const valO = get("O");
    const oStr = String(valO).toLowerCase().trim();
    const oIsFixOrTheo = (oStr === "fix rate" || oStr === "theoretical");

    if (typeof valO === 'number') {
      if (!(valO > 0)) log(rNum, "O", "Should equal to L5, \"fix rate\" or \"theoretical\"");
      if (!L_isHyphen && valO !== L5) log(rNum, "O", "Should equal to L5, or if L= \"-\" then, O= \"fix rate\" or O=\"theoretical\"");
    } else {
      if (!oIsFixOrTheo) log(rNum, "O", "Should equal to L5, \"fix rate\" or \"theoretical\"");
      if (L_isHyphen && !oIsFixOrTheo) log(rNum, "O", "Should equal to L5, or if L= \"-\" then, O= \"fix rate\" or O=\"theoretical\"");
    }

    // Column P Condition: Basic Amount
    if (!(typeof get("P") === 'number' && get("P") > 0)) log(rNum, "P", "Should be a number >0");

    // Column Z Conditions: Reference/Comparative Rate
    const valZ = get("Z");
    if (valZ === "") log(rNum, "Z", "Should be a number >0, \"fix rate\" or \"theoretical\"");
    if (typeof valZ === 'number' && !L_isHyphen && valZ !== L6) {
      log(rNum, "Z", "Should equal to L6, or if L= \"-\" then, O= \"fix rate\" or O=\"theoretical\"");
    }

    // Standard Formula Output Verification (Columns driven by Formulas)
    ["AA", "AB", "AC", "AG", "AJ"].forEach(c => {
      if (String(get(c)).trim() === "") log(rNum, c, "Should not be empty (c/o fx)");
    });

    // Multi-period Consistency (Source Column Comparisons)
    ["AF", "AI"].forEach(c => {
      if (String(get(c)).trim() === "") log(rNum, c, "Should not be empty if E has entry");
    });

    // Variances / Consumption Flags: Alerts if >30% change (AH=Vol variance, AK=Cost variance)
    ["AH", "AK"].forEach(c => {
      const v = get(c);
      if (typeof v === 'number') {
        if (v > 0.3 || v < -0.3) log(rNum, c, "Variance alert: Value is outside +/- 30% threshold.");
      }
    });
  }
}

function confirmFetchOverwrite(tabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  const ui = SpreadsheetApp.getUi();
  
  if (!sheet) return false;

  const lastSheetRow = sheet.getLastRow();
  let hasExistingData = false;

  // 1. Only bother checking if the sheet has rows up to or past the dataStartRow
  if (lastSheetRow >= CONFIG.dataStartRow) {
    // 2. Fetch all values specifically in Column E (index 5)
    const colE_values = sheet.getRange(CONFIG.dataStartRow, 5, lastSheetRow - CONFIG.dataStartRow + 1, 1).getValues();
    
    // 3. Smart check: Ignore blanks, unchecked boxes (false), null, and empty spaces
    hasExistingData = colE_values.some(row => {
      const val = row[0];
      if (val === "" || val === null || val === undefined || val === false) return false;
      return String(val).trim() !== ""; // Returns true ONLY if real data exists
    });
  }

  // 4. Show modal ONLY if we confirmed there's actual data in Col E
  if (hasExistingData) {
    const res = ui.alert(
      'Confirm Overwrite', 
      `Data already exists in "${tabName}". Overwrite?`, 
      ui.ButtonSet.YES_NO
    );
    // Return false if they click NO or close the dialog
    if (res !== ui.Button.YES) return false; 
  }

  // 5. Proceed as normal if there is no data OR if they clicked YES
  return true;
}



/**
 * RE-INITIALIZATION: 
 * If you ever need to reset to the original file, 
 * run the "resetDatabaseID" function at the bottom.
 */


function recordActivePBTT() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); 
  } catch (e) {
    SpreadsheetApp.getUi().alert("Server Busy. Please try again.");
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterDbId = "1hMMUd4ho50HP63dc2fRAo--iK-m7YotamkKtsDGT_Us"; 

    // Define the specific columns to check for each tab
    const tabValidationMaps = {
      "Elec":  ["L","O","P","Q","Z","AA","AB","AC","AG","AH","AJ","AK"],
      "Water": ["L","O","P","S","T","U","V","W","X","Z","AA","AB","AC","AG","AH","AJ","AK"],
      "LPG":   ["L","M","N","O","P","Q","Z","AA","AB","AC","AG","AH","AJ","AK"]
    };

    // 1. GET THE PERMANENT REF# FROM INSTRUCTIONS
    const instSheet = ss.getSheetByName("Instructions");
    const currentRef = instSheet ? instSheet.getRange("C14").getValue().toString().trim() : "";
    if (currentRef === "") {
      SpreadsheetApp.getUi().alert("❌ ERROR: No Reference Number found in 'Instructions' tab C14.");
      return;
    }

    // ===============================================
    // 2. PERIOD LOCK CHECKER
    // ===============================================
    const dbPeriodTab = "dvPeriod";
    try {
      const dbSs = SpreadsheetApp.openById(masterDbId);
      const periodSheet = dbSs.getSheetByName(dbPeriodTab);
      const periodData = periodSheet.getDataRange().getValues();
      
      let activeFound = false;
      let allowed = false;
      let expiredDates = []; // Store expiration dates if they fail the date check

      for (let i = 1; i < periodData.length; i++) {
        // Checking column D (index 3) for "Active"
        if (String(periodData[i][3]).trim() === "Active") {
          activeFound = true;
          const today = new Date(); 
          today.setHours(0, 0, 0, 0);
          
          const lockD = new Date(periodData[i][2]); 
          lockD.setHours(0, 0, 0, 0);
          
          if (today < lockD) {
            // We found an 'Active' period that is NOT expired.
            // Permission to proceed granted!
            allowed = true;
            break; // Stop looking, as we only need one valid Active period to allow submission.
          } else { 
            // It is Active but Expired. Add date to list for our alert message.
            expiredDates.push(Utilities.formatDate(lockD, "Asia/Manila", "MMM d, yyyy")); 
          }
        }
      }
      
      if (!activeFound) {
        SpreadsheetApp.getUi().alert("🚫 NO ACTIVE PERIOD\n\nPlease set a period to 'Active' in the database (dvPeriod tab) first before proceeding.");
        return;
      }
      
      if (!allowed) {
        // All active periods were found to be expired
        const expiredDatesMsg = expiredDates.join("\n- ");
        SpreadsheetApp.getUi().alert(`🚫 PERIOD LOCKED\n\nAll 'Active' periods have already expired. Expiration dates:\n- ${expiredDatesMsg}`);
        return;
      }

    } catch (err) { 
      SpreadsheetApp.getUi().alert("❌ Validation Error: " + err.message); 
      return; 
    }

    // --- 3. TAB-SPECIFIC VALIDATION PASS ---
    for (let tabName in tabValidationMaps) {
      let currentSheet = ss.getSheetByName(tabName);
      if (!currentSheet) continue; 

      let lastRow = currentSheet.getLastRow();
      let startRow = 13;
      if (lastRow < startRow) continue;

      // Fetch up to Column AK (37 columns)
      let dataRange = currentSheet.getRange(startRow, 1, lastRow - startRow + 1, 37).getValues();
      let requiredCols = tabValidationMaps[tabName];

      for (let i = 0; i < dataRange.length; i++) {
        let rowData = dataRange[i];
        let valA = String(rowData[0] || "").toLowerCase();

        // STOP checking this tab if we hit the TOTAL row
        if (valA.includes("total") && !valA.includes("sub")) break;

        let rawValE = rowData[4]; 
        let valE = (rawValE === undefined || rawValE === null) ? "" : String(rawValE).trim();
        
        // If Column E has data, validate the tab-specific columns
        if (valE !== "") {
          for (let colLetter of requiredCols) {
            let colIdx = colToIdx(colLetter); // Ensure colToIdx() helper exists in your script
            let rawCellVal = rowData[colIdx];
            
            // Convert to string safely (preserves numeric 0 and boolean false)
            let cellValue = (rawCellVal === undefined || rawCellVal === null) ? "" : String(rawCellVal).trim();

            // Check if purely blank. 0, 0.00, and "-" are no longer blocked!
            if (cellValue === "") {
              SpreadsheetApp.getUi().alert(
                `🚫 INCOMPLETE DATA\n\n` +
                `Tab: [${tabName}]\n` +
                `Row: ${i + startRow}\n` +
                `Column: ${colLetter}\n\n` +
                `This cell is completely blank but is required because Column E has a value. Please fix before recording.`
              );
              return; // STOP the entire process
            }
          }
        }
      }
    }

    // 4. FIND THE CORRECT SHEET TO EXTRACT HEADERS (Prop Name, Dates)
    let activeSheet = ss.getActiveSheet();
    let headerSheet = tabValidationMaps[activeSheet.getName()] ? activeSheet : 
                      Object.keys(tabValidationMaps).map(n => ss.getSheetByName(n)).find(s => s !== null);
    
    if (!headerSheet) {
      SpreadsheetApp.getUi().alert("🚫 Error: No utility tabs (Elec, Water, LPG) found.");
      return;
    }

    // Assumes processSheetHeaders() is available
    const extractedData = processSheetHeaders(headerSheet);
    if (!extractedData) return;

    // 5. DUPLICATE CHECK
    const props = PropertiesService.getScriptProperties();
    let activeDB_ID = props.getProperty("ACTIVE_DB_ID") || masterDbId;
    let db = SpreadsheetApp.openById(activeDB_ID);
    let dSh = db.getSheetByName("PBTT Submission");
    let lastRowDb = dSh.getLastRow();

    if (lastRowDb >= 5) {
      const dbRangeValues = dSh.getRange(5, 2, lastRowDb - 4, 10).getValues();
      const rawStart = headerSheet.getRange("E7").getValue();
      const rawEnd = headerSheet.getRange("E8").getValue();
      const fmt = (d) => d ? Utilities.formatDate(new Date(d), "Asia/Manila", "yyyy-MM-dd") : "";
      const currentStart = fmt(rawStart);
      const currentEnd = fmt(rawEnd);

      for (let i = 0; i < dbRangeValues.length; i++) {
        const row = dbRangeValues[i];
        
        if (String(row[9]).trim() === currentRef) {
          SpreadsheetApp.getUi().alert(`🚫 BLOCK: Reference ${currentRef} is already recorded.`);
          return;
        }
        
        const isDataMatch = (
          String(row[0]).trim().toLowerCase() === String(extractedData[0]).trim().toLowerCase() && 
          String(row[1]).trim().toLowerCase() === String(extractedData[1]).trim().toLowerCase() && 
          fmt(row[4]) === currentStart && fmt(row[5]) === currentEnd
        );
        if (isDataMatch) {
          SpreadsheetApp.getUi().alert("🚫 DUPLICATE ERROR: Billing period already submitted.");
          return;
        }
      }
    }

    // 6. DATABASE SUBMISSION
    const timestamp = Utilities.formatDate(new Date(), "Asia/Manila", "MMM d, yyyy hh:mm a");
    const userEmail = Session.getActiveUser().getEmail();
    const ssUrl = ss.getUrl();
    const ssName = ss.getName();

    const finalRow = [timestamp, ...extractedData, ssName, ssUrl, userEmail, currentRef];
    dSh.appendRow(finalRow);

    SpreadsheetApp.getUi().alert(`✅ SUCCESS: Submitted successfully.`);

  } catch (x) {
    SpreadsheetApp.getUi().alert("System Error: " + x.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper: Column Letter to 0-based Index
 */
function colToIdx(letter) {
  let column = 0, length = letter.length;
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column - 1;
}
/**
 * Handles cleaning, concatenating, deduping values, 
 * and checking for correct dates.
 */
function processSheetHeaders(sheet) {
    const config = [
        { cell: "E4", label: "PROPERTY NAME", type: "text" },
        { cell: "E6", label: "BILLER/PAYEE COMPANY:", type: "text" },
        { cell: "E5", label: "LOCATION", type: "text", sourceRange: "B13:B" },
        { cell: "E11", label: "PROVIDER & ACCOUNT NO:", type: "text", sourceRange: "Y13:Y" },
        { cell: "E7", label: "START DATE", type: "date" },
        { cell: "E8", label: "END DATE", type: "date" },
    ];

    const results = [];
    const missing = [];

    // --- 1. DATE VALIDATION ---
    const startDateValue = sheet.getRange("E7").getValue();
    const endDateValue = sheet.getRange("E8").getValue();

    if (
        !(startDateValue instanceof Date) ||
        isNaN(startDateValue) ||
        !(endDateValue instanceof Date) ||
        isNaN(endDateValue)
    ) {
        SpreadsheetApp.getUi().alert("❌ ERROR: Start Date or End Date is empty or invalid.");
        return null;
    }

    const startDate = new Date(startDateValue);
    const endDate = new Date(endDateValue);
    const today = new Date();

    // Basic logical check
    if (endDate <= startDate) {
        SpreadsheetApp.getUi().alert("❌ DATE ERROR: End Date (E8) must be after Start Date (E7).");
        return null;
    }

    // --- NEW MONTH-MATCH WARNING ---
    const currentMonth = today.getMonth(); // 0-11
    const currentYear = today.getFullYear();
    const endMonth = endDate.getMonth();
    const endYear = endDate.getFullYear();

    // If month or year don't match today, trigger warning
    if (currentMonth !== endMonth || currentYear !== endYear) {
        const formattedEnd = Utilities.formatDate(endDate, "GMT+8", "MMMM yyyy");
        const ui = SpreadsheetApp.getUi();
        const response = ui.alert(
            "⚠️ CHECK DATE PERIOD",
            `The End Date is currently set to: ${formattedEnd}.\n\n` +
                `Note: This does NOT match today's month.\n` +
                `Is this period correct for your submission?`,
            ui.ButtonSet.YES_NO
        );

        if (response !== ui.Button.YES) {
            return null; // Stop submission if user clicks No
        }
    }

    // --- 2. HEADER DATA EXTRACTION ---
    for (let item of config) {
        let val = sheet.getRange(item.cell).getValue();

        if (item.sourceRange) {
            let rawItems = [];
            if (val)
                val.toString()
                    .split(",")
                    .forEach((p) => rawItems.push(p.trim()));

            const lastR = sheet.getLastRow();
            const colLetter = item.sourceRange.substring(0, 1);

            if (lastR >= 13) {
                // Get Column A (for TOTAL check)
                const colAValues = sheet
                    .getRange("A13:A" + lastR)
                    .getValues()
                    .flat();

                // Get the source column values (B or Y)
                const sourceValues = sheet
                    .getRange(colLetter + "13:" + colLetter + lastR)
                    .getValues()
                    .flat();

                for (let i = 0; i < colAValues.length; i++) {
                    const colAValue = colAValues[i];

                    // STOP if TOTAL is found in Column A
                    if (colAValue && colAValue.toString().trim().toUpperCase() === "TOTAL") {
                        break;
                    }

                    const cellValue = sourceValues[i];
                    if (cellValue && cellValue.toString().trim() !== "") {
                        cellValue
                            .toString()
                            .split(",")
                            .forEach((p) => rawItems.push(p.trim()));
                    }
                }
            }

            // Deduplicate and clean
            const uniqueItems = Array.from(new Set(rawItems.map((s) => s.trim()))).filter(Boolean);

            val = uniqueItems.join(", ");
        }

        if (!val && val !== 0) missing.push(item.label);
        if (item.type === "date" && val)
            val = Utilities.formatDate(new Date(val), Session.getScriptTimeZone(), "MMM d, yyyy");
        results.push(val);
    }

    if (missing.length > 0) {
        SpreadsheetApp.getUi().alert("Missing Header Info: " + missing.join(", "));
        return null;
    }
    return results;
}
/**
 * Calculates total cell count (Max Rows * Max Cols) across all tabs in a file.
 */
function getTotalCellCount(ss) {
  let total = 0;
  const sheets = ss.getSheets();
  sheets.forEach(sh => {
    total += (sh.getMaxRows() * sh.getMaxColumns());
  });
  return total;
}


/**
 * NEW LOGIC: When database is full, create a new one.
 * It will ALWAYS pull the header from the original MASTER FILE Row 4
 * and place it into Row 4 of the new file.
 */
function rotateToNewDatabase(oldDb, oldSheet) {
  const folder = DriveApp.getFolderById(BACKUP_FOLDER_ID);
  const time = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd_HHmmss");
  const newName = "PBTT_Submission_Database_" + time;

  // 1. Create the new spreadsheet file
  const newFile = SpreadsheetApp.create(newName);
  const newFileId = newFile.getId();
  
  // 2. Move to the backup folder
  const driveFile = DriveApp.getFileById(newFileId);
  folder.addFile(driveFile);
  DriveApp.getRootFolder().removeFile(driveFile);

  // 3. Set up the new sheet
  const targetSheetName = "PBTT Submission";
  const newSheet = newFile.insertSheet(targetSheetName);

  // --- HARDCODED MASTER HEADER FETCH ---
  // We use PBTT_DB_ID (your master) to ensure we always get the Row 4 labels
  try {
    const masterSS = SpreadsheetApp.openById(PBTT_DB_ID);
    const masterSheet = masterSS.getSheetByName(targetSheetName);
    
    // We assume the header is roughly 15 columns wide (A to O) 
    // based on your recordActivePBTT data extraction
    const headerWidth = Math.max(masterSheet.getLastColumn(), 15);
    const masterHeaderRange = masterSheet.getRange(4, 1, 1, headerWidth);
    const targetRange = newSheet.getRange(4, 1, 1, headerWidth);
    
    // Copy Values from Master Row 4
    const headerValues = masterHeaderRange.getValues();
    targetRange.setValues(headerValues);
    
    // Copy Styles (Backgrounds, Bold, etc.) from Master Row 4
    targetRange.setBackgrounds(masterHeaderRange.getBackgrounds());
    targetRange.setFontColors(masterHeaderRange.getFontColors());
    targetRange.setFontWeights(masterHeaderRange.getFontWeights());
    targetRange.setHorizontalAlignments(masterHeaderRange.getHorizontalAlignments());
    
    console.log("Successfully copied Row 4 header from Master ID to Row 4 of new file.");
  } catch (e) {
    console.error("Could not fetch master header: " + e.message);
    // Fallback: If Master is unreachable, we try to grab it from the full sheet (oldSheet)
    const fallbackWidth = oldSheet.getLastColumn() || 15;
    const vals = oldSheet.getRange(4, 1, 1, fallbackWidth).getValues();
    newSheet.getRange(4, 1, 1, fallbackWidth).setValues(vals);
  }

  // Delete the blank "Sheet1" that comes with every new spreadsheet
  const defaultSheet = newFile.getSheetByName("Sheet1");
  if (defaultSheet) newFile.deleteSheet(defaultSheet);

  // 4. Update the Registry file
  try {
    const regSs = SpreadsheetApp.openById(BACKUP_REGISTRY_ID);
    const regSh = regSs.getSheetByName("Backup Files") || regSs.insertSheet("Backup Files");
    regSh.appendRow([new Date(), "NEW ACTIVE DB: " + newName, newFile.getUrl()]);
  } catch (e) {
    console.warn("Registry update failed, but file was rotated.");
  }

  // 5. Update Script Properties so future submissions go to the NEW file
  PropertiesService.getScriptProperties().setProperty("ACTIVE_DB_ID", newFileId);

  return newFileId;
}
/**
 * Run this to reset the database submission.
 */
function fullResetDatabasePointer() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("ACTIVE_DB_ID"); // Removes the link to the deleted/new file
  SpreadsheetApp.getUi().alert("Reset successful. The script is now looking at the original MASTER file again.");
}

function checkCurrentDbSize() {
  const props = PropertiesService.getScriptProperties();
  const activeDB_ID = props.getProperty("ACTIVE_DB_ID") || PBTT_DB_ID;
  const db = SpreadsheetApp.openById(activeDB_ID);
  
  const count = getTotalCellCount(db);
  const formattedCount = count.toLocaleString();
  const percent = ((count / 10000000) * 100).toFixed(2);
  
  SpreadsheetApp.getUi().alert(
    `Database Stats:\n\n` +
    `File: ${db.getName()}\n` +
    `Total Cells Used: ${formattedCount}\n` +
    `Capacity Used: ${percent}%`
  );
}
