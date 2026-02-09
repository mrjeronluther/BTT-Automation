/* =================================
CONFIGURATION & MAPPING
================================= */
const SOURCE_DB_URL = "https://docs.google.com/spreadsheets/d/1JNYCjZfGYyVTxYkrws4D7SnrO-hIjDJ61sFqj0WBWEE/edit";
const PBTT_DB_ID    = "16Oai_3c4H_E2wgC-CUkSk1Eez90_KdtlaqHnHJFclBQ";

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
        .addItem("1. Fetch Data", "fetchElec")
        .addItem("2. Run Formulas", "runFormulaElec")
        .addSeparator()
        .addItem("Scan Tab", "scanElecTab")
        .addItem("Clear Tab", "clearElec"))
    .addSubMenu(ui.createMenu("💧 Water")
        .addItem("1. Fetch Data", "fetchWater")
        .addItem("2. Run Formulas", "runFormulaWater")
        .addSeparator()
        .addItem("Scan Tab", "scanWaterTab")
        .addItem("Clear Tab", "clearWater"))
    .addSubMenu(ui.createMenu("🔥 LPG")
        .addItem("1. Fetch Data", "fetchLPG")
        .addItem("2. Run Formulas", "runFormulaLPG")
        .addSeparator()
        .addItem("Scan Tab", "scanLPGTab")
        .addItem("Clear Tab", "clearLPG"))
    .addSeparator()
    .addItem("📤 Submit Active PBTT", "recordActivePBTT")
    .addToUi();
}

/* =================================
2. HELPER: COL LETTER TO INDEX
================================= */
function colToIdx(letter) {
  let column = 0, length = letter.length;
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
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

  const sourceLink = sheet.getRange("A1").getValue();
  if (!sourceLink) { SpreadsheetApp.getUi().alert("Paste SOURCE LINK in A1."); return; }

  let sourceSS;
  try { sourceSS = SpreadsheetApp.openByUrl(sourceLink); } 
  catch (e) { SpreadsheetApp.getUi().alert("Cannot open source link."); return; }
  
  const sourceSheet = sourceSS.getSheetByName(tabName);
  if (!sourceSheet) { SpreadsheetApp.getUi().alert(`Tab "${tabName}" not found in source.`); return; }

  const lastSourceRow = sourceSheet.getLastRow();
  const lastSourceCol = sourceSheet.getLastColumn();
  if (lastSourceRow < CONFIG.dataStartRow) return;

  const rawData = sourceSheet.getRange(CONFIG.dataStartRow, 1, lastSourceRow - CONFIG.dataStartRow + 1, lastSourceCol).getValues();

  let lastNonEmptyIndexInA = -1;
  for (let i = rawData.length - 1; i >= 0; i--) {
    if (String(rawData[i][0]).trim() !== "") { lastNonEmptyIndexInA = i; break; }
  }

  const processedData = (lastNonEmptyIndexInA !== -1) ? rawData.slice(0, lastNonEmptyIndexInA + 1) : [];
  const rowsNeeded = processedData.length;
  const destWidth = sheet.getMaxColumns(); 
  const pasteArray = [];

  const skipLetters = EXCLUSIONS[tabName] || [];
  const skipIndices = skipLetters.map(letter => colToIdx(letter));
  const mapping = FETCH_MAPS[tabName] || {};

  for (let i = 0; i < rowsNeeded; i++) {
    let sourceRow = processedData[i];
    let destRow = new Array(destWidth).fill(""); 
    const valA_source = String(sourceRow[0] || "").toLowerCase();
    
    if (valA_source.includes("total") || valA_source.includes("sub-total")) {
      destRow[0] = sourceRow[0]; 
    } else {
      // Step 1: Standard Copy (obeying exclusions)
      for (let c = 0; c < sourceRow.length; c++) {
        if (skipIndices.includes(c)) continue;
        if (c < destWidth) destRow[c] = sourceRow[c];
      }

      // Step 2: Custom Mapping (Fetching source L to AF, etc)
      Object.keys(mapping).forEach(targetCol => {
        let sourceCol = mapping[targetCol];
        let tIdx = colToIdx(targetCol);
        let sIdx = colToIdx(sourceCol);
        if (sourceRow[sIdx] !== undefined) destRow[tIdx] = sourceRow[sIdx];
      });
    }
    pasteArray.push(destRow);
  }

  const currentLastRow = sheet.getLastRow();
  const clearHeight = Math.max(currentLastRow - CONFIG.dataStartRow + 1, 1);
  sheet.getRange(CONFIG.dataStartRow, 1, clearHeight, destWidth).clearContent();
  
  if (pasteArray.length > 0) {
    sheet.getRange(CONFIG.dataStartRow, 1, pasteArray.length, pasteArray[0].length).setValues(pasteArray);
  }
  SpreadsheetApp.getActive().toast(`Data fetch complete for ${tabName}. Values mapped as requested.`, "Success");
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

  if (!valL5 || !valL6) {
    SpreadsheetApp.getUi().alert("❌ Action Blocked: L5 and L6 references are required.");
    return;
  }

  // N10 Check only for LPG
  if (tabName === "LPG") {
    const valN10 = sheet.getRange("N10").getValue();
    if (!valN10) {
      SpreadsheetApp.getUi().alert("❌ Action Blocked: N10 is required for LPG formulas.");
      return;
    }
  }

  if (valL5 <= valL6) {
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
    P: (r) => `=IFERROR(IF(O${r}="fix rate", "Put/input", ROUND(ROUND(O${r}, 2) * ROUND(L${r}, 3), 2)),"-")`,
    S: (r) => `=IF(NOT(ISNUMBER($U$10)),"-",$U$10)`,
    T: (r) => `=IFERROR(S${r}*L${r},"-")`,
    U: (r) => `=IFERROR(L${r}+T${r},"-")`,
    V: (r) =>`=IF(OR(J${r}="Fix Rate", O${r}="Fix Rate"), "-", IF(AND(ISNUMBER(P${r}), ISNUMBER(S${r})), P${r}*S${r}, "-"))`,
    W: (r) => `=IFERROR(V${r}+P${r},"-")`,
    X: (r) => `=IFERROR(IF(J${r}="fix rate", ROUND(P${r}*1.12, 3), ROUND(W${r}*1.12, 3)), "-")`,
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

  // Find stopRow
  let stopRow = lastRow;
  for (let i = CONFIG.dataStartRow - 1; i < lastRow; i++) {
    if (String(fullDataA[i][0]).toLowerCase().trim() === "total") {
      stopRow = i + 1;
      break; 
    }
  }

  // --- STEP 1: APPLY ROW-LEVEL LOGIC (With Blank Row handling) ---
  for (let i = CONFIG.dataStartRow - 1; i < stopRow; i++) {
    const r = i + 1;
    const labelA = String(fullDataA[i][0]).toLowerCase().trim();
    const valE = String(fullDataE[i][0]).trim();

    // Skip Totals and Subtotals from clearing logic
    if (labelA.includes("total")) continue;

    // IF COL E IS BLANK: Wipe the entire row content and skip formula application
    if (valE === "") {
      sheet.getRange(r, 1, 1, sheet.getLastColumn()).clearContent();
      continue;
    }

    const rowData = sheet.getRange(r, 1, 1, 35).getValues()[0];
    const valO = String(rowData[14] || "").trim();
    const valP = String(rowData[15] || "").trim();
    const valZ = String(rowData[25] || "").trim();
    const valK = String(rowData[10] || "").toLowerCase();

    let targetCols = Object.keys(activeMap);
    
    // Safety check for manual entries
    if (valO !== "") targetCols = targetCols.filter(c => c !== "O"); 
    if (valP !== "") targetCols = targetCols.filter(c => c !== "P"); 
    if (valZ !== "") targetCols = targetCols.filter(c => c !== "Z");
    if (valK.includes("theoretical")) targetCols = targetCols.filter(c => c !== "L");

    targetCols.forEach(colKey => {
      sheet.getRange(`${colKey}${r}`).setFormula(activeMap[colKey](r));
    });
  }

  // --- STEP 2: CALCULATE SUMS ---
  const sumCols = ["L", "N", "P", "Q", "AA", "AB", "AF", "AG", "AI", "AJ"];
  let sectionStartRow = CONFIG.dataStartRow;
  let subTotalRowsFound = [];

  for (let i = CONFIG.dataStartRow - 1; i < stopRow; i++) {
    const rowLabel = String(fullDataA[i][0]).toLowerCase().trim();
    const r = i + 1;

    if (rowLabel.includes("sub total") || rowLabel.includes("sub-total")) {
      const rangeEnd = r - 1;
      sumCols.forEach(col => sheet.getRange(`${col}${r}`).setFormula(`=SUM(${col}${sectionStartRow}:${col}${rangeEnd})`));
      subTotalRowsFound.push(r);
      sectionStartRow = r + 1;
    }

    if (rowLabel === "total" && subTotalRowsFound.length > 0) {
      sumCols.forEach(col => {
        let refs = subTotalRowsFound.map(subR => `${col}${subR}`).join(",");
        sheet.getRange(`${col}${r}`).setFormula(`=SUM(${refs})`);
      });
    }
  }

  // --- STEP 3: CLEANUP NUMERIC VALUES BELOW TOTAL ---
  const lastSheetRow = sheet.getLastRow();
  if (lastSheetRow > stopRow) {
    const footerRange = sheet.getRange(stopRow + 1, 1, lastSheetRow - stopRow, sheet.getLastColumn());
    const footerValues = footerRange.getValues();
    const cleanedFooter = footerValues.map(row => row.map(cell => (typeof cell === 'number' && cell !== "") ? "" : cell));
    footerRange.setValues(cleanedFooter);
  }

  // Final Formatting
  SpreadsheetApp.flush();
  ["P", "Q", "J", "L", "AF", "AI", "AJ", "AG"].forEach(c => {
    sheet.getRange(`${c}${CONFIG.dataStartRow}:${c}${stopRow}`).setNumberFormat("#,##0.00");
  });
  
  SpreadsheetApp.getActive().toast(`Formula logic and row cleaning complete for ${tabName}.`, "Success");
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
6. UPDATED SCAN TAB (Checklist Integration)
================================= */
function scanTab(tabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.dataStartRow) return;

  // 1. SETUP LOGGING
  let logSheet = ss.getSheetByName("IssueLogs");
  if (!logSheet) {
    logSheet = ss.insertSheet("IssueLogs");
    logSheet.appendRow(["Timestamp", "Tab", "Cell", "Column Label", "Error Message", "Remarks"]);
    logSheet.getRange(1, 1, 1, 6).setFontWeight("bold");
  } else {
    const logLastRow = logSheet.getLastRow();
    if (logLastRow > 1) logSheet.getRange(2, 1, logLastRow, 6).clearContent();
  }

  // 2. FETCH DATA & CONFIG VALUES
  const dataRange = sheet.getRange(CONFIG.dataStartRow, 1, lastRow - CONFIG.dataStartRow + 1, 37); // Cols A to AK
  const dataValues = dataRange.getValues();
  const headers = sheet.getRange(CONFIG.headerRow, 1, 1, 37).getValues()[0];
  
  const valL5 = sheet.getRange("L5").getValue();
  const valL6 = sheet.getRange("L6").getValue();
  const timestamp = new Date();
  const logs = [];

  // Reset colors for the range
  dataRange.setBackground(null);

  // 3. GLOBAL CHECKS (Item 1 & 2)
  if (!(valL5 > 0)) logs.push([timestamp, tabName, "L5", "Rate Reference", "Must be a number > 0", "Critical Configuration"]);
  if (!(valL6 > 0)) logs.push([timestamp, tabName, "L6", "Base Reference", "Must be a number > 0", "Critical Configuration"]);

  // 4. ROW-BY-ROW SCANNING
  for (let i = 0; i < dataValues.length; i++) {
    const row = dataValues[i];
    const rowNum = CONFIG.dataStartRow + i;
    const labelA = String(row[0]).toLowerCase();
    
    if (labelA.includes("total") || labelA === "") continue;

    const rowIssues = [];
    const getVal = (colLet) => row[colToIdx(colLet)];
    const logError = (colLet, msg) => {
      rowIssues.push([timestamp, tabName, `${colLet}${rowNum}`, headers[colToIdx(colLet)], msg, "Condition Mismatch"]);
      sheet.getRange(`${colLet}${rowNum}`);
    };

    // Helper logical variables
    const hasE = String(getVal("E")).trim() !== "";
    const L_val = getVal("L");
    const O_val = getVal("O");
    const Z_val = getVal("Z");

    // -- VALIDATION RULES --

    // Item 3, 4, 5, 6 (J, K, L Not blank if E exists)
    ["J", "K", "L"].forEach(col => {
      if (hasE && String(getVal(col)).trim() === "") logError(col, "Should not be blank if E has entry");
    });

    // Item 7 & 8 (LPG only: M, N > 0 if L is number)
    if (tabName === "LPG") {
      ["M", "N"].forEach(col => {
        let v = getVal(col);
        if (typeof L_val === 'number' && (!(typeof v === 'number' && v > 0))) logError(col, "Should be a number > 0 if L is numeric");
      });
    }

    // Item 9 & 10 (Column O Logic)
    const O_clean = String(O_val).toLowerCase().trim();
    const isTheoreticalOrFix = (O_clean === "fix rate" || O_clean === "theoretical");
    
    if (typeof O_val === 'number') {
      if (O_val <= 0) logError("O", "Number must be > 0");
    } else if (!isTheoreticalOrFix) {
      logError("O", 'Must be >0, "fix rate" or "theoretical"');
    }
    
    if (String(L_val).trim() !== "-") {
       if (O_val != valL5) logError("O", `Value must match L5 (${valL5})`);
    } else {
       if (!isTheoreticalOrFix) logError("O", 'L is "-", O must be "fix rate" or "theoretical"');
    }

    // Item 11 (P > 0)
    if (!(getVal("P") > 0)) logError("P", "Should be a number > 0");

    // Item 12 (Elec, LPG only: Q > 0)
    if (tabName !== "Water") {
      if (!(getVal("Q") > 0)) logError("Q", "Should be a number > 0");
    }

    // Item 14-18 (Water only: S, T, U, V, W not empty)
    if (tabName === "Water") {
      ["S", "T", "U", "V", "W"].forEach(col => {
        if (String(getVal(col)).trim() === "") logError(col, "Field required (c/o fx)");
      });
      if (!(getVal("X") > 0)) logError("X", "Should be a number > 0"); // Item 19
    }

    // Item 21 & 22 (Column Z Logic)
    if (String(getVal("Z")).trim() === "") logError("Z", "Field required (c/o fx)");
    if (String(L_val).trim() !== "-") {
      if (Z_val != valL6) logError("Z", `Value must match L6 (${valL6})`);
    }

    // Item 23, 24, 25, 29, 32 (General Not Empty Checks)
    ["AA", "AB", "AC", "AG", "AJ"].forEach(col => {
      if (String(getVal(col)).trim() === "") logError(col, "Should not be empty (c/o fx)");
    });

    // Item 28 & 31 (AF, AI not empty if E has entry)
    ["AF", "AI"].forEach(col => {
      if (hasE && String(getVal(col)).trim() === "") logError(col, "Should not be empty if E has entry");
    });

    // Item 30 & 33 (Variances ±30%)
    ["AH", "AK"].forEach(col => {
      let v = getVal(col);
      if (typeof v === 'number' && (v > 0.30 || v < -0.30)) logError(col, "Variance outside ±30% limit");
    });

    if (rowIssues.length > 0) logs.push(...rowIssues);
  }

  // 5. OUTPUT RESULTS
  if (logs.length > 0) {
    logSheet.getRange(2, 1, logs.length, 6).setValues(logs);
    SpreadsheetApp.getUi().alert(`Scan Complete: ${logs.length} checklist issues identified. Cells marked in Red.`);
  } else {
    SpreadsheetApp.getUi().alert("Scan Complete: No checklist violations found.");
  }
}
function confirmFetchOverwrite(tabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  const ui = SpreadsheetApp.getUi();
  if(!sheet) return false;
  if (sheet.getLastRow() >= CONFIG.dataStartRow) {
    const res = ui.alert('Confirm Overwrite', `Data already exists in "${tabName}". Overwrite?`, ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) return false;
  }
  return true;
}

function recordActivePBTT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const [p, l, s, e] = ["E4","E5","E7","E8"].map(x => sheet.getRange(x).getValue());
  if (!p || !l || !s || !e) { SpreadsheetApp.getUi().alert("Missing Header Info (E4, E5, E7, E8)"); return; }
  try {
    const db = SpreadsheetApp.openById(PBTT_DB_ID);
    let dSh = db.getSheetByName("ACTIVE PBTT") || db.insertSheet("ACTIVE PBTT");
    dSh.appendRow([new Date(), p, l, s, e, ss.getUrl(), Session.getActiveUser().getEmail()]);
    sheet.getRangeList(["E7", "E8"]);
    SpreadsheetApp.getUi().alert("Recorded!");
  } catch(x) { SpreadsheetApp.getUi().alert("Err: "+x.message); }
}
