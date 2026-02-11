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
        .addItem("2. Run Formulas", "runFormulaElec"))
        //.addItem("Clear Tab", "clearElec")
    .addSubMenu(ui.createMenu("💧 Water")
        .addItem("1. Fetch Data", "fetchWater")
        .addItem("2. Run Formulas", "runFormulaWater"))
        //.addItem("Clear Tab", "clearWater")
    .addSubMenu(ui.createMenu("🔥 LPG")
        .addItem("1. Fetch Data", "fetchLPG")
        .addItem("2. Run Formulas", "runFormulaLPG"))
        //.addItem("Clear Tab", "clearLPG")
    .addSeparator()
    .addItem("🔍 Scan All Tabs (Elec, Water, LPG)", "scanAllTabs") // ONE BUTTON TO RULE THEM ALL
    .addItem("📤 Submit Active PBTT", "recordActivePBTT")
    .addToUi();
}

/* =================================
   2. SCAN WRAPPER & LOG CLEANING
================================= */
function scanAllTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabsToScan = ["Elec", "Water", "LPG"];
  const logSheetNames = ["Basic Anomaly", "KA Anomaly"]; //Basic Anomaly

  // 1. CLEAR BOTH LOG SHEETS ONCE AT THE START
  logSheetNames.forEach(name => {
    let s = ss.getSheetByName(name) || ss.insertSheet(name);
    const lastR = s.getLastRow();
    if (lastR > 1) s.getRange(2, 1, lastR - 1, 6).clearContent();
    if (lastR === 0) s.appendRow(["Timestamp", "Tab", "Cell", "Column Label", "Error Message", "Remarks"]);
  });

  // 2. SCAN EACH TAB ONE BY ONE (APPENDING TO LOGS)
  tabsToScan.forEach(tabName => {
    scanTab(tabName, false); // 'false' prevents log clearing inside the loop
  });

  SpreadsheetApp.getUi().alert("✅ Full Scan Complete. Please check the 'Basic Anomaly' and 'KA Anomaly' sheets.");
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

// Check if L5 and L6 are numbers and not empty
if (valL5 === "" || valL6 === "" || isNaN(valL5) || isNaN(valL6)) {
  SpreadsheetApp.getUi().alert("❌ Action Blocked: L5 and L6 must contain numeric values.");
  return;
}

// N10 Check only for LPG
if (tabName === "LPG") {
  const valN10 = sheet.getRange("N10").getValue();
  if (valN10 === "" || isNaN(valN10)) {
    SpreadsheetApp.getUi().alert("❌ Action Blocked: N10 must contain a numeric value for LPG formulas.");
    return;
  }
}

// U10 Check only for Water
if (tabName === "Water") {
  const valU10 = sheet.getRange("U10").getValue();
  if (valU10 === "" || isNaN(valU10)) {
    SpreadsheetApp.getUi().alert("❌ Action Blocked: U10 must contain a numeric value for Water formulas.");
    return;
  }
}

// Logical comparison
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

  const standardLogSheet = setupLogSheet("Basic Anomaly");
  const kaLogSheet = setupLogSheet("KA Anomaly");

  const dataRange = sheet.getRange(CONFIG.dataStartRow, 1, lastRow - CONFIG.dataStartRow + 1, 37);
  const dataValues = dataRange.getValues();
  const headers = sheet.getRange(CONFIG.headerRow, 1, 1, 37).getValues()[0];
  const valL5 = sheet.getRange("L5").getValue();
  const valL6 = sheet.getRange("L6").getValue();
  const rawE4 = sheet.getRange("E4").getValue();
  
  const kaRefMap = getKAData(); // This is the updated Col E Logic function
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
    const labelA = String(row[0]).toLowerCase().trim();
    if (labelA.includes("total") || labelA === "") continue;

    // KA VALIDATION
    if (kaRefMap) {
      const valE = String(row[colToIdx("E")] || "").trim();
      const valF = String(row[colToIdx("F")] || "").trim().toUpperCase();
      const valG = String(row[colToIdx("G")] || "").trim().toUpperCase();
      const hasKA = (valF === "KA" || valG === "KA");

      const matchedKey = findReferenceKey(valE, kaRefMap);
      const validCategories = matchedKey ? kaRefMap[matchedKey] : [];
      const headerE4 = superClean(rawE4); 
      let isMatch = false;

      if (validCategories.length > 0) {
        for (let k = 0; k < validCategories.length; k++) {
          let keyword = superClean(validCategories[k]);
          if (keyword !== "" && (headerE4.includes(keyword) || keyword.includes(headerE4))) {
            isMatch = true;
            break;
          }
        }
      }

      if (isMatch) {
        if (!hasKA) logHelper(row, rowNum, "F", 'user need to put "KA"', `Match: [${valE}]`, kaLogs);
      } else {
        if (hasKA) logHelper(row, rowNum, "F", 'user need to remove "KA"', `No match for [${valE}] in E4`, kaLogs);
      }
    }

    // COMMON CHECKLIST
    runCommonChecklist(row, rowNum, (r, c, m, res, arr) => logHelper(row, r, c, m, res, arr), valL5, valL6);

    // TAB SPECIFIC CHECKS
    switch(tabName) {
      case "Elec":
        if (!(typeof row[colToIdx("Q")] === 'number' && row[colToIdx("Q")] > 0)) logHelper(row, rowNum, "Q", "Should be a number >0");
        break;
      case "Water":
        ["S", "T", "U", "V", "W"].forEach(c => { if (String(row[colToIdx(c)]).trim() === "") logHelper(row, rowNum, c, "Missing calculation"); });
        if (!(typeof row[colToIdx("X")] === 'number' && row[colToIdx("X")] > 0)) logHelper(row, rowNum, "X", "Should be a number >0");
        break;
      case "LPG":
        const vL = row[colToIdx("L")];
        if (typeof vL === 'number') {
          if (!(typeof row[colToIdx("M")] === 'number' && row[colToIdx("M")] > 0)) logHelper(row, rowNum, "M", "Should be a number if L is number, should be >0");
          if (!(typeof row[colToIdx("N")] === 'number' && row[colToIdx("N")] > 0)) logHelper(row, rowNum, "N", "Should be a number if L is number, should be >0");
        }
        if (!(typeof row[colToIdx("Q")] === 'number' && row[colToIdx("Q")] > 0)) logHelper(row, rowNum, "Q", "Should be a number >0");
        break;
    }
  }

  // --- WRITE TO LOGS (APPEND MODE) ---
  if (issueLogs.length > 0) {
    const sIdx = standardLogSheet.getLastRow() + 1;
    standardLogSheet.getRange(sIdx, 1, issueLogs.length, 6).setValues(issueLogs);
  }
  if (kaLogs.length > 0) {
    const kIdx = kaLogSheet.getLastRow() + 1;
    kaLogSheet.getRange(kIdx, 1, kaLogs.length, 6).setValues(kaLogs);
  }



  const totalErrors = issueLogs.length + kaLogs.length;
  if (totalErrors > 0) {
    SpreadsheetApp.getUi().alert(`Scan Complete: ${issueLogs.length} standard issues and ${kaLogs.length} KA issues logged.`);
  } else {
    SpreadsheetApp.getUi().alert(`✅ Scan Success in ${tabName}.`);
  }
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
    if (!sheet) return null;
    const lastR = sheet.getLastRow();
    if (lastR < 2) return {};

    const rawData = sheet.getRange(2, 2, lastR - 1, 4).getValues(); // Cols B to E
    const propertyMap = {};

    rawData.forEach(r => {
      const mainProp = superClean(r[0]); // Col B
      const category = superClean(r[1]); // Col C
      const iterations = String(r[3] || ""); // Col E

      const addKey = (key) => {
        if (!key) return;
        if (!propertyMap[key]) propertyMap[key] = [];
        if (!propertyMap[key].includes(category)) propertyMap[key].push(category);
      };

      addKey(mainProp);
      if (iterations) {
        iterations.split(",").forEach(part => addKey(superClean(part)));
      }
    });
    return propertyMap;
  } catch (e) {
    console.error("KA Ref Error: " + e.message);
    return null;
  }
}


/* =================================
REFACTORED: THE "COMMON" CHECKLIST (ALL TABS)
================================= */
function runCommonChecklist(row, rNum, log, L5, L6) {
  const get = (let) => row[colToIdx(let)];
  const valE = String(get("E")).trim();
  const valL = get("L");
  const L_isHyphen = (String(valL).trim() === "-");
  
  // Logic: ONLY if Col E has entry
  if (valE !== "") {

    // J, K, L Conditions
    ["J", "K", "L"].forEach(c => {
      if (String(get(c)).trim() === "") log(rNum, c, "Should not be blank if E has entry");
    });

    // Column O Conditions
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

    // Column P Condition
    if (!(typeof get("P") === 'number' && get("P") > 0)) log(rNum, "P", "Should be a number >0");

    // Column Z Conditions
    const valZ = get("Z");
    if (valZ === "") log(rNum, "Z", "Should be a number >0, \"fix rate\" or \"theoretical\"");
    if (typeof valZ === 'number' && !L_isHyphen && valZ !== L6) {
      log(rNum, "Z", "Should equal to L6, or if L= \"-\" then, O= \"fix rate\" or O=\"theoretical\"");
    }

    // Standard Non-Empty Logic (AA, AB, AC, AG, AJ)
    ["AA", "AB", "AC", "AG", "AJ"].forEach(c => {
      if (String(get(c)).trim() === "") log(rNum, c, "Should not be empty (c/o fx)");
    });

    // Entries dependent on Col E (AF, AI)
    ["AF", "AI"].forEach(c => {
      if (String(get(c)).trim() === "") log(rNum, c, "Should not be empty if E has entry");
    });

    // Variances (AH, AK)
    ["AH", "AK"].forEach(c => {
      const v = get(c);
      if (typeof v === 'number') {
        if (v > 0.3 || v < -0.3) log(rNum, c, "Should not be >30%, should not be <-30%");
      }
    });
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
