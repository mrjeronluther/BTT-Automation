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
    "AF": "L",
    "AH": "P"
  }
};

// COLUMNS TO BE LEFT BLANK DURING FETCH (To be filled by Run Formula)
const EXCLUSIONS = {
  "Elec": ["K", "L", "O", "P", "Q", "Z", "AA", "AB", "AC", "AG", "AH", "AJ", "AK"],
  "Water": ["K", "L", "O", "P", "Q", "S","T", "U", "V", "W", "X", "Z", "AA", "AB", "AC", "AG", "AH", "AJ", "AK"],
  "LPG": ["K", "L", "O", "P", "Q", "Z", "AA", "AB", "AC", "AG", "AH", "AJ", "AK"]
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
4. RUN FORMULAS (UPDATED)
================================= */
function applyFormulasToSheet(tabName) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return;

  const valL5 = sheet.getRange("L5").getValue();
  if (valL5 === "" || valL5 === null) {
    SpreadsheetApp.getUi().alert("❌ Action Blocked: L5 (Rate) is required.");
    return;
  }

  // --- FORMULA MAPS ---
  const formulaMapElec = {
    L: r => `=IFERROR((K${r}-J${r})*I${r},"-")`,
    O: r => `=$L$5`, 
    P: r => `=IFERROR(IF(O${r}="fix rate","Put/input",ROUND(L${r}*O${r}, 2)),"-")`,
    Q: r => `=IFERROR(ROUND(P${r}*1.12, 2), "-")`,
    Z: r => `=$L$6`,
    AA: r => `=IFERROR(L${r}*Z${r},"-")`,
    AB: r => `=IFERROR(P${r}-AA${r},"-")`,
    AC: r => `=IFERROR((O${r}-Z${r})/Z${r},"-")`,
    AG: r => `=IFERROR(L${r}-AF${r},"-")`,
    AH: r => `=IFERROR(AG${r}/AF${r},"-")`,
    AJ: r => `=IFERROR(P${r}-AI${r},"-")`,          
    AK: r => `=IFERROR(AI${r}/AJ${r},"-")`          
  };

  const formulaMapWater = {

    L: r => `=IFERROR(K${r}-J${r}, "-")`,
    O: r => `=IF(NOT(ISNUMBER($L$5)),"-",$L$5)`,
    P: r => `=IFERROR(IF(O${r}="fix rate", "Put/input", ROUND(ROUND(O${r}, 2) * ROUND(L${r}, 3), 2)),"-")`,
    S: r => `=IF(NOT(ISNUMBER($U$10)),"-",$U$10)`, 
    T: r => `=IFERROR(S${r}*L${r},"-")`,
    U: r => `=IFERROR(L${r}+T${r},"-")`,
    
    V: r => `=IF(OR(J${r}="Fix Rate", O${r}="Fix Rate"), "-", IF(AND(ISNUMBER(P${r}), ISNUMBER(S${r})), P${r}*S${r}, "-"))`,

    W: r => `=IFERROR(V${r}+P${r},"-")`,
    X: r => `=IFERROR(IF(J${r}="fix rate", ROUND(P${r}*1.12, 3), ROUND(W${r}*1.12, 3)), "-")`,

    Z: r => `=IF(NOT(ISNUMBER($L$6)),"-",$L$6)`,
    AA: r => `=IFERROR(L${r}*Z${r},"-")`,
    AB: r => `=IFERROR(W${r}-AA${r},"-")`,
    AC: r => `=IFERROR((O${r}-Z${r})/Z${r},"-")`,
    AG: r => `=IFERROR(L${r}-AF${r},"-")`,
    AH: r => `=IFERROR(AG${r}/AF${r},"-")`,
    AJ: r => `=IFERROR(W${r}-AI${r},"-")`,
    AK: r => `=IFERROR(AI${r}/AJ${r},"-")`
  };

  const formulaMapLPG = {
    L: r => `=IFERROR(K${r}-J${r}, "-")`,
    M: r => `=if(not(isnumber($N$10)),".",$N$10)`,
    
    N: r => `=iferror(L${r}*M${r},"-")`,
    
    O: r => `=IF(NOT(ISNUMBER($L$5)),"-",$L$5)`,
    P: r => `=IFERROR(IF(O${r}="fix rate","Put/input", ROUND(N${r}*O${r}, 2)),"-")`,
    Q: r => `=IFERROR(ROUND(P${r}*1.12, 2), "-")`,
    Z: r => `=IF(NOT(ISNUMBER($L$6)),"-",$L$6)`,

    AA: r => `=IFERROR(N${r}*Z${r},"-")`,
    AB: r => `=IFERROR(P${r}-AA${r},"-")`,
    AC: r => `=IFERROR((O${r}-Z${r})/Z${r},"-")`, 

    AG: r => `=IFERROR(N${r}-AE${r},"-")`,
    AH: r => `=IFERROR(AG${r}/AF${r},"-")`,
    AJ: r => `=IFERROR(P${r}-AI${r},"-")`,
    AK: r => `=IFERROR(AI${r}/AJ${r},"-")`
  };

  let activeMap = (tabName === "Water") ? formulaMapWater : (tabName === "LPG" ? formulaMapLPG : formulaMapElec);
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.dataStartRow) return;

  const dataRange = sheet.getRange(CONFIG.dataStartRow, 1, lastRow - CONFIG.dataStartRow + 1, 1).getValues();
  
  let loopLimit = 0;
  for(let i = 0; i < dataRange.length; i++){
    if(String(dataRange[i][0]).trim().toUpperCase() === "TOTAL") { loopLimit = i; break; }
    loopLimit = i + 1;
  }

  const checkValues = sheet.getRange(CONFIG.dataStartRow, 1, loopLimit, 34).getValues();

  for (let i = 0; i < loopLimit; i++) {
    const r = CONFIG.dataStartRow + i;
    const valA = String(checkValues[i][0] || "").toLowerCase();
    if (valA.includes("total") || valA.includes("sub-total")) continue;

    const valK = String(checkValues[i][10] || "").toLowerCase(); 
    const valO = String(checkValues[i][14] || "").trim(); // Col O index is 14
    const valP = String(checkValues[i][15] || "").trim(); // Col O index is 15
    const valZ = String(checkValues[i][25] || "").trim(); // Col O index is 25
    const valJK = String(checkValues[i][9] || "") + " " + valK;

    let targetCols = Object.keys(activeMap);

    // --- NEW LOGIC: PRESERVE COL O IF IT HAS A VALUE ---
    if (valO !== "") { 
      targetCols = targetCols.filter(c => c !== "O"); 
    }
    if (valP !== "") { 
      targetCols = targetCols.filter(c => c !== "P"); 
    }

    if (valZ !== "") { 
      targetCols = targetCols.filter(c => c !== "Z"); 
    }
    
    
    // Theoretical exception for Col L
    if (valJK.includes("theoretical") || valK.includes("theoretical")) { 
      targetCols = targetCols.filter(c => c !== "L"); 
    }
    
    targetCols.forEach(colKey => {
      sheet.getRange(`${colKey}${r}`).setFormula(activeMap[colKey](r));
    });
  }
// =================================
  // FINAL "FORCED" FORMATTING
  // =================================
  if (loopLimit > 0) {
    // 1. Flush pending formulas so the formatting applies to final results
    SpreadsheetApp.flush();

    // 2. Identify specifically what format belongs where
    // Helper function to apply and force format
    const forceFormat = (col, format) => {
      let range = sheet.getRange(CONFIG.dataStartRow, colToIdx(col) + 1, loopLimit);
      range.clearFormat(); // Remove previous stuck formatting
      range.setNumberFormat(format);
    };

    // Columns P & Q: Decimal with Commas (Amount/Rate)
    forceFormat("P", "#,##0.00");
    forceFormat("Q", "#,##0.00");

    // Column J: History Reference (Amount/Consump)
    forceFormat("J", "#,##0.00");

    // Columns AC, AG, AJ: Percentages
    // Note: ensure your formulas in AG/AJ result in 0.05 for 5%
    forceFormat("AC", "0.00%"); // (Rate Variance %)
    forceFormat("AH", "0.00%"); // (Rate Variance %)
    forceFormat("AG", "#,##0.00"); // (Consumption Variance %)
    forceFormat("AJ", "#,##0.00"); // (Amount Variance %)
    forceFormat("AK", "0.00%"); // (Rate Variance %)
    

    // Column L: Consumption Total
    forceFormat("L", "#,##0.00");

    // Optional: Clean look for non-calculating rows (like those returning "-")
    // Re-flush to ensure the numbers "settle" into the grid
    SpreadsheetApp.flush();
  }

  SpreadsheetApp.flush(); 
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

function scanTab(tabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.dataStartRow) return;

  // 1. SETUP & CLEAR LOGGING SHEET (Row 2 onwards)
  let logSheet = ss.getSheetByName("IssueLogs");
  if (!logSheet) {
    logSheet = ss.insertSheet("IssueLogs");
    logSheet.appendRow(["Timestamp", "Tab", "Cell", "Column Label", "Cell entry issue", "Remarks"]);
    logSheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#f3f3f3");
  } else {
    // Clear everything from row 2 down to ensure fresh logs
    const logLastRow = logSheet.getLastRow();
    if (logLastRow > 1) {
      logSheet.getRange(2, 1, logLastRow, 6).clearContent();
    }
  }

  // 2. DEFINE TARGETS (AH is index 33, AK is index 36)
  const threshold = 0.30; 
  const targetColIndices = [33, 36]; // Indices for AH and AK
  const targetColLetters = { 33: "AH", 36: "AK" };

  // 3. RESET COLORS FIRST
  // This ensures colors are "gone only if they don't meet the condition"
  targetColIndices.forEach(idx => {
    sheet.getRange(CONFIG.dataStartRow, idx + 1, lastRow - CONFIG.dataStartRow + 1, 1).setBackground(null);
  });

  // Get data for values and headers
  const dataRange = sheet.getRange(CONFIG.dataStartRow, 1, lastRow - CONFIG.dataStartRow + 1, 37);
  const dataValues = dataRange.getValues();
  const headerValues = sheet.getRange(CONFIG.headerRow, 1, 1, 37).getValues()[0];

  const logs = [];
  const timestamp = new Date();

  // 4. SCAN DATA
  for (let i = 0; i < dataValues.length; i++) {
    const rowNum = CONFIG.dataStartRow + i;
    
    // Skip if column A contains "total"
    const rowLabel = String(dataValues[i][0]).toLowerCase();
    if (rowLabel.includes("total")) continue;

    targetColIndices.forEach(idx => {
      const val = dataValues[i][idx];
      const colLetter = targetColLetters[idx];
      
      // Condition: More than 30% or Below -30%
      if (typeof val === 'number' && (val > threshold || val < -threshold)) {
        
        // 1. Set Background to Red (stays until next scan fixes it)
        sheet.getRange(rowNum, idx + 1).setBackground("#f4cccc"); 

        // 2. Add to logs array
        logs.push([
          timestamp,
          tabName,
          `${colLetter}${rowNum}`,
          headerValues[idx],                       // Column Label
          (val * 100).toFixed(2) + "%",            // Cell entry issue (The % value)
          `Variance is outside ±30% range`         // Remarks
        ]);
      }
    });
  }

  // 5. WRITE NEW LOGS
  if (logs.length > 0) {
    logSheet.getRange(2, 1, logs.length, 6).setValues(logs);
    SpreadsheetApp.getUi().alert(`Scan Complete: ${logs.length} issues identified and logged.`);
  } else {
    SpreadsheetApp.getUi().alert("Scan Complete: No variances exceeding ±30% found.");
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
    sheet.getRangeList(["E7", "E8"]).clearContent();
    SpreadsheetApp.getUi().alert("Recorded!");
  } catch(x) { SpreadsheetApp.getUi().alert("Err: "+x.message); }
}
