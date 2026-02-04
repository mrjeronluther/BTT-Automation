/* =================================
CONFIGURATION
================================= */
const SOURCE_DB_URL = "https://docs.google.com/spreadsheets/d/1JNYCjZfGYyVTxYkrws4D7SnrO-hIjDJ61sFqj0WBWEE/edit";
const PBTT_DB_ID    = "16Oai_3c4H_E2wgC-CUkSk1Eez90_KdtlaqHnHJFclBQ";

const CONFIG = {
  headerRow: 12,
  dataStartRow: 13,
  minCols: 34
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
2. FETCH DATA
================================= */
function fetchDataOnly(tabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return;

  const sourceLink = sheet.getRange("A5").getValue();
  if (!sourceLink) { SpreadsheetApp.getUi().alert("Paste SOURCE LINK in A5."); return; }

  let sourceSS;
  try { sourceSS = SpreadsheetApp.openByUrl(sourceLink); } 
  catch (e) { SpreadsheetApp.getUi().alert("Cannot open source link."); return; }
  
  const sourceSheet = sourceSS.getSheetByName(tabName);
  if (!sourceSheet) { SpreadsheetApp.getUi().alert(`Tab "${tabName}" not found in source.`); return; }

  const lastSourceRow = sourceSheet.getLastRow();
  const lastSourceCol = sourceSheet.getLastColumn();
  if (lastSourceRow < CONFIG.dataStartRow) return;

  let rawData = sourceSheet.getRange(CONFIG.dataStartRow, 1, lastSourceRow - CONFIG.dataStartRow + 1, lastSourceCol).getValues();

  let rowLimit = rawData.length;
  for (let i = 0; i < rawData.length; i++) {
    if (String(rawData[i][0]).toUpperCase().trim() === "TOTAL") {
      rowLimit = i + 1; 
      break;
    }
  }
  const processedData = rawData.slice(0, rowLimit);
  const rowsNeeded = processedData.length;
  const destWidth = sheet.getMaxColumns(); 
  const pasteArray = [];

  for (let i = 0; i < rowsNeeded; i++) {
    let sourceRow = processedData[i];
    let destRow = new Array(destWidth).fill(""); // Initializes entire row as blank

    for (let c = 0; c < sourceRow.length; c++) {
      /**
       * EXCLUSION LOGIC
       * We skip these columns so we don't overwrite formulas or manual inputs in destination:
       * 10 (K), 11 (L), 14 (O), 15 (P), 25 (Z)
       */
      if (c === 10 || c === 11 || c === 14 || c === 15 || c === 25) {
        continue;
      }
      
      if (c < destWidth) destRow[c] = sourceRow[c];
    }

    // SHARED MAPPING
    // 1. Move Source K (Index 10) to Target J (Index 9) for previous reading ref
    if (sourceRow.length > 10) destRow[9] = sourceRow[10];

    // 2. Map historical usage/billing for Elec & Water only (if needed)
    // For LPG, Target K remains blank (index 10 was skipped in the loop and not mapped here)
    if (tabName !== "LPG") {
       if (sourceRow.length > 11) destRow[30] = sourceRow[11]; // Source L -> Dest AE
       if (sourceRow.length > 15) destRow[33] = sourceRow[15]; // Source P -> Dest AH
    } else {
       // LPG specific mapping for last month consumption (Optional, remove if LPG doesn't use variance)
       if (sourceRow.length > 11) destRow[30] = sourceRow[11]; 
    }

    pasteArray.push(destRow);
  }

  const clearHeight = Math.max(sheet.getLastRow() - CONFIG.dataStartRow + 1, rowsNeeded);
  if(clearHeight > 0) {
    sheet.getRange(CONFIG.dataStartRow, 1, clearHeight, destWidth).clearContent();
  }
  
  if (pasteArray.length > 0) {
    sheet.getRange(CONFIG.dataStartRow, 1, pasteArray.length, pasteArray[0].length).setValues(pasteArray);
  }

  SpreadsheetApp.getActive().toast(`Fetched ${tabName} data. Target K is now blank.`, "Success");
}

/* =================================
3. RUN FORMULAS (MASTER)
================================= */
function applyFormulasToSheet(tabName) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return;

  const valL2 = sheet.getRange("L5").getValue(); // Base Rate ref
  const valL3 = sheet.getRange("L6").getValue(); // Secondary Ref
  if (valL2 === "" || valL3 === "") {
    SpreadsheetApp.getUi().alert("❌ Action Blocked: L5 and L6 are required.");
    return;
  }

  // Elec Map
  const formulaMapElec = {
    L: r => `=IFERROR((K${r}-J${r})*I${r},0)`,
    O: r => `=$L$5`, 
    P: r => `=IFERROR(IF(O${r}="fix rate","Put/input",ROUND(L${r}*O${r},2)),"0")`,
    Q: r => `=IFERROR(ROUND(P${r}*1.12, 3), "-")`,
    G: r => `=IFERROR(ROUND(P${r}*1.12,2),"-")`,
    Z: r => `=$L$6`,
    AF: r => `=IFERROR(L${r}-AE${r},"-")`,
    AG: r => `=IFERROR(AF${r}/AE${r},"-")`,
    AI: r => `=IFERROR(P${r}-AH${r},"-")`,
    AJ: r => `=IFERROR(AI${r}/AH${r},"-")`,
    AA: r => `=IFERROR(L${r}*Z${r},"-")`,
    AB: r => `=IFERROR(P${r}-AA${r},"-")`,
    AC: r => `=IFERROR(AB${r}/Q${r},"-")`
  };

  // Water Map
  const formulaMapWater = {
    L: r => `=K${r}-J${r}`,
    O: r => `=IF(NOT(ISNUMBER($L$5)),"0",$L$5)`,
    P: r => `=IFERROR(IF(O${r}="fix rate", "Put/input",ROUND(ROUND(O${r}, 2) * ROUND(L${r}, 3), 2)),"-")`,
    S: r => `=IF(NOT(ISNUMBER($U$10)),"0",$U$10)`, 
    T: r => `=IFERROR(S${r}*L${r},"0")`,
    U: r => `=IFERROR(L${r}+T${r},"0")`,
    V: r => `=IFERROR(P${r}*S${r},"0")`,
    W: r => `=IFERROR(V${r}+P${r},"0")`,
    X: r => `=IFERROR(ROUND(W${r}*1.12, 3), "-")`,
    Z: r => `=IF(NOT(ISNUMBER($L$6)),".",$L$6)`,
    AA: r => `=IFERROR(L${r}*Z${r},"-")`,
    AB: r => `=IFERROR(W${r}-AA${r},"-")`,
    AC: r => `=IFERROR(AB${r}/Q${r},"-")`,
    AF: r => `=IFERROR(L${r}-AE${r},"-")`,
    AG: r => `=IFERROR(AF${r}/AE${r},"-")`,
    AI: r => `=IFERROR(W${r}-AH${r},"-")`,
    AJ: r => `=IFERROR(AI${r}/AH${r},"-")`
  };

  // LPG Map
  const formulaMapLPG = {
    L: r => `=K${r}-J${r}`,
    M: r => `=$N$10*L${r}`,
    N: r => `=IFERROR(L${r}*M${r}, "0")`,
    O: r => `=IF(NOT(ISNUMBER($L$5)),"0",$L$5)`,
    P: r => `=IFERROR(IF(O${r}="fix rate","Put/input",N${r}*O${r}),"-")`,
    Q: r => `=IFERROR(ROUND(P${r}*1.12, 3), "-")`,
    Z: r => `=IF(NOT(ISNUMBER($L$6)),".",$L$6)`,
    AA: r => `=IFERROR(N${r}*Z${r},"-")`,
    AB: r => `=IFERROR(P${r}-AA${r},"-")`,
    AC: r => `=IFERROR(AB${r}/Q${r},"-")`,
    AF: r => `=IFERROR(N${r}-AE${r},"-")`,
    AG: r => `=IFERROR(AF${r}/AE${r},"-")`,
    AI: r => `=IFERROR(P${r}-AH${r},"-")`,
    AJ: r => `=IFERROR(AI${r}/AH${r},"-")`
  };

  // Select the correct map
  let activeMap;
  if (tabName === "Water") activeMap = formulaMapWater;
  else if (tabName === "LPG") activeMap = formulaMapLPG;
  else activeMap = formulaMapElec;

  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(CONFIG.dataStartRow, 1, lastRow - CONFIG.dataStartRow + 1, 1).getValues();
  let loopLimit = dataRange.length;
  for(let i = 0; i < dataRange.length; i++){
    if(String(dataRange[i][0]).trim().toUpperCase() === "TOTAL") { loopLimit = i; break; }
  }

  // Get data up to Col Z to check for existing manual values
  const checkValues = sheet.getRange(CONFIG.dataStartRow, 1, loopLimit, 26).getValues();

  for (let i = 0; i < loopLimit; i++) {
    const r = CONFIG.dataStartRow + i;
    
    const valJ = String(checkValues[i][9] || "");
    const valK = String(checkValues[i][10] || "").toLowerCase();
    const valN = String(checkValues[i][13] || "").toLowerCase();
    const valO = String(checkValues[i][14] || "").toLowerCase();
    const valP = checkValues[i][15]; 
    const valZ = checkValues[i][25]; 

    const valJK = valJ + " " + valK;
    let targetCols = [];

    // --- CONDITION LOGIC ---
    if (valO.includes("fix rate")) { 
      // Water/LPG run P only, Elec runs P&Q
      targetCols = (tabName === "Elec") ? ["P", "Q", "Z"] : ["P", "Z"]; 
    } 
    else if (valJK.includes("theoretical")) { 
       // Water/LPG run P only, Elec runs P&Q
      targetCols = (tabName === "Elec") ? ["O", "Z", "P", "Q"] : ["O", "Z", "P"]; 
    } 
    else if (valN.includes("special rate") || valO.includes("special rate")) { 
      targetCols = (tabName === "Elec") ? ["L", "O", "P", "Q", "Z"] : ["L", "O", "P", "Z"]; 
    } 
    else { 
      targetCols = Object.keys(activeMap); 
    }

    // MANDATORY CALCS (variance, subtotals etc)
    const alwaysRun = ["M", "N", "AA", "AB", "AC", "S", "T", "U", "V", "W", "X"];
    alwaysRun.forEach(col => {
      if (activeMap[col] && !targetCols.includes(col)) targetCols.push(col);
    });

    // --- OVERWRITE PROTECTION ---

    // 1. Protection for O and Z (Never overwrite if they have data)
    if (checkValues[i][14] !== "") targetCols = targetCols.filter(c => c !== "O"); // check original O
    if (checkValues[i][25] !== "") targetCols = targetCols.filter(c => c !== "Z"); // check original Z

    // 2. Protection for P if "Fix Rate" exists and P is not empty
    if (valO.includes("fix rate") && (valP !== "" && valP !== null)) {
      targetCols = targetCols.filter(c => c !== "P");
    }

    // 3. General column safety (e.g. theoretical row shouldn't calc usage in L)
    if (valK.includes("theoretical")) targetCols = targetCols.filter(c => c !== "L");

    // APPLY FORMULAS
    targetCols.forEach(colKey => {
      if (activeMap[colKey]) sheet.getRange(`${colKey}${r}`).setFormula(activeMap[colKey](r));
    });
  }

  // Formatting variances
  sheet.getRange(CONFIG.dataStartRow, 33, loopLimit).setNumberFormat("0.00%"); 
  sheet.getRange(CONFIG.dataStartRow, 36, loopLimit).setNumberFormat("0.00%"); 
  SpreadsheetApp.flush(); 
}
/* =================================
4. UTILS
================================= */
function logToIssueTab(sourceTab, rowNum, colName, value) {
  const ss = SpreadsheetApp.getActive();
  let logSheet = ss.getSheetByName("IssueLogs") || ss.insertSheet("IssueLogs");
  if (logSheet.getLastRow() === 0) logSheet.appendRow(["Timestamp", "Source Tab", "Row", "Col", "Value", "Issue"]);
  const formattedVal = (typeof value === 'number') ? (value * 100).toFixed(2) + "%" : value;
  logSheet.appendRow([new Date(), sourceTab, rowNum, colName, formattedVal, value > 0.3 ? "Exceeded +30%" : "Below -30%"]);
}

function clearTabData(tabName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
  if (sheet && sheet.getLastRow() >= CONFIG.dataStartRow) {
    sheet.getRange(CONFIG.dataStartRow, 1, sheet.getLastRow() - CONFIG.dataStartRow + 1, sheet.getLastColumn()).clearContent();
  }
}

function scanTab(tabName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if(lastRow >= CONFIG.dataStartRow) {
    sheet.getRange(CONFIG.dataStartRow, 1, lastRow, sheet.getLastColumn()).setBackground(null);
    [5, 11, 31, 34].forEach(c => {
       const vals = sheet.getRange(CONFIG.dataStartRow, c, lastRow-CONFIG.dataStartRow+1).getValues();
       for(let i=0; i<vals.length; i++) if(!vals[i][0]) sheet.getRange(CONFIG.dataStartRow+i, c).setBackground("#fff176");
    });
  }
}

/* =================================
5. SUBMIT PBTT
================================= */
function recordActivePBTT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const [p, l, s, e] = ["E1","E2","E4","E5"].map(x => sheet.getRange(x).getValue());
  if (!p || !l || !s || !e) return;
  try {
    const db = SpreadsheetApp.openById(PBTT_DB_ID);
    let dSh = db.getSheetByName("ACTIVE PBTT") || db.insertSheet("ACTIVE PBTT");
    dSh.appendRow([new Date(), p, l, s, e, ss.getUrl(), Session.getActiveUser().getEmail()]);
    sheet.getRangeList(["E1","E2","E4","E5","E7"]).clearContent();
    SpreadsheetApp.getUi().alert("Recorded!");
  } catch(x) { SpreadsheetApp.getUi().alert("Err: "+x.message); }
}

/* =================================
6. TRIGGERS
================================= */
function fetchElec() { fetchDataOnly("Elec"); }
function runFormulaElec() { applyFormulasToSheet("Elec"); }
function clearElec() { clearTabData("Elec"); }
function scanElecTab() { scanTab("Elec"); }

function fetchWater() { fetchDataOnly("Water"); }
function runFormulaWater() { applyFormulasToSheet("Water"); }
function clearWater() { clearTabData("Water"); }
function scanWaterTab() { scanTab("Water"); }

function fetchLPG() { fetchDataOnly("LPG"); }
function runFormulaLPG() { applyFormulasToSheet("LPG"); }
function clearLPG() { clearTabData("LPG"); }
function scanLPGTab() { scanTab("LPG"); }
