/* =================================
CONFIGURATION & MAPPING
================================= */
const PBTT_DB_ID = "1hMMUd4ho50HP63dc2fRAo--iK-m7YotamkKtsDGT_Us";
const BACKUP_REGISTRY_ID = "10-ywOh509BNRMd0C-Mb8b5gibbu62D_K8U8cWYcV59U";
const BACKUP_FOLDER_ID = "1aokNFrCuVdLWs4AylG7LNekCtfQ5B1-p";
const CELL_LIMIT_MAX = 10000000; // Google's absolute limit (10M)
const CELL_ROTATION_LIMIT = 8000000; // Accurate threshold to trigger rotation (80%)

/**
 * Helper to check if a Property Name qualifies for validation bypass.
 * Bypasses: "CT - Capital Town" and "MG - Maple Grove"
 */
function isBypassedProperty(propertyName) {
  if (!propertyName) return false;
  const clean = superClean(propertyName);
  const bypassKeywords = ["capital town", "maple grove"];
  return bypassKeywords.some(keyword => clean.includes(keyword));
}

/**
 * Strict helper to match user Property Name to Master File Tab Name.
 * Enforces that if PROPERTY contains "MREIT", the TAB must also contain "MREIT".
 */
function isPropertyAndTabMatch(userProperty, tabName) {
  if (!userProperty || !tabName) return false;
  const cleanProp = superClean(userProperty);
  const cleanTab = superClean(tabName);
  if (!cleanProp || !cleanTab) return false;

  const propHasMreit = cleanProp.includes("mreit");
  const tabHasMreit = cleanTab.includes("mreit");

  // RULE: If Property contains MREIT, the Tab MUST also contain MREIT (and vice-versa)
  if (propHasMreit !== tabHasMreit) {
    return false;
  }

  // Handle MREIT specific abbreviations
  if (propHasMreit) {
    if ((cleanProp.includes("iloilo") || cleanProp.includes("ilo")) && 
        (cleanTab.includes("iloilo") || cleanTab.includes("ilo"))) return true;

    if ((cleanProp.includes("eastwood") || cleanProp.includes("ew")) && 
        (cleanTab.includes("eastwood") || cleanTab.includes("ew"))) return true;

    if ((cleanProp.includes("mckinley") || cleanProp.includes("mkh")) && 
        (cleanTab.includes("mckinley") || cleanTab.includes("mkh"))) return true;

    return cleanTab === cleanProp || cleanTab.includes(cleanProp) || cleanProp.includes(cleanTab);
  }

  // Standard non-MREIT tab matching
  return cleanTab === cleanProp || cleanTab.includes(cleanProp) || cleanProp.includes(cleanTab);
}

/**
 * Helper to map special Property Names to their designated Tab Names in the master file.
 * - MREIT EASTWOOD -> MREIT_EW
 * - MREIT MCKINLEY -> MREIT_MKH
 * - MREIT ILOILO   -> MREIT_ILO
 */
function getSpecialPropertyTabName(propertyName) {
  if (!propertyName) return null;
  const clean = superClean(propertyName);
  if (clean.includes("mreit")) {
    if (clean.includes("eastwood") || clean.includes("ew")) return "MREIT_EW";
    if (clean.includes("mckinley") || clean.includes("mkh")) return "MREIT_MKH";
    if (clean.includes("iloilo") || clean.includes("ilo")) return "MREIT_ILO";
  }
  return null;
}

/**
 * Helper to scan all three utility tabs to collect a set of base tenant names
 * that have been tagged with "_Affiliates" in any sheet.
 */
function getGlobalAffiliates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabs = ["Elec", "Water", "LPG"];
  const affiliates = new Set();

  tabs.forEach(tabName => {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.dataStartRow) return;

    const colEIndex = colToIdx("E") + 1;
    const data = sheet.getRange(CONFIG.dataStartRow, colEIndex, lastRow - CONFIG.dataStartRow + 1, 1).getValues();
    
    data.forEach(row => {
      const valE = String(row[0] || "").trim();
      if (valE.includes('_')) {
        const parts = valE.split('_');
        const suffix = parts.pop().trim().toLowerCase();
        if (suffix === "affiliates") {
          const baseName = parts.join('_').trim().toLowerCase();
          if (baseName) {
            affiliates.add(baseName);
          }
        }
      }
    });
  });
  return affiliates;
}

/**
 * Helper to verify if the user is attempting to run functions on the master template.
 * Blocks execution and prompts the user to make a copy.
 */
function isMasterFileBlocked() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterTemplateId = "1MS_GtdErDFaxI8ZLjIsTFTK4oWXrf2ULapKjwp_qCWw";
  
  if (ss.getId() === masterTemplateId) {
    SpreadsheetApp.getUi().alert(
      "⚠️ Action Blocked on Master File",
      "This is the master template file. You must make a copy of this spreadsheet (File > Make a copy) first before running any setups, formulas, or utilities.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return true; // Blocked
  }
  return false; // Not blocked
}

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
  "Elec": ["K", "L", "O", "P", "Q", "Z", "AA", "AB", "AC", "AG", "AH", "AJ", "AK","AD"],
  "Water": ["K", "L", "O", "P", "Q", "S", "T", "U", "V", "W", "X", "Z", "AA", "AB", "AC", "AG", "AH", "AJ", "AK","AD"],
  "LPG": ["K", "L", "O", "P", "M", "N", "Q", "Z", "AA", "AB", "AC", "AG", "AH", "AJ", "AK","AD"]
};

/* =================================
1. MENU
================================= */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Utility Manager")
    .addItem("🛠️ Setup", "INSTALL_SYSTEM")
    .addSubMenu(ui.createMenu("⚡ Electricity")
      .addItem("1. Fetch Data", "masterFetchElec")
      .addItem("2. Run Formulas", "runFormulaElec"))

    .addSubMenu(ui.createMenu("💧 Water")
      .addItem("1. Fetch Data", "masterFetchWater")
      .addItem("2. Run Formulas", "runFormulaWater"))

    .addSubMenu(ui.createMenu("🔥 LPG")
      .addItem("1. Fetch Data", "masterFetchLPG")
      .addItem("2. Run Formulas", "runFormulaLPG"))

    .addSeparator()
    .addItem("📤 Submit Active PBTT", "recordActivePBTT")

    .addToUi();
}

function masterFetchElec() {
  INITIALIZE_SYSTEM_BUTTON();
  fetchElec();
}
function masterFetchWater() {
  fetchWater();
}
function masterFetchLPG() {
  fetchLPG();
}

function scanAllTabs() {
  if (isMasterFileBlocked()) return false;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabsToScan = ["Elec", "Water", "LPG"];

  const extData = getExternalValidationData();
  if (!extData) {
    SpreadsheetApp.getUi().alert("🚫 Validation Error: Could not download external validation data. Scan halted and submission blocked.");
    return false;
  }

  // --- 1. REQUIREMENT CHECKER ---
  const requirements = {
    "Elec": { config: ["L5", "L6"], cols: ["L", "O", "P", "Q", "Z", "AA", "AB", "AC", "AG", "AH", "AJ", "AK"] },
    "Water": { config: ["L5", "L6", "U10"], cols: ["L", "O", "P", "S", "T", "U", "V", "W", "X", "Z", "AA", "AB", "AC", "AG", "AH", "AJ", "AK"] },
    "LPG": { config: ["L5", "L6", "N10"], cols: ["L", "M", "N", "O", "P", "Q", "Z", "AA", "AB", "AC", "AG", "AH", "AJ", "AK"] }
  };

  let stopErrors = [];

  tabsToScan.forEach(tabName => {
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) return;

    let missing = [];
    const req = requirements[tabName];

    const startRow = CONFIG.dataStartRow;
    const lastRow = Math.max(startRow, sheet.getLastRow());

    const colEAndADData = sheet.getRange(startRow, 5, lastRow - startRow + 1, 26).getValues();

    let configsChecked = false;

    colEAndADData.forEach((row, idx) => {
      let colEValue = row[0];
      let valAD = String(row[25] || "").trim();
      let actualRow = startRow + idx;

      if (colEValue !== "") {
        if (valAD.toUpperCase() === "MONITORING") return;

        if (!configsChecked) {
          req.config.forEach(c => {
            if (sheet.getRange(c).getValue() === "") missing.push(`Cell ${c}`);
          });
          configsChecked = true;
        }

        // Check if Column J is THEORETICAL
        let valJ = String(sheet.getRange("J" + actualRow).getValue() || "").trim().toUpperCase();
        let colsToCheck = req.cols.slice();

        if (valJ === "THEORETICAL") {
          // Do not require Col L for Elec, Water, and LPG
          colsToCheck = colsToCheck.filter(c => c !== "L");

          if (tabName === "Elec") {
            // Elec: Col P is required
            if (!colsToCheck.includes("P")) colsToCheck.push("P");
          } else if (tabName === "Water" || tabName === "LPG") {
            // Water and LPG: Col O is required, do not require L-dependent columns
            if (!colsToCheck.includes("O")) colsToCheck.push("O");
            colsToCheck = colsToCheck.filter(c => !["P", "S", "T", "U", "V", "W", "X", "M", "N"].includes(c));
          }
        }

        colsToCheck.forEach(col => {
          if (sheet.getRange(col + actualRow).getValue() === "") {
            missing.push(`${col}${actualRow}`);
          }
        });
      }
    });

    if (missing.length > 0) {
      stopErrors.push(`[${tabName}]: ${missing.join(", ")}`);
    }
  });

  if (stopErrors.length > 0) {
    SpreadsheetApp.getUi().alert("🚫 SCAN CANCELLED - DATA MISSING\n\n" + stopErrors.join("\n\n"));
    return false;
  }

  // --- 2. CLEAR LOGS ---
  const logSheetNames = ["Basic Anomalies", "Client Rate Anomalies"];
  logSheetNames.forEach(name => {
    let s = ss.getSheetByName(name) || ss.insertSheet(name);
    if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 6).clearContent();
    if (s.getLastRow() === 0) s.appendRow(["Timestamp", "Tab", "Cell", "Column Label", "Error Message", "Remarks"]);
  });

  // --- 2.1 VALIDATE INSTRUCTIONS CELL C24 AGAINST MASTER TAB NAMES (WITH BYPASS) ---
  const instSheet = ss.getSheetByName("Instructions");
  if (instSheet && extData && extData.propertyTabs) {
    const c24Val = String(instSheet.getRange("C24").getValue() || "").trim();
    const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MMM d, yyyy");
    const basicAnomaliesSheet = ss.getSheetByName("Basic Anomalies");

    if (c24Val === "") {
      basicAnomaliesSheet.appendRow([timestamp, "Instructions", "C24", "Property Selection", "Instructions cell C24 is empty.", "Please input property selection"]);
    } else if (!isBypassedProperty(c24Val)) {
      const isMatch = extData.propertyTabs.some(tabName => isPropertyAndTabMatch(c24Val, tabName));
      if (!isMatch) {
        basicAnomaliesSheet.appendRow([
          timestamp, 
          "Instructions", 
          "C24", 
          "Property Selection", 
          `Property "${c24Val}" does not match any Tab Name in master file.`, 
          "Tab Name lookup failed in file 12OOOzMVeWPb6SKJyNu3tewSPKrbu3s93jJA3SmPNSY4"
        ]);
      }
    }
  }

  // --- 3. RUN SCANS SILENTLY ---
  tabsToScan.forEach(tabName => {
    scanTab(tabName, false, extData); 
  });

  // --- 4. SHOW FINAL SUMMARY MODAL ---
  const stdTotal = Math.max(0, ss.getSheetByName("Basic Anomalies").getLastRow() - 1);
  const kaTotal = Math.max(0, ss.getSheetByName("Client Rate Anomalies").getLastRow() - 1);
  showScanSuccessModal(tabsToScan, stdTotal, kaTotal);

  return (stdTotal === 0 && kaTotal === 0);
}

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
  if (isMasterFileBlocked()) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const masterDbId = "1hMMUd4ho50HP63dc2fRAo--iK-m7YotamkKtsDGT_Us";
  const props = PropertiesService.getScriptProperties();
  const activeDB_ID = props.getProperty("ACTIVE_DB_ID") || masterDbId;

  const instructionSheet = ss.getSheetByName("Instructions");
  if (!instructionSheet) {
    ui.alert("❌ ERROR: 'Instructions' tab not found.");
    return;
  }

  const currentRef = instructionSheet.getRange("C7").getValue().toString().trim();

  if (currentRef === "") {
    ui.alert("❌ ERROR: Cell C7 in 'Instructions' tab is empty. Please enter a Reference Number first.");
    return;
  }

  try {
    const dbSs = SpreadsheetApp.openById(activeDB_ID);
    const subTab = dbSs.getSheetByName("PBTT Submission");

    if (subTab) {
      const lastDbRow = subTab.getLastRow();
      if (lastDbRow > 1) {
        const recordedRefs = subTab.getRange(2, 11, lastDbRow - 1, 1).getValues().flat();
        const recordedRefsStr = recordedRefs.map(r => String(r).trim());

        if (recordedRefsStr.includes(currentRef)) {
          ui.alert(
            `🚫 FETCH BLOCKED\n\n` +
            `You cannot use this file more than 1.\n\n` +
            `Make a copy of the file "(Master) BTT Template" instead.`
          );
          return;
        }
      }
    }
  } catch (err) {
    console.error("Database Validation Error: " + err.message);
    ui.alert("⚠️ Database connection warning: Could not verify Reference Status.");
  }

  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return;

  const dataStartRow = 13;

  const sourceLink = sheet.getRange("A1").getValue();
  if (!sourceLink) {
    ui.alert("Paste SOURCE LINK in cell C7 in Instructions Tab (or ensure A1 references it).");
    return;
  }

  let sourceSS;
  try {
    sourceSS = SpreadsheetApp.openByUrl(sourceLink);
  } catch (e) {
    ui.alert("Cannot open source link.");
    return;
  }

  if (sourceSS.getId() === ss.getId()) {
    ui.alert("FETCH CANCELLED: You are using the current spreadsheet's URL. Please use an external source link.");
    return;
  }

  const sourceSheet = sourceSS.getSheetByName(tabName);
  if (!sourceSheet) {
    ui.alert(`Tab "${tabName}" not found in source.`);
    return;
  }

  const lastSourceRow = sourceSheet.getLastRow();
  const lastSourceCol = Math.max(sourceSheet.getLastColumn(), 29);
  if (lastSourceRow < dataStartRow) return;

  const rawData = sourceSheet.getRange(dataStartRow, 1, lastSourceRow - dataStartRow + 1, lastSourceCol).getValues();

  // Ensure the sheet has at least 38 columns (up to Column AL)
  if (sheet.getMaxColumns() < 38) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 38 - sheet.getMaxColumns());
  }

  const destWidth = sheet.getMaxColumns();
  const pasteArray = [];

  const skipIndices = (EXCLUSIONS[tabName] || []).map(letter => colToIdx(letter));
  const mapping = FETCH_MAPS[tabName] || {};

  let totalFound = false;
  let rowCounter = 0;

  let lastTotalIndex = -1;
  for (let i = rawData.length - 1; i >= 0; i--) {
    let checkVal = String(rawData[i][0] || "").trim().toLowerCase();
    let isSub = checkVal.includes("subtotal") || checkVal.includes("sub-total") || checkVal.includes("sub total");

    if (checkVal.includes("total") && !isSub) {
      lastTotalIndex = i;
      break;
    }
  }

  for (let i = 0; i < rawData.length; i++) {
    let sourceRow = rawData[i];
    let valA_source = String(sourceRow[0] || "").trim();
    let valA_lower = valA_source.toLowerCase();

    let isSubTotal = valA_lower.includes("subtotal") || valA_lower.includes("sub-total") || valA_lower.includes("sub total");
    let isTerminatingTotal = (i === lastTotalIndex);
    let isIntermediateTotal = (!isTerminatingTotal && valA_lower.includes("total") && !isSubTotal);

    let rowHasData = sourceRow.some(cell => String(cell).trim() !== "");

    if (rowHasData || isSubTotal || isIntermediateTotal || isTerminatingTotal) {
      if (isTerminatingTotal) {
        if (pasteArray.length > 0 && !pasteArray[pasteArray.length - 1].every(cell => cell === "")) {
          pasteArray.push(new Array(destWidth).fill(""));
        }
      }

      let destRow = new Array(destWidth).fill("");

      if (isSubTotal || isIntermediateTotal || isTerminatingTotal) {
        destRow[0] = valA_source;
      } else {
        rowCounter++;
        destRow[0] = rowCounter;
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

    if (isTerminatingTotal) {
      totalFound = true;
      for (let s = 0; s < 3; s++) pasteArray.push(new Array(destWidth).fill(""));
      let sigRow = new Array(destWidth).fill("");
      sigRow[0] = "Prepared By:"; sigRow[7] = "Checked By:"; sigRow[28] = "Noted By:";
      pasteArray.push(sigRow);
      break;
    }
  }

  if (!totalFound) {
    if (pasteArray.length > 0 && !pasteArray[pasteArray.length - 1].every(cell => cell === "")) {
      pasteArray.push(new Array(destWidth).fill(""));
    }
    let dummyTotalRow = new Array(destWidth).fill("");
    dummyTotalRow[0] = "TOTAL";
    pasteArray.push(dummyTotalRow);
    for (let s = 0; s < 3; s++) pasteArray.push(new Array(destWidth).fill(""));
    let dummySig = new Array(destWidth).fill("");
    dummySig[0] = "Prepared By:"; dummySig[7] = "Checked By:"; dummySig[28] = "Noted By:";
    pasteArray.push(dummySig);
  }

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

  SpreadsheetApp.getActive().toast(`Fetch complete. Mapped Columns apply flawlessly to intermediate totals. Unused trailing data ignored.`, "Success", 5);
}

/* =================================
4. RUN FORMULAS
================================= */
const sumColsELEC = ["L", "N", "P", "Q", "AA", "AB", "AF", "AG", "AI", "AJ"];
const sumColsWAT = ["L", "N", "P", "Q", "X", "AA", "AB", "AF", "AG", "AI", "AJ"];
const sumColsLPG = ["L", "N", "P", "Q", "AA", "AB", "AF", "AG", "AI", "AJ"];

function applyFormulasToSheet(tabName) {
  if (isMasterFileBlocked()) return;

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(tabName);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet "${tabName}" not found.`);
    return;
  }

  const valL5 = sheet.getRange("L5").getValue();
  const valL6 = sheet.getRange("L6").getValue();

  if (valL5 === "" || valL6 === "" || isNaN(valL5) || isNaN(valL6)) {
    SpreadsheetApp.getUi().alert("❌ Action Blocked: L5 and L6 must contain numeric values.");
    return;
  }

  if (Number(valL5) <= Number(valL6)) {
    SpreadsheetApp.getUi().alert("❌ Action Blocked: L5 (Current Rate) must be greater than L6 (Previous Rate).");
    return;
  }

  let activeMap;
  let activeSumCols;

  if (tabName === "Water") {
    activeMap = formulaMapWater;
    activeSumCols = sumColsWAT;
    const valU10 = sheet.getRange("U10").getValue();
    if (valU10 === "" || isNaN(valU10)) {
      SpreadsheetApp.getUi().alert("❌ Action Blocked: U10 must contain a numeric value for Water formulas.");
      return;
    }
  } else if (tabName === "LPG") {
    activeMap = formulaMapLPG;
    activeSumCols = sumColsLPG;
    const valN10 = sheet.getRange("N10").getValue();
    if (valN10 === "" || isNaN(valN10)) {
      SpreadsheetApp.getUi().alert("❌ Action Blocked: N10 must contain a numeric value for LPG formulas.");
      return;
    }
  } else {
    activeMap = formulaMapElec;
    activeSumCols = sumColsELEC;
  }

  const lastRow = sheet.getLastRow();
  const fullDataA = sheet.getRange(1, 1, lastRow, 1).getValues();
  const colEIdx = colToIdx("E");
  const fullDataE = sheet.getRange(1, colEIdx + 1, lastRow, 1).getValues();

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
      sheet.getRange(r, 1, 1, 35).clearContent();
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
    if (valP !== "" && valP !== "Put/input") targetCols = targetCols.filter(c => c !== "P");
    if (valZ !== "") targetCols = targetCols.filter(c => c !== "Z");
    if (valJ.includes("theoretical") || valK.includes("theoretical")) {
      targetCols = targetCols.filter(c => c !== "L");
    }

    targetCols.forEach(colKey => {
      sheet.getRange(`${colKey}${r}`).setFormula(activeMap[colKey](r));
    });
  }

  let sectionStartRow = CONFIG.dataStartRow;
  let subTotalRowsFound = [];

  for (let i = CONFIG.dataStartRow - 1; i < stopRow; i++) {
    const rowLabel = String(fullDataA[i][0]).toLowerCase().trim();
    const r = i + 1;
    const normalizedLabel = rowLabel.replace(/[^a-z]/g, "");

    if (normalizedLabel.includes("subtotal")) {
      const rangeEnd = r - 1;
      activeSumCols.forEach(col => {
        sheet.getRange(`${col}${r}`).setFormula(`=SUM(${col}${sectionStartRow}:${col}${rangeEnd})`);
      });
      subTotalRowsFound.push(r);
      sectionStartRow = r + 1;
    }

    if (normalizedLabel === "total") {
      activeSumCols.forEach(col => {
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
    const cleanedFooter = footerValues.map(row =>
      row.map(cell => (typeof cell === 'number' && cell !== "") ? "" : cell)
    );
    footerRange.setValues(cleanedFooter);
  }

  SpreadsheetApp.flush();

  ["P", "Q", "AF", "AI", "AJ", "AG"].forEach(c => {
    sheet.getRange(`${c}${CONFIG.dataStartRow}:${c}${stopRow}`).setNumberFormat("#,##0.00");
  });

  ["J", "K", "L"].forEach(c => {
    sheet.getRange(`${c}${CONFIG.dataStartRow}:${c}${stopRow}`).setNumberFormat("#,##0.0000");
  });

  ["AH", "AK", "AC"].forEach(c => {
    sheet.getRange(`${c}${CONFIG.dataStartRow}:${c}${stopRow}`).setNumberFormat("0.00%");
  });

  SpreadsheetApp.getActive().toast(`Logic applied successfully to ${tabName}.`, "Success");
}

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
  AK: (r) => `=IFERROR(AJ${r}/AI${r},"-")`,
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
  AK: (r) => `=IFERROR(AJ${r}/AI${r},"-")`,
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
  AG: (r) => `=IFERROR(N${r}-AF${r},"-")`,
  AH: (r) => `=IFERROR(AG${r}/AF${r},"-")`,
  AJ: (r) => `=IFERROR(P${r}-AI${r},"-")`,
  AK: (r) => `=IFERROR(AJ${r}/AI${r},"-")`,
};

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
  if (isMasterFileBlocked()) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
  if (sheet && sheet.getLastRow() >= CONFIG.dataStartRow) {
    sheet.getRange(CONFIG.dataStartRow, 1, sheet.getLastRow() - CONFIG.dataStartRow + 1, sheet.getMaxColumns()).clearContent();
  }
}

/* =================================
7. FINAL SCAN TAB
================================= */
function scanTab(tabName, shouldClearLogs = true, extData = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return;

  if (!extData) {
    extData = getExternalValidationData();
  }

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

  const scanCols = Math.max(sheet.getLastColumn(), 38);
  const dataRange = sheet.getRange(CONFIG.dataStartRow, 1, lastRow - CONFIG.dataStartRow + 1, scanCols);
  const dataValues = dataRange.getValues();
  const headers = sheet.getRange(CONFIG.headerRow, 1, 1, scanCols).getValues()[0];
  const valL5 = sheet.getRange("L5").getValue();
  const valL6 = sheet.getRange("L6").getValue();
  const rawE4 = sheet.getRange("E4").getValue();

  const kaRefMap = getKAData();
  const globalAffiliates = getGlobalAffiliates(); 
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

    const valE = String(row[colToIdx("E")] || "").trim();
    const labelA = String(row[0] || "").trim();

    if (valE === "") continue;

    const normalizedLabelA = labelA.toLowerCase().replace(/[^a-z]/g, "");
    if (normalizedLabelA.includes("total")) continue;

    const valAD = String(row[colToIdx("AD")] || "").trim();
    if (valAD.toUpperCase() === "MONITORING") {
      const valP = row[colToIdx("P")];
      const isZeroP = (valP === 0 || String(valP).trim() === "0" || (!isNaN(valP) && Number(valP) === 0 && String(valP).trim() !== ""));
      
      // If REMARKS is MONITORING and Column P is 0, record as an anomaly
      if (isZeroP) {
        logHelper(row, rowNum, "P", 'Cannot use "MONITORING" in Remarks when Column P is 0');
      }
      continue;
    }

    // --- STEP 1: RUN CHECKLIST (Pass tabName for theoretical checks) ---
    runCommonChecklist(row, rowNum, (r, c, m, res, arr) => logHelper(row, r, c, m, res, arr), valL5, valL6, globalAffiliates, tabName);

    // --- STEP 2: KA VALIDATION ---
    if (kaRefMap) {
      const valF = String(row[colToIdx("F")] || "").trim().toUpperCase();
      const valG = String(row[colToIdx("G")] || "").trim().toUpperCase();
      const hasKA = (valF === "KA" || valF === "KA&SR" || valG === "KA");

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
        if (!hasKA) logHelper(row, rowNum, "F", 'user need to put "KA"', `DB match: [${valE}]`, kaLogs);
      } else {
        if (hasKA) logHelper(row, rowNum, "F", 'user need to remove "KA"', `No DB entry found for [${valE}] in Site [${headerE4}]`, kaLogs);
      }
    }

    // --- STEP 3: TAB SPECIFIC CALCULATIONS ---
    const valJ_Row = String(row[colToIdx("J")] || "").trim().toUpperCase();
    const isRowTheo = (valJ_Row === "THEORETICAL");

    switch (tabName) {
      case "Elec": {
        const valAD_Q = String(row[colToIdx("AD")] || "").trim().toLowerCase();
        const skipQ = valAD_Q.includes("inaccesible meter") || valAD_Q.includes("inaccessible meter") || valAD_Q.includes("minimal usage");
        if (!skipQ) {
          if (!(typeof row[colToIdx("Q")] === 'number' && row[colToIdx("Q")] > 0)) logHelper(row, rowNum, "Q", "Amount should be a number > 0");
        }
        break;
      }
      case "Water": {
        if (!isRowTheo) {
          ["S", "T", "U", "V", "W"].forEach(c => { if (String(row[colToIdx(c)]).trim() === "") logHelper(row, rowNum, c, "Formula output missing"); });
          const valAD_X = String(row[colToIdx("AD")] || "").trim().toLowerCase();
          const skipX = valAD_X.includes("inaccesible meter") || valAD_X.includes("inaccessible meter") || valAD_X.includes("minimal usage");
          if (!skipX) {
            if (!(typeof row[colToIdx("X")] === 'number' && row[colToIdx("X")] > 0)) logHelper(row, rowNum, "X", "VAT amount missing");
          }
        }
        break;
      }
      case "LPG": {
        if (!isRowTheo) {
          const vL = row[colToIdx("L")];
          const valAD_LPG = String(row[colToIdx("AD")] || "").trim().toLowerCase();
          const skipLPG = valAD_LPG.includes("inaccesible meter") || valAD_LPG.includes("inaccessible meter") || valAD_LPG.includes("minimal usage");
          if (typeof vL === 'number' && !skipLPG) {
            if (!(typeof row[colToIdx("M")] === 'number' && row[colToIdx("M")] > 0)) logHelper(row, rowNum, "M", "Multiplier missing");
            if (!(typeof row[colToIdx("N")] === 'number' && row[colToIdx("N")] > 0)) logHelper(row, rowNum, "N", "Consumption amount error");
          }
        }
        break;
      }
    }

    // --- STEP 4: EXTERNAL VALIDATION (Match Property Tab First, then Columns) ---
    if (extData && extData.ss) {
      let userProperty = String(sheet.getRange("E4").getValue() || "").trim();
      if (!userProperty) {
        const iSheet = ss.getSheetByName("Instructions");
        if (iSheet) userProperty = String(iSheet.getRange("C24").getValue() || "").trim();
      }

      const cleanProp = superClean(userProperty);

      // BYPASS: Skip external checks if property is Capital Town or Maple Grove
      if (!isBypassedProperty(userProperty)) {

        if (cleanProp && !extData.cache[cleanProp]) {
          let matchedSheet = null;
          const allSheets = extData.ss.getSheets();
          for (let s of allSheets) {
            if (isPropertyAndTabMatch(userProperty, s.getName())) {
              matchedSheet = s;
              break;
            }
          }

          if (matchedSheet) {
            const extLastRow = matchedSheet.getLastRow();
            if (extLastRow >= 3) {
              const extHeaders = matchedSheet.getRange(2, 1, 1, matchedSheet.getLastColumn()).getValues()[0];
              let tenantCol = -1;
              let codeCol = -1;
              let propCol = -1;

              for (let c = 0; c < extHeaders.length; c++) {
                const hClean = superClean(extHeaders[c]);
                if (hClean === "property" || hClean.includes("property")) propCol = c;
                if (hClean === "tenant name" || hClean.includes("tenant name")) tenantCol = c;
                if (hClean === "customer code" || hClean.includes("customer code")) codeCol = c;
              }

              if (tenantCol !== -1 && codeCol !== -1) {
                const extRows = matchedSheet.getRange(3, 1, extLastRow - 2, matchedSheet.getLastColumn()).getValues();
                const records = extRows.map(r => ({
                  property: (propCol !== -1) ? String(r[propCol] || "").trim() : matchedSheet.getName(),
                  tenantName: String(r[tenantCol] || "").trim(),
                  customerCode: String(r[codeCol] || "").trim()
                })).filter(r => r.tenantName || r.customerCode);

                extData.cache[cleanProp] = { sheetName: matchedSheet.getName(), records: records, hasPropCol: (propCol !== -1), valid: true };
              } else {
                extData.cache[cleanProp] = { sheetName: matchedSheet.getName(), error: "Missing 'TENANT NAME' or 'CUSTOMER CODE' in Row 2 headers.", valid: false };
              }
            } else {
              extData.cache[cleanProp] = { sheetName: matchedSheet.getName(), error: "Property tab has no data in Row 3+.", valid: false };
            }
          } else {
            extData.cache[cleanProp] = { error: `No Tab matching Property "${userProperty}" in master file.`, valid: false };
          }
        }

        const propData = extData.cache[cleanProp];

        if (!propData || !propData.valid) {
          logHelper(row, rowNum, "E", propData ? propData.error : `Property "${userProperty}" tab not found in master file.`);
        } else {
          let partnerColIdx = -1;
          let codeColIdx = -1;
          for (let c = 0; c < headers.length; c++) {
            const hClean = superClean(headers[c]);
            if (hClean.includes("retail partner") || hClean.includes("tenant name")) partnerColIdx = c;
            if (hClean.includes("tenant code") || hClean.includes("customer code")) codeColIdx = c;
          }
          if (partnerColIdx === -1) partnerColIdx = colToIdx("E");

          const alScanIdx = colToIdx("AL");
          const userRetailPartner = String(row[partnerColIdx] || "").trim();
          
          // Strictly read Tenant Code from Column AL (or dynamic TENANT CODE header)
          let userTenantCode = String(row[alScanIdx] !== undefined && row[alScanIdx] !== null ? row[alScanIdx] : "").trim();
          if (!userTenantCode && codeColIdx !== -1) {
            userTenantCode = String(row[codeColIdx] || "").trim();
          }

          let basePartner = userRetailPartner;
          if (userRetailPartner.includes('_')) {
            const parts = userRetailPartner.split('_');
            if (parts[parts.length - 1].trim().toLowerCase() === "affiliates") {
              basePartner = parts.slice(0, -1).join('_').trim();
            }
          }

          const cleanPartner = superClean(userRetailPartner);
          const cleanBase = superClean(basePartner);
          const cleanCode = superClean(userTenantCode);

          const partnerExists = propData.records.some(rec => {
            const cRec = superClean(rec.tenantName);
            return cRec && (cRec === cleanPartner || cRec === cleanBase);
          });

          const codeExists = propData.records.some(rec => {
            const cRec = superClean(rec.customerCode);
            return cRec && cRec === cleanCode;
          });

          if (!partnerExists) {
            logHelper(row, rowNum, "E", `RETAIL PARTNER "${userRetailPartner}" does not exist in master Tab "${propData.sheetName}" under TENANT NAME.`);
          }
          if (!codeExists) {
            logHelper(row, rowNum, "AL", `TENANT CODE "${userTenantCode}" does not exist in master Tab "${propData.sheetName}" under CUSTOMER CODE.`);
          }

          if (partnerExists && codeExists) {
            const pairExists = propData.records.some(rec => {
              const cRecTenant = superClean(rec.tenantName);
              const cRecCode = superClean(rec.customerCode);
              const tMatch = cRecTenant && (cRecTenant === cleanPartner || cRecTenant === cleanBase);
              const cMatch = cRecCode && cRecCode === cleanCode;
              return tMatch && cMatch;
            });

            if (!pairExists) {
              logHelper(
                row, 
                rowNum, 
                "E", 
                `Binding mismatch: RETAIL PARTNER "${userRetailPartner}" and TENANT CODE "${userTenantCode}" are not paired together in Tab "${propData.sheetName}".`, 
                `Property Tab Row Pairing Invalid`
              );
            }
          }
        }
      }
    }
  }

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
   OTHER HELPERS
================================= */
function superClean(val) {
  if (!val) return "";
  let str = String(val).toLowerCase();
  str = str.replace(/[^a-z0-9\s]/g, ' ').replace(/[\s\u00A0]+/g, ' ').trim();
  return str;
}

function getKAData() {
  const KA_REF_URL = "https://docs.google.com/spreadsheets/d/1jY-9FMha3x972o4Gz1d6DVD36d3ppjHW_WM1DHJz6ag/edit";
  try {
    const ss = SpreadsheetApp.openByUrl(KA_REF_URL);
    const sheet = ss.getSheetByName("Data");
    if (!sheet) throw new Error("Master sheet 'Data' not found.");

    const lastR = sheet.getLastRow();
    if (lastR < 2) return {};

    const rawData = sheet.getRange(2, 1, lastR - 1, 5).getValues();
    const propertyMap = {};

    for (let i = 0; i < rawData.length; i++) {
      const row = rawData[i];

      const valA = String(row[0] || "").trim();
      const valB = String(row[1] || "").trim();
      const valE = String(row[4] || "").trim();
      const category = superClean(row[2]);

      const currentRowNum = i + 2;

      if (valE !== "" && valA === "") {
        const specificMsg = `🛑 MASTER DATABASE ERROR (Row ${currentRowNum})\n\nColumn E contains values, but Column A is blank. You must input a number first in Column A of the Master File to proceed.`;
        SpreadsheetApp.getUi().alert(specificMsg);
        throw new Error("Aborted: Missing number in Master Column A.");
      }

      if (valB !== "" && valA === "") {
        const errorMsg = `🛑 MASTER DATABASE ERROR\n\nRow ${currentRowNum} has a Main Name (Col B) but is missing an Identifier in Column A.\n\nPlease fix the Master File to proceed.`;
        SpreadsheetApp.getUi().alert(errorMsg);
        throw new Error("Master Data Violation: Missing Column A.");
      }

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
    if (e.message.includes("Aborted") || e.message.includes("Violation")) throw e;
    console.error("KA Ref Error: " + e.message);
    return null;
  }
}

/* =================================
REFACTORED: THE "COMMON" CHECKLIST (ALL TABS)
================================= */
function runCommonChecklist(row, rNum, log, L5, L6, globalAffiliates = new Set(), tabName = "") {
  const get = (colLetter) => row[colToIdx(colLetter)];

  const valA = String(get("A") || "").trim();
  const valE = String(get("E") || "").trim();

  const valAD = String(get("AD") || "").trim();
  if (valAD.toUpperCase() === "MONITORING") {
    return;
  }

  const currentTenant = valE.toLowerCase().trim();
  const currentBase = valE.includes('_') ? valE.split('_').slice(0, -1).join('_').trim().toLowerCase() : currentTenant;
  const isAffiliate = (valE.includes('_') && valE.split('_').pop().trim().toLowerCase() === "affiliates") || globalAffiliates.has(currentBase);

  if (valE !== "" && valA === "") {
    const errorMsg = `CRITICAL DATA ERROR\n\nRow ${rNum} has a Tenant Name in Column E ("${valE}") but the identifier in Column A is blank.\n\nPROCESS HALTED: Every tenant must have an Row Number in Column A to continue.`;
    SpreadsheetApp.getUi().alert(errorMsg);
    throw new Error(`Execution stopped at row ${rNum}: Missing Col A with populated Col E.`);
  }

  const valJ_Val = String(get("J") || "").trim().toUpperCase();
  const isTheo = (valJ_Val === "THEORETICAL");

  const valL = get("L");
  const L_isHyphen = (String(valL).trim() === "-");

  if (valE !== "") {
    // If Column J is THEORETICAL, do not require Column L
    const mandatoryReadingCols = isTheo ? ["J", "K"] : ["J", "K", "L"];
    mandatoryReadingCols.forEach(c => {
      if (String(get(c)).trim() === "") log(rNum, c, "Should not be blank if E has entry");
    });

    // Check specific requirements for THEORETICAL
    if (isTheo) {
      if (tabName === "Elec" && String(get("P")).trim() === "") {
        log(rNum, "P", "Column P is required when Column J is THEORETICAL");
      }
      if ((tabName === "Water" || tabName === "LPG") && String(get("O")).trim() === "") {
        log(rNum, "O", "Column O is required when Column J is THEORETICAL");
      }
    }

    const valF = String(get("F") || "").trim().toUpperCase();
    const valG = String(get("G") || "").trim().toUpperCase();
    const hasKA = (valF === "KA" || valF === "KA&SR" || valG === "KA");
    const hasSR = (valF === "SR" || valF === "KA&SR" || valG === "SR");

    const allowedFValues = ["KA", "SR", "REG", "KA&SR"];
    if (!allowedFValues.includes(valF)) {
      log(rNum, "F", "Column F must be \"KA\", \"SR\", \"REG\", or \"KA&SR\" when Column E has a value");
    }

    const valO = get("O");
    const oStr = String(valO).toLowerCase().trim();
    const oIsFixOrTheo = (oStr === "fix rate" || oStr === "theoretical");

    if (hasKA) {
      if (typeof valO === 'number') {
        if (!(valO > 0)) {
          log(rNum, "O", "KA Billing Rate must be a valid positive number, \"fix rate\" or \"theoretical\"");
        }
      } else {
        if (!oIsFixOrTheo) {
          log(rNum, "O", "KA Billing Rate must be a valid positive number, \"fix rate\" or \"theoretical\"");
        }
      }
    } else if (hasSR) {
      if (typeof valO === 'number') {
        if (!(valO >= L6)) {
          log(rNum, "O", "SR Billing Rate must not be less than L6 (Previous Rate)");
        }
      } else {
        if (!oIsFixOrTheo) {
          log(rNum, "O", "SR Billing Rate must be a valid number not less than L6, \"fix rate\" or \"theoretical\"");
        }
      }
    } else if (!isAffiliate) {
      if (typeof valO === 'number') {
        if (!(valO > 0)) log(rNum, "O", "Should equal to L5, \"fix rate\" or \"theoretical\"");
        if (!L_isHyphen && valO !== L5) log(rNum, "O", "Should equal to L5, or if L= \"-\" then, O= \"fix rate\" or O=\"theoretical\"");
      } else {
        if (!oIsFixOrTheo) log(rNum, "O", "Should equal to L5, \"fix rate\" or \"theoretical\"");
        if (L_isHyphen && !oIsFixOrTheo) log(rNum, "O", "Should equal to L5, or if L= \"-\" then, O= \"fix rate\" or O=\"theoretical\"");
      }
    }

    // Column P Condition: Required for standard rows, or Elec when THEORETICAL
    if (!isTheo || tabName === "Elec") {
      const valP = get("P");
      const valAD_P = String(get("AD") || "").trim().toLowerCase();
      
      // Case-insensitive check for allowed remarks (including common variants)
      const allowedRemarks = [
        "defective meter",
        "inaccessible meter",
        "inaccesible meter",
        "very minimal usage",
        "minimal usage"
      ];
      const hasAllowedRemark = allowedRemarks.some(kw => valAD_P.includes(kw));
      const isZeroValue = (valP === 0 || String(valP).trim() === "0" || (!isNaN(valP) && Number(valP) === 0 && String(valP).trim() !== ""));

      // If value is 0 and remarks qualify, exclude from anomalies
      if (isZeroValue && hasAllowedRemark) {
        // Excluded from anomalies
      } else {
        if (!(typeof valP === 'number' && valP > 0)) {
          log(rNum, "P", "Should be a number >0");
        }
      }
    }

    const valZ = get("Z");
    if (valZ === "") log(rNum, "Z", "Should be a number >0, \"fix rate\" or \"theoretical\"");
    if (typeof valZ === 'number' && !L_isHyphen && valZ !== L6) {
      log(rNum, "Z", "Should equal to L6, or if L= \"-\" then, O= \"fix rate\" or O=\"theoretical\"");
    }

    ["AA", "AB", "AC", "AG", "AJ"].forEach(c => {
      if (String(get(c)).trim() === "") log(rNum, c, "Should not be empty (c/o fx)");
    });

    ["AF", "AI"].forEach(c => {
      if (String(get(c)).trim() === "") log(rNum, c, "Should not be empty if E has entry");
    });

    const valAD_Exp = String(get("AD") || "").trim();
    const cleanAD = valAD_Exp.toLowerCase().replace(/[^a-z0-9]/g, "");
    
    const placeholderBlacklist = [
      "na", "none", "nil", "notapplicable", "notaplicable", "no", "ok", "okay",
      "test", "testing", "tbd", "tbc", "noted", "done", "noneed", "notneeded",
      "nocomment", "noidea", "nothing", "yes", "asdf"
    ];

    const isPlaceholder = placeholderBlacklist.includes(cleanAD);
    const isLongEnough = cleanAD.length >= 5;

    const hasValidExplanation = (valAD_Exp !== "" && !isPlaceholder && isLongEnough);

    if (!hasValidExplanation) {
      ["AH", "AK"].forEach(c => {
        const v = get(c);
        if (typeof v === 'number') {
          if (v > 0.3 || v < -0.3) {
            let errorMsg = "Variance alert: Value is outside +/- 30% threshold.";
            if (valAD_Exp === "") {
              errorMsg = "Variance alert: High variance detected (+/- 30%). Please provide an explanation in Column AD.";
            } else if (isPlaceholder || !isLongEnough) {
              errorMsg = "Variance alert: Incomplete explanation. Please describe the specific reason for this variance in Column AD.";
            }
            log(rNum, c, errorMsg);
          }
        }
      });
    }
  }
}

function confirmFetchOverwrite(tabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  const ui = SpreadsheetApp.getUi();

  if (!sheet) return false;

  const lastSheetRow = sheet.getLastRow();
  let hasExistingData = false;

  if (lastSheetRow >= CONFIG.dataStartRow) {
    const colE_values = sheet.getRange(CONFIG.dataStartRow, 5, lastSheetRow - CONFIG.dataStartRow + 1, 1).getValues();
    hasExistingData = colE_values.some(row => {
      const val = row[0];
      if (val === "" || val === null || val === undefined || val === false) return false;
      return String(val).trim() !== "";
    });
  }

  if (hasExistingData) {
    const res = ui.alert(
      'Confirm Overwrite',
      `Data already exists in "${tabName}". Overwrite?`,
      ui.ButtonSet.YES_NO
    );
    if (res !== ui.Button.YES) return false;
  }

  return true;
}

/* =================================
8. RECORD & SUBMIT ACTIVE PBTT
================================= */
function recordActivePBTT() {
  if (isMasterFileBlocked()) return;

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Server Busy. Please try again.");
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    // 1. RUN SYSTEM SCAN TO DETECT ANOMALIES PRIOR TO SUBMITTING
    const isClean = scanAllTabs();
    if (!isClean) {
      ui.alert(
        "🚫 SUBMISSION BLOCKED\n\n" +
        "Anomalies or configuration gaps were discovered during the system scan.\n\n" +
        "Please inspect the 'Basic Anomalies' and 'Client Rate Anomalies' logs, resolve all items, and attempt submission again."
      );
      return;
    }

    const masterDbId = "1hMMUd4ho50HP63dc2fRAo--iK-m7YotamkKtsDGT_Us";

const tabValidationMaps = {
      "Elec": ["B", "D", "F", "H", "I", "J", "K", "L", "O", "P", "Q", "Y", "Z", "AA", "AB", "AC", "AG", "AH", "AJ", "AK", "AL"],
      "Water": ["B", "D", "F", "H", "J", "K", "L", "O", "P", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AG", "AH", "AJ", "AK", "AL"],
      "LPG": ["B", "D", "F", "H", "J", "K", "L", "M", "N", "O", "P", "Q", "Y", "Z", "AA", "AB", "AC", "AG", "AH", "AJ", "AK", "AL"]
    };

    // ===============================================
    // 1. REF# & DATE SETUP
    // ===============================================
    const instSheet = ss.getSheetByName("Instructions");
    if (!instSheet) {
      ui.alert("❌ ERROR: 'Instructions' tab not found.");
      return;
    }

    const currentRef = instSheet.getRange("C7").getValue().toString().trim();
    if (currentRef === "") {
      ui.alert("❌ ERROR: No Reference Number found in 'Instructions' tab C7.");
      return;
    }

    const rawTargetStart = instSheet.getRange("C26").getValue();
    const rawTargetEnd = instSheet.getRange("C27").getValue();

    const normalizeDate = (d) => {
      if (!d || !(d instanceof Date) || isNaN(d.getTime())) return null;
      const n = new Date(d);
      n.setHours(0, 0, 0, 0);
      return n.getTime();
    };

    const targetStartInfo = normalizeDate(rawTargetStart);
    const targetEndInfo = normalizeDate(rawTargetEnd);

    if (!targetStartInfo || !targetEndInfo) {
      ui.alert("❌ ERROR: Invalid or missing billing dates in 'Instructions' tab (C26/C27).");
      return;
    }

    // ===============================================
    // 2. PERIOD STATUS VALIDATION
    // ===============================================
    const dbPeriodTab = "dvPeriod";
    let periodFound = false;
    let periodIsActive = false;
    let periodIsLocked = false;
    let lockDateFormatted = "";

    try {
      const dbSs = SpreadsheetApp.openById(masterDbId);
      const periodSheet = dbSs.getSheetByName(dbPeriodTab);
      const periodData = periodSheet.getDataRange().getValues();

      const today = new Date();
      today.setHours(0, 0, 0, 0);

      for (let i = 1; i < periodData.length; i++) {
        const row = periodData[i];
        const dbStart = normalizeDate(row[0]);
        const dbEnd = normalizeDate(row[1]);

        if (dbStart === targetStartInfo && dbEnd === targetEndInfo) {
          periodFound = true;
          const status = String(row[3]).trim();
          if (status === "Active") {
            periodIsActive = true;
          }

          if (row[2]) {
            const lockDate = new Date(row[2]);
            lockDate.setHours(0, 0, 0, 0);
            const bypassTag = String(row[4] || "").trim();

            if (today >= lockDate) {
              if (bypassTag === "Bypass") {
                periodIsLocked = false;
              } else {
                periodIsLocked = true;
                lockDateFormatted = Utilities.formatDate(lockDate, "Asia/Manila", "MMM d, yyyy");
              }
            }
          }
          break;
        }
      }

      if (!periodFound) {
        ui.alert(
          "🚫 CONFIGURATION ERROR\n\n" +
          "The dates in Instructions C26 & C27 do not match any known period in the master database.\n" +
          "Please verify your billing start/end dates."
        );
        return;
      }

      if (!periodIsActive) {
        ui.alert(
          "🚫 SUBMISSION BLOCKED\n\n" +
          "The billing period defined is set to 'Inactive' in the system.\n\n" +
          "Please contact the administrator."
        );
        return;
      }

      if (periodIsLocked) {
        ui.alert(
          `🚫 PERIOD LOCKED\n\n` +
          `The active period defined is reached its Lock Date on ${lockDateFormatted}.\n` +
          `Submission is blocked.\n\n` +
          `To submit, a 'Bypass' tag is required from Admin.`
        );
        return;
      }

    } catch (err) {
      ui.alert("❌ Validation Connection Error: " + err.message);
      return;
    }

    // ===============================================
    // 3. TAB SPECIFIC VALIDATION & EXTERNAL TAB/COLUMN CHECK
    // ===============================================
    const extValidationFileId = "12OOOzMVeWPb6SKJyNu3tewSPKrbu3s93jJA3SmPNSY4";
    let extSS;
    try {
      extSS = SpreadsheetApp.openById(extValidationFileId);
    } catch (extErr) {
      ui.alert("❌ External Validation Error: Cannot access master spreadsheet '12OOOzMVeWPb6SKJyNu3tewSPKrbu3s93jJA3SmPNSY4'.");
      return;
    }

    const extSheetsList = extSS.getSheets();
    const propertyRecordsCache = {};

    for (let tabName in tabValidationMaps) {
      let currentSheet = ss.getSheetByName(tabName);
      if (!currentSheet) continue;

      let lastRow = currentSheet.getLastRow();
      let startRow = 13;
      if (lastRow < startRow) continue;

      let userProperty = String(currentSheet.getRange("E4").getValue() || "").trim();
      if (!userProperty && instSheet) {
        userProperty = String(instSheet.getRange("C24").getValue() || "").trim();
      }

      if (!userProperty) {
        ui.alert(
          `🚫 PROPERTY NAME MISSING\n\n` +
          `Tab: [${tabName}]\n\n` +
          `Please provide a valid PROPERTY NAME in cell E4 or Instructions C24 before submitting.`
        );
        return;
      }

      const cleanUserProp = superClean(userProperty);
      const isPropertyBypassed = isBypassedProperty(userProperty);

      let propSheetName = "";

      // Check external file if not bypassed
      if (!isPropertyBypassed) {
        let targetPropertySheet = null;
        for (let s of extSheetsList) {
          if (isPropertyAndTabMatch(userProperty, s.getName())) {
            targetPropertySheet = s;
            break;
          }
        }

        if (!targetPropertySheet) {
          ui.alert(
            `🚫 INVALID PROPERTY (TAB NOT FOUND)\n\n` +
            `Property: "${userProperty}"\n` +
            `Tab in Template: [${tabName}]\n\n` +
            `No matching Tab Name for "${userProperty}" was found in the master validation file.\n\n` +
            `Submission is blocked.`
          );
          return;
        }

        propSheetName = targetPropertySheet.getName();

        if (!propertyRecordsCache[propSheetName]) {
          const extLastRow = targetPropertySheet.getLastRow();
          if (extLastRow < 3) {
            ui.alert(
              `🚫 EMPTY MASTER DIRECTORY\n\n` +
              `The master sheet for Tab "${propSheetName}" contains no tenant records (Row 3+ is empty).`
            );
            return;
          }

          const extHeaders = targetPropertySheet.getRange(2, 1, 1, targetPropertySheet.getLastColumn()).getValues()[0];
          let tenantColIdx = -1;
          let codeColIdx = -1;
          let propColIdx = -1;

          for (let c = 0; c < extHeaders.length; c++) {
            const hClean = superClean(extHeaders[c]);
            if (hClean === "property" || hClean.includes("property")) propColIdx = c;
            if (hClean === "tenant name" || hClean.includes("tenant name")) tenantColIdx = c;
            if (hClean === "customer code" || hClean.includes("customer code")) codeColIdx = c;
          }

          if (tenantColIdx === -1 || codeColIdx === -1) {
            ui.alert(
              `🚫 MASTER HEADER CONFIGURATION ERROR\n\n` +
              `Master Tab: [${propSheetName}]\n\n` +
              `Could not locate Row 2 headers for 'TENANT NAME' and/or 'CUSTOMER CODE'.\n` +
              `Please ensure Row 2 contains these exact column names.`
            );
            return;
          }

          const extRawData = targetPropertySheet.getRange(3, 1, extLastRow - 2, targetPropertySheet.getLastColumn()).getValues();
          const records = [];

          for (let r = 0; r < extRawData.length; r++) {
            const row = extRawData[r];
            const tVal = String(row[tenantColIdx] || "").trim();
            const cVal = String(row[codeColIdx] || "").trim();
            const pVal = (propColIdx !== -1) ? String(row[propColIdx] || "").trim() : propSheetName;
            
            if (tVal || cVal) {
              records.push({
                property: pVal,
                tenantName: tVal,
                customerCode: cVal
              });
            }
          }
          propertyRecordsCache[propSheetName] = records;
        }
      }

      const matchingPropRecords = propertyRecordsCache[propSheetName] || [];

      // Ensure range spans past Column AL (Column 38)
      const fetchCols = Math.max(currentSheet.getLastColumn(), 38);
      const tabHeaders = currentSheet.getRange(CONFIG.headerRow, 1, 1, fetchCols).getValues()[0];
      
      let partnerColIdx = -1;
      let codeColIdx = -1;

      for (let c = 0; c < tabHeaders.length; c++) {
        const hClean = superClean(tabHeaders[c]);
        if (hClean.includes("retail partner") || hClean.includes("tenant name")) partnerColIdx = c;
        if (hClean.includes("tenant code") || hClean.includes("customer code")) codeColIdx = c;
      }
      if (partnerColIdx === -1) partnerColIdx = colToIdx("E");

      let dataRange = currentSheet.getRange(startRow, 1, lastRow - startRow + 1, fetchCols).getValues();
      let displayRange = currentSheet.getRange(startRow, 1, lastRow - startRow + 1, fetchCols).getDisplayValues();

      const alIdx = colToIdx("AL");
      const eIdx = colToIdx("E");

      for (let i = 0; i < dataRange.length; i++) {
        let rowData = dataRange[i];
        let valA = String(rowData[0] || "").toLowerCase().trim();

        if (valA.includes("total") && !valA.includes("sub")) break;
        if (valA.includes("subtotal") || valA.includes("sub-total") || valA.includes("sub total")) continue;

        // Robust reader: Reads raw data or displayed formula/formatted value
        const getVal = (colIndex) => {
          if (colIndex < 0 || colIndex >= rowData.length) return "";
          const raw = rowData[colIndex];
          const disp = displayRange[i] ? displayRange[i][colIndex] : "";
          if (raw !== undefined && raw !== null && String(raw).trim() !== "") return String(raw).trim();
          if (disp !== undefined && disp !== null && String(disp).trim() !== "") return String(disp).trim();
          return "";
        };

        let valE = getVal(partnerColIdx !== -1 ? partnerColIdx : eIdx);

        if (valE !== "") {
          const valAD = String(rowData[colToIdx("AD")] || "").trim();
          if (valAD.toUpperCase() === "MONITORING") {
            const rawP = rowData[colToIdx("P")];
            const isZeroP = (rawP === 0 || String(rawP).trim() === "0" || (!isNaN(rawP) && Number(rawP) === 0 && String(rawP).trim() !== ""));
            
            // Block submission if MONITORING is used with a 0 value in Column P
            if (isZeroP) {
              ui.alert(
                `🚫 INVALID ENTRY\n\n` +
                `Tab: [${tabName}]\n` +
                `Row: ${i + startRow}\n` +
                `Column: P / AD\n\n` +
                `User cannot use "MONITORING" in Remarks if the value in Column P is 0.`
              );
              return;
            }
            continue;
          }

          // Check if Column J is THEORETICAL
          const rawValJ = rowData[colToIdx("J")];
          const valJ = (rawValJ === undefined || rawValJ === null) ? "" : String(rawValJ).trim().toUpperCase();
          const isRowTheo = (valJ === "THEORETICAL");

          // Dynamically adjust required columns for THEORETICAL
          let requiredCols = tabValidationMaps[tabName].slice();
          if (isRowTheo) {
            // Do not require Column L for Elec, Water, and LPG
            requiredCols = requiredCols.filter(col => col !== "L");

            if (tabName === "Elec") {
              // Elec: Column P is required
              if (!requiredCols.includes("P")) requiredCols.push("P");
            } else if (tabName === "Water" || tabName === "LPG") {
              // Water and LPG: Column O is required, do not require L-dependent columns
              if (!requiredCols.includes("O")) requiredCols.push("O");
              requiredCols = requiredCols.filter(col => !["P", "S", "T", "U", "V", "W", "X", "M", "N"].includes(col));
            }
          }

          // --- EXTERNAL TENANT & CODE CHECK (SKIPPED IF BYPASSED) ---
          if (!isPropertyBypassed) {
            const userRetailPartner = valE;

            // Strictly read Tenant Code from Column AL (or dynamic TENANT CODE header)
            let userTenantCode = getVal(alIdx);
            if (!userTenantCode && codeColIdx !== -1) {
              userTenantCode = getVal(codeColIdx);
            }

            let basePartner = userRetailPartner;
            if (userRetailPartner.includes('_')) {
              const parts = userRetailPartner.split('_');
              if (parts[parts.length - 1].trim().toLowerCase() === "affiliates") {
                basePartner = parts.slice(0, -1).join('_').trim();
              }
            }

            const cleanPartner = superClean(userRetailPartner);
            const cleanBasePartner = superClean(basePartner);
            const cleanCode = superClean(userTenantCode);

            const partnerExists = matchingPropRecords.some(rec => {
              const cRecTenant = superClean(rec.tenantName);
              return cRecTenant && (cRecTenant === cleanPartner || cRecTenant === cleanBasePartner);
            });

            if (!partnerExists) {
              ui.alert(
                `🚫 INVALID RETAIL PARTNER\n\n` +
                `Tab: [${tabName}]\n` +
                `Row: ${i + startRow}\n\n` +
                `RETAIL PARTNER "${userRetailPartner}" does not exist in master Tab "${propSheetName}" (TENANT NAME).\n\n` +
                `Submission is blocked.`
              );
              return;
            }

            if (!userTenantCode) {
              ui.alert(
                `🚫 MISSING TENANT CODE\n\n` +
                `Tab: [${tabName}]\n` +
                `Row: ${i + startRow} (Column AL)\n\n` +
                `TENANT CODE is blank for "${userRetailPartner}".\n\n` +
                `Please input the tenant code in Column AL (e.g. MREI028).`
              );
              return;
            }

            const codeExists = matchingPropRecords.some(rec => {
              const cRecCode = superClean(rec.customerCode);
              return cRecCode && cRecCode === cleanCode;
            });

            if (!codeExists) {
              ui.alert(
                `🚫 INVALID TENANT CODE\n\n` +
                `Tab: [${tabName}]\n` +
                `Row: ${i + startRow}\n\n` +
                `TENANT CODE "${userTenantCode}" does not exist in master Tab "${propSheetName}" (CUSTOMER CODE).\n\n` +
                `Submission is blocked.`
              );
              return;
            }

            const pairMatches = matchingPropRecords.some(rec => {
              const cRecTenant = superClean(rec.tenantName);
              const cRecCode = superClean(rec.customerCode);
              const cRecProp = superClean(rec.property);
              
              const tMatch = cRecTenant && (cRecTenant === cleanPartner || cRecTenant === cleanBasePartner);
              const cMatch = cRecCode && cRecCode === cleanCode;
              
              // If the master tab has a PROPERTY column, verify it matches the selected Property
              const pMatch = !rec.property || !cleanUserProp || cRecProp === cleanUserProp || (getSpecialPropertyTabName(userProperty) !== null);

              return tMatch && cMatch && pMatch;
            });

            if (!pairMatches) {
              ui.alert(
                `🚫 TENANT BINDING MISMATCH\n\n` +
                `Tab: [${tabName}]\n` +
                `Row: ${i + startRow}\n\n` +
                `RETAIL PARTNER "${userRetailPartner}" and TENANT CODE "${userTenantCode}" do not correspond to the same row for Property "${userProperty}" in master Tab "${propSheetName}".\n\n` +
                `Submission is blocked.`
              );
              return;
            }
          }

          // --- LPG CONVERSION FACTOR CHECK ---
          if (tabName === "LPG") {
            const valN10 = Number(currentSheet.getRange("N10").getValue());
            const valFVal = String(rowData[colToIdx("F")] || "").trim().toUpperCase();
            const valMVal = Number(rowData[colToIdx("M")]);
            
            if (valFVal === "KA" && valMVal !== valN10) {
              ui.alert(
                `🚫 INVALID CONVERSION FACTOR\n\n` +
                `Tab: [LPG]\n` +
                `Row: ${i + startRow}\n\n` +
                `Tenant is marked as KA, but the "Conversion Factor CBM to KG" in Column M (${valMVal}) is different from the standard "Conversion Factor" in N10 (${valN10}).\n\n` +
                `Please change Column F to "KA&SR" to allow a different Conversion Factor.`
              );
              return;
            }
          }

          for (let colLetter of requiredCols) {
            let colIdx = colToIdx(colLetter);
            let rawCellVal = rowData[colIdx];
            let cellValue = getVal(colIdx);
            let visibleCellVal = (displayRange[i] && displayRange[i][colIdx] ? displayRange[i][colIdx] : "").trim();


            if (cellValue === "") {
              ui.alert(
                `🚫 INCOMPLETE DATA\n\n` +
                `Tab: [${tabName}]\n` +
                `Row: ${i + startRow}\n` +
                `Column: ${colLetter}\n\n` +
                `Required field is blank.`
              );
              return;
            }

            if (colLetter === "O" && (cellValue === "0" || cellValue === "0.00" || rawCellVal === 0)) {
              ui.alert(
                `🚫 INVALID DATA\n\n` +
                `Tab: [${tabName}]\n` +
                `Row: ${i + startRow}\n` +
                `Column: O\n\n` +
                `Value cannot be exactly 0 (zero) when Column E has data.`
              );
              return;
            }

            // 3. Checker for Col L: Cannot be negative (if provided)
            if (colLetter === "L" && cellValue !== "") {
              let numVal = Number(cellValue);
              if (!isNaN(numVal) && numVal < 0) {
                ui.alert(
                  `🚫 INVALID CONSUMPTION\n\n` +
                  `Tab: [${tabName}]\n` +
                  `Row: ${i + startRow}\n` +
                  `Column: L\n\n` +
                  `Value (${cellValue}) cannot be a negative number.\n` +
                  `Please check if the Current Reading is lower than the Previous Reading.`
                );
                return;
              }
            }

            if (colLetter === "P") {
              if (visibleCellVal.includes("%")) continue;

              const adIdx = colToIdx("AD");
              const rawAD = rowData[adIdx];
              const valAD = (rawAD === undefined || rawAD === null) ? "" : String(rawAD).trim().toLowerCase();
              
              const allowedRemarks = [
                "defective meter",
                "inaccessible meter",
                "inaccesible meter",
                "very minimal usage",
                "minimal usage"
              ];
              const hasAllowedRemark = allowedRemarks.some(kw => valAD.includes(kw));

              let numVal = Number(cellValue);
              const isZeroValue = (cellValue === "0" || (!isNaN(numVal) && numVal === 0 && cellValue !== ""));

              // Bypass if 0 and Remarks match approved reasons
              if (isZeroValue && hasAllowedRemark) continue;

              if (isNaN(numVal) || numVal === 0) {
                ui.alert(
                  `🚫 INVALID ENTRY\n\n` +
                  `Tab: [${tabName}]\n` +
                  `Row: ${i + startRow}\n` +
                  `Column: P\n\n` +
                  `Make sure the value is not equal to 0 or set as percentage (%)`
                );
                return;
              }
            }

            if (colLetter === "F") {
              const valFCheck = cellValue.toUpperCase();
              const allowedFValuesSubmit = ["KA", "SR", "REG", "KA&SR"];
              if (!allowedFValuesSubmit.includes(valFCheck)) {
                ui.alert(
                  `🚫 INVALID CATEGORY\n\n` +
                  `Tab: [${tabName}]\n` +
                  `Row: ${i + startRow}\n` +
                  `Column: F\n\n` +
                  `Value ("${cellValue}") must be exactly "KA", "SR", "REG", or "KA&SR".`
                );
                return;
              }
            }

          }
        }
      }
    }

    // ===============================================
    // 4. DATA EXTRACTION AND SUBMISSION
    // ===============================================
    let activeSheet = ss.getActiveSheet();
    let headerSheet = tabValidationMaps[activeSheet.getName()] ? activeSheet :
      Object.keys(tabValidationMaps).map(n => ss.getSheetByName(n)).find(s => s !== null);

    if (!headerSheet) {
      ui.alert("🚫 Error: No utility tabs (Elec, Water, LPG) found.");
      return;
    }

    const extractedData = processSheetHeaders(headerSheet);
    if (!extractedData) return;

    const props = PropertiesService.getScriptProperties();
    let activeDB_ID = props.getProperty("ACTIVE_DB_ID") || masterDbId;
    let db = SpreadsheetApp.openById(activeDB_ID);
    let dSh = db.getSheetByName("PBTT Submission");

    const timestamp = Utilities.formatDate(new Date(), "Asia/Manila", "MMM d, yyyy hh:mm a");
    const userEmail = Session.getActiveUser().getEmail();
    const activeFileName = ss.getName();
    const ssUrl = ss.getUrl();

    const finalRow = [timestamp, ...extractedData, activeFileName, ssUrl, userEmail, currentRef];

    const dbData = dSh.getDataRange().getValues();
    let rowIndexToOverwrite = -1;
    const refColumnIndex = finalRow.length - 1;

    for (let r = 1; r < dbData.length; r++) {
      let existingRef = String(dbData[r][refColumnIndex] || "").trim();
      if (existingRef === currentRef) {
        rowIndexToOverwrite = r + 1;
        break;
      }
    }

    if (rowIndexToOverwrite > 0) {
      dSh.getRange(rowIndexToOverwrite, 1, 1, finalRow.length).setValues([finalRow]);
      ui.alert(`✅ SUCCESS: Submission updated (Overwrite existing Ref# ${currentRef}).`);
    } else {
      dSh.appendRow(finalRow);
      ui.alert(`✅ SUCCESS: New submission recorded successfully.`);
    }

  } catch (x) {
    SpreadsheetApp.getUi().alert("System Error: " + x.message);
  } finally {
    lock.releaseLock();
  }
}

/* =================================
9. INITIALIZATION & DATA SYNC
================================= */
function INITIALIZE_SYSTEM_BUTTON() {
  if (isMasterFileBlocked()) return;

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  ss.toast("🔄 Syncing databases and checking Ref#...", "System Status", 3);

  try {
    syncDataAcrossFiles();
    syncKAData();
    generateUniqueAlphanumericRef();
  } catch (e) {
    ui.alert("❌ Error during initialization: " + e.message);
  }
}

function INSTALL_SYSTEM() {
  if (isMasterFileBlocked()) return;

  const ui = SpreadsheetApp.getUi();
  
  try {
    syncDataAcrossFiles();           
    syncKAData();

    generateUniqueAlphanumericRef();

    const functionName = 'runStartupSequence';
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => { 
      if (t.getHandlerFunction() === functionName) {
        ScriptApp.deleteTrigger(t); 
      }
    });

    ui.alert("🚀 UPDATE COMPLETE\n\n- Tabs 'dvPeriod', 'dvGen', and 'KA_DATA' overwritten.\n- Ref# is secured in cell C7.");
    
  } catch (e) {
    ui.alert("❌ Action Failed: " + e.message);
  }
}

function runStartupSequence() {
  syncDataAcrossFiles();
  syncKAData(); 
  generateUniqueAlphanumericRef();
}

function syncKAData() {
  const sourceId = "1jY-9FMha3x972o4Gz1d6DVD36d3ppjHW_WM1DHJz6ag";
  const targetSS = SpreadsheetApp.getActiveSpreadsheet(); 

  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(15000)) throw new Error("Database is busy during KA_DATA update."); 

    const sourceSS = SpreadsheetApp.openById(sourceId);
    const sourceSheet = sourceSS.getSheetByName("Data");
    if (!sourceSheet) return;

    let targetSheet = targetSS.getSheetByName("KA_DATA");
    if (!targetSheet) {
      targetSheet = targetSS.insertSheet("KA_DATA");
    }

    const sourceRange = sourceSheet.getDataRange();
    const sourceData = sourceRange.getValues();
    const sourceFormats = sourceRange.getNumberFormats(); 
    const numRows = sourceData.length;
    
    if (numRows > 0) {
      const numCols = sourceData[0].length;
      targetSheet.clear(); 
      
      if (targetSheet.getMaxRows() < numRows) {
        targetSheet.insertRowsAfter(targetSheet.getMaxRows(), numRows - targetSheet.getMaxRows());
      }
      if (targetSheet.getMaxColumns() < numCols) {
        targetSheet.insertColumnsAfter(targetSheet.getMaxColumns(), numCols - targetSheet.getMaxColumns());
      }

      const targetRange = targetSheet.getRange(1, 1, numRows, numCols);
      targetRange.setValues(sourceData);
      targetRange.setNumberFormats(sourceFormats);
    }
    SpreadsheetApp.flush();
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function syncDataAcrossFiles() {
  const sourceId = "1hMMUd4ho50HP63dc2fRAo--iK-m7YotamkKtsDGT_Us";
  const targetSS = SpreadsheetApp.getActiveSpreadsheet(); 
  const tabsToSync = ["dvPeriod", "dvGen"]; 
  const timeZone = Session.getScriptTimeZone();

  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(15000)) throw new Error("Database is busy."); 

    const sourceSS = SpreadsheetApp.openById(sourceId);

    tabsToSync.forEach(tabName => {
      const sourceSheet = sourceSS.getSheetByName(tabName);
      const targetSheet = targetSS.getSheetByName(tabName);

      if (sourceSheet && targetSheet) {
        let sourceRange = sourceSheet.getDataRange();
        let sourceData = sourceRange.getValues();
        const numRows = sourceData.length;
        const numCols = sourceData[0].length;
        
        if (numRows > 0) {
          if (numRows > 1) { 
            for (let i = 1; i < numRows; i++) {
              if (sourceData[i][0] instanceof Date) sourceData[i][0] = Utilities.formatDate(sourceData[i][0], timeZone, "MMM d, yyyy");
              if (tabName === "dvPeriod") {
                if (sourceData[i][1] instanceof Date) sourceData[i][1] = Utilities.formatDate(sourceData[i][1], timeZone, "MMM d, yyyy");
                if (sourceData[i][2] instanceof Date) sourceData[i][2] = Utilities.formatDate(sourceData[i][2], timeZone, "MMM d, yyyy");
              }
            }
          }

          targetSheet.clear(); 
          
          if (targetSheet.getMaxRows() < numRows) targetSheet.insertRowsAfter(targetSheet.getMaxRows(), numRows - targetSheet.getMaxRows());
          if (targetSheet.getMaxColumns() < numCols) targetSheet.insertColumnsAfter(targetSheet.getMaxColumns(), numCols - targetSheet.getMaxColumns());

          targetSheet.getRange(1, 1, numRows, numCols).setValues(sourceData);
        }
      }
    });
    SpreadsheetApp.flush();
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function generateUniqueAlphanumericRef() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const instSheet = ss.getSheetByName("Instructions");
  if (!instSheet) return;

  const cell = instSheet.getRange("C7");
  const existingRef = cell.getValue().toString().trim();
  
  if (existingRef !== "" && existingRef !== null) {
    console.log("Ref# already exists. Generation skipped to prevent overwrite.");
    return; 
  }

  const masterId = "1hMMUd4ho50HP63dc2fRAo--iK-m7YotamkKtsDGT_Us";
  try {
    const masterSS = SpreadsheetApp.openById(masterId);
    const dbSheet = masterSS.getSheetByName("PBTT Submission");
    const lastRow = dbSheet.getLastRow();
    
    let existingRefs = new Set();
    if (lastRow >= 5) {
      const data = dbSheet.getRange(5, 11, lastRow - 4, 1).getValues();
      existingRefs = new Set(data.flat().map(v => String(v).trim()));
    }

    const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
    let newRef = "";
    let isUnique = false;

    while (!isUnique) {
      let result = "";
      for (let i = 0; i < 6; i++) result += chars.charAt(Math.floor(Math.random() * chars.length));
      newRef = "Ref#" + result;
      if (!existingRefs.has(newRef)) isUnique = true;
    }
    
    if (cell.getValue().toString().trim() === "") {
       cell.setValue(newRef);
    }
    
  } catch (e) {
    console.error("Ref# Gen Error: " + e.toString());
  }
}

function processSheetHeaders(sheet) {
  const sourceFileUrl = sheet.getRange("A1").getValue().toString().trim();
  const textToCheck = sourceFileUrl.toUpperCase();

  if (
    sourceFileUrl === "" ||
    !(sourceFileUrl.toLowerCase().includes("http") || textToCheck === "N/A" || textToCheck === "NA")
  ) {
    SpreadsheetApp.getUi().alert(
      "🚫 SUBMISSION BLOCKED\n\n" +
      "A valid SOURCE FILE URL is missing in cell C20 in Instruction Tab or A1 in Utilities Tab.\n" +
      "Please provide a valid URL, or enter 'N/A' | 'NA' before submitting."
    );
    return null;
  }

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

  const startDateValue = sheet.getRange("E7").getValue();
  const endDateValue = sheet.getRange("E8").getValue();

  if (!(startDateValue instanceof Date) || isNaN(startDateValue) ||
    !(endDateValue instanceof Date) || isNaN(endDateValue)) {
    SpreadsheetApp.getUi().alert("❌ ERROR: Start Date or End Date is empty or invalid.");
    return null;
  }

  const startDate = new Date(startDateValue);
  const endDate = new Date(endDateValue);
  const today = new Date();

  if (endDate <= startDate) {
    SpreadsheetApp.getUi().alert("❌ DATE ERROR: End Date (E8) must be after Start Date (E7).");
    return null;
  }

  const currentMonth = today.getMonth();
  const currentYear = today.getFullYear();
  const endMonth = endDate.getMonth();
  const endYear = endDate.getFullYear();

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
    if (response !== ui.Button.YES) return null;
  }

  for (let item of config) {
    let finalVal = null;

    if (item.cell === "E11") {
      let rawItems = [];
      const tabsToCheck = ["Elec", "Water", "LPG"];
      const ss = sheet.getParent();

      tabsToCheck.forEach(tabName => {
        let utilSheet = ss.getSheetByName(tabName);
        if (!utilSheet) return;

        let e11Val = utilSheet.getRange("E11").getValue();
        if (e11Val) {
          e11Val.toString().split(",").forEach(v => rawItems.push(v.trim()));
        }

        let lastR = utilSheet.getLastRow();
        if (lastR >= 13) {
          let colA = utilSheet.getRange("A13:A" + lastR).getValues().flat();
          let colY = utilSheet.getRange("Y13:Y" + lastR).getValues().flat();

          for (let i = 0; i < colA.length; i++) {
            let aVal = colA[i] ? colA[i].toString().trim().toUpperCase() : "";
            if (aVal.includes("TOTAL") && !aVal.includes("SUB")) break;

            let yVal = colY[i];
            if (yVal && yVal.toString().trim() !== "") {
              yVal.toString().split(",").forEach(v => rawItems.push(v.trim()));
            }
          }
        }
      });
      let uniqueItems = Array.from(new Set(rawItems)).filter(Boolean);
      finalVal = uniqueItems.join(", ");
    }
    else if (item.sourceRange) {
      let rawItems = [];
      let mainVal = sheet.getRange(item.cell).getValue();

      if (mainVal) {
        mainVal.toString().split(",").forEach(v => rawItems.push(v.trim()));
      }

      let colLetter = item.sourceRange.substring(0, 1);
      let lastR = sheet.getLastRow();
      if (lastR >= 13) {
        let colA = sheet.getRange("A13:A" + lastR).getValues().flat();
        let colSource = sheet.getRange(colLetter + "13:" + colLetter + lastR).getValues().flat();

        for (let i = 0; i < colA.length; i++) {
          let aVal = colA[i] ? colA[i].toString().trim().toUpperCase() : "";
          if (aVal.includes("TOTAL") && !aVal.includes("SUB")) break;

          let sVal = colSource[i];
          if (sVal && sVal.toString().trim() !== "") {
            sVal.toString().split(",").forEach(v => rawItems.push(v.trim()));
          }
        }
      }
      let uniqueItems = Array.from(new Set(rawItems)).filter(Boolean);
      finalVal = uniqueItems.join(", ");
    }
    else {
      finalVal = sheet.getRange(item.cell).getValue();
      if (item.type === "date" && finalVal) {
        finalVal = Utilities.formatDate(new Date(finalVal), Session.getScriptTimeZone(), "MMM d, yyyy");
      }
    }

    if ((finalVal === "" || finalVal === undefined || finalVal === null) && finalVal !== 0) {
      missing.push(item.label);
    }

    results.push(finalVal);
  }

  if (missing.length > 0) {
    SpreadsheetApp.getUi().alert("🚫 MISSING HEADER INFO:\n\n" + missing.join("\n"));
    return null;
  }

  return results;
}

function getTotalCellCount(ss) {
  let total = 0;
  const sheets = ss.getSheets();
  sheets.forEach(sh => {
    total += (sh.getMaxRows() * sh.getMaxColumns());
  });
  return total;
}

function rotateToNewDatabase(oldDb, oldSheet) {
  const folder = DriveApp.getFolderById(BACKUP_FOLDER_ID);
  const time = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd_HHmmss");
  const newName = "PBTT_Submission_Database_" + time;

  const newFile = SpreadsheetApp.create(newName);
  const newFileId = newFile.getId();

  const driveFile = DriveApp.getFileById(newFileId);
  folder.addFile(driveFile);
  DriveApp.getRootFolder().removeFile(driveFile);

  const targetSheetName = "PBTT Submission";
  const newSheet = newFile.insertSheet(targetSheetName);

  try {
    const masterSS = SpreadsheetApp.openById(PBTT_DB_ID);
    const masterSheet = masterSS.getSheetByName(targetSheetName);

    const headerWidth = Math.max(masterSheet.getLastColumn(), 15);
    const masterHeaderRange = masterSheet.getRange(4, 1, 1, headerWidth);
    const targetRange = newSheet.getRange(4, 1, 1, headerWidth);

    const headerValues = masterHeaderRange.getValues();
    targetRange.setValues(headerValues);

    targetRange.setBackgrounds(masterHeaderRange.getBackgrounds());
    targetRange.setFontColors(masterHeaderRange.getFontColors());
    targetRange.setFontWeights(masterHeaderRange.getFontWeights());
    targetRange.setHorizontalAlignments(masterHeaderRange.getHorizontalAlignments());

    console.log("Successfully copied Row 4 header from Master ID to Row 4 of new file.");
  } catch (e) {
    console.error("Could not fetch master header: " + e.message);
    const fallbackWidth = oldSheet.getLastColumn() || 15;
    const vals = oldSheet.getRange(4, 1, 1, fallbackWidth).getValues();
    newSheet.getRange(4, 1, 1, fallbackWidth).setValues(vals);
  }

  const defaultSheet = newFile.getSheetByName("Sheet1");
  if (defaultSheet) newFile.deleteSheet(defaultSheet);

  try {
    const regSs = SpreadsheetApp.openById(BACKUP_REGISTRY_ID);
    const regSh = regSs.getSheetByName("Backup Files") || regSs.insertSheet("Backup Files");
    regSh.appendRow([new Date(), "NEW ACTIVE DB: " + newName, newFile.getUrl()]);
  } catch (e) {
    console.warn("Registry update failed, but file was rotated.");
  }

  PropertiesService.getScriptProperties().setProperty("ACTIVE_DB_ID", newFileId);
  return newFileId;
}

function fullResetDatabasePointer() {
  if (isMasterFileBlocked()) return;

  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("ACTIVE_DB_ID");
  SpreadsheetApp.getUi().alert("Reset successful. The script is now looking at the original MASTER file again.");
}

function checkCurrentDbSize() {
  if (isMasterFileBlocked()) return;

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

/**
 * Fetches the master validation spreadsheet and its property Tab Names.
 * File ID: 12OOOzMVeWPb6SKJyNu3tewSPKrbu3s93jJA3SmPNSY4
 */
function getExternalValidationData() {
  const extId = "12OOOzMVeWPb6SKJyNu3tewSPKrbu3s93jJA3SmPNSY4";
  try {
    const ss = SpreadsheetApp.openById(extId);
    const sheets = ss.getSheets();
    const propertyTabs = sheets.map(s => s.getName());
    return { ss, propertyTabs, cache: {} };
  } catch (e) {
    console.error("Failed to fetch external validation data: " + e.message);
    return null;
  }
}





/**
 * MASTER FUNCTION: INITIALIZE
 * For manual syncing and checking Ref#.
 */
function INITIALIZE_SYSTEM_BUTTON() {
  if (isMasterFileBlocked()) return;

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  ss.toast("🔄 Syncing databases and checking Ref#...", "System Status", 3);

  try {
    syncDataAcrossFiles();           // Existing DB Sync
    syncKAData();                    // NEW: Gets data & format for KA_DATA
    generateUniqueAlphanumericRef(); // This function will exit if C7 is already filled.
    //ui.alert("✅ SYSTEM INITIALIZED\n\nDatabase Tabs updated and Ref# verified.");
  } catch (e) {
    ui.alert("❌ Error during initialization: " + e.message);
  }
}

/**
 * INSTALLER FUNCTION
 * This performs the one-time setup and immediate data sync.
 */
function INSTALL_SYSTEM() {
  if (isMasterFileBlocked()) return;

  const ui = SpreadsheetApp.getUi();
  
  try {
    // 1. Force immediate data overwrite for data-only tabs
    syncDataAcrossFiles();           
    syncKAData(); // Appended fetch into one-time install    

    // 2. Try to generate Ref#. If C7 is already filled, this does nothing.
    generateUniqueAlphanumericRef();

    // 3. REMOVE the Trigger (so it no longer runs on open)
    const functionName = 'runStartupSequence';
    const triggers = ScriptApp.getProjectTriggers();
    // This loops through and deletes the trigger
    triggers.forEach(t => { 
      if (t.getHandlerFunction() === functionName) {
        ScriptApp.deleteTrigger(t); 
      }
    });

    ui.alert("🚀 UPDATE COMPLETE\n\n- Tabs 'dvPeriod', 'dvGen', and 'KA_DATA' overwritten.\n- Ref# is secured in cell C7.");
    
  } catch (e) {
    ui.alert("❌ Action Failed: " + e.message);
  }
}
/**
 * STARTUP SEQUENCE
 */
function runStartupSequence() {
  syncDataAcrossFiles();
  syncKAData(); 
  generateUniqueAlphanumericRef();
}

/**
 * NEW FUNCTION: SYNC DATA & FORMATTING TO "KA_DATA"
 * Grabs ALL values and number formatting (dates, %, etc) and overwrites "KA_DATA"
 */
function syncKAData() {
  const sourceId = "1jY-9FMha3x972o4Gz1d6DVD36d3ppjHW_WM1DHJz6ag";
  const targetSS = SpreadsheetApp.getActiveSpreadsheet(); 

  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(15000)) throw new Error("Database is busy during KA_DATA update."); 

    // Open Source Spreadsheet & Tab
    const sourceSS = SpreadsheetApp.openById(sourceId);
    const sourceSheet = sourceSS.getSheetByName("Data");
    
    // Safety check - Ensures Source Data Tab actually exists
    if (!sourceSheet) return;

    // Connect to KA_DATA Tab (or auto-create if missing)
    let targetSheet = targetSS.getSheetByName("KA_DATA");
    if (!targetSheet) {
      targetSheet = targetSS.insertSheet("KA_DATA");
    }

    const sourceRange = sourceSheet.getDataRange();
    
    // FETCH BOTH VALUES AND FORMATS (Percentage, Dates, Decimals, etc.)
    const sourceData = sourceRange.getValues();
    const sourceFormats = sourceRange.getNumberFormats(); 
    
    const numRows = sourceData.length;
    
    if (numRows > 0) {
      const numCols = sourceData[0].length;
      
      // Wipe target before rewrite
      targetSheet.clear(); 
      
      // Expand tab capacity exactly to the Source's limits
      if (targetSheet.getMaxRows() < numRows) {
        targetSheet.insertRowsAfter(targetSheet.getMaxRows(), numRows - targetSheet.getMaxRows());
      }
      if (targetSheet.getMaxColumns() < numCols) {
        targetSheet.insertColumnsAfter(targetSheet.getMaxColumns(), numCols - targetSheet.getMaxColumns());
      }

      // Feed full dataset cleanly over into target, THEN apply the format rules
      const targetRange = targetSheet.getRange(1, 1, numRows, numCols);
      targetRange.setValues(sourceData);            // Copies raw numbers/strings
      targetRange.setNumberFormats(sourceFormats);  // Applies the %, Date, Text formulas
    }
    SpreadsheetApp.flush();
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

/**
 * 1. SYNC DATA ACROSS FILES (RESTRICTED TO DATA TABS ONLY)
 * Clears and overwrites specific tabs. Does NOT touch the 'Instructions' sheet.
 */
function syncDataAcrossFiles() {
  const sourceId = "1hMMUd4ho50HP63dc2fRAo--iK-m7YotamkKtsDGT_Us";
  const targetSS = SpreadsheetApp.getActiveSpreadsheet(); 
  
  const tabsToSync = ["dvPeriod", "dvGen"]; 
  const timeZone = Session.getScriptTimeZone();

  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(15000)) throw new Error("Database is busy."); 

    const sourceSS = SpreadsheetApp.openById(sourceId);

    tabsToSync.forEach(tabName => {
      const sourceSheet = sourceSS.getSheetByName(tabName);
      const targetSheet = targetSS.getSheetByName(tabName);

      if (sourceSheet && targetSheet) {
        let sourceRange = sourceSheet.getDataRange();
        let sourceData = sourceRange.getValues();
        const numRows = sourceData.length;
        const numCols = sourceData[0].length;
        
        if (numRows > 0) {
          // Legacy Date Formatting mapped here manually
          if (numRows > 1) { 
            for (let i = 1; i < numRows; i++) {
              if (sourceData[i][0] instanceof Date) sourceData[i][0] = Utilities.formatDate(sourceData[i][0], timeZone, "MMM d, yyyy");
              if (tabName === "dvPeriod") {
                if (sourceData[i][1] instanceof Date) sourceData[i][1] = Utilities.formatDate(sourceData[i][1], timeZone, "MMM d, yyyy");
                if (sourceData[i][2] instanceof Date) sourceData[i][2] = Utilities.formatDate(sourceData[i][2], timeZone, "MMM d, yyyy");
              }
            }
          }

          targetSheet.clear(); 
          
          if (targetSheet.getMaxRows() < numRows) targetSheet.insertRowsAfter(targetSheet.getMaxRows(), numRows - targetSheet.getMaxRows());
          if (targetSheet.getMaxColumns() < numCols) targetSheet.insertColumnsAfter(targetSheet.getMaxColumns(), numCols - targetSheet.getMaxColumns());

          targetSheet.getRange(1, 1, numRows, numCols).setValues(sourceData);
        }
      }
    });
    SpreadsheetApp.flush();
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

/**
 * 2. GENERATE PERMANENT REF# (PERSISTENT LOGIC)
 */
function generateUniqueAlphanumericRef() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const instSheet = ss.getSheetByName("Instructions");
  if (!instSheet) return;

  const cell = instSheet.getRange("C7");
  const existingRef = cell.getValue().toString().trim();
  
  if (existingRef !== "" && existingRef !== null) {
    console.log("Ref# already exists. Generation skipped to prevent overwrite.");
    return; 
  }

  const masterId = "1hMMUd4ho50HP63dc2fRAo--iK-m7YotamkKtsDGT_Us";
  try {
    const masterSS = SpreadsheetApp.openById(masterId);
    const dbSheet = masterSS.getSheetByName("PBTT Submission");
    const lastRow = dbSheet.getLastRow();
    
    let existingRefs = new Set();
    if (lastRow >= 5) {
      const data = dbSheet.getRange(5, 11, lastRow - 4, 1).getValues();
      existingRefs = new Set(data.flat().map(v => String(v).trim()));
    }

    const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
    let newRef = "";
    let isUnique = false;

    while (!isUnique) {
      let result = "";
      for (let i = 0; i < 6; i++) result += chars.charAt(Math.floor(Math.random() * chars.length));
      newRef = "Ref#" + result;
      if (!existingRefs.has(newRef)) isUnique = true;
    }
    
    if (cell.getValue().toString().trim() === "") {
       cell.setValue(newRef);
    }
    
  } catch (e) {
    console.error("Ref# Gen Error: " + e.toString());
  }
}
