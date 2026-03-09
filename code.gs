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
    .addItem("🛠️ Setup", "INSTALL_SYSTEM")
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
    
    // Set boundaries for our rows based on your CONFIG and the sheet's current size
    const startRow = CONFIG.dataStartRow;
    const lastRow = Math.max(startRow, sheet.getLastRow()); 
    
    // Get all values for Column E to efficiently see which rows are populated
    // (Start row, column 5=E, number of rows down, 1 column wide)
    const colEData = sheet.getRange(startRow, 5, lastRow - startRow + 1, 1).getValues();
    
    let configsChecked = false; // Flag to ensure we don't redundantly check static config cells

    // Loop through every single row
    colEData.forEach((row, idx) => {
      let colEValue = row[0];
      let actualRow = startRow + idx;
      
      // ONLY trigger the requirement check if Column E HAS A VALUE
      if (colEValue !== "") {
        
        // 1. Check Configuration cells (We only need to alert about these once per sheet)
        if (!configsChecked) {
          req.config.forEach(c => { 
            if(sheet.getRange(c).getValue() === "") missing.push(`Cell ${c}`); 
          });
          configsChecked = true;
        }
        
        // 2. Check the specific column requirements for this exact row
        req.cols.forEach(col => { 
          // Reads: (e.g.) sheet.getRange("L" + 10).getValue()
          if(sheet.getRange(col + actualRow).getValue() === "") {
            // Log as L10, P12, etc., so the user knows exactly which row failed
            missing.push(`${col}${actualRow}`); 
          } 
        });
      }
    });
    
    // Group and add error alerts per tab if anything was found
    if (missing.length > 0) {
      stopErrors.push(`[${tabName}]: ${missing.join(", ")}`);
    }
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
  const ui = SpreadsheetApp.getUi();
  
  // ==========================================
  // 1. DATABASE VALIDATION (Check REF #)
  // ==========================================
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
          return; // STOP EXECUTION
        }
      }
    }
  } catch (err) {
    console.error("Database Validation Error: " + err.message);
    ui.alert("⚠️ Database connection warning: Could not verify Reference Status.");
  }

  // ==========================================
  // EXISTING FETCH LOGIC (With source linkage safeguards)
  // ==========================================
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

  // --- STAGE 2: PROCESSING ---
  if (sheet.getMaxColumns() < 29) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 29 - sheet.getMaxColumns());
  }

  const destWidth = sheet.getMaxColumns(); 
  const pasteArray = [];
  
  // Requires global variables: EXCLUSIONS, FETCH_MAPS, colToIdx externally defined!
  const skipIndices = (EXCLUSIONS[tabName] || []).map(letter => colToIdx(letter));
  const mapping = FETCH_MAPS[tabName] || {};

  let totalFound = false;
  let rowCounter = 0; 

  // --- 1. Identify the *Absolute Last* TOTAL row by looping from bottom to top ---
  let lastTotalIndex = -1;
  for (let i = rawData.length - 1; i >= 0; i--) {
    let checkVal = String(rawData[i][0] || "").trim().toLowerCase();
    let isSub = checkVal.includes("subtotal") || checkVal.includes("sub-total") || checkVal.includes("sub total");
    
    // Grabs final bottom match:
    if (checkVal.includes("total") && !isSub) {
      lastTotalIndex = i;
      break; 
    }
  }

  // --- 2. Iterate array pulling everything systematically into destRows mapping ---
  for (let i = 0; i < rawData.length; i++) {
    let sourceRow = rawData[i];
    let valA_source = String(sourceRow[0] || "").trim();
    let valA_lower = valA_source.toLowerCase();
    
    let isSubTotal = valA_lower.includes("subtotal") || valA_lower.includes("sub-total") || valA_lower.includes("sub total");
    let isTerminatingTotal = (i === lastTotalIndex); 
    let isIntermediateTotal = (!isTerminatingTotal && valA_lower.includes("total") && !isSubTotal);

    let rowHasData = sourceRow.some(cell => String(cell).trim() !== "");
    
    if (rowHasData || isSubTotal || isIntermediateTotal || isTerminatingTotal) {
      
      // Inject aesthetic space/gap above Final Terminating row
      if (isTerminatingTotal) {
        if (pasteArray.length > 0 && !pasteArray[pasteArray.length - 1].every(cell => cell === "")) {
          pasteArray.push(new Array(destWidth).fill(""));
        }
      }

      let destRow = new Array(destWidth).fill("");
      
      // Assing Label Name in Column 1 naturally
      if (isSubTotal || isIntermediateTotal || isTerminatingTotal) {
        destRow[0] = valA_source; 
      } else {
        rowCounter++;
        destRow[0] = rowCounter; 
      }

      // Loop directly through and filter
      for (let c = 1; c < sourceRow.length; c++) {
        if (skipIndices.includes(c)) continue;
        if (c < destWidth) destRow[c] = sourceRow[c];
      }

      // Explicit target mappings applying uniformly
      Object.keys(mapping).forEach(targetCol => {
        let sIdx = colToIdx(mapping[targetCol]);
        let tIdx = colToIdx(targetCol);
        if (sourceRow[sIdx] !== undefined) destRow[tIdx] = sourceRow[sIdx];
      });

      pasteArray.push(destRow);
    }

    // BREAK SCRIPT LOGIC & ATTACH FOOTER 
    // Wait until mapping is finished inside the array cleanly before stopping it and writing the footprint
    if (isTerminatingTotal) {
      totalFound = true;
      
      for (let s = 0; s < 3; s++) pasteArray.push(new Array(destWidth).fill("")); // Sign spaces

      let sigRow = new Array(destWidth).fill("");
      sigRow[0] = "Prepared By:"; sigRow[7] = "Checked By:"; sigRow[28] = "Noted By:";
      pasteArray.push(sigRow);
      break; 
    }
  }

  // Check fallback protocol incase data lacked standard totals universally 
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
  
  SpreadsheetApp.getActive().toast(`Fetch complete. Mapped Columns apply flawlessly to intermediate totals. Unused trailing data ignored.`, "Success", 5);
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

    // Read column E and A values first
    const valE = String(row[colToIdx("E")] || "").trim();
    const labelA = String(row[0] || "").trim();
    
    // --- NEW: STRICT COLUMN E SKIP LOGIC ---
    // If Column E is empty, skip to the next row immediately
    if (valE === "") continue;

    // Continue to ignore "Total" rows at the bottom
    const normalizedLabelA = labelA.toLowerCase().replace(/[^a-z]/g, "");
    if (normalizedLabelA.includes("total")) continue;

    // --- STEP 1: RUN CHECKLIST (Now safely checking ONLY valid rows) ---
    runCommonChecklist(row, rowNum, (r, c, m, res, arr) => logHelper(row, r, c, m, res, arr), valL5, valL6);

    // --- STEP 2: KA VALIDATION ---
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

    // --- STEP 3: TAB SPECIFIC CALCULATIONS ---
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
    // Wait for up to 30 seconds for other processes to finish.
    lock.waitLock(30000); 
  } catch (e) {
    SpreadsheetApp.getUi().alert("Server Busy. Please try again.");
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // ⚠️ ID of the Master Database
    const masterDbId = "1hMMUd4ho50HP63dc2fRAo--iK-m7YotamkKtsDGT_Us"; 

    // Define the specific columns to check for each tab
    const tabValidationMaps = {
      "Elec":  ["B","C","D","H","I","J","K","L","O","P","Q","Y","Z","AA","AB","AC","AG","AH","AJ","AK"],
      "Water": ["B","C","D","H","J","K","L","O","P","S","T","U","V","W","X","Y","Z","AA","AB","AC","AG","AH","AJ","AK"],
      "LPG":   ["B","C","D","H","J","K","L","M","N","O","P","Q","Y","Z","AA","AB","AC","AG","AH","AJ","AK"]
    };

    const ui = SpreadsheetApp.getUi();

    // ===============================================
    // 1. REF# & DATE SETUP
    // ===============================================
    const instSheet = ss.getSheetByName("Instructions");
    if (!instSheet) {
      ui.alert("❌ ERROR: 'Instructions' tab not found.");
      return;
    }

    // Ref Check (C7)
    const currentRef = instSheet.getRange("C7").getValue().toString().trim();
    if (currentRef === "") {
      ui.alert("❌ ERROR: No Reference Number found in 'Instructions' tab C7.");
      return;
    }

    // Get Target Dates (C26, C27)
    const rawTargetStart = instSheet.getRange("C26").getValue();
    const rawTargetEnd = instSheet.getRange("C27").getValue();

    // Helper to normalize dates (strip time) for accurate comparison
    const normalizeDate = (d) => {
      if (!d || !(d instanceof Date) || isNaN(d.getTime())) return null;
      const n = new Date(d);
      n.setHours(0, 0, 0, 0); // Reset time to midnight
      return n.getTime(); // Use numeric time for easy comparison
    };

    const targetStartInfo = normalizeDate(rawTargetStart);
    const targetEndInfo = normalizeDate(rawTargetEnd);

    if (!targetStartInfo || !targetEndInfo) {
      ui.alert("❌ ERROR: Invalid or missing billing dates in 'Instructions' tab (C26/C27).");
      return;
    }

    // ===============================================
    // 2. PERIOD STATUS VALIDATION (Match C26/C27 + Check Lock & Bypass)
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

      // Iterate DB to find the specific period in C26/C27
      for (let i = 1; i < periodData.length; i++) {
        const row = periodData[i];
        
        // Col A (Start) & Col B (End)
        const dbStart = normalizeDate(row[0]);
        const dbEnd = normalizeDate(row[1]);

        // MATCH FOUND?
        if (dbStart === targetStartInfo && dbEnd === targetEndInfo) {
          periodFound = true;
          
          // 1. Check Status (Col D)
          const status = String(row[3]).trim();
          if (status === "Active") {
            periodIsActive = true;
          }

          // 2. Check Lock Date (Col C) & Bypass (Col E)
          if (row[2]) {
             const lockDate = new Date(row[2]);
             lockDate.setHours(0,0,0,0);
             
             // Get Bypass Value from Column E (Index 4)
             const bypassTag = String(row[4] || "").trim();

             // Logic: If Today >= Lock Date
             if (today >= lockDate) {
               if (bypassTag === "Bypass") {
                 // ALLOWED: Bypass is active, ignore lock.
                 periodIsLocked = false; 
               } else {
                 // BLOCKED: Lock date met and NO Bypass tag.
                 periodIsLocked = true;
                 lockDateFormatted = Utilities.formatDate(lockDate, "Asia/Manila", "MMM d, yyyy");
               }
             }
          }
          break; // Stop looking once the matching period is found
        }
      }

      // -- VALIDATE FINDINGS --
      
      // Error 1: Dates don't match any row in DB
      if (!periodFound) {
        ui.alert(
          "🚫 CONFIGURATION ERROR\n\n" +
          "The dates in Instructions C26 & C27 do not match any known period in the master database.\n" +
          "Please verify your billing start/end dates."
        );
        return;
      }

      // Error 2: Period found but marked Inactive/Closed
      if (!periodIsActive) {
        ui.alert(
          "🚫 SUBMISSION BLOCKED\n\n" +
          "The billing period defined is set to 'Inactive' in the system.\n\n" +
          "Please contact the administrator."
        );
        return;
      }

      // Error 3: Locked (Date Passed AND No Bypass)
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
    // 3. TAB SPECIFIC VALIDATION
    // ===============================================
    for (let tabName in tabValidationMaps) {
      let currentSheet = ss.getSheetByName(tabName);
      if (!currentSheet) continue; 

      let lastRow = currentSheet.getLastRow();
      let startRow = 13;
      if (lastRow < startRow) continue;

      // Fetch up to Column AK (37 columns)
      let dataRange = currentSheet.getRange(startRow, 1, lastRow - startRow + 1, 37).getValues();
      let displayRange = currentSheet.getRange(startRow, 1, lastRow - startRow + 1, 37).getDisplayValues();
      
      let requiredCols = tabValidationMaps[tabName];

      for (let i = 0; i < dataRange.length; i++) {
        let rowData = dataRange[i];
        let valA = String(rowData[0] || "").toLowerCase();

        // STOP checking if we hit TOTAL row
        if (valA.includes("total") && !valA.includes("sub")) break;

        let rawValE = rowData[4]; 
        let valE = (rawValE === undefined || rawValE === null) ? "" : String(rawValE).trim();
        
        // If Column E has data, validate the specific required columns
        if (valE !== "") {
          for (let colLetter of requiredCols) {
            let colIdx = colToIdx(colLetter); // Requires external helper colToIdx
            let rawCellVal = rowData[colIdx];
            let cellValue = (rawCellVal === undefined || rawCellVal === null) ? "" : String(rawCellVal).trim();
            
            // Get visible display value specifically to detect characters like "%" natively formatted
            let visibleCellVal = (displayRange[i][colIdx] || "").trim();

            // 1. Existing Checker: Must not be blank
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

            // 2. Checker for Col O: Cannot be 0 (but accepts other valid characters/symbols)
            if (colLetter === "O" && (cellValue === "0" || cellValue === "0.00" || rawCellVal === 0)) {
               ui.alert(
                `🚫 INVALID DATA\n\n` +
                `Tab: [${tabName}]\n` +
                `Row: ${i + startRow}\n` +
                `Column: O\n\n` +
                `Value cannot be exactly 0 (zero) when Column E has data. It can be any other valid character.`
              );
              return; 
            }

            // 3. UPDATED Checker for Col P: 
            // - ALLOWED completely if it contains "%".
            // - If NO "%", it MUST be a valid number and CANNOT be exactly 0.
            if (colLetter === "P") {
              
              if (visibleCellVal.includes("%")) {
                // Allowed blindly: continue seamlessly to next loop iteration
                continue; 
              } else {
                // Since there is no %, strict check to ensure it's a number and not 0
                let numVal = Number(cellValue); 
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

            }

          }
        }
      }
    }

    // ===============================================
    // 4. DATA EXTRACTION AND SUBMISSION (UPDATED)
    // ===============================================
    
    // Find Header Sheet
    let activeSheet = ss.getActiveSheet();
    let headerSheet = tabValidationMaps[activeSheet.getName()] ? activeSheet : 
                      Object.keys(tabValidationMaps).map(n => ss.getSheetByName(n)).find(s => s !== null);
    
    if (!headerSheet) {
      ui.alert("🚫 Error: No utility tabs (Elec, Water, LPG) found.");
      return;
    }

    // Process Headers (Assumes existing helper)
    const extractedData = processSheetHeaders(headerSheet);
    if (!extractedData) return;

    // Connect to DB
    const props = PropertiesService.getScriptProperties();
    let activeDB_ID = props.getProperty("ACTIVE_DB_ID") || masterDbId;
    let db = SpreadsheetApp.openById(activeDB_ID);
    let dSh = db.getSheetByName("PBTT Submission");

    // Prepare Payload
    const timestamp = Utilities.formatDate(new Date(), "Asia/Manila", "MMM d, yyyy hh:mm a");
    const userEmail = Session.getActiveUser().getEmail();
    const activeFileName = ss.getName();
    const ssUrl = ss.getUrl();

    // Final Array [Timestamp, ...Data, File, Url, Email, Ref#]
    // currentRef is the LAST item in the array.
    const finalRow = [timestamp, ...extractedData, activeFileName, ssUrl, userEmail, currentRef];

    // --- OVERWRITE LOGIC START ---
    
    const dbData = dSh.getDataRange().getValues();
    let rowIndexToOverwrite = -1;
    // Assume currentRef is in the last column of the data being submitted
    const refColumnIndex = finalRow.length - 1; 

    // Loop through DB to see if REF# already exists (Start at 1 to skip Header)
    for (let r = 1; r < dbData.length; r++) {
      let existingRef = String(dbData[r][refColumnIndex] || "").trim();
      
      // Match Found?
      if (existingRef === currentRef) {
        rowIndexToOverwrite = r + 1; // 0-based array to 1-based row index
        break;
      }
    }

    if (rowIndexToOverwrite > 0) {
      // Overwrite existing row
      dSh.getRange(rowIndexToOverwrite, 1, 1, finalRow.length).setValues([finalRow]);
      ui.alert(`✅ SUCCESS: Submission updated (Overwrite existing Ref# ${currentRef}).`);
    } else {
      // Create new row
      dSh.appendRow(finalRow);
      ui.alert(`✅ SUCCESS: New submission recorded successfully.`);
    }
    
  } catch (x) {
    SpreadsheetApp.getUi().alert("System Error: " + x.message);
  } finally {
    lock.releaseLock();
  }
}

// Helper: Column Letter to Index
function colToIdx(char) {
    let sum = 0;
    for (let i = 0; i < char.length; i++) {
        sum *= 26;
        sum += char.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
    }
    return sum - 1;
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

   const sourceFileUrl = sheet.getRange("A1").getValue().toString().trim();
    const textToCheck = sourceFileUrl.toUpperCase(); // Used to easily check N/A or NA regardless of capitalization

    // Block IF: it is completely empty OR it does NOT contain 'http' AND is NOT 'N/A' AND is NOT 'NA'
    if (
        sourceFileUrl === "" || 
        !(sourceFileUrl.toLowerCase().includes("http") || textToCheck === "N/A" || textToCheck === "NA")
    ) {
        SpreadsheetApp.getUi().alert(
            "🚫 SUBMISSION BLOCKED\n\n" +
            "A valid SOURCE FILE URL is missing in cell C20 in Instruction Tab or A1 in Utilities Tab.\n" +
            "Please provide a valid URL, or enter 'N/A' | 'NA' before submitting."
        );
        return null; // Stop the process completely
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

    // --- 1. DATE VALIDATION ---
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

    // --- MONTH-MATCH WARNING ---
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

    // --- 2. HEADER DATA EXTRACTION ---
    for (let item of config) {
        let finalVal = null;
        
        if (item.cell === "E11") {
            // ========================================================
            // SPAN ALL 3 TABS FOR: Provider & Account No (E11 + Y13:Y)
            // ========================================================
            let rawItems = [];
            const tabsToCheck = ["Elec", "Water", "LPG"];
            const ss = sheet.getParent();
            
            tabsToCheck.forEach(tabName => {
                let utilSheet = ss.getSheetByName(tabName);
                if (!utilSheet) return; // Skip if tab is missing
                
                // 1. Grab E11 from this tab
                let e11Val = utilSheet.getRange("E11").getValue();
                if (e11Val) {
                    e11Val.toString().split(",").forEach(v => rawItems.push(v.trim()));
                }
                
                // 2. Grab Y13:Y from this tab (Stop at "TOTAL" row)
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
            // Clean & Deduplicate collected Accounts
            let uniqueItems = Array.from(new Set(rawItems)).filter(Boolean);
            finalVal = uniqueItems.join(", ");
            
        } 
        else if (item.sourceRange) {
            // ========================================================
            // NORMAL ARRAY LOOP: Just the current Active Tab (e.g. Location B13:B)
            // ========================================================
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
            // ========================================================
            // BASIC SINGLE CELLS: Property Name, Dates, Payee, etc.
            // ========================================================
            finalVal = sheet.getRange(item.cell).getValue();
            if (item.type === "date" && finalVal) {
                finalVal = Utilities.formatDate(new Date(finalVal), Session.getScriptTimeZone(), "MMM d, yyyy");
            }
        }

        // --- Verify if empty (Treat '0' as valid text/number) ---
        if ((finalVal === "" || finalVal === undefined || finalVal === null) && finalVal !== 0) {
            missing.push(item.label);
        }

        results.push(finalVal);
    }

    // Reject Submission if Required Header data is missing
    if (missing.length > 0) {
        SpreadsheetApp.getUi().alert("🚫 MISSING HEADER INFO:\n\n" + missing.join("\n"));
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
