// File: MJM_SheetUtils.gs
// Description: Contains utility functions for Google Sheets interaction specific to the
// MJM Application Tracker module, such as managing the main spreadsheet, the "Applications"
// data sheet, and its specific formatting.
// Relies on constants from Global_Constants.gs and MJM_Config.gs.

/**
 * Converts a 1-based column index to its letter representation (e.g., 1 -> A, 27 -> AA).
 * @param {number} column The 1-based column index.
 * @return {string} The column letter(s).
 */
function MJM_columnToLetter(column) {
  let temp, letter = '';
  if (typeof column !== 'number' || column < 1) return ''; // Basic validation
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Gets or creates the main application spreadsheet object.
 * Uses APP_SPREADSHEET_ID or APP_TARGET_FILENAME from Global_Constants.gs.
 * @return {GoogleAppsScript.Spreadsheet.Spreadsheet | null} The spreadsheet object or null on failure.
 */
function MJM_getOrCreateSpreadsheet_Core() {
  let ss = null;
  const fixedId = (typeof APP_SPREADSHEET_ID !== 'undefined' ? APP_SPREADSHEET_ID : ""); // From Global_Constants.gs
  const targetFilename = (typeof APP_TARGET_FILENAME !== 'undefined' ? APP_TARGET_FILENAME : "Comprehensive Job Suite Data_FallbackName"); // From Global_Constants.gs, with a fallback
  const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : false);

  if (DEBUG) Logger.log(`[DEBUG] MJM_SheetUtils (Core): Attempting to get SS. FixedID: "${fixedId}", TargetFilename: "${targetFilename}"`);

  // Try opening by fixed ID first
  if (fixedId && fixedId.trim() !== "" && !String(fixedId).toUpperCase().includes("YOUR_")) {
    try {
      ss = SpreadsheetApp.openById(fixedId);
      if(DEBUG) Logger.log(`  MJM_SheetUtils (Core): Opened spreadsheet by Fixed ID: "${ss.getName()}" (ID: ${fixedId}).`);
    } catch (e) {
      Logger.log(`[ERROR] MJM_SheetUtils (Core): FIXED ID FAIL - Could not open spreadsheet with ID "${fixedId}". Error: ${e.message}. Will attempt by filename if configured.`);
      // If fixedId fails, ss remains null, and we proceed to try by filename (if configured)
    }
  }
  
  // If ss is still null (meaning fixedId was not set, was invalid, or opening by ID failed) AND a target filename is provided
  if (!ss && targetFilename && targetFilename.trim() !== "") {
    if(DEBUG && fixedId) Logger.log(`  MJM_SheetUtils (Core): Fixed ID method failed or not primary. Attempting to find/create by name: "${targetFilename}".`);
    else if(DEBUG) Logger.log(`  MJM_SheetUtils (Core): No Fixed ID. Attempting to find/create by name: "${targetFilename}".`);
    try {
      const files = DriveApp.getFilesByName(targetFilename);
      if (files.hasNext()) {
        const file = files.next();
        ss = SpreadsheetApp.open(file);
        if(DEBUG) Logger.log(`  MJM_SheetUtils (Core): Found and opened existing spreadsheet by name: "${ss.getName()}" (ID: ${ss.getId()}).`);
        if (files.hasNext()) Logger.log(`  [WARN] MJM_SheetUtils (Core): Multiple files found with the name "${targetFilename}". Used the first one found.`);
      } else {
        if(DEBUG) Logger.log(`  MJM_SheetUtils (Core): No spreadsheet found by name "${targetFilename}". Creating new one.`);
        ss = SpreadsheetApp.create(targetFilename);
        Logger.log(`[INFO] MJM_SheetUtils (Core): Successfully created new spreadsheet: "${ss.getName()}" (ID: ${ss.getId()}). This new sheet will contain "Sheet1".`);
      }
    } catch (eDrive) {
      Logger.log(`[ERROR] MJM_SheetUtils (Core): DRIVE/OPEN/CREATE by NAME FAILED for "${targetFilename}". Error: ${eDrive.message}.`);
      return null; // Critical failure if cannot open/create by name
    }
  } else if (!ss) { 
     // If ss is still null here, it means neither fixedId nor targetFilename was validly configured or functional.
     Logger.log("[ERROR] MJM_SheetUtils (Core): CRITICAL - NEITHER APP_SPREADSHEET_ID NOR APP_TARGET_FILENAME yielded a spreadsheet. Please check configuration in Global_Constants.gs.");
     return null;
  }
  return ss;
}

/**
 * Gets or creates a specific sheet tab within the provided main application spreadsheet.
 * Handles initial "Sheet1" renaming if applicable. Calls specific MJM formatting routines.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} mainSpreadsheet The main spreadsheet object.
 * @param {string} targetSheetName The desired name for the sheet tab.
 * @param {number} [desiredIndex=null] Optional. 0-based index for where to insert/move the sheet.
 * @return {{spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, sheet: GoogleAppsScript.Spreadsheet.Sheet | null, newSheetCreatedThisOp: boolean}}
 */
function MJM_getOrCreateSheet_V2(mainSpreadsheet, targetSheetName, desiredIndex = null) {
  const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : false);
  let newSheetCreatedThisOperation = false;

  if (!mainSpreadsheet) { 
    Logger.log(`[ERROR] MJM_SheetUtils (V2): mainSpreadsheet object is null. Cannot get/create sheet: "${targetSheetName}".`); 
    return { spreadsheet: null, sheet: null, newSheetCreatedThisOp: false }; 
  }
  if (!targetSheetName || typeof targetSheetName !== 'string' || targetSheetName.trim() === "") {
    Logger.log(`[ERROR] MJM_SheetUtils (V2): targetSheetName is invalid. Spreadsheet: "${mainSpreadsheet.getName()}".`);
    return { spreadsheet: mainSpreadsheet, sheet: null, newSheetCreatedThisOp: false };
  }
  if (DEBUG) Logger.log(`  MJM_SheetUtils (V2): Seeking/Creating sheet "${targetSheetName}" in SS "${mainSpreadsheet.getName()}". Desired Index: ${desiredIndex}`);

  let targetSheet = mainSpreadsheet.getSheetByName(targetSheetName);

  if (!targetSheet) {
    if(DEBUG) Logger.log(`    Sheet tab "${targetSheetName}" not found. Creating...`);
    try {
      const sheets = mainSpreadsheet.getSheets();
      const defaultSheet1 = mainSpreadsheet.getSheetByName("Sheet1"); // Standard default name

      // Scenario 1: The spreadsheet was just created by MJM_getOrCreateSpreadsheet_Core and *only* contains "Sheet1".
      if (sheets.length === 1 && defaultSheet1 && defaultSheet1.getName() === "Sheet1") {
        if (targetSheetName === "Sheet1") { // If target IS "Sheet1"
            targetSheet = defaultSheet1;
            if(DEBUG) Logger.log(`    Target is "Sheet1", which is the only sheet. Using it (effectively new for this script).`);
            // sheet is considered "newly handled" as we are claiming/formatting it for the first time.
        } else { // Target is something else, so rename the existing "Sheet1"
            targetSheet = defaultSheet1;
            targetSheet.setName(targetSheetName);
            if(DEBUG) Logger.log(`    Renamed initial "Sheet1" to "${targetSheetName}".`);
        }
        newSheetCreatedThisOperation = true; // In both sub-cases, this is effectively a new sheet setup
      } else { // Spreadsheet has multiple sheets, or "Sheet1" is not its only sheet, or it might not exist
        // Insert the new sheet. If desiredIndex is specified, try to insert there.
        let insertOptions = {}; // Pass empty options object. DO NOT pass null or undefined.
        if (desiredIndex !== null && typeof desiredIndex === 'number' && desiredIndex >= 0) {
            targetSheet = mainSpreadsheet.insertSheet(targetSheetName, desiredIndex, insertOptions);
            if(DEBUG) Logger.log(`    Inserted new sheet "${targetSheetName}" at index ${desiredIndex}.`);
        } else {
            targetSheet = mainSpreadsheet.insertSheet(targetSheetName, insertOptions);
            if(DEBUG) Logger.log(`    Inserted new sheet "${targetSheetName}" at default position.`);
        }
        newSheetCreatedThisOperation = true;

        // Clean up "Sheet1" ONLY if it exists, is NOT our target sheet, AND other sheets exist (we didn't just create the only one)
        const nowDefaultSheet1 = mainSpreadsheet.getSheetByName('Sheet1'); // Re-fetch in case it was the one just inserted
        if (nowDefaultSheet1 && nowDefaultSheet1.getSheetId() !== targetSheet.getSheetId() && mainSpreadsheet.getSheets().length > 1) {
          try {
              mainSpreadsheet.deleteSheet(nowDefaultSheet1); 
              if(DEBUG) Logger.log(`    Removed leftover default 'Sheet1' after creating "${targetSheetName}".`);
          } catch (eDel) {
              if(DEBUG) Logger.log(`[WARN] MJM_SheetUtils (V2): Could not delete leftover 'Sheet1': ${eDel.message}`);
          }
        }
      }
    } catch (eCreate) {
        Logger.log(`[ERROR] MJM_SheetUtils (V2): Error creating/renaming sheet "${targetSheetName}": ${eCreate.message}\n${eCreate.stack || ''}`);
        return { spreadsheet: mainSpreadsheet, sheet: null, newSheetCreatedThisOp: false };
    }
  } else {
    if(DEBUG) Logger.log(`    Found existing sheet tab "${targetSheetName}".`);
    // If sheet exists and a desiredIndex is specified, attempt to move it.
    if (desiredIndex !== null && typeof desiredIndex === 'number' && desiredIndex >= 0) {
        const currentSheetIndex = targetSheet.getIndex() -1; // getIndex is 1-based
        if (currentSheetIndex !== desiredIndex) {
            try {
                mainSpreadsheet.setActiveSheet(targetSheet); // moveActiveSheet operates on the active sheet
                mainSpreadsheet.moveActiveSheet(desiredIndex + 1); // moveActiveSheet index is 1-based
                if(DEBUG) Logger.log(`    Moved existing sheet "${targetSheetName}" from index ${currentSheetIndex} to ${desiredIndex}.`);
            } catch (eMove) {
                 Logger.log(`[WARN] MJM_SheetUtils (V2): Could not move existing sheet "${targetSheetName}" to index ${desiredIndex}. Error: ${eMove.message}`);
            }
        }
    }
  }
  
  // Specific MJM formatting if applicable
  // APP_TRACKER_SHEET_TAB_NAME is a Global Constant
  if (targetSheet && targetSheet.getName() === APP_TRACKER_SHEET_TAB_NAME) {
    MJM_setupApplicationsSheetFormatting(targetSheet, newSheetCreatedThisOperation); // Pass the flag
  }
  // Other specific sheet formatting calls could go here (e.g., for Leads if MJM_Leads_SheetUtils calls this)

  return { spreadsheet: mainSpreadsheet, sheet: targetSheet, newSheetCreatedThisOp: newSheetCreatedThisOperation };
}

/**
 * Sets up specific formatting for the "Applications" data sheet (APP_TRACKER_SHEET_TAB_NAME).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The "Applications" sheet object.
 * @param {boolean} [isNewlyHandledSheet=false] Flag indicating if the sheet was just created/renamed/cleared by the calling function.
 */
function MJM_setupApplicationsSheetFormatting(sheet, isNewlyHandledSheet = false) {
  // APP_TRACKER_SHEET_TAB_NAME from Global_Constants.gs
  if (!sheet || typeof sheet.getName !== 'function' || sheet.getName() !== APP_TRACKER_SHEET_TAB_NAME) { 
    Logger.log(`[WARN] MJM_SheetUtils (AppsFormat): Called with invalid sheet ("${sheet ? sheet.getName() : 'null'}"). Expected "${APP_TRACKER_SHEET_TAB_NAME}". Skipping format.`);
    return; // Critical: exit if sheet object is bad or not the correct sheet
  }
  const DEBUG = GLOBAL_DEBUG_MODE;
  if(DEBUG) Logger.log(`  MJM_SheetUtils (AppsFormat): Applying formatting to sheet: "${sheet.getName()}". IsNewlyHandled: ${isNewlyHandledSheet}`);

  // Headers from MJM_Config.gs
  let appSheetHeaders = new Array(MJM_APP_TOTAL_COLUMNS).fill('');
  appSheetHeaders[MJM_APP_PROCESSED_TIMESTAMP_COL - 1] = "Processed Timestamp"; 
  appSheetHeaders[MJM_APP_EMAIL_DATE_COL - 1] = "Email Date";
  appSheetHeaders[MJM_APP_PLATFORM_COL - 1] = "Platform"; 
  appSheetHeaders[MJM_APP_COMPANY_COL - 1] = "Company Name";
  appSheetHeaders[MJM_APP_JOB_TITLE_COL - 1] = "Job Title"; 
  appSheetHeaders[MJM_APP_STATUS_COL - 1] = "Status";
  appSheetHeaders[MJM_APP_PEAK_STATUS_COL - 1] = "Peak Status"; 
  appSheetHeaders[MJM_APP_LAST_UPDATE_DATE_COL - 1] = "Last Update Email Date";
  appSheetHeaders[MJM_APP_EMAIL_SUBJECT_COL - 1] = "Email Subject"; 
  appSheetHeaders[MJM_APP_EMAIL_LINK_COL - 1] = "Email Link";
  appSheetHeaders[MJM_APP_EMAIL_ID_COL - 1] = "Email ID";

  // If sheet is newly handled or appears empty, prepare it fully
  // Check getLastRow first. If sheet is truly empty, getLastColumn might be 0.
  let isEmptyLooking = sheet.getLastRow() === 0;
  if (!isEmptyLooking && sheet.getLastRow() === 1) {
      const lastCol = sheet.getLastColumn();
      if (lastCol > 0) { // Ensure there's at least one column to check
        isEmptyLooking = sheet.getRange(1,1,1,lastCol).getValues()[0].every(c => String(c).trim() === "");
      } else { // lastCol is 0, means sheet is effectively 0x0 or 1x0, so empty
        isEmptyLooking = true; 
      }
  }

  if (isNewlyHandledSheet || isEmptyLooking) {
    if(DEBUG) Logger.log(`    Applications sheet "${sheet.getName()}" determined new/empty. Full clear and setup.`);
    
    // Make sheet active before clearing, sometimes helps with context
    try { SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet); } catch(e) { /* ignore if headless */}
    
    // Order of clear operations might matter if the sheet object reference gets stale.
    // Safest might be to clear content, then formats, then resize, then set headers.
    sheet.clearContents(); // Clear content first
    sheet.clearFormats();  // Then clear formats
    
    // Ensure sheet dimensions after clearing (clear can sometimes reduce max rows/cols)
    const minRowsNeeded = 100; // For data and operations
    const minColsNeeded = MJM_APP_TOTAL_COLUMNS;
    if(sheet.getMaxRows() < minRowsNeeded) sheet.insertRowsAfter(Math.max(1,sheet.getMaxRows()), minRowsNeeded - Math.max(1,sheet.getMaxRows()));
    if(sheet.getMaxColumns() < minColsNeeded) sheet.insertColumnsAfter(Math.max(1,sheet.getMaxColumns()), minColsNeeded - Math.max(1,sheet.getMaxColumns()));
    
    sheet.getRange(1, 1, 1, MJM_APP_TOTAL_COLUMNS).setValues([appSheetHeaders]);
  } else { 
      // Sheet has content. Ensure header row isn't blank from a partial previous error.
      const firstRowValues = sheet.getRange(1, 1, 1, MJM_APP_TOTAL_COLUMNS).getValues()[0];
      if (firstRowValues.every(cell => String(cell).trim() === '')) {
          sheet.getRange(1, 1, 1, MJM_APP_TOTAL_COLUMNS).setValues([appSheetHeaders]);
          if(DEBUG) Logger.log(`    Applications sheet "${sheet.getName()}" had content but blank header; headers re-applied.`);
      }
  }

  // Apply common formatting (headers, widths, data rows, banding etc.)
  const headerRange = sheet.getRange(1, 1, 1, MJM_APP_TOTAL_COLUMNS);
  headerRange.setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  try { if (sheet.getRowHeight(1) !== 40) sheet.setRowHeight(1, 40); } catch(e) {/*ignore*/}

  try { /* ... column width setting logic as before ... */
    sheet.setColumnWidth(MJM_APP_PROCESSED_TIMESTAMP_COL, 160); sheet.setColumnWidth(MJM_APP_EMAIL_DATE_COL, 120);
    sheet.setColumnWidth(MJM_APP_PLATFORM_COL, 100); sheet.setColumnWidth(MJM_APP_COMPANY_COL, 200);
    sheet.setColumnWidth(MJM_APP_JOB_TITLE_COL, 250); sheet.setColumnWidth(MJM_APP_STATUS_COL, 150);
    sheet.setColumnWidth(MJM_APP_PEAK_STATUS_COL, 150); sheet.setColumnWidth(MJM_APP_LAST_UPDATE_DATE_COL, 160);
    sheet.setColumnWidth(MJM_APP_EMAIL_SUBJECT_COL, 300); sheet.setColumnWidth(MJM_APP_EMAIL_LINK_COL, 100);
    sheet.setColumnWidth(MJM_APP_EMAIL_ID_COL, 200);
  } catch (e) { Logger.log(`[WARN] MJM_SheetUtils (AppsFormat): Col width error: ${e.message}`); }

  const maxRows = sheet.getMaxRows();
  if (maxRows > 1) {
      const dataFormatRangeRowCount = Math.min(1000, maxRows - 1); 
      if (dataFormatRangeRowCount > 0) {
        const dataArea = sheet.getRange(2, 1, dataFormatRangeRowCount, MJM_APP_TOTAL_COLUMNS);
        try {
            dataArea.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment('top');
            sheet.setRowHeightsForced(2, dataFormatRangeRowCount, 30);
            sheet.getRange(2, MJM_APP_EMAIL_LINK_COL, dataFormatRangeRowCount, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        } catch(e) { if(DEBUG) Logger.log(`    Data area formatting error (sheet might be too small or protected): ${e.message}`);}
      }
  }

  try { /* ... Banding logic as before ... */
      const bandings = sheet.getBandings(); bandings?.forEach(b=>b.remove());
      sheet.getRange(1, 1, sheet.getMaxRows(), MJM_APP_TOTAL_COLUMNS)
           .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
           .setHeaderRowColor("#D0E4F5").setFirstRowColor("#FFFFFF").setSecondRowColor("#F3F3F3");
  } catch(e){ Logger.log(`[WARN] MJM_SheetUtils (AppsFormat): Banding error: ${e.message}`);}

  try{if(sheet.getFrozenRows() !== 1) sheet.setFrozenRows(1);} catch(e){Logger.log(`[WARN] AppsFormat: SetFrozenRows error: ${e}`);}
  
  if (MJM_APP_PEAK_STATUS_COL > 0 && MJM_APP_PEAK_STATUS_COL <= sheet.getMaxColumns()) {
      try{if(!sheet.isColumnHiddenByUser(MJM_APP_PEAK_STATUS_COL)) sheet.hideColumns(MJM_APP_PEAK_STATUS_COL);}catch(e){Logger.log(`[WARN] AppsFormat: HidePeakCol error: ${e}`);}
  }
  
  const currentMaxConfiguredCols = sheet.getMaxColumns();
  if (currentMaxConfiguredCols > MJM_APP_TOTAL_COLUMNS) {
    try { sheet.deleteColumns(MJM_APP_TOTAL_COLUMNS + 1, currentMaxConfiguredCols - MJM_APP_TOTAL_COLUMNS); } 
    catch (eDelCols) { 
        try { sheet.hideColumns(MJM_APP_TOTAL_COLUMNS + 1, currentMaxConfiguredCols - MJM_APP_TOTAL_COLUMNS); }
        catch(eHideCols){ Logger.log(`[WARN] AppsFormat: Hide/Del extra cols: ${eHideCols.message}`); }
    }
  }
  if(DEBUG) Logger.log(`  MJM_SheetUtils (AppsFormat): Formatting complete for sheet "${sheet.getName()}".`);
}


/**
 * Deletes the default "Sheet1" if it exists, is not a designated important sheet,
 * and other sheets are present in the spreadsheet.
 * Call this LAST in the main setup routine.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} mainSpreadsheet The main application spreadsheet.
 */
function MJM_cleanupDefaultSheet1(mainSpreadsheet) {
    const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : false);
    if (!mainSpreadsheet) { /* ... */ return; }
    const defaultSheetInstance = mainSpreadsheet.getSheetByName("Sheet1");
    if (defaultSheetInstance) {
        const importantApplicationSheetNames = [ APP_TRACKER_SHEET_TAB_NAME, LEADS_SHEET_TAB_NAME, DASHBOARD_TAB_NAME, DASHBOARD_HELPER_SHEET_NAME,PROFILE_DATA_SHEET_NAME, JD_ANALYSIS_SHEET_NAME, BULLET_SCORING_RESULTS_SHEET_NAME ];
        if (!importantApplicationSheetNames.includes("Sheet1")) {
            if (mainSpreadsheet.getSheets().length > 1) {
                try { mainSpreadsheet.deleteSheet(defaultSheetInstance); if(DEBUG) Logger.log(`  MJM_cleanupDefaultSheet1: Deleted leftover "Sheet1".`); }
                catch (e) { Logger.log(`[WARN] MJM_cleanupDefaultSheet1: Could not delete "Sheet1": ${e.message}`); }
            } else { if(DEBUG) Logger.log(`  MJM_cleanupDefaultSheet1: "Sheet1" is only sheet, not deleting.`);}
        } else { if(DEBUG) Logger.log(`  MJM_cleanupDefaultSheet1: "Sheet1" is important, not deleting.`);}
    } else { if(DEBUG) Logger.log(`  MJM_cleanupDefaultSheet1: Default "Sheet1" not found.`);}
}
