// File: RTS_SheetSetup.gs
// Description: Contains functions for setting up and formatting the master profile data sheet
// (PROFILE_DATA_SHEET_NAME) within a given spreadsheet.
// This is now intended to be called by MJM's main setup routine (runFullProjectInitialSetup).
// Relies on global constants from Global_Constants.gs for styling, sheet name, and profile structure.

/**
 * Sets up or reformats the master profile data sheet ("MasterProfile") within the provided spreadsheet.
 * Creates the sheet if it doesn't exist, defines section headers based on PROFILE_STRUCTURE,
 * applies formatting, and adds placeholder rows for data entry.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} targetSpreadsheet The spreadsheet object where the profile sheet should be set up.
 * @return {GoogleAppsScript.Spreadsheet.Sheet | null} The formatted profile sheet object, or null on critical error.
 */
function RTS_setupMasterResumeSheet(targetSpreadsheet) {
  const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : true);
  let ui;
  try {
    if (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getActiveSpreadsheet()) {
      ui = SpreadsheetApp.getUi();
    }
  } catch (e) { /* UI not essential for core logic if called from non-UI context */
    Logger.log("[WARN] RTS_SheetSetup: Could not get UI context. Alerts will be skipped if this instance of 'ui' is null.");
  }

  if (DEBUG) Logger.log(`--- RTS_SheetSetup: Starting setup for Profile Sheet in Spreadsheet: "${targetSpreadsheet ? targetSpreadsheet.getName() : 'NULL SPREADSHEET PROVIDED'}" ---`);

  if (!targetSpreadsheet) {
    Logger.log("[ERROR] RTS_SheetSetup: Target spreadsheet object not provided. Cannot set up profile sheet.");
    if (ui) ui.alert("Setup Error (RTS Profile Sheet)", "Target spreadsheet object was missing. Profile sheet setup failed.", ui.ButtonSet.OK);
    return null;
  }

  // Constants from Global_Constants.gs:
  const profileSheetNameGlobal = (typeof PROFILE_DATA_SHEET_NAME !== 'undefined' ? PROFILE_DATA_SHEET_NAME : "MasterProfile_FallbackName");
  const profileStructureGlobal = (typeof PROFILE_STRUCTURE !== 'undefined' ? PROFILE_STRUCTURE : []);
  const numBulletColsGlobal = (typeof NUM_DEDICATED_PROFILE_BULLET_COLUMNS !== 'undefined' ? NUM_DEDICATED_PROFILE_BULLET_COLUMNS : 3);
  const headerBgColorGlobal = (typeof PROFILE_HEADER_BACKGROUND_COLOR !== 'undefined' ? PROFILE_HEADER_BACKGROUND_COLOR : "#E1BEE7");
  const headerFontColorGlobal = (typeof PROFILE_HEADER_FONT_COLOR !== 'undefined' ? PROFILE_HEADER_FONT_COLOR : "#4A148C");
  const subHeaderBgColorGlobal = (typeof PROFILE_SUB_HEADER_BACKGROUND_COLOR !== 'undefined' ? PROFILE_SUB_HEADER_BACKGROUND_COLOR : "#F3E5F5");
  const borderColorGlobal = (typeof PROFILE_BORDER_COLOR !== 'undefined' ? PROFILE_BORDER_COLOR : "#BDBDBD");

  if (profileStructureGlobal.length === 0) {
    Logger.log("[ERROR] RTS_SheetSetup: PROFILE_STRUCTURE in Global_Constants.gs is empty or not defined. Cannot structure sheet.");
    if (ui) ui.alert("Configuration Error", "PROFILE_STRUCTURE is not defined in Global_Constants.gs. Profile sheet setup failed.", ui.ButtonSet.OK);
    return null;
  }

  if (DEBUG) Logger.log(`  RTS_SheetSetup: Target Profile Sheet Name: "${profileSheetNameGlobal}", Bullet Cols: ${numBulletColsGlobal}`);

  let sheetToFormat = targetSpreadsheet.getSheetByName(profileSheetNameGlobal);
  let sheetWasNewlyHandled = false; // Tracks if sheet was created, renamed, or cleared this run

  if (sheetToFormat) {
    if(DEBUG) Logger.log(`  RTS_SheetSetup: Sheet "${profileSheetNameGlobal}" already exists.`);
    // Only prompt to clear if UI is available. If run headless, default to clearing for a consistent setup.
    let userConsentToClear = SpreadsheetApp.getUi().Button.YES; // Default for headless run
    if (ui) {
        userConsentToClear = ui.alert(
            "Confirm Profile Sheet Reformat",
            `The sheet named "${profileSheetNameGlobal}" already exists in spreadsheet "${targetSpreadsheet.getName()}".\n\nDo you want to CLEAR its contents and re-apply the standard format?\n\nWARNING: This will ERASE any existing data in the "${profileSheetNameGlobal}" tab.`,
            ui.ButtonSet.YES_NO
        );
    }
    
    if (userConsentToClear === SpreadsheetApp.getUi().Button.YES) {
      Logger.log(`  RTS_SheetSetup: User consented (or headless default) to clear existing sheet "${profileSheetNameGlobal}". Clearing...`);
      sheetToFormat.clear(); // Clears content, formats, notes, etc.
      // Ensure columns/rows are visible for reformatting after a clear
      if(sheetToFormat.getMaxColumns() > 0) sheetToFormat.showColumns(1, sheetToFormat.getMaxColumns());
      if(sheetToFormat.getMaxRows() > 0) sheetToFormat.showRows(1, sheetToFormat.getMaxRows());
      sheetWasNewlyHandled = true;
    } else {
      Logger.log(`  RTS_SheetSetup: User chose NOT to clear existing sheet "${profileSheetNameGlobal}". Setup for this sheet aborted.`);
      if (ui) ui.alert("Profile Sheet Reformat Cancelled", `Reformatting of sheet "${profileSheetNameGlobal}" was cancelled by the user. No changes made to this sheet.`, ui.ButtonSet.OK);
      return sheetToFormat; // Return existing sheet as is
    }
  } else {
    if(DEBUG) Logger.log(`  RTS_SheetSetup: Sheet "${profileSheetNameGlobal}" not found. Creating it.`);
    const defaultSheet = targetSpreadsheet.getSheetByName("Sheet1");
    // If target SS has only "Sheet1" and target name is different, rename "Sheet1"
    if (targetSpreadsheet.getSheets().length === 1 && defaultSheet && defaultSheet.getName() === "Sheet1" && profileSheetNameGlobal !== "Sheet1") {
        defaultSheet.setName(profileSheetNameGlobal);
        sheetToFormat = defaultSheet;
        sheetToFormat.clear(); // Clear the renamed sheet
        Logger.log(`  RTS_SheetSetup: Renamed existing "Sheet1" to "${profileSheetNameGlobal}" and cleared it.`);
    } else { // Insert a new sheet
        sheetToFormat = targetSpreadsheet.insertSheet(profileSheetNameGlobal);
        Logger.log(`  RTS_SheetSetup: Inserted new sheet named "${profileSheetNameGlobal}".`);
        // If a "Sheet1" still exists and it's not our target sheet, and there are now multiple sheets, delete the old "Sheet1".
        const lingeringDefaultSheet = targetSpreadsheet.getSheetByName("Sheet1");
        if (lingeringDefaultSheet && lingeringDefaultSheet.getSheetId() !== sheetToFormat.getSheetId() && targetSpreadsheet.getSheets().length > 1) {
            try { targetSpreadsheet.deleteSheet(lingeringDefaultSheet); Logger.log(`  RTS_SheetSetup: Deleted original "Sheet1" as it was not the target.`); }
            catch(eDel) { Logger.log(`[WARN] RTS_SheetSetup: Could not delete original "Sheet1" after creating target sheet: ${eDel.message}`); }
        }
    }
    sheetWasNewlyHandled = true;
  }

  if (!sheetToFormat) { // Should be extremely rare given the logic above
    Logger.log("[CRITICAL ERROR] RTS_SheetSetup: Could not obtain a valid sheet object to format after creation/find attempts.");
    if (ui) ui.alert("Profile Sheet Error", "A critical error occurred: Could not obtain a sheet object for the profile data.", ui.ButtonSet.OK);
    return null;
  }
  
  try { targetSpreadsheet.setActiveSheet(sheetToFormat); } catch(e) { /* non-critical if fails */ }

  // --- Apply Structure and Formatting ---
  Logger.log(`  RTS_SheetSetup: Applying structure and formatting to sheet "${profileSheetNameGlobal}"...`);

  let maxDynamicHeaders = 0;
  profileStructureGlobal.forEach(section => {
    let headersForSection = section.headers ? [...section.headers] : [];
    const ucTitle = section.title.toUpperCase();
    if (ucTitle === "EXPERIENCE" || ucTitle === "LEADERSHIP & UNIVERSITY INVOLVEMENT") {
      for (let i = 1; i <= numBulletColsGlobal; i++) if(!headersForSection.includes(`Responsibility${i}`)) headersForSection.push(`Responsibility${i}`);
    } else if (ucTitle === "PROJECTS") {
      for (let i = 1; i <= numBulletColsGlobal; i++) if(!headersForSection.includes(`DescriptionBullet${i}`)) headersForSection.push(`DescriptionBullet${i}`);
    }
    if (headersForSection.length > maxDynamicHeaders) maxDynamicHeaders = headersForSection.length;
  });
  const numColsToSetup = Math.max(2, maxDynamicHeaders); // At least Key/Value for Personal Info

  if (sheetToFormat.getMaxColumns() < numColsToSetup) sheetToFormat.insertColumns(sheetToFormat.getMaxColumns() + 1, numColsToSetup - sheetToFormat.getMaxColumns());
  else if (sheetToFormat.getMaxColumns() > numColsToSetup) sheetToFormat.hideColumns(numColsToSetup + 1, sheetToFormat.getMaxColumns() - numColsToSetup);
  if (sheetToFormat.getMaxRows() < 250) sheetToFormat.insertRows(sheetToFormat.getMaxRows() + 1, 250 - sheetToFormat.getMaxRows());

  sheetToFormat.setColumnWidth(1, 230); // Main ID column (Section title, Key, Category)
  if (numColsToSetup >= 2) sheetToFormat.setColumnWidth(2, 300); // Primary value/detail column
  for (let c = 3; c <= numColsToSetup; c++) sheetToFormat.setColumnWidth(c, 180); // Subsequent columns

  let currentSheetRow = 1;
  profileStructureGlobal.forEach(sectionConfig => {
    const sectionTitleUC = sectionConfig.title.toUpperCase();
    const mainHeaderDisplayRange = sheetToFormat.getRange(currentSheetRow, 1, 1, numColsToSetup);
    mainHeaderDisplayRange.setValue(sectionTitleUC)
      .setBackground(headerBgColorGlobal).setFontColor(headerFontColorGlobal)
      .setFontWeight("bold").setFontSize(12).setHorizontalAlignment("left").setVerticalAlignment("middle")
      .setBorder(true, true, true, true, null, null, borderColorGlobal, SpreadsheetApp.BorderStyle.SOLID_THIN)
      .mergeAcross();
    currentSheetRow++;

    let fullHeadersForThisSection = sectionConfig.headers ? [...sectionConfig.headers] : [];
    if (sectionTitleUC === "EXPERIENCE" || sectionTitleUC === "LEADERSHIP & UNIVERSITY INVOLVEMENT") {
      for (let i=1; i<=numBulletColsGlobal; i++) if(!fullHeadersForThisSection.includes(`Responsibility${i}`)) fullHeadersForThisSection.push(`Responsibility${i}`);
    } else if (sectionTitleUC === "PROJECTS") {
      for (let i=1; i<=numBulletColsGlobal; i++) if(!fullHeadersForThisSection.includes(`DescriptionBullet${i}`)) fullHeadersForThisSection.push(`DescriptionBullet${i}`);
    }

    if (sectionTitleUC === "PERSONAL INFO") {
      const suggestedPIKeys = ["Full Name", "Location", "Phone", "Email", "LinkedIn URL", "GitHub URL", "Portfolio URL"]; // Example starter keys
      suggestedPIKeys.forEach(piKey => {
        sheetToFormat.getRange(currentSheetRow, 1).setValue(piKey).setBackground(subHeaderBgColorGlobal).setBorder(null, null, true, null, null, null, borderColorGlobal, SpreadsheetApp.BorderStyle.DOTTED);
        sheetToFormat.getRange(currentSheetRow, 2).setValue("").setBorder(null, null, true, null, null, null, borderColorGlobal, SpreadsheetApp.BorderStyle.DOTTED); // For Value
        // Apply border to other columns in PI section if numColsToSetup > 2 for consistency
        if (numColsToSetup > 2) sheetToFormat.getRange(currentSheetRow, 3, 1, numColsToSetup - 2).setBorder(null, null, true, null, null, null, borderColorGlobal, SpreadsheetApp.BorderStyle.DOTTED);
        currentSheetRow++;
      });
    } else if (sectionTitleUC === "SUMMARY") {
      const summaryInstructionRange = sheetToFormat.getRange(currentSheetRow, 1, 1, numColsToSetup);
      summaryInstructionRange.mergeAcross()
                           .setValue("(Enter your professional summary here. You can use Column A and wrap text, or merge cells if preferred.)")
                           .setFontStyle("italic").setBackground("#FFFDE7") // Light yellow for instruction/placeholder
                           .setBorder(null, null, true, null, null, null, borderColorGlobal, SpreadsheetApp.BorderStyle.DOTTED);
      currentSheetRow++;
    } else if (fullHeadersForThisSection.length > 0) { // For other tabular sections
      sheetToFormat.getRange(currentSheetRow, 1, 1, fullHeadersForThisSection.length).setValues([fullHeadersForThisSection])
        .setBackground(subHeaderBgColorGlobal).setFontWeight("bold").setFontSize(10)
        .setBorder(null, null, true, null, null, null, borderColorGlobal, SpreadsheetApp.BorderStyle.SOLID_MEDIUM); // Thicker bottom border for headers
      currentSheetRow++;
    }
    
    // Add blank data placeholder rows under tabular sections
    const numBlankDataRowsToAdd = (sectionTitleUC !== "PERSONAL INFO" && sectionTitleUC !== "SUMMARY" && fullHeadersForThisSection.length > 0) ? 2 : (sectionTitleUC === "SUMMARY" ? 1:0); // Extra space for summary
    if (numBlankDataRowsToAdd > 0) {
        if ((currentSheetRow + numBlankDataRowsToAdd -1) <= sheetToFormat.getMaxRows()) { // Check bounds
             const blankRange = sheetToFormat.getRange(currentSheetRow, 1, numBlankDataRowsToAdd, numColsToSetup);
             blankRange.clearContent().setBackground(null); // Clear any inherited format
             // Apply light dotted bottom border to placeholder data rows
             blankRange.setBorder(null, null, true, null, null, null, borderColorGlobal, SpreadsheetApp.BorderStyle.DOTTED);
        }
      currentSheetRow += numBlankDataRowsToAdd;
    }
    currentSheetRow++; // Blank spacer row between sections
  });
  
  Logger.log(`[SUCCESS] RTS_SheetSetup: Profile Sheet "${profileSheetNameGlobal}" in "${targetSpreadsheet.getName()}" structure applied/re-applied.`);
  if (ui && sheetWasNewlyHandled) {
      ui.alert(
          "Profile Sheet Setup Complete",
          `Sheet "${profileSheetNameGlobal}" in spreadsheet "${targetSpreadsheet.getName()}" has been (re)structured.\n\nPlease review and fill in your master profile details in this tab.`,
          ui.ButtonSet.OK // Corrected ui.alert
      );
  }
  return sheetToFormat;
}
