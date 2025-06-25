// File: MJM_Leads_SheetUtils.gs (or MRM_Leads_SheetUtils.gs)
// Description: Contains utility functions specific to the MJM "Potential Job Leads" sheet,
// such as initial formatting, header mapping, retrieving processed email IDs, and writing job lead data or errors.
// Relies on constants from Global_Constants.gs and MJM_Config.gs.

/**
 * Sets up initial formatting for the "Potential Job Leads" sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} leadsSheet The "Potential Job Leads" sheet object.
 * @param {string[]} leadsHeadersArray The array of header names for the leads sheet (from MJM_Config.gs).
 */
function MJM_Leads_setupSheetFormatting(leadsSheet, leadsHeadersArray) {
  const moduleName = "Leads Sheet Util";
  // LEADS_SHEET_TAB_NAME from Global_Constants.gs
  if (!leadsSheet || leadsSheet.getName() !== LEADS_SHEET_TAB_NAME) {
    Logger.log(`[WARN] ${moduleName}: MJM_Leads_setupSheetFormatting called with invalid sheet object or name. Expected "${LEADS_SHEET_TAB_NAME}".`);
    return;
  }
  if (!leadsHeadersArray || !Array.isArray(leadsHeadersArray) || leadsHeadersArray.length === 0) {
      Logger.log(`[ERROR] ${moduleName}: leadsHeadersArray not provided or empty. Cannot format sheet "${leadsSheet.getName()}".`);
      return;
  }
  Logger.log(`[INFO] ${moduleName}: Applying formatting to sheet "${leadsSheet.getName()}".`);

  // Clear existing format only if sheet is considered new/empty (first row is blank or doesn't match first header)
  // This avoids wiping out formatting if user customized it, unless sheet is being freshly setup.
  const firstHeaderCell = leadsSheet.getRange(1,1);
  if (leadsSheet.getLastRow() === 0 || (leadsSheet.getLastRow() === 1 && firstHeaderCell.isBlank())) {
    leadsSheet.clearFormats(); // Clear formats for a truly new/empty sheet
    Logger.log(`  Sheet "${leadsSheet.getName()}" appeared empty, cleared existing formats before applying new ones.`);
  }

  // Ensure headers are present and styled
  const headerRange = leadsSheet.getRange(1, 1, 1, leadsHeadersArray.length);
  const currentHeaders = headerRange.getValues()[0];
  let headersMatch = currentHeaders.length === leadsHeadersArray.length && currentHeaders.every((val, index) => String(val).trim() === leadsHeadersArray[index]);

  if (!headersMatch || firstHeaderCell.isBlank()) { // If headers don't match or sheet was truly empty
    leadsSheet.getRange(1, 1, 1, leadsHeadersArray.length).setValues([leadsHeadersArray]);
    Logger.log(`  Headers set/verified for "${leadsSheet.getName()}".`);
  }
  headerRange.setFontWeight("bold").setHorizontalAlignment("center").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  if (leadsSheet.getFrozenRows() < 1) leadsSheet.setFrozenRows(1);

  // Apply/Re-apply banding (clearing old ones first is good practice if run multiple times)
  try {
      const existingBandings = leadsSheet.getBandings();
      for (let i = 0; i < existingBandings.length; i++) { existingBandings[i].remove(); }
      leadsSheet.getRange(1, 1, leadsSheet.getMaxRows(), leadsHeadersArray.length)
                .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
                .setHeaderRowColor("#C9DAF8") // Example: Light Cornflower Blue
                .setFirstRowColor("#FFFFFF")
                .setSecondRowColor("#F3F3F3");
      Logger.log(`  Row banding applied to "${leadsSheet.getName()}".`);
  } catch(e) { Logger.log(`[WARN] ${moduleName}: Banding error for "${leadsSheet.getName()}": ${e.toString()}`); }

  // Column Widths (adjust these as needed)
  const columnWidthsLeads = {
      "Date Added": 100, "Job Title": 220, "Company": 180, "Location": 150,
      "Source Email Subject": 250, "Link to Job Posting": 280, "Status": 100,
      "Source Email ID": 150, "Processed Timestamp": 120, "Notes": 300
  };
  leadsHeadersArray.forEach((headerName, index) => {
    const columnIndex = index + 1;
    if (columnWidthsLeads[headerName]) {
      try { leadsSheet.setColumnWidth(columnIndex, columnWidthsLeads[headerName]); }
      catch (e) { Logger.log(`[WARN] ${moduleName}: Error setting width for column "${headerName}": ${e.toString()}`); }
    }
  });
  Logger.log(`  Column widths applied to "${leadsSheet.getName()}".`);

  // Hide any extra columns to the right
  const maxSheetCols = leadsSheet.getMaxColumns();
  if (maxSheetCols > leadsHeadersArray.length) {
    try { leadsSheet.hideColumns(leadsHeadersArray.length + 1, maxSheetCols - leadsHeadersArray.length); }
    catch(e) { Logger.log(`[WARN] ${moduleName}: Error hiding extra columns: ${e.toString()}`); }
  }
}

/**
 * Retrieves a specific sheet (by name) from a given spreadsheet object and maps its header names to column numbers.
 * Ensures the sheet is formatted if it's being newly recognized.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet The spreadsheet object.
 * @param {string} sheetName The name of the sheet to retrieve (e.g., LEADS_SHEET_TAB_NAME from Global_Constants.gs).
 * @param {string[]} expectedHeadersArray The array of expected header names for this sheet (e.g., MJM_LEADS_SHEET_HEADERS from MJM_Config.gs).
 * @return {{sheet: GoogleAppsScript.Spreadsheet.Sheet | null, headerMap: Object}}
 */
function MJM_Leads_getSheetAndHeaderMap(spreadsheet, sheetName, expectedHeadersArray) {
  const moduleName = "Leads Sheet Util (getSheetAndHeaderMap)";
  if (!spreadsheet) {
    Logger.log(`[ERROR] ${moduleName}: Spreadsheet object is null. Cannot get sheet "${sheetName}".`);
    return { sheet: null, headerMap: {} };
  }
  if (!sheetName || !expectedHeadersArray || expectedHeadersArray.length === 0) {
      Logger.log(`[ERROR] ${moduleName}: sheetName or expectedHeadersArray is invalid.`);
      return { sheet: null, headerMap: {} };
  }

  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`[INFO] ${moduleName}: Sheet "${sheetName}" not found in "${spreadsheet.getName()}". Will be created by MJM_Leads_setupSheetFormatting if called via setup.`);
    // Setup flow (MJM_runInitialSetup_JobLeadsModule) should have created it. If called outside setup, this might be an issue.
    // For processing, we expect it to exist.
    return { sheet: null, headerMap: {} }; // If sheet doesn't exist, cannot map headers.
  }

  const headersFromSheet = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerMap = {};
  let headersAreValid = true;

  // Create map and validate against expectedHeadersArray
  expectedHeadersArray.forEach(expectedHeader => {
      const indexInSheet = headersFromSheet.findIndex(h => String(h).trim() === expectedHeader);
      if (indexInSheet !== -1) {
          headerMap[expectedHeader] = indexInSheet + 1; // 1-based column index
      } else {
          Logger.log(`[WARN] ${moduleName}: Expected header "${expectedHeader}" NOT FOUND in sheet "${sheet.getName()}".`);
          headersAreValid = false; // Mark as invalid if a crucial header is missing
      }
  });

  // Additionally, map any other headers found in the sheet, even if not in expectedHeadersArray (might be user-added)
  headersFromSheet.forEach((h, i) => {
      const trimmedHeader = String(h).trim();
      if (trimmedHeader !== "" && !headerMap[trimmedHeader]) { // If not empty and not already mapped
          headerMap[trimmedHeader] = i + 1;
      }
  });

  if (Object.keys(headerMap).length === 0 || !headersAreValid) { // If no headers mapped or critical ones missing
    Logger.log(`[ERROR] ${moduleName}: Header map for "${sheet.getName()}" is empty or crucial headers are missing. Please ensure headers match MJM_LEADS_SHEET_HEADERS.`);
    return { sheet: sheet, headerMap: {} }; // Return sheet but empty map to signal error
  }

  if (GLOBAL_DEBUG_MODE) Logger.log(`[DEBUG] ${moduleName}: Header map for "${sheet.getName()}": ${JSON.stringify(headerMap)}`);
  return { sheet: sheet, headerMap: headerMap };
}

/**
 * Retrieves a set of all unique email IDs from the "Source Email ID" column of the leads sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} leadsSheet The "Potential Job Leads" sheet object.
 * @param {Object} headerMap An object mapping header names to column numbers for the leads sheet.
 * @return {Set<string>} A set of processed email IDs.
 */
function MJM_Leads_getProcessedEmailIdsFromSheet(leadsSheet, headerMap) {
  const ids = new Set();
  const moduleName = "Leads Sheet Util (getProcessedEmailIds)";
  const emailIdColHeader = "Source Email ID"; // This must match a header in MJM_LEADS_SHEET_HEADERS

  if (!leadsSheet) { Logger.log(`[WARN] ${moduleName}: Sheet object is null.`); return ids; }
  if (!headerMap || !headerMap[emailIdColHeader]) {
    Logger.log(`[WARN] ${moduleName}: "${emailIdColHeader}" not found in headerMap for "${leadsSheet.getName()}". HeaderMap: ${JSON.stringify(headerMap)}`);
    return ids;
  }

  const emailIdColNum = headerMap[emailIdColHeader];
  const lastRow = leadsSheet.getLastRow();
  if (lastRow < 2) return ids; // No data rows if only header

  try {
    const emailIdValues = leadsSheet.getRange(2, emailIdColNum, lastRow - 1, 1).getValues();
    emailIdValues.forEach(row => {
      if (row[0] && String(row[0]).trim() !== "") ids.add(String(row[0]).trim());
    });
  } catch (e) {
    Logger.log(`[ERROR] ${moduleName}: Error reading email IDs from col ${emailIdColNum} in "${leadsSheet.getName()}": ${e.toString()}`);
  }
  return ids;
}

/**
 * Appends a new row with job lead data to the "Potential Job Leads" sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} leadsSheet The sheet object.
 * @param {GoogleAppsScript.Gmail.GmailMessage} message The source Gmail message (used for subject, ID).
 * @param {Object} jobData An object containing the job lead details extracted by AI (jobTitle, company, location, linkToJobPosting).
 * @param {Object} headerMap An object mapping header names to column numbers for the sheet.
 */
function MJM_Leads_writeJobToSheet(leadsSheet, message, jobData, headerMap) {
  const moduleName = "Leads Sheet Util (writeJobToSheet)";
  if (!leadsSheet) { Logger.log(`[ERROR] ${moduleName}: Sheet object is null.`); return; }
  if (!headerMap || Object.keys(headerMap).length === 0) { Logger.log(`[ERROR] ${moduleName}: headerMap is invalid.`); return; }
  if (!message || !jobData) { Logger.log(`[ERROR] ${moduleName}: Message or jobData is null.`); return; }

  // MJM_LEADS_SHEET_HEADERS from MJM_Config.gs defines the canonical order and completeness
  const newRowArray = new Array(MJM_LEADS_SHEET_HEADERS.length).fill("");

  MJM_LEADS_SHEET_HEADERS.forEach((header, index) => {
    const colNumInMap = headerMap[header]; // This util now expects caller to use MJM_Leads_getSheetAndHeaderMap which ensures map covers these headers.
    if (!colNumInMap) {
      // This shouldn't happen if headerMap is generated from MJM_LEADS_SHEET_HEADERS correctly
      // Or if the sheet on Drive has been manually altered.
      if (GLOBAL_DEBUG_MODE) Logger.log(`  [DEBUG] ${moduleName}: Header "${header}" not in provided headerMap for writing. Will be blank in sheet.`);
      return; // Skip if header somehow not in map (should ideally not happen)
    }
    // The position in newRowArray is `index` (from iterating MJM_LEADS_SHEET_HEADERS).
    // The actual column on the sheet this header corresponds to is headerMap[header]. We build the row in the canonical order.
    switch (header) {
      case "Date Added":            newRowArray[index] = new Date(); break;
      case "Job Title":             newRowArray[index] = jobData.jobTitle || "N/A"; break;
      case "Company":               newRowArray[index] = jobData.company || "N/A"; break;
      case "Location":              newRowArray[index] = jobData.location || "N/A"; break;
      case "Source Email Subject":  newRowArray[index] = message.getSubject(); break;
      case "Link to Job Posting":   newRowArray[index] = jobData.linkToJobPosting || "N/A"; break;
      case "Status":                newRowArray[index] = "New"; break; // Default status
      case "Source Email ID":       newRowArray[index] = message.getId(); break;
      case "Processed Timestamp":   newRowArray[index] = new Date(); break;
      case "Notes":                 newRowArray[index] = jobData.notes || ""; break; // Allow 'notes' field from jobData if AI provides
      default: break; // Should not happen if MJM_LEADS_SHEET_HEADERS is exhaustive
    }
  });

  try {
    leadsSheet.appendRow(newRowArray);
    if (GLOBAL_DEBUG_MODE) Logger.log(`  Appended lead: "${jobData.jobTitle || 'N/A'}" to "${leadsSheet.getName()}".`);
  } catch (e) {
    Logger.log(`[ERROR] ${moduleName}: Failed to append row for lead "${jobData.jobTitle || 'N/A'}": ${e.toString()}`);
  }
}

/**
 * Appends an error entry to the "Potential Job Leads" sheet when processing a lead fails.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} leadsSheet The sheet object.
 * @param {GoogleAppsScript.Gmail.GmailMessage} message The Gmail message that caused the error.
 * @param {string} errorType A short description of the error type (e.g., "Gemini API Error").
 * @param {string} errorDetails Additional details about the error.
 * @param {Object} headerMap An object mapping header names to column numbers for the sheet.
 */
function MJM_Leads_writeErrorEntryToSheet(leadsSheet, message, errorType, errorDetails, headerMap) {
  const moduleName = "Leads Sheet Util (writeErrorEntry)";
  if (!leadsSheet) { Logger.log(`[ERROR] ${moduleName}: Sheet object is null.`); return; }
  if (!headerMap || Object.keys(headerMap).length === 0) { Logger.log(`[ERROR] ${moduleName}: headerMap is invalid.`); return; }
  
  const detailsStr = String(errorDetails || "No details").substring(0, 1000);
  Logger.log(`[WARN] ${moduleName}: Writing error for msg ${message ? message.getId() : 'UnknownMsg'}: ${errorType}. Details: ${detailsStr}`);

  // MJM_LEADS_SHEET_HEADERS from MJM_Config.gs defines the canonical order
  const errorRowArray = new Array(MJM_LEADS_SHEET_HEADERS.length).fill("");

  MJM_LEADS_SHEET_HEADERS.forEach((header, index) => {
    switch (header) {
      case "Date Added":            errorRowArray[index] = new Date(); break;
      case "Job Title":             errorRowArray[index] = "PROCESSING ERROR"; break;
      case "Company":               errorRowArray[index] = errorType.substring(0,200); break;
      case "Source Email Subject":  errorRowArray[index] = message ? message.getSubject() : "N/A"; break;
      case "Status":                errorRowArray[index] = "Error"; break;
      case "Source Email ID":       errorRowArray[index] = message ? message.getId() : "N/A"; break;
      case "Processed Timestamp":   errorRowArray[index] = new Date(); break;
      case "Notes":                 errorRowArray[index] = `Type: ${errorType}. Details: ${detailsStr}`; break;
      default:                      errorRowArray[index] = "N/A"; break; // For Location, Link, etc.
    }
  });

  try {
    leadsSheet.appendRow(errorRowArray);
  } catch (e) {
    Logger.log(`[ERROR] ${moduleName}: Failed to append error row: ${e.toString()}`);
  }
}
