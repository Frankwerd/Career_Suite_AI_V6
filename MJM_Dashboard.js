// File: MJM_Dashboard.gs
// Description: Manages the MJM Dashboard sheet, its charts, and the DashboardHelperData sheet.
// Implements specific formulas for DashboardHelperData.
// Formats Dashboard for correct visibility and layout.
// Relies on constants from Global_Constants.gs and MJM_Config.gs.

/**
 * Gets or creates the MJM Dashboard Sheet and ensures it's the first tab.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet The main application spreadsheet.
 * @return {GoogleAppsScript.Spreadsheet.Sheet | null} The dashboard sheet object.
 */
function MJM_getOrCreateDashboardSheet(spreadsheet) {
  if (!spreadsheet) { Logger.log("[ERROR] MJM_Dashboard (GetDash): Spreadsheet object missing."); return null; }
  // DASHBOARD_TAB_NAME from Global_Constants.gs
  const dashboardSheetName = (typeof DASHBOARD_TAB_NAME !== 'undefined' ? DASHBOARD_TAB_NAME : "Dashboard_FallbackName");
  let dashboardSheet = spreadsheet.getSheetByName(dashboardSheetName);

  if (!dashboardSheet) {
    dashboardSheet = spreadsheet.insertSheet(dashboardSheetName, 0); // Insert at index 0 (first tab)
    Logger.log(`[INFO] MJM_Dashboard (GetDash): Created new sheet "${dashboardSheetName}" at first position.`);
  }
  try { // Ensure it's the first tab if it exists or was just created
    if (dashboardSheet.getIndex() !== 1) { // getIndex() is 1-based
      spreadsheet.setActiveSheet(dashboardSheet);
      spreadsheet.moveActiveSheet(1); // moveActiveSheet uses 1-based position
    }
    Logger.log(`[INFO] MJM_Dashboard (GetDash): Sheet "${dashboardSheetName}" ensured as first tab.`);
  } catch (e) { Logger.log(`[WARN] MJM_Dashboard (GetDash): Could not move/activate dashboard sheet. Error: ${e.message}`); }
  return dashboardSheet;
}

/**
 * Gets or creates the MJM Dashboard Helper Sheet. Makes it visible for formula setting.
 * The calling function (MJM_formatDashboardSheet) is responsible for hiding it afterwards.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet The main application spreadsheet.
 * @return {GoogleAppsScript.Spreadsheet.Sheet | null} The helper sheet object or null.
 */
function MJM_getOrCreateHelperSheet(spreadsheet) {
  if (!spreadsheet) { Logger.log("[ERROR] MJM_Dashboard (GetHelper): Spreadsheet object missing."); return null; }
  // DASHBOARD_HELPER_SHEET_NAME from Global_Constants.gs
  const helperSheetName = (typeof DASHBOARD_HELPER_SHEET_NAME !== 'undefined' ? DASHBOARD_HELPER_SHEET_NAME : "DashboardHelper_FallbackName");
  const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : false);
  let helperSheet = spreadsheet.getSheetByName(helperSheetName);

  if (!helperSheet) {
    helperSheet = spreadsheet.insertSheet(helperSheetName);
    Logger.log(`[INFO] MJM_Dashboard (GetHelper): Created new helper sheet "${helperSheetName}".`);
  }
  try {
    if (helperSheet.isSheetHidden()) { // Ensure visible for formula writing
      helperSheet.showSheet();
      if (DEBUG) Logger.log(`  MJM_Dashboard (GetHelper): Helper sheet "${helperSheetName}" was hidden, now shown for setup.`);
    } else {
      if (DEBUG) Logger.log(`  MJM_Dashboard (GetHelper): Helper sheet "${helperSheetName}" is already visible.`);
    }
  } catch (e) {
    Logger.log(`[WARN] MJM_Dashboard (GetHelper): Could not ensure helper sheet "${helperSheetName}" is visible: ${e.message}`);
  }
  return helperSheet;
}

function MJM_formatDashboardSheet(dashboardSheet, appDataSheetName) {
  const functionNameForLog = "MJM_formatDashboardSheet";
  // DASHBOARD_TAB_NAME and other constants from Global_Constants.gs
  const dashboardTabNameConst = (typeof DASHBOARD_TAB_NAME !== 'undefined' ? DASHBOARD_TAB_NAME : "Dashboard_FallbackName");
  const helperSheetNameConst = (typeof DASHBOARD_HELPER_SHEET_NAME !== 'undefined' ? DASHBOARD_HELPER_SHEET_NAME : "DashboardHelper_FallbackName");
  const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : false);

  Logger.log(`[INFO] ${functionNameForLog}: Starting formatting for sheet: "${dashboardSheet ? dashboardSheet.getName() : 'NULL dashboardSheet'}", using AppDataSheetName: "${appDataSheetName}".`);

  if (!dashboardSheet || dashboardSheet.getName() !== dashboardTabNameConst) {
    const errorMsg = `[ERROR] ${functionNameForLog}: Invalid or incorrect dashboardSheet provided. Expected name: "${dashboardTabNameConst}", Got: "${dashboardSheet ? dashboardSheet.getName() : 'null'}". Aborting format.`;
    Logger.log(errorMsg);
    return; 
  }
  if (!appDataSheetName || typeof appDataSheetName !== 'string' || appDataSheetName.trim() === "") {
    const errorMsg = `[ERROR] ${functionNameForLog}: appDataSheetName (for 'Applications' data source) not provided or invalid. Cannot create formulas. Aborting format.`;
    Logger.log(errorMsg);
    return;
  }

  let helperSheet = null; 
  const ss = dashboardSheet.getParent(); 
  Logger.log(`  [INFO] ${functionNameForLog}: Parent spreadsheet obtained: "${ss.getName()}". Attempting to get helper sheet: "${helperSheetNameConst}".`);
  
  try {
    helperSheet = ss.getSheetByName(helperSheetNameConst);
    if (helperSheet) {
      Logger.log(`  [INFO] ${functionNameForLog}: Helper sheet "${helperSheet.getName()}" found.`);
      if (helperSheet.isSheetHidden()) {
        helperSheet.showSheet();
        if (DEBUG) Logger.log(`    ${functionNameForLog}: Helper sheet "${helperSheet.getName()}" was hidden, now shown for setup.`);
      }
    } else {
      Logger.log(`  [ERROR] ${functionNameForLog}: Helper sheet "${helperSheetNameConst}" NOT FOUND.`);
    }
  } catch (eGetHelper) {
      Logger.log(`  [ERROR] ${functionNameForLog}: Exception while trying to get/show helper sheet "${helperSheetNameConst}": ${eGetHelper.toString()}`);
  }

  // 1. Robustly Clear and Prepare Dashboard Sheet
  if (DEBUG) Logger.log(`  ${functionNameForLog}: Step 1 - Showing all, clearing charts, ensuring min dimensions for Dashboard sheet.`);
  try { if (dashboardSheet.getMaxRows() > 0) dashboardSheet.showRows(1, dashboardSheet.getMaxRows()); SpreadsheetApp.flush(); } catch (e) { Logger.log(`[WARN] ${functionNameForLog} (DashShowRowsErr): ${e.message}`); }
  try { if (dashboardSheet.getMaxColumns() > 0) dashboardSheet.showColumns(1, dashboardSheet.getMaxColumns()); SpreadsheetApp.flush(); } catch (e) { Logger.log(`[WARN] ${functionNameForLog} (DashShowColsErr): ${e.message}`); }

  const existingChartsOnDashboard = dashboardSheet.getCharts();
  if (existingChartsOnDashboard && existingChartsOnDashboard.length > 0) {
    if (DEBUG) Logger.log(`    Found ${existingChartsOnDashboard.length} existing charts on dashboard. Removing them...`);
    for (let i = 0; i < existingChartsOnDashboard.length; i++) {
      try { dashboardSheet.removeChart(existingChartsOnDashboard[i]); } 
      catch (eChartRemove) { Logger.log(`[WARN] ${functionNameForLog}: Could not remove an existing chart: ${eChartRemove.message}`); }
    }
    SpreadsheetApp.flush(); 
    if (DEBUG) Logger.log("    All existing charts removed from dashboard.");
  } else {
    if (DEBUG) Logger.log("    No existing charts found on dashboard to remove.");
  }
  
  const MIN_DASH_ROWS_NEEDED = 50;
  const MIN_DASH_COLS_NEEDED = 13; // A to M
  if (dashboardSheet.getMaxRows() < MIN_DASH_ROWS_NEEDED) dashboardSheet.insertRowsAfter(Math.max(1, dashboardSheet.getMaxRows()), MIN_DASH_ROWS_NEEDED - Math.max(1, dashboardSheet.getMaxRows()));
  if (dashboardSheet.getMaxColumns() < MIN_DASH_COLS_NEEDED) dashboardSheet.insertColumnsAfter(Math.max(1, dashboardSheet.getMaxColumns()), MIN_DASH_COLS_NEEDED - Math.max(1, dashboardSheet.getMaxColumns()));
  if (DEBUG) Logger.log(`    Dashboard sheet dimensions ensured: Min Rows=${MIN_DASH_ROWS_NEEDED}, Min Cols=${MIN_DASH_COLS_NEEDED}.`);

  try { dashboardSheet.setHiddenGridlines(true); } catch (e) { Logger.log(`[WARN] ${functionNameForLog}: Gridline hide error: ${e.message}`); }

  // 2. Styling & Layout Constants (Define these at the top or ensure they are globally accessible if moved to Global_Constants.gs)
  const TEAL_ACCENT_BG = "#26A69A"; 
  const HEADER_TEXT_COLOR = "#FFFFFF"; 
  const LIGHT_GREY_CARD_BG = "#F5F5F5";
  const DARK_GREY_TEXT = "#424242"; 
  const CARD_BORDER_COLOR = "#BDBDBD"; 
  const VALUE_TEXT_COLOR = TEAL_ACCENT_BG;
  const METRIC_FONT_SIZE = 15; 
  const METRIC_FONT_WEIGHT = "bold"; 
  const LABEL_FONT_WEIGHT = "bold";
  const MANUAL_REVIEW_VALUE_COLOR = (typeof MJM_MANUAL_REVIEW_VALUE_COLOR !== 'undefined') ? MJM_MANUAL_REVIEW_VALUE_COLOR : "#FF8F00"; // From potential MJM_Config or default
  const PALE_YELLOW_BG = (typeof MJM_PALE_YELLOW_BG !== 'undefined') ? MJM_PALE_YELLOW_BG : "#FFFDE7"; // Example: Or #FFF9C4
  const PALE_ORANGE_BG = (typeof MJM_PALE_ORANGE_BG !== 'undefined') ? MJM_PALE_ORANGE_BG : "#fff3e0"; // Example: Or #fff3e0

 const DASH_LAYOUT_SPACER_COL_A_WIDTH = 20, DASH_LAYOUT_LABEL_COL_WIDTH = 150, 
        DASH_LAYOUT_VALUE_COL_WIDTH = 75, DASH_LAYOUT_INTER_CARD_SPACER_WIDTH = 15;
  
  const DASH_HEADER_LAST_MERGE_COLUMN_INDEX = 13; // Column M (NEW constant for header span)
  const DASH_CONTENT_LAST_COL_INDEX = 12;         // Actual content typically up to Column L
  const DASH_CONTENT_VISIBLE_ROWS = 45;   

  // 3. Main Dashboard Header
  if (DEBUG) Logger.log(`  ${functionNameForLog}: Setting main header.`);
  dashboardSheet.getRange(1, 1, 1, DASH_HEADER_LAST_MERGE_COLUMN_INDEX).merge() 
    .setValue("MJM - Job Application Dashboard").setBackground(TEAL_ACCENT_BG).setFontColor(HEADER_TEXT_COLOR)
    .setFontSize(18).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle");
  dashboardSheet.setRowHeight(1, 45); dashboardSheet.setRowHeight(2, 10);

  // 4. Scorecard Section Title
  dashboardSheet.getRange("B3").setValue("Key Metrics Overview:").setFontSize(14).setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT);
  dashboardSheet.setRowHeight(3, 30); dashboardSheet.setRowHeight(4, 10);

  // 5. Prepare Column Letter References for Formulas
  const appShtFormulaRef = `'${appDataSheetName}'`; 
  const appCompColLetter = MJM_columnToLetter(MJM_APP_COMPANY_COL);
  const appStatColLetter = MJM_columnToLetter(MJM_APP_STATUS_COL);
  const appPeakColLetter = MJM_columnToLetter(MJM_APP_PEAK_STATUS_COL);
  const appEmailDateColLetter = MJM_columnToLetter(MJM_APP_EMAIL_DATE_COL); 
  const appPlatformColLetter = MJM_columnToLetter(MJM_APP_PLATFORM_COL);   
  const appJobTitleColLetter = MJM_columnToLetter(MJM_APP_JOB_TITLE_COL);   

  // 6. Set Scorecard Formulas & Basic Styling 
  if (DEBUG) Logger.log(`  ${functionNameForLog}: Setting scorecard formulas and labels.`);
  
  // --- Row 5 ---
  dashboardSheet.getRange("B5").setValue("Total Applications"); 
  dashboardSheet.getRange("C5").setFormula(`=IFERROR(COUNTA(${appShtFormulaRef}!${appCompColLetter}2:${appCompColLetter}), 0)`);
  dashboardSheet.getRange("E5").setValue("Peak Interviews"); 
  dashboardSheet.getRange("F5").setFormula(`=IFERROR(COUNTIF(${appShtFormulaRef}!${appPeakColLetter}2:${appPeakColLetter},"${MJM_APP_INTERVIEW_STATUS}"), 0)`);
  dashboardSheet.getRange("H5").setValue("Interview Rate (Peak)"); 
  dashboardSheet.getRange("I5").setFormula(`=IFERROR(F5/C5, 0)`); 
  dashboardSheet.getRange("K5").setValue("Offer Rate (Peak)"); 
  dashboardSheet.getRange("L5").setFormula(`=IFERROR(F7/C5, 0)`);

  // --- Row 7 ---
  dashboardSheet.getRange("B7").setValue("Active Applications"); 
  dashboardSheet.getRange("C7").setFormula(`=IFERROR(COUNTIFS(${appShtFormulaRef}!${appStatColLetter}2:${appStatColLetter}, "<>"&"", ${appShtFormulaRef}!${appStatColLetter}2:${appStatColLetter}, "<>${MJM_APP_REJECTED_STATUS}", ${appShtFormulaRef}!${appStatColLetter}2:${appStatColLetter}, "<>${MJM_APP_ACCEPTED_STATUS}"), 0)`);
  dashboardSheet.getRange("E7").setValue("Peak Offers"); 
  dashboardSheet.getRange("F7").setFormula(`=IFERROR(COUNTIF(${appShtFormulaRef}!${appPeakColLetter}2:${appPeakColLetter},"${MJM_APP_OFFER_STATUS}"), 0)`);
  dashboardSheet.getRange("H7").setValue("Current Interviews"); 
  dashboardSheet.getRange("I7").setFormula(`=IFERROR(COUNTIF(${appShtFormulaRef}!${appStatColLetter}2:${appStatColLetter},"${MJM_APP_INTERVIEW_STATUS}"), 0)`);
  dashboardSheet.getRange("K7").setValue("Current Assessments"); 
  dashboardSheet.getRange("L7").setFormula(`=IFERROR(COUNTIF(${appShtFormulaRef}!${appStatColLetter}2:${appStatColLetter},"${MJM_APP_ASSESSMENT_STATUS}"), 0)`);
    
  // --- Row 9 (Revised Layout) ---
  dashboardSheet.getRange("B9").setValue("Total Rejections"); 
  dashboardSheet.getRange("C9").setFormula(`=IFERROR(COUNTIF(${appShtFormulaRef}!${appStatColLetter}2:${appStatColLetter},"${MJM_APP_REJECTED_STATUS}"),0)`);
  dashboardSheet.getRange("E9").setValue("Apps Viewed (Peak)"); 
  dashboardSheet.getRange("F9").setFormula(`=IFERROR(COUNTIF(${appShtFormulaRef}!${appPeakColLetter}2:${appPeakColLetter},"${MJM_APP_VIEWED_STATUS}"), 0)`);
  dashboardSheet.getRange("H9").setValue("Manual Review"); 
  const manualReviewFormula = `=IFERROR(SUM(ARRAYFORMULA(N( REGEXMATCH(TRIM(${appShtFormulaRef}!${appCompColLetter}2:${appCompColLetter}), "^${RegExp.escape(MANUAL_REVIEW_NEEDED_TEXT)}$") + REGEXMATCH(TRIM(${appShtFormulaRef}!${appJobTitleColLetter}2:${appJobTitleColLetter}), "^${RegExp.escape(MANUAL_REVIEW_NEEDED_TEXT)}$") + REGEXMATCH(TRIM(${appShtFormulaRef}!${appStatColLetter}2:${appStatColLetter}), "^${RegExp.escape(MANUAL_REVIEW_NEEDED_TEXT)}$") > 0 ))), 0)`;
  dashboardSheet.getRange("I9").setFormula(manualReviewFormula); 
  dashboardSheet.getRange("K9").setValue("Direct Reject Rate");
  const directRejectFormula = `=IFERROR(COUNTIFS(${appShtFormulaRef}!${appStatColLetter}2:${appStatColLetter},"${MJM_APP_REJECTED_STATUS}", ${appShtFormulaRef}!${appPeakColLetter}2:${appPeakColLetter},"${MJM_APP_DEFAULT_STATUS}") / C5, 0)`;
  dashboardSheet.getRange("L9").setFormula(directRejectFormula); 

  // 7. Set Scorecard Row Heights & Common Styling
  [5, 7, 9].forEach(rIdx => dashboardSheet.setRowHeight(rIdx, 40)); 
  [6, 8].forEach(rIdx => dashboardSheet.setRowHeight(rIdx, 10)); 
  dashboardSheet.setRowHeight(10, 15); 

  const scCardPairs = ["B5:C5", "E5:F5", "H5:I5", "K5:L5", "B7:C7", "E7:F7", "H7:I7", "K7:L7","B9:C9", "E9:F9", "H9:I9", "K9:L9"];

const scLblCells = [
    "B5", "E5", "H5", "K5", // K5 is back
    "B7", "E7", "H7", "K7",
    "B9", "E9", "H9", "K9" 
]; 
const scValCells = [
    "C5", "F5", "I5", "L5", // L5 is back
    "C7", "F7", "I7", "L7",
    "C9", "F9", "I9", "L9" 
];
  
scCardPairs.forEach(rgStr => 
  dashboardSheet.getRange(rgStr)
    // ... (setBackground, setBorder, setVerticalAlignment as before) ...
    .setBackground(LIGHT_GREY_CARD_BG) 
    .setBorder(true, true, true, true, true, true, CARD_BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID_THIN)
    .setVerticalAlignment("middle")
);
  
scLblCells.forEach(cAddr => 
  dashboardSheet.getRange(cAddr)
    // ... (setFontWeight, setFontColor, setHorizontalAlignment as before) ...
    .setFontWeight(LABEL_FONT_WEIGHT)
    .setFontColor(DARK_GREY_TEXT)
    .setHorizontalAlignment("left")
);
  
scValCells.forEach(cAddr => { 
  const cell = dashboardSheet.getRange(cAddr); 
  cell.setFontSize(METRIC_FONT_SIZE)
      .setFontWeight(METRIC_FONT_WEIGHT)
      .setHorizontalAlignment("center")
      .setFontColor(VALUE_TEXT_COLOR); 
  
  // Specific number formats
  if (cAddr === "I5" || cAddr === "L5" || cAddr === "L9" ) { // Rate cells (Interview Rate, Offer Rate, Direct Reject Rate)
      cell.setNumberFormat("0.00%"); 
  } else { 
      cell.setNumberFormat("0");
  }
});
  
// --- APPLY CUSTOM COLORS (remains the same) ---
dashboardSheet.getRange("H9:I9").setBackground(PALE_YELLOW_BG); // Manual Review
dashboardSheet.getRange("I9").setFontColor(MANUAL_REVIEW_VALUE_COLOR); 
dashboardSheet.getRange("K9:L9").setBackground(PALE_ORANGE_BG); // Direct Reject Rate
    
  // APPLY CUSTOM COLORS
  dashboardSheet.getRange("H9:I9").setBackground(PALE_YELLOW_BG); // Manual Review
  dashboardSheet.getRange("I9").setFontColor(MANUAL_REVIEW_VALUE_COLOR); 
  dashboardSheet.getRange("K9:L9").setBackground(PALE_ORANGE_BG); // Direct Reject Rate

  // 8. Chart Section Titles
  dashboardSheet.getRange("B11").setValue("Application Platform & Weekly Trends").setFontSize(12).setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT);
  dashboardSheet.setRowHeight(11, 25); dashboardSheet.setRowHeight(12, 5);
  dashboardSheet.getRange("B28").setValue("Application Funnel Analysis").setFontSize(12).setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT);
  dashboardSheet.setRowHeight(28, 25); dashboardSheet.setRowHeight(29, 5);

  // 9. Set Dashboard Column Widths
  dashboardSheet.setColumnWidth(1, DASH_LAYOUT_SPACER_COL_A_WIDTH); 
  [2, 5, 8, 11].forEach(cI => dashboardSheet.setColumnWidth(cI, DASH_LAYOUT_LABEL_COL_WIDTH)); 
  [3, 6, 9, 12].forEach(cI => dashboardSheet.setColumnWidth(cI, DASH_LAYOUT_VALUE_COL_WIDTH));  
  [4, 7, 10].forEach(cI => dashboardSheet.setColumnWidth(cI, DASH_LAYOUT_INTER_CARD_SPACER_WIDTH)); 
  dashboardSheet.setColumnWidth(13, DASH_LAYOUT_SPACER_COL_A_WIDTH); 

  // 10. Setup Helper Sheet Formulas
  if (helperSheet) { 
    if (DEBUG) Logger.log(`  ${functionNameForLog}: Setting precise formulas in helper sheet "${helperSheet.getName()}". Source sheet: '${appDataSheetName}'`);
    helperSheet.clearContents(); 

    helperSheet.getRange("A1").setValue("Platform"); helperSheet.getRange("B1").setValue("Count");
    helperSheet.getRange("D1").setValue("Week Starting"); helperSheet.getRange("E1").setValue("Applications");
    helperSheet.getRange("G1").setValue("Stage"); helperSheet.getRange("H1").setValue("Count");
    helperSheet.getRange("J1").setValue("RAW_VALID_DATES_FOR_WEEKLY"); helperSheet.getRange("K1").setValue("CALCULATED_WEEK_STARTS");
    helperSheet.getRange("A1:K1").setFontWeight("bold");
    
    helperSheet.getRange("A2").setFormula(`=IFERROR(QUERY(${appShtFormulaRef}!${appPlatformColLetter}2:${appPlatformColLetter}, "SELECT ${appPlatformColLetter}, COUNT(${appPlatformColLetter}) WHERE ${appPlatformColLetter} IS NOT NULL AND ${appPlatformColLetter} <> '' GROUP BY ${appPlatformColLetter} ORDER BY COUNT(${appPlatformColLetter}) DESC LABEL ${appPlatformColLetter} '', COUNT(${appPlatformColLetter}) ''", 0), {"No Platforms",0})`);
  
    helperSheet.getRange("J2").setFormula(`=IFERROR(FILTER(${appShtFormulaRef}!${appEmailDateColLetter}2:${appEmailDateColLetter}, ISNUMBER(${appShtFormulaRef}!${appEmailDateColLetter}2:${appEmailDateColLetter})), {"";""})`);
    helperSheet.getRange("J2:J").setNumberFormat("yyyy-mm-dd hh:mm:ss");
    
    helperSheet.getRange("K2").setFormula(`=ARRAYFORMULA(IF(ISBLANK(J2:J), "", IF(IFERROR(VALUE(J2:J), 0) > 0, DATE(YEAR(J2:J), MONTH(J2:J), DAY(J2:J) - WEEKDAY(J2:J, 2) + 1), "")))`);
    helperSheet.getRange("K2:K").setNumberFormat("yyyy-mm-dd");
    
    helperSheet.getRange("D2").setFormula(`=IFERROR(SORT(UNIQUE(FILTER(K2:K, K2:K<>""))), {"No Dates";""})`);
    helperSheet.getRange("D2:D").setNumberFormat("yyyy-mm-dd");
    helperSheet.getRange("E2").setFormula(`=ARRAYFORMULA(IF(TRIM(D2:D)="", "", COUNTIF(K2:K, TEXT(D2:D,"yyyy-mm-dd"))))`);
    helperSheet.getRange("E2:E").setNumberFormat("0");

    const funnelStageNames = [MJM_APP_DEFAULT_STATUS, MJM_APP_VIEWED_STATUS, MJM_APP_ASSESSMENT_STATUS, MJM_APP_INTERVIEW_STATUS, MJM_APP_OFFER_STATUS];
    helperSheet.getRange(2, 7, funnelStageNames.length, 1).setValues(funnelStageNames.map(s => [s]));
    helperSheet.getRange("H2").setFormula(`=IFERROR(COUNTA(${appShtFormulaRef}!${appCompColLetter}2:${appCompColLetter}), 0)`);
    for (let i = 1; i < funnelStageNames.length; i++) { 
      helperSheet.getRange(i + 2, 8).setFormula(`=IFERROR(COUNTIF(${appShtFormulaRef}!${appPeakColLetter}2:${appPeakColLetter}, G${i + 2}), 0)`);
    }

    if (DEBUG) Logger.log(`  ${functionNameForLog}: Helper sheet formulas set successfully.`);
    Utilities.sleep(1500); 
    try { 
        if (!helperSheet.isSheetHidden()) { 
            helperSheet.hideSheet(); 
            if (DEBUG) Logger.log(`  ${functionNameForLog}: Helper sheet "${helperSheet.getName()}" successfully hidden.`); 
        } else {
            if (DEBUG) Logger.log(`  ${functionNameForLog}: Helper sheet "${helperSheet.getName()}" was already hidden.`);
        }
    } catch (eHide) { 
        Logger.log(`[WARN] ${functionNameForLog}: Could not hide helper sheet "${helperSheet.getName()}": ${eHide.message}`); 
    }
  } else { 
    const errorMsgForHelper = `[CRITICAL ERROR] ${functionNameForLog}: Helper sheet ("${helperSheetNameConst}") was NOT AVAILABLE when attempting to set formulas. Formulas NOT set. Dashboard will be incorrect.`;
    Logger.log(errorMsgForHelper);
    // Consider throwing an error if this state is absolutely critical for subsequent steps.
    // throw new Error(errorMsgForHelper); 
  }

  // 11. Final Dashboard Visibility
  if (DEBUG) Logger.log(`  ${functionNameForLog}: Setting Dashboard visibility.`);
  const currentMaxDashRows = dashboardSheet.getMaxRows();
  if (currentMaxDashRows > DASH_CONTENT_VISIBLE_ROWS) {
    try { dashboardSheet.hideRows(DASH_CONTENT_VISIBLE_ROWS + 1, currentMaxDashRows - DASH_CONTENT_VISIBLE_ROWS); }
    catch (eHR) { Logger.log(`[WARN] ${functionNameForLog} (Dash Format - HideRows): ${eHR.message}`); }
  }
  const currentMaxDashCols = dashboardSheet.getMaxColumns();
  const lastVisibleColPlusSpacer = DASH_CONTENT_LAST_COL_INDEX + 1; 
  if (currentMaxDashCols > lastVisibleColPlusSpacer) {
    try { dashboardSheet.hideColumns(lastVisibleColPlusSpacer + 1, currentMaxDashCols - lastVisibleColPlusSpacer); }
    catch (eHC) { Logger.log(`[WARN] ${functionNameForLog} (Dash Format - HideCols): ${eHC.message}`); }
  }
  Logger.log(`[INFO] ${functionNameForLog}: Formatting completed for sheet "${dashboardSheet.getName()}".`);
} // End of MJM_formatDashboardSheet

/**
 * Main function to orchestrate updates to dashboard elements and charts.
 * It ensures the helper sheet formulas have had a chance to calculate, then
 * rebuilds/updates the charts to point to the (potentially new) data in the helper sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet The main application spreadsheet.
 */
function MJM_updateDashboardMetrics(spreadsheet) {
  if (!spreadsheet) { Logger.log("[ERROR] MJM_Dashboard (UpdateMetrics): Spreadsheet object missing."); return; }
  const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : false);
  const SCRIPT_START_TIME_METRICS = new Date();
  if (DEBUG) Logger.log(`\n==== STARTING MJM DASHBOARD METRICS & CHART UPDATE (${SCRIPT_START_TIME_METRICS.toLocaleTimeString()}) ====`);

  // DASHBOARD_TAB_NAME, DASHBOARD_HELPER_SHEET_NAME from Global_Constants.gs
  const dashboardSheet = spreadsheet.getSheetByName(DASHBOARD_TAB_NAME);
  const helperSheet = spreadsheet.getSheetByName(DASHBOARD_HELPER_SHEET_NAME);

  if (!dashboardSheet) { Logger.log(`[ERROR] MJM_Dashboard (UpdateMetrics): Dashboard sheet "${DASHBOARD_TAB_NAME}" MISSING. Cannot update charts.`); return; }
  if (!helperSheet) { Logger.log(`[ERROR] MJM_Dashboard (UpdateMetrics): Helper sheet "${DASHBOARD_HELPER_SHEET_NAME}" MISSING. Chart data sources likely invalid.`); return; }

  // Ensure helper sheet is temporarily visible if charts need to read from it and it was hidden
  let helperWasHidden = false;
  try {
    if (helperSheet.isSheetHidden()) {
      helperSheet.showSheet();
      helperWasHidden = true;
      SpreadsheetApp.flush(); // Try to ensure sheet is visible before chart access
      Utilities.sleep(500); // Give it a moment
      if (DEBUG) Logger.log(`  MJM_Dashboard (UpdateMetrics): Helper sheet "${helperSheet.getName()}" was hidden, temporarily shown for chart updates.`);
    }
  } catch (eShow) {
    Logger.log(`[WARN] MJM_Dashboard (UpdateMetrics): Could not show helper sheet "${helperSheet.getName()}" for chart updates: ${eShow.message}. Charts may fail.`);
  }

  if (DEBUG) Logger.log(`  MJM_Dashboard (UpdateMetrics): Scorecards on Dashboard & Helper sheet data are formula-driven. Ensuring chart objects are correctly configured...`);
  try {
    // Call individual chart update functions, passing the sheet objects
    MJM_updatePlatformDistributionChart(dashboardSheet, helperSheet);
    MJM_updateApplicationsOverTimeChart(dashboardSheet, helperSheet);
    MJM_updateApplicationFunnelChart(dashboardSheet, helperSheet);

    if (DEBUG) Logger.log(`  MJM_Dashboard (UpdateMetrics): Chart object update/creation process complete.`);
  } catch (e) {
    Logger.log(`[ERROR] MJM_Dashboard (UpdateMetrics): Error during chart update/creation calls: ${e.toString()}\nStack: ${e.stack}`);
  }

  // Re-hide helper sheet if it was temporarily shown
  if (helperWasHidden) {
    try {
      Utilities.sleep(500); // Give charts a moment to render if they depend on visible data (less likely for GAS charts)
      if (!helperSheet.isSheetHidden()) helperSheet.hideSheet();
      if (DEBUG) Logger.log(`  MJM_Dashboard (UpdateMetrics): Helper sheet "${helperSheet.getName()}" re-hidden.`);
    } catch (eHide) {
      Logger.log(`[WARN] MJM_Dashboard (UpdateMetrics): Could not re-hide helper sheet after chart updates: ${eHide.message}`);
    }
  }
  if (DEBUG) Logger.log(`\n==== MJM DASHBOARD METRICS & CHART UPDATE FINISHED (Total time: ${(new Date().getTime() - SCRIPT_START_TIME_METRICS.getTime()) / 1000}s) ====`);
}

// --- INDIVIDUAL CHART UPDATE/CREATION FUNCTIONS ---

/**
 * Updates or creates the Platform Distribution Pie Chart on the dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboardSheet The dashboard sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} helperSheet The helper sheet containing chart data.
 */
function MJM_updatePlatformDistributionChart(dashboardSheet, helperSheet) {
  const CHART_TITLE = "Platform Distribution (MJM)";
  const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : false); // From Global_Constants.gs
  const functionNameForLog = "MJM_updatePlatformDistributionChart";

  if (DEBUG) Logger.log(`  [INFO] ${functionNameForLog}: Starting update/creation for chart: "${CHART_TITLE}"`);

  if (!dashboardSheet || typeof dashboardSheet.getCharts !== 'function') {
      Logger.log(`[ERROR] ${functionNameForLog}: Invalid dashboardSheet object provided.`);
      return;
  }
  if (!helperSheet || typeof helperSheet.getRange !== 'function') {
      Logger.log(`[ERROR] ${functionNameForLog}: Invalid helperSheet object provided.`);
      return;
  }

  // --- Validate Data Source ---
  // Expects A1 = "Platform", B1 = "Count" in helperSheet
  const headerA1 = String(helperSheet.getRange("A1").getValue()).trim();
  const headerB1 = String(helperSheet.getRange("B1").getValue()).trim();

  if (headerA1.toUpperCase() !== "PLATFORM" || headerB1.toUpperCase() !== "COUNT") {
    Logger.log(`[WARN] ${functionNameForLog} ("${CHART_TITLE}"): Helper sheet headers incorrect. Expected A1='Platform', B1='Count'. Found A1='${headerA1}', B1='${headerB1}'. Chart might be incorrect or not generate.`);
    // Optionally, remove an existing chart if headers are critically wrong, or just let it try and potentially fail.
    // For now, we'll let it proceed and rely on lastDataRowInHelperColA check.
  }

  // Determine the last row of actual data in column A (Platform names) of the helper sheet
  const colAValues = helperSheet.getRange("A1:A" + helperSheet.getMaxRows()).getValues(); // Get all values in Col A
  let lastDataRowInHelperColA = 0;
  for (let i = colAValues.length - 1; i >= 0; i--) { // Iterate backwards
    if (colAValues[i][0] !== null && String(colAValues[i][0]).trim() !== "") {
      lastDataRowInHelperColA = i + 1; // i is 0-based, row is 1-based
      break;
    }
  }
  
  if (DEBUG) Logger.log(`    ${functionNameForLog} ("${CHART_TITLE}"): Determined last data row in HelperSheet!A as: ${lastDataRowInHelperColA}`);

  if (lastDataRowInHelperColA < 2) { // Need at least a header row (row 1) and one data row (row 2)
    if (DEBUG) Logger.log(`    ${functionNameForLog} ("${CHART_TITLE}"): Insufficient data for chart. Found ${lastDataRowInHelperColA} relevant row(s) in HelperSheet Col A (need >=2 for header + data). Chart will be removed/not created.`);
    // If data is bad/missing, remove any existing chart with this title to avoid confusion.
    dashboardSheet.getCharts().forEach(chart => { 
        if (chart.getOptions().get('title') === CHART_TITLE) {
            try { dashboardSheet.removeChart(chart); } catch(eRm) { Logger.log(`      Error removing old chart: ${eRm}`);}
        }
    });
    return; // Exit if no data to plot beyond headers
  }

  // Define the data range: Column A (Platform) and Column B (Count) from row 1 to lastDataRowInHelperColA
  const dataRange = helperSheet.getRange(1, 1, lastDataRowInHelperColA, 2); // e.g., A1:B5
  
  if (DEBUG) {
    Logger.log(`    ${functionNameForLog} ("${CHART_TITLE}"): Data range for chart: ${helperSheet.getName()}!${dataRange.getA1Notation()}`);
    const platformRangeValues = dataRange.getValues(); // Get values for logging
    Logger.log(`    ${functionNameForLog} ("${CHART_TITLE}"): Platform Data to be plotted (all rows): ${JSON.stringify(platformRangeValues)}`);
    if (platformRangeValues.length > 0 && platformRangeValues[0].length > 1) {
        Logger.log(`      Header for X-axis (A1 Expected: "Platform"): ${platformRangeValues[0][0]} (Type: ${typeof platformRangeValues[0][0]})`);
        Logger.log(`      Header for Y-axis (B1 Expected: "Count"): ${platformRangeValues[0][1]} (Type: ${typeof platformRangeValues[0][1]})`);
    }
    if (platformRangeValues.length > 1 && platformRangeValues[1] && platformRangeValues[1].length > 1) {
        Logger.log(`      First data platform (A2): ${platformRangeValues[1][0]} (Type: ${typeof platformRangeValues[1][0]})`);
        Logger.log(`      First data count (B2): ${platformRangeValues[1][1]} (Type: ${typeof platformRangeValues[1][1]})`);
    }
  }

  // Chart placement and size (these should match how you formatted the dashboard visually)
  const anchorRow = 13;    // Row where the top-left of the chart is anchored
  const anchorCol = 2;     // Column where the top-left of the chart is anchored (e.g., B=2)
  const chartWidth = 460;  // Adjust as needed
  const chartHeight = 280; // Adjust as needed

  // Find if a chart with this title and position already exists
  let existingChart = null;
  const charts = dashboardSheet.getCharts();
  for (let i = 0; i < charts.length; i++) {
    const chart = charts[i];
    const options = chart.getOptions();
    const containerInfo = chart.getContainerInfo();
    if (options.get('title') === CHART_TITLE && 
        containerInfo && // Ensure containerInfo is not null
        containerInfo.getAnchorRow() === anchorRow && 
        containerInfo.getAnchorColumn() === anchorCol) {
      existingChart = chart;
      if (DEBUG) Logger.log(`    ${functionNameForLog} ("${CHART_TITLE}"): Found existing chart to modify.`);
      break;
    }
  }

  const chartBuilder = (existingChart ? existingChart.modify() : dashboardSheet.newChart())
    .setChartType(Charts.ChartType.PIE)
    .addRange(dataRange) // This adds the data from HelperSheet!A1:B<lastDataRow>
    .setOption('title', CHART_TITLE)
    .setOption('pieHole', 0.4) // For a donut chart effect
    .setOption('width', chartWidth)
    .setOption('height', chartHeight)
    .setOption('legend', { position: Charts.Position.RIGHT, textStyle: { fontSize: 10 } })
    .setOption('pieSliceText', 'percentage') // Show percentage on slices
    .setOption('sliceVisibilityThreshold', 0) // Show all slices, even small ones
    // Optional: Define custom colors for slices if needed (matches number of data rows - 1)
    // .setOption('colors', ['#4285F4', '#DB4437', '#F4B400', '#0F9D58', '#AB47BC']) // Example colors
    .setPosition(anchorRow, anchorCol, 0, 0); // Anchor to cell B13 (row 13, col 2), no offset

  try {
    if (existingChart) {
      dashboardSheet.updateChart(chartBuilder.build());
      if(DEBUG) Logger.log(`  [SUCCESS] ${functionNameForLog}: Chart "${CHART_TITLE}" UPDATED successfully.`);
    } else {
      dashboardSheet.insertChart(chartBuilder.build());
      if(DEBUG) Logger.log(`  [SUCCESS] ${functionNameForLog}: Chart "${CHART_TITLE}" CREATED successfully.`);
    }
  } catch (e) {
    Logger.log(`[ERROR] ${functionNameForLog} ("${CHART_TITLE}"): Failed to build/insert/update chart. Error: ${e.message}\nStack: ${e.stack || 'No stack'}`);
  }
} // End of MJM_updatePlatformDistributionChart

/**
 * Updates or creates the Applications Over Time Line Chart on the dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboardSheet The dashboard sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} helperSheet The helper sheet.
 */
function MJM_updateApplicationsOverTimeChart(dashboardSheet, helperSheet) {
  const CHART_TITLE = "Applications Over Time (MJM Weekly)";
  const DEBUG = GLOBAL_DEBUG_MODE;
  if (DEBUG) Logger.log(`    Updating/Creating Chart: ${CHART_TITLE}`);

  // Data for this chart from Helper Sheet: D1="Week Starting", E1="Applications". Data starts D2:E2
  const lastDataRowInHelperColD = helperSheet.getRange("D:D").getValues().filter(String).length; // Count of non-empty cells in Col D

  if (lastDataRowInHelperColD < 2 || // Need header + at least one data row
    String(helperSheet.getRange("D1").getValue()).trim() !== "Week Starting" ||
    String(helperSheet.getRange("E1").getValue()).trim() !== "Applications") {
    if (DEBUG) Logger.log(`      ${CHART_TITLE}: Insufficient/invalid data in "${helperSheet.getName()}" Col D/E. Check headers D1,E1 & data rows. Found ${lastDataRowInHelperColD} rows in D.`);
    dashboardSheet.getCharts().forEach(chart => { if (chart.getOptions().get('title') === CHART_TITLE) dashboardSheet.removeChart(chart); });
    return;
  }
  const dataRange = helperSheet.getRange(1, 4, lastDataRowInHelperColD, 2); // Range D1:E<last_data_row_in_D>
  if (DEBUG) Logger.log(`      ${CHART_TITLE}: Data range for chart: ${helperSheet.getName()}!${dataRange.getA1Notation()}`);

  const anchorRow = 13, anchorCol = 8, chartWidth = 460, chartHeight = 280; // Target: H13

  let existingChart = dashboardSheet.getCharts().find(c =>
    c.getOptions().get('title') === CHART_TITLE &&
    c.getContainerInfo().getAnchorRow() === anchorRow &&
    c.getContainerInfo().getAnchorColumn() === anchorCol
  );

  const chartBuilder = (existingChart ? existingChart.modify() : dashboardSheet.newChart())
    .setChartType(Charts.ChartType.LINE)
    .addRange(dataRange)
    .setOption('title', CHART_TITLE)
    .setOption('hAxis', { title: 'Week Starting', format: 'M/d', textStyle: { fontSize: 10 } })
    .setOption('vAxis', { title: 'No. of Applications', viewWindow: { min: 0, max: 150 }, textStyle: { fontSize: 10 } })
    .setOption('legend', { position: 'none' })
    .setOption('colors', ['#26A69A']) // Use a color from your theme
    .setOption('width', chartWidth)
    .setOption('height', chartHeight)
    .setPosition(anchorRow, anchorCol, 0, 0);

  try {
    if (existingChart) dashboardSheet.updateChart(chartBuilder.build());
    else dashboardSheet.insertChart(chartBuilder.build());
    Logger.log(`[INFO] MJM_Dashboard (Chart): "${CHART_TITLE}" ${existingChart ? 'updated' : 'created'}.`);
  } catch (e) {
    Logger.log(`[ERROR] MJM_Dashboard (Chart - ${CHART_TITLE}): Failed to build/insert/update: ${e.message}\n${e.stack || ''}`);
  }
}

/**
 * Updates or creates the Application Funnel Column Chart on the dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboardSheet The dashboard sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} helperSheet The helper sheet.
 */
function MJM_updateApplicationFunnelChart(dashboardSheet, helperSheet) {
  const CHART_TITLE = "Application Funnel (MJM Peak Stages)";
  const DEBUG = GLOBAL_DEBUG_MODE;
  if (DEBUG) Logger.log(`    Updating/Creating Chart: ${CHART_TITLE}`);

  // Data for this chart from Helper Sheet: G1="Stage", H1="Count". Data from G2:H...
  const lastDataRowInHelperColG = helperSheet.getRange("G:G").getValues().filter(String).length;
  if (lastDataRowInHelperColG < 2 || String(helperSheet.getRange("G1").getValue()).trim() !== "Stage") {
    if (DEBUG) Logger.log(`      ${CHART_TITLE}: Insufficient/invalid data in "${helperSheet.getName()}" Col G. Need G1='Stage' & >1 data row. Found ${lastDataRowInHelperColG} rows in G.`);
    dashboardSheet.getCharts().forEach(chart => { if (chart.getOptions().get('title') === CHART_TITLE) dashboardSheet.removeChart(chart); });
    return;
  }
  const dataRange = helperSheet.getRange(1, 7, lastDataRowInHelperColG, 2); // Range G1:H<last_data_row_in_G>
  if (DEBUG) Logger.log(`      ${CHART_TITLE}: Data range for chart: ${helperSheet.getName()}!${dataRange.getA1Notation()}`);

  const anchorRow = 30, anchorCol = 2, chartWidth = 460, chartHeight = 280; // Target: B30

  let existingChart = dashboardSheet.getCharts().find(c =>
    c.getOptions().get('title') === CHART_TITLE &&
    c.getContainerInfo().getAnchorRow() === anchorRow &&
    c.getContainerInfo().getAnchorColumn() === anchorCol
  );

  const chartBuilder = (existingChart ? existingChart.modify() : dashboardSheet.newChart())
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dataRange)
    .setOption('title', CHART_TITLE)
    .setOption('hAxis', { title: 'Application Stage', slantedText: true, slantedTextAngle: 30, textStyle: { fontSize: 10 } })
    .setOption('vAxis', { title: 'Number of Applications', viewWindow: { min: 0 }, textStyle: { fontSize: 10 } })
    .setOption('legend', { position: 'none' })
    .setOption('colors', ['#26A69A'])
    .setOption('bar', { groupWidth: '60%' })
    .setOption('width', chartWidth)
    .setOption('height', chartHeight)
    .setPosition(anchorRow, anchorCol, 0, 0);

  try {
    if (existingChart) dashboardSheet.updateChart(chartBuilder.build());
    else dashboardSheet.insertChart(chartBuilder.build());
    Logger.log(`[INFO] MJM_Dashboard (Chart): "${CHART_TITLE}" ${existingChart ? 'updated' : 'created'}.`);
  } catch (e) {
    Logger.log(`[ERROR] MJM_Dashboard (Chart - ${CHART_TITLE}): Failed to build/insert/update: ${e.message}\n${e.stack || ''}`);
  }
}
