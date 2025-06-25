// File: MJM_UI.gs
// Description: Contains the single onOpen(e) function for creating custom menus
// for the entire integrated "Comprehensive AI Job Suite" application.
// Also includes UI-triggered wrapper functions for multi-stage processes like RTS.

/**
 * Creates the main custom menu when the spreadsheet is opened.
 * This is the ONLY onOpen(e) function in the project.
 * @param {Object} e The event object from the onOpen trigger.
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  // APP_NAME from Global_Constants.gs, or a default.
  const topMenuName = (typeof APP_NAME !== 'undefined' ? `⚙️ ${APP_NAME}` : '⚙️ AI Job Suite Tools');

  ui.createMenu(topMenuName)
    .addSubMenu(ui.createMenu('▶️ Initial Project Setup')
      .addItem('RUN FULL PROJECT SETUP (All Modules)', 'runFullProjectInitialSetup') // From MJM_main.gs
      .addSeparator()
      .addItem('Setup: MJM App Tracker Only', 'MJM_setupAppTrackerModule') // From MJM_main.gs
      .addItem('Setup: MJM Job Leads Only', 'MJM_setupLeadsModule')       // From MJM_Leads_Main.gs
      .addItem('Setup: RTS Profile Sheet Only', 'RTS_setupMasterResumeSheet') // From RTS_SheetSetup.gs
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('MJM: Manual Processing')
      .addItem('Process Application Update Emails', 'MJM_processJobApplicationEmails') // From MJM_main.gs
      .addItem('Process Job Lead Emails', 'MJM_processJobLeads')                   // From MJM_Leads_Main.gs
      .addItem('Mark Stale Applications', 'MJM_markStaleApplicationsAsRejected')  // From MJM_main.gs
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('RTS: Resume Tailoring')
      .addItem('STEP 1: Analyze JD & Score Profile Bullets', 'RTS_triggerStage1FromUi_Wrapper') // Wrapper function defined below
      .addItem('STEP 2: Tailor Selected Bullets (After Manual YES)', 'RTS_triggerStage2FromUi_Wrapper') // Wrapper
      .addItem('STEP 3: Generate Tailored Resume Document', 'RTS_triggerStage3FromUi_Wrapper')   // Wrapper
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 Admin & Configuration')
      .addItem('Set SHARED Gemini API Key', 'setSharedGeminiApiKey_UI')    // From MJM_AdminUtils.gs
      .addItem('Set SHARED Groq API Key', 'setSharedGroqApiKey_UI')        // From MJM_AdminUtils.gs
      .addItem('Show All User Properties', 'showAllUserProperties')          // From MJM_AdminUtils.gs
      .addSeparator()
      .addItem('TEMP: Set Hardcoded Gemini Key', 'TEMPORARY_manualSetSharedGeminiApiKey') // From MJM_AdminUtils.gs
      .addItem('TEMP: Set Hardcoded Groq Key', 'TEMPORARY_manualSetSharedGroqApiKey')       // From MJM_AdminUtils.gs
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('🧪 Run System Tests')
      .addItem('RTS Test: Stage 1 (Analyze & Score)', 'RTS_testStage1_AnalysisAndScoring')  // From RTS_TestCalls.gs
      .addItem('RTS Test: Stage 2 (Tailor Selected)', 'RTS_testStage2_TailoringSelected')   // From RTS_TestCalls.gs
      .addItem('RTS Test: Stage 3 (Generate Document)', 'RTS_testStage3_GenerateDocument') // From RTS_TestCalls.gs
      .addSeparator()
      .addItem('RTS Test: Format Master Profile Doc Only', 'RTS_testMasterProfileFormattingOnly') // From RTS_TestCalls.gs
      .addItem('RTS Test: Core AI Logic (Groq)', 'RTS_testCoreAIFunctionality')                // From RTS_TestCalls.gs
      // Add MJM-specific test menu items if you create test functions for them
      // .addSeparator()
      // .addItem('MJM Test: Process Sample App Email', 'MJM_testProcessSampleAppEmail')
    )
    .addToUi();
}

// --- UI WRAPPER FUNCTIONS for RTS Workflow ---
// These handle UI prompts/alerts and then call the core RTS logic functions.

/**
 * UI-triggered wrapper for RTS Stage 1: Prompts for JD, then calls core Stage 1 logic.
 */
function RTS_triggerStage1FromUi_Wrapper() {
  const ui = SpreadsheetApp.getUi();
  const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : true);

  // Get the current spreadsheet ID (the one MJM is operating on)
  let currentSpreadsheetId = null;
  if (typeof APP_SPREADSHEET_ID !== 'undefined' && APP_SPREADSHEET_ID && !APP_SPREADSHEET_ID.toUpperCase().includes("YOUR_")) {
    currentSpreadsheetId = APP_SPREADSHEET_ID;
  } else if (SpreadsheetApp.getActiveSpreadsheet()) {
    currentSpreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  }

  if (!currentSpreadsheetId) {
    ui.alert("Error", "Could not determine the active spreadsheet ID. Please ensure the script is bound to a sheet or APP_SPREADSHEET_ID is set.");
    return;
  }
  if (DEBUG) Logger.log(`RTS_triggerStage1FromUi_Wrapper: Operating on Spreadsheet ID: ${currentSpreadsheetId}`);

  const jdResponse = ui.prompt(
    "RTS Stage 1: Job Description",
    "Paste the full Job Description text below:",
    ui.ButtonSet.OK_CANCEL
  );

  if (jdResponse.getSelectedButton() === ui.Button.OK) {
    const jdText = jdResponse.getResponseText();
    if (jdText && jdText.trim() !== "") {
      ui.alert(
        "Processing RTS Stage 1...",
        "Analyzing Job Description and scoring profile bullets against it.\nThis may take several minutes.\n\nPlease check the Execution Transcript (View > Execution transcript) for progress. You'll get another alert on completion.",
        ui.ButtonSet.OK
      );
      // Call the core Stage 1 function from RTS_Main.gs (ensure it's globally named RTS_runStage1_AnalyzeAndScore)
      const result = RTS_runStage1_AnalyzeAndScore(jdText, currentSpreadsheetId);

      if (result && result.success) {
        // PROFILE_DATA_SHEET_NAME, JD_ANALYSIS_SHEET_NAME, BULLET_SCORING_RESULTS_SHEET_NAME, USER_SELECT_YES_VALUE from Global_Constants.gs
        ui.alert("RTS Stage 1 Complete!", `${result.message}\n\nPlease review the "${BULLET_SCORING_RESULTS_SHEET_NAME}" tab in this spreadsheet and mark bullets with "${USER_SELECT_YES_VALUE}" in the 'SelectToTailor(Manual)' column for Stage 2 processing.`, ui.ButtonSet.OK);
      } else {
        ui.alert("RTS Stage 1 Failed", `Process did not complete successfully.\nMessage: ${result ? result.message : 'Unknown error.'}${result && result.details ? `\nDetails: ${JSON.stringify(result.details)}` : ""}`, ui.ButtonSet.OK);
      }
    } else {
      ui.alert("Input Error", "No Job Description text was provided for Stage 1.");
    }
  } else {
    ui.alert("Cancelled", "RTS Stage 1 process was cancelled by the user.");
  }
}

/**
 * UI-triggered wrapper for RTS Stage 2: Calls core Stage 2 logic.
 */
function RTS_triggerStage2FromUi_Wrapper() {
  const ui = SpreadsheetApp.getUi();
  const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : true);
  let currentSpreadsheetId = null; // Same logic as in Stage 1 wrapper to get ID

  if (typeof APP_SPREADSHEET_ID !== 'undefined' && APP_SPREADSHEET_ID && !APP_SPREADSHEET_ID.toUpperCase().includes("YOUR_")) currentSpreadsheetId = APP_SPREADSHEET_ID;
  else if (SpreadsheetApp.getActiveSpreadsheet()) currentSpreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  if (!currentSpreadsheetId) { ui.alert("Error", "Spreadsheet ID undetermined."); return; }
  if (DEBUG) Logger.log(`RTS_triggerStage2FromUi_Wrapper: Operating on Spreadsheet ID: ${currentSpreadsheetId}`);

  // BULLET_SCORING_RESULTS_SHEET_NAME from Global_Constants.gs
  ui.alert(
    "Processing RTS Stage 2...",
    `Tailoring bullets marked with '${USER_SELECT_YES_VALUE}' in the "${BULLET_SCORING_RESULTS_SHEET_NAME}" sheet.\nThis may take a few minutes per selected bullet.\n\nCheck Execution Transcript for progress. You'll get a completion alert.`,
    ui.ButtonSet.OK
  );
  // Call the core Stage 2 function from RTS_Main.gs
  const result = RTS_runStage2_TailorSelectedBullets(currentSpreadsheetId);

  if (result && result.success) {
    ui.alert("RTS Stage 2 Complete!", `${result.message}\n\nThe "TailoredBulletText(Stage2)" column in "${BULLET_SCORING_RESULTS_SHEET_NAME}" has been updated.`, ui.ButtonSet.OK);
  } else {
    ui.alert("RTS Stage 2 Failed", result ? result.message : "Unknown error.", ui.ButtonSet.OK);
  }
}

/**
 * UI-triggered wrapper for RTS Stage 3: Calls core Stage 3 logic.
 */
function RTS_triggerStage3FromUi_Wrapper() {
  const ui = SpreadsheetApp.getUi();
  const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : true);
  let currentSpreadsheetId = null; 
if (typeof APP_SPREADSHEET_ID !== 'undefined' && APP_SPREADSHEET_ID && !APP_SPREADSHEET_ID.toUpperCase().includes("YOUR_")) {
    currentSpreadsheetId = APP_SPREADSHEET_ID;
} else if (SpreadsheetApp.getActiveSpreadsheet()) {
    currentSpreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
}

if (!currentSpreadsheetId) { 
    ui.alert("Error", "Spreadsheet ID for Stage 3 could not be determined. Ensure the script is bound to a sheet or APP_SPREADSHEET_ID is properly configured in Global_Constants.gs.", ui.ButtonSet.OK);
    return; 
}
  if (DEBUG) Logger.log(`RTS_triggerStage3FromUi_Wrapper: Operating on Spreadsheet ID: ${currentSpreadsheetId}`);

  ui.alert(
    "Processing RTS Stage 3...",
    "Assembling final resume data and generating the Google Document.\nThis may take a moment.\n\nCheck Execution Transcript for progress. You'll get a completion alert with the document link.",
    ui.ButtonSet.OK
  );
  // Call the core Stage 3 function from RTS_Main.gs
  const result = RTS_runStage3_BuildAndGenerateDocument(currentSpreadsheetId);

  if (result && result.success && result.docUrl) {
    ui.alert("RTS Stage 3 Complete!", `Document generated: ${result.docUrl}\n\n${result.message}`, ui.ButtonSet.OK);
  } else {
    ui.alert("RTS Stage 3 Failed", `Process did not complete successfully or document URL was not returned.\nMessage: ${result ? result.message : 'Unknown error.'}`, ui.ButtonSet.OK);
  }
}
