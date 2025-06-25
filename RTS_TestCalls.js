// File: RTS_TestCalls.gs
// Description: Contains various test functions for the AI Resume Tailoring (RTS) components.
// These tests are now part of the integrated MJM project and can be run from its UI.
// The onOpen() function from the original RTS has been REMOVED, as menu creation
// is handled by a single onOpen() in MJM_UI.gs (or similar).
// Relies on constants from Global_Constants.gs and potentially RTS_Constants.gs for isolated test SSIDs.

// Helper to get the target spreadsheet ID for RTS tests.
// Prefers RTS_STANDALONE_TEST_SPREADSHEET_ID if defined in RTS_Constants.gs for isolated testing,
// otherwise falls back to the main application spreadsheet ID from Global_Constants.gs or the active sheet.
function _RTS_getTestSpreadsheetId() {
  // RTS_STANDALONE_TEST_SPREADSHEET_ID could be defined in a (now very slim) RTS_Constants.gs for this specific purpose
  if (typeof RTS_STANDALONE_TEST_SPREADSHEET_ID !== 'undefined' && RTS_STANDALONE_TEST_SPREADSHEET_ID && !RTS_STANDALONE_TEST_SPREADSHEET_ID.toUpperCase().includes("YOUR_")) {
    Logger.log(`RTS_TestCalls: Using RTS_STANDALONE_TEST_SPREADSHEET_ID: ${RTS_STANDALONE_TEST_SPREADSHEET_ID}`);
    return RTS_STANDALONE_TEST_SPREADSHEET_ID;
  } else if (APP_SPREADSHEET_ID && !APP_SPREADSHEET_ID.toUpperCase().includes("YOUR_")) { // From Global_Constants.gs
    Logger.log(`RTS_TestCalls: Using global APP_SPREADSHEET_ID: ${APP_SPREADSHEET_ID}`);
    return APP_SPREADSHEET_ID;
  } else if (SpreadsheetApp.getActiveSpreadsheet()) {
    const activeId = SpreadsheetApp.getActiveSpreadsheet().getId();
    Logger.log(`RTS_TestCalls: Using Active Spreadsheet ID: ${activeId}`);
    return activeId;
  } else {
    Logger.log("[ERROR] RTS_TestCalls: Could not determine a valid spreadsheet ID for testing.");
    SpreadsheetApp.getUi().alert("Test Error", "Spreadsheet ID for RTS tests could not be determined. Configure APP_SPREADSHEET_ID or ensure script is bound to a sheet.");
    return null;
  }
}

/**
 * LEGACY TEST: Kept for reference if original orchestrateTailoringProcess is maintained.
 * This test will likely FAIL unless RTS_orchestrateTailoringProcess (old one) is defined globally.
 */
function RTS_testFullResumeTailoring_Legacy() {
  Logger.log("--- RTS LEGACY Full Resume Tailoring Test ---");
  Logger.log("NOTE: This test relies on 'RTS_orchestrateTailoringProcess' (the original one) being defined and functional.");

  const testSpreadsheetId = _RTS_getTestSpreadsheetId();
  if (!testSpreadsheetId) return;

  // PROFILE_DATA_SHEET_NAME from Global_Constants.gs
  const masterProfile = RTS_getMasterProfileData(testSpreadsheetId, PROFILE_DATA_SHEET_NAME);
  if (!masterProfile?.personalInfo) { Logger.log("FAILURE: Failed to load Master Profile for legacy test."); return; }
  Logger.log(`SUCCESS: Master Profile loaded for legacy test. Name: ${masterProfile.personalInfo.fullName}`);

  const sampleJD = `Job Title: Legacy Test Analyst...`; // Abridged JD

  if (typeof RTS_orchestrateTailoringProcess === 'function') { // Check if the legacy function exists
    const result = RTS_orchestrateTailoringProcess(masterProfile, sampleJD); // Assumes old function name
    // ... (rest of original legacy test logic using 'result') ...
    if (result && result.tailoredResumeObject) {
      // RESUME_TEMPLATE_DOC_ID from Global_Constants.gs
      const docUrl = RTS_createFormattedResumeDoc(result.tailoredResumeObject, "Legacy Test Resume", RESUME_TEMPLATE_DOC_ID);
      Logger.log(docUrl ? `SUCCESS (Legacy Test): Doc URL: ${docUrl}` : "FAILURE (Legacy Test): Doc creation failed.");
    } else { Logger.log("FAILURE (Legacy Test): Orchestration failed."); }
  } else {
    Logger.log("SKIPPING Legacy Test: RTS_orchestrateTailoringProcess function not found.");
  }
  Logger.log("--- RTS LEGACY Full Resume Tailoring Test Finished ---");
}

/**
 * Tests ONLY document formatting using master profile data from the test spreadsheet.
 */
function RTS_testMasterProfileFormattingOnly() {
  Logger.log("--- RTS MASTER PROFILE FORMATTING ONLY Test ---");
  const testSpreadsheetId = _RTS_getTestSpreadsheetId();
  if (!testSpreadsheetId) return;

  const masterProfile = RTS_getMasterProfileData(testSpreadsheetId, PROFILE_DATA_SHEET_NAME); // Global const
  if (!masterProfile?.personalInfo) { Logger.log("FAILURE: Could not load Master Profile."); return; }
  Logger.log(`SUCCESS: Master Profile loaded for formatting test. Name: ${masterProfile.personalInfo.fullName}`);

  const docTitle = `Master Profile Format Test - ${masterProfile.personalInfo.fullName}`;
  // RESUME_TEMPLATE_DOC_ID from Global_Constants.gs
  const docUrl = RTS_createFormattedResumeDoc(masterProfile, docTitle, RESUME_TEMPLATE_DOC_ID);
  Logger.log(docUrl ? `SUCCESS: Master Profile doc created: ${docUrl}` : "FAILURE: Master Profile doc creation failed.");
  Logger.log("--- RTS MASTER PROFILE FORMATTING ONLY Test Finished ---");
}

/**
 * Tests core RTS AI functionalities (JD analysis, matching, tailoring, summary).
 */
function RTS_testCoreAIFunctionality() {
  Logger.log("--- RTS Core AI Functionality Test (Groq) ---");
  const testSpreadsheetId = _RTS_getTestSpreadsheetId();
  if (!testSpreadsheetId) return;

  const masterProfile = RTS_getMasterProfileData(testSpreadsheetId, PROFILE_DATA_SHEET_NAME); // Global const
  if (!masterProfile?.personalInfo) { Logger.log("FAILURE: Could not load Master Profile for AI tests."); return; }
  const sampleBullet = (masterProfile.sections.find(s => s.title === "EXPERIENCE")?.items[0]?.responsibilities[0]) || "Developed innovative solutions.";

  const sampleJDShort = `Job Title: AI Test Engineer. Responsibilities: Test AI models. Skills: Python, Testing.`;
  Logger.log("  Testing RTS_analyzeJobDescription...");
  const jdAnalysis = RTS_analyzeJobDescription(sampleJDShort); // RTS_ function from RTS_TailoringLogic.gs
  if (!jdAnalysis || jdAnalysis.error) { Logger.log(`FAIL: RTS_analyzeJobDescription. ${JSON.stringify(jdAnalysis)}`); return; }
  Logger.log(`  SUCCESS: JD Analysis (first 100 chars): ${JSON.stringify(jdAnalysis).substring(0,100)}...`);

  Logger.log("  Testing RTS_matchResumeSection...");
  const matchResult = RTS_matchResumeSection(sampleBullet, jdAnalysis); // RTS_ function
  Logger.log(matchResult && !matchResult.error ? `  SUCCESS: Match Score: ${matchResult.relevanceScore}` : `FAIL: RTS_matchResumeSection. ${JSON.stringify(matchResult)}`);

  Logger.log("  Testing RTS_tailorBulletPoint...");
  const tailoredBullet = RTS_tailorBulletPoint(sampleBullet, jdAnalysis, jdAnalysis.jobTitle || "Target Role"); // RTS_ function
  Logger.log(tailoredBullet && !tailoredBullet.startsWith("ERROR:") ? `  SUCCESS: Tailored: "${tailoredBullet}"` : `FAIL: RTS_tailorBulletPoint. ${tailoredBullet}`);

  Logger.log("  Testing RTS_generateTailoredSummary...");
  const summary = RTS_generateTailoredSummary(sampleBullet, jdAnalysis, masterProfile.personalInfo.fullName); // RTS_ function
  Logger.log(summary && !summary.startsWith("ERROR:") ? `  SUCCESS: Summary: "${summary}"` : `FAIL: RTS_generateTailoredSummary. ${summary}`);
  Logger.log("--- RTS Core AI Functionality Test Finished ---");
}

// --- STAGED PROCESSING TESTS (using global RTS_runStageX functions) ---

function RTS_testStage1_AnalysisAndScoring() { // Renamed from _New for consistency
  Logger.log("--- RTS Testing STAGE 1: JD Analysis & Scoring ---");
  const testSpreadsheetId = _RTS_getTestSpreadsheetId();
  if (!testSpreadsheetId) return;

  const sampleJD = `Job Title: Junior RTS Tester. Responsibilities: Execute test scripts. Requirements: Detail-oriented.`;
  Logger.log(`  Using Sample JD for Stage 1 Test. Target SS: ${testSpreadsheetId}`);
  // RTS_runStage1_AnalyzeAndScore is the global function from RTS_Main.gs
  const result = RTS_runStage1_AnalyzeAndScore(sampleJD, testSpreadsheetId);
  Logger.log(`Result from Stage 1: ${JSON.stringify(result)}`);
  if (result?.success) {
    // JD_ANALYSIS_SHEET_NAME, BULLET_SCORING_RESULTS_SHEET_NAME from Global_Constants.gs
    Logger.log(`SUCCESS: Stage 1 completed. Check sheets "${JD_ANALYSIS_SHEET_NAME}" & "${BULLET_SCORING_RESULTS_SHEET_NAME}" in Spreadsheet ID: ${testSpreadsheetId}`);
  } else Logger.log("FAILURE: Stage 1 did not complete successfully.");
}

function RTS_testStage2_TailoringSelected() {
  Logger.log("--- RTS Testing STAGE 2: Tailoring Selected Bullets ---");
  const testSpreadsheetId = _RTS_getTestSpreadsheetId();
  if (!testSpreadsheetId) return;
  Logger.log(`  IMPORTANT: Ensure Stage 1 run & bullets selected with '${USER_SELECT_YES_VALUE}' in "${BULLET_SCORING_RESULTS_SHEET_NAME}" in SS ID: ${testSpreadsheetId}`);

  // RTS_runStage2_TailorSelectedBullets from RTS_Main.gs; USER_SELECT_YES_VALUE, BULLET_SCORING_RESULTS_SHEET_NAME from Global_Constants.gs
  const result = RTS_runStage2_TailorSelectedBullets(testSpreadsheetId);
  Logger.log(`Result from Stage 2: ${JSON.stringify(result)}`);
  if (result?.success) Logger.log(`SUCCESS: Stage 2 completed. Check "TailoredBulletText(Stage2)" column in "${BULLET_SCORING_RESULTS_SHEET_NAME}".`);
  else Logger.log("FAILURE: Stage 2 did not complete successfully.");
}

function RTS_testStage3_GenerateDocument() {
  Logger.log("--- RTS Testing STAGE 3: Assemble & Generate Document ---");
  const testSpreadsheetId = _RTS_getTestSpreadsheetId();
  if (!testSpreadsheetId) return;
  Logger.log(`  IMPORTANT: Ensure Stage 1 & 2 successfully run (with selections) for SS ID: ${testSpreadsheetId}`);
  // Prerequisite check from original test can be added back here if needed

  // RTS_runStage3_BuildAndGenerateDocument from RTS_Main.gs
  const result = RTS_runStage3_BuildAndGenerateDocument(testSpreadsheetId);
  Logger.log(`Result from Stage 3: ${JSON.stringify(result?.message)} URL: ${result?.docUrl}`); // Avoid logging full resume object
  if (result?.success && result.docUrl) Logger.log(`SUCCESS: Stage 3 completed! Document URL: ${result.docUrl}`);
  else Logger.log("FAILURE: Stage 3 failed or document URL missing.");
}

// NO onOpen() function here. It's handled globally in MJM_UI.gs.
