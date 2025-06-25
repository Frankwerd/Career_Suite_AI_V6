// File: RTS_Constants.gs
// Description: Contains constants *exclusively* for standalone testing or highly specific
// internal details of the Resume Tailoring Script (RTS) components, IF ANY.
// Most operational constants for RTS are now in Global_Constants.gs as they are
// part of the integrated application's configuration or shared workflow.

// --- STANDALONE RTS TESTING CONFIGURATION (Optional - For Developer Use Only) ---
// For: RTS Core (Developer Standalone Testing ONLY - if you want to test RTS logic against a completely separate spreadsheet)
// If you don't need to test RTS against a separate spreadsheet instance, you can delete these.
// If kept, these would be used *only* by functions within RTS_TestCalls.gs that are specifically designed
// for isolated component testing and explicitly use these constants instead of the global ones.

/*
const RTS_STANDALONE_TEST_SPREADSHEET_ID = "YOUR_SEPARATE_RTS_TEST_SPREADSHEET_ID_HERE"; // << REVIEW/DELETE: ID of a spreadsheet ONLY for RTS testing
const RTS_STANDALONE_PROFILE_SHEET_NAME = "ResumeData"; // << REVIEW/DELETE: Tab name in the standalone test sheet for profile data.
                                                          // Should align with PROFILE_DATA_SHEET_NAME from Global_Constants.gs in concept.
const RTS_STANDALONE_JD_ANALYSIS_SHEET_NAME = "JDAnalysisData_Test"; // << REVIEW/DELETE: Different name to avoid clash if testing within main sheet
const RTS_STANDALONE_SCORING_SHEET_NAME = "BulletScoringResults_Test"; // << REVIEW/DELETE
*/


// --- LEGACY CONSTANTS (Commented out - For Historical Reference Only if needed) ---
// For: RTS Core (Historical Reference - All Operational Legacy Constants Should Be Obsolete)
/*
// These were constants from an older version of the RTS script.
// They are likely no longer used by the current staged processing logic.
// Kept here (commented out) only if needed for understanding legacy test calls or previous logic.
// Otherwise, they can be safely deleted.

const ORCH_API_CALL_DELAY = 10000;
const ORCH_SKIP_TAILORING_IF_SCORE_BELOW = 0.3;
const ORCH_FINAL_INCLUSION_THRESHOLD = 0.7;
const ORCH_MAX_BULLETS_TO_SCORE_PER_JOB = 3;
const ORCH_MAX_BULLETS_TO_INCLUDE_PER_JOB = 2;
// ... any other legacy ORCH_ constants from the original RTS Constants.gs ...
*/

// --- IMPORTANT NOTES ---
//
// 1. PRIMARY SPREADSHEET ID:
//    The constant 'MASTER_RESUME_SPREADSHEET_ID' from the original RTS_Constants.gs is now OBSOLETE
//    for runtime operations within the integrated MJM application. The integrated application uses
//    'APP_SPREADSHEET_ID' or 'APP_TARGET_FILENAME' from 'Global_Constants.gs' to identify the
//    single, main spreadsheet. RTS functions will be passed this spreadsheet ID.
//
// 2. SHEET NAMES (for Profile, JD Analysis, Scoring):
//    Constants like 'MASTER_RESUME_DATA_SHEET_NAME', 'JD_ANALYSIS_SHEET_NAME',
//    'BULLET_SCORING_RESULTS_SHEET_NAME' have been moved to 'Global_Constants.gs'
//    (e.g., as 'PROFILE_DATA_SHEET_NAME', 'JD_ANALYSIS_SHEET_NAME', etc.) because these sheets
//    will now exist as tabs within the main application spreadsheet.
//
// 3. RESUME/PROFILE STRUCTURE:
//    'RESUME_STRUCTURE', 'NUM_DEDICATED_BULLET_COLUMNS', and related styling constants
//    are now in 'Global_Constants.gs' (e.g., 'PROFILE_STRUCTURE') as they define a core data
//    structure and its setup, which is now part of the integrated application.
//
// 4. API KEYS & LLM MODELS:
//    GROQ/Gemini API key property names and default model names are in 'Global_Constants.gs'.
//
// 5. RTS LOGIC CONFIG (Delays, Thresholds):
//    Staging delays (API_CALL_DELAY_STAGE1_MS, etc.) and Stage 3 thresholds are now in
//    'Global_Constants.gs' (e.g., 'RTS_API_CALL_DELAY_STAGE1_MS') so they can be configured centrally.
//
// This file, 'RTS_Constants.gs', should ideally be very small or empty if all operational
// needs are met by 'Global_Constants.gs'. The commented-out sections above are examples of
// what might remain for very specific, isolated developer testing or historical reference.
