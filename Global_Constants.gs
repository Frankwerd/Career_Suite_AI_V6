// File: Global_Constants.gs
// Description: ALL Core Constants for the integrated "Comprehensive AI-Powered Job Search & Application Suite".
// Comments denote the primary module associated with each constant group.
// Master Job Manager (MJM) & Resume Tailor Script (RTS)

// --- General Application Configuration ---
// For: BOTH (Shared)
const GLOBAL_DEBUG_MODE = true; // Set to false for production to reduce logging
const APP_NAME = "Comprehensive AI Job Suite"; // << REVIEW / REPLACE >> Optional name for UI/Logs

// --- Primary Spreadsheet Identity (The ONE Spreadsheet for this entire application) ---
// For: BOTH (Shared Core of the Application)
// Option 1: If you use a FIXED ID for your main spreadsheet
const APP_SPREADSHEET_ID = ""; // << REVIEW / REPLACE: YOUR_MAIN_SPREADSHEET_ID_HERE or leave blank
// Option 2: If your script finds/creates the spreadsheet by name (used if APP_SPREADSHEET_ID is blank)
const APP_TARGET_FILENAME = "Comprehensive AI Job Suite Data"; // << REVIEW / REPLACE: Name of your single, main spreadsheet

// --- Core Sheet Tab Names (within APP_SPREADSHEET_ID or APP_TARGET_FILENAME) ---
// For: BOTH (Shared - as MJM will trigger RTS, and RTS results will be in these tabs within MJM's sheet)
const PROFILE_DATA_SHEET_NAME = "MasterProfile";          // Tab for master resume/profile data (Origin: RTS)
const JD_ANALYSIS_SHEET_NAME = "JDAnalysisData";           // Tab for storing AI analysis of Job Descriptions (Origin: RTS)
const BULLET_SCORING_RESULTS_SHEET_NAME = "BulletScoringResults"; // Tab for storing resume bullet scores (Origin: RTS)

// For: MJM Core (but RTS might indirectly use leads data if tailoring is initiated from a Lead)
const APP_TRACKER_SHEET_TAB_NAME = "Applications";         // Tab for tracking job applications
const LEADS_SHEET_TAB_NAME = "Potential Job Leads";        // Tab for incoming job leads

// For: MJM Core (Dashboard specific)
const DASHBOARD_TAB_NAME = "Dashboard";
const DASHBOARD_HELPER_SHEET_NAME = "DashboardHelperData";

// --- Master Profile / Resume Data Structure & Setup (Origin: RTS, now Global) ---
// For: BOTH (Shared - Defines profile structure; MJM will initiate its setup, RTS will parse it)
const NUM_DEDICATED_PROFILE_BULLET_COLUMNS = 3;
const PROFILE_STRUCTURE = [
  { title: "PERSONAL INFO", headers: ["Key", "Value"] },
  { title: "SUMMARY", headers: null },
  { title: "EXPERIENCE", headers: ["JobTitle", "Company", "Location", "StartDate", "EndDate"] },
  { title: "EDUCATION", headers: ["Institution", "Degree", "Location", "StartDate", "EndDate", "GPA", "RelevantCoursework"] },
  { title: "TECHNICAL SKILLS & CERTIFICATES", headers: ["CategoryName", "SkillItem", "Details", "Issuer", "IssueDate"] },
  { title: "PROJECTS", headers: ["ProjectName", "Organization", "Role", "StartDate", "EndDate", "Technologies", "GitHubName1", "GitHubURL1", "Impact", "FutureDevelopment"] },
  { title: "LEADERSHIP & UNIVERSITY INVOLVEMENT", headers: ["Organization", "Role", "Location", "StartDate", "EndDate", "Description"] },
  { title: "HONORS & AWARDS", headers: ["AwardName", "Details", "Date"] }
];
const PROFILE_SCHEMA_VERSION = "2.0.0_Integrated"; // Updated schema version
// Styling Constants for Profile Sheet Setup (if setup is managed centrally)
const PROFILE_HEADER_BACKGROUND_COLOR = "#E1BEE7"; // << REVIEW / REPLACE e.g., light purple
const PROFILE_HEADER_FONT_COLOR = "#4A148C";       // << REVIEW / REPLACE e.g., dark purple
const PROFILE_SUB_HEADER_BACKGROUND_COLOR = "#F3E5F5"; // << REVIEW / REPLACE e.g., lighter purple
const PROFILE_BORDER_COLOR = "#BDBDBD";             // << REVIEW / REPLACE e.g., grey

// --- Document Generation (Origin: RTS, now Global) ---
// For: RTS Core (but relevant globally if MJM initiates doc generation)
const RESUME_TEMPLATE_DOC_ID = "18eX765FWVBHpOZ2jzxwNdGVKSMOzmyPJmS_Kfzp8JSk"; // << REVIEW / REPLACE !!!

// --- API Key Property Names (for PropertiesService) ---
// For: BOTH (Shared - any module needing these APIs will use these property names)
const SHARED_GEMINI_API_KEY_PROPERTY = 'SHARED_GEMINI_API_KEY';
const SHARED_GROQ_API_KEY_PROPERTY = 'SHARED_GROQ_API_KEY';

// --- Default LLM Model Names ---
// For: BOTH (Shared - providing defaults if specific functions don't override)
const DEFAULT_GEMINI_MODEL = "gemini-1.5-flash-latest";
const DEFAULT_GROQ_MODEL = "gemma2-9b-it"; // << REVIEW / CONFIRM - e.g., mixtral-8x7b-32768, llama3-8b-8192, llama3-70b-8192, gemma2-9b-it

// --- Resume Tailoring Logic Configuration (Origin: RTS, now Global as part of integrated workflow) ---
// For: RTS Core (parameters for tailoring stages)
const RTS_API_CALL_DELAY_STAGE1_MS = 2000;
const RTS_API_CALL_DELAY_STAGE2_MS = 2000;
const RTS_API_CALL_DELAY_STAGE3_SUMMARY_MS = 2000;
const RTS_STAGE3_FINAL_INCLUSION_SCORE_THRESHOLD = 0.01; // Low threshold to include many scored items initially
const RTS_STAGE3_MAX_BULLETS_PER_JOB = 5;
const RTS_STAGE3_MAX_BULLETS_PER_PROJECT = 5;
const RTS_STAGE3_MAX_HIGHLIGHTS_FOR_SUMMARY = 4;

// Placeholder for user selection 'YES' value (used in BulletScoringResults sheet)
// For: RTS Core
const USER_SELECT_YES_VALUE = "YES"; // Value user types to select a bullet for tailoring

// Placeholder value if Gemini (or other parsers) cannot determine company/title/status
// For: BOTH (Shared - MJM parsing uses it, RTS parsing might implicitly align)
const MANUAL_REVIEW_NEEDED_TEXT = "N/A - Manual Review";
