// File: MJM_Config.gs (or MJM_Module_Config.gs)
// Description: Contains configuration constants *exclusively* for the
// Master Job Manager's Application Tracking and Lead Generation modules.
// Shared/Global constants are defined in Global_Constants.gs.

// --- Gmail Label & Filter Configuration (Specific to MJM Email Parsing Modules) ---
// For: MJM Core (App Tracker & Leads Modules)

const MJM_MASTER_GMAIL_LABEL_PARENT = "Master Job Manager"; // Top-level parent for all MJM-specific labels

// Labels for the MJM Application Tracker module
const MJM_TRACKER_GMAIL_LABEL_PARENT = MJM_MASTER_GMAIL_LABEL_PARENT + "/Job Application Tracker";
const MJM_TRACKER_GMAIL_LABEL_TO_PROCESS = MJM_TRACKER_GMAIL_LABEL_PARENT + "/To Process";
const MJM_TRACKER_GMAIL_LABEL_PROCESSED = MJM_TRACKER_GMAIL_LABEL_PARENT + "/Processed";
const MJM_TRACKER_GMAIL_LABEL_MANUAL_REVIEW = MJM_TRACKER_GMAIL_LABEL_PARENT + "/Manual Review Needed";

// Filter Query for MJM Application Tracker module (to catch application updates)
const MJM_TRACKER_GMAIL_FILTER_QUERY_APP_UPDATES = 'subject:("your application" OR "application to" OR "application for" OR "application update" OR "thank you for applying" OR "thanks for applying" OR "thank you for your interest" OR "received your application")';

// Labels for the MJM Job Leads Tracker module
const MJM_LEADS_GMAIL_LABEL_PARENT = MJM_MASTER_GMAIL_LABEL_PARENT + "/Job Application Potential";
const MJM_LEADS_GMAIL_LABEL_NEEDS_PROCESS = MJM_LEADS_GMAIL_LABEL_PARENT + "/NeedsProcess";
const MJM_LEADS_GMAIL_LABEL_DONE_PROCESS = MJM_LEADS_GMAIL_LABEL_PARENT + "/DoneProcess";
const MJM_LEADS_GMAIL_FILTER_QUERY = 'subject:("job alert") OR subject:(jobalert)'; // Gmail filter query for leads

// UserProperty KEYS for storing specific Label IDs for MJM Leads Tracker filter creation
const MJM_LEADS_USER_PROPERTY_NEEDS_PROCESS_LABEL_ID = 'mjmLeadsGmailNeedsProcessLabelId'; // Prefixed for clarity
const MJM_LEADS_USER_PROPERTY_DONE_PROCESS_LABEL_ID = 'mjmLeadsGmailDoneProcessLabelId';   // Prefixed for clarity

// --- "Applications" Sheet (MJM App Tracker): Column Indices & Settings ---
// For: MJM Core (App Tracker Module only)
const MJM_APP_PROCESSED_TIMESTAMP_COL = 1;
const MJM_APP_EMAIL_DATE_COL = 2;
const MJM_APP_PLATFORM_COL = 3;
const MJM_APP_COMPANY_COL = 4;
const MJM_APP_JOB_TITLE_COL = 5;
const MJM_APP_STATUS_COL = 6;
const MJM_APP_PEAK_STATUS_COL = 7;
const MJM_APP_LAST_UPDATE_DATE_COL = 8;
const MJM_APP_EMAIL_SUBJECT_COL = 9;
const MJM_APP_EMAIL_LINK_COL = 10;
const MJM_APP_EMAIL_ID_COL = 11;
const MJM_APP_TOTAL_COLUMNS = 11; // Total columns in the "Applications" sheet

// --- "Applications" Sheet (MJM App Tracker): Status Values & Hierarchy ---
// For: MJM Core (App Tracker Module only)
// Note: MANUAL_REVIEW_NEEDED_TEXT is a global constant (from Global_Constants.gs)
const MJM_APP_DEFAULT_STATUS = "Applied";
const MJM_APP_REJECTED_STATUS = "Rejected";
const MJM_APP_OFFER_STATUS = "Offer Received";
const MJM_APP_ACCEPTED_STATUS = "Offer Accepted";
const MJM_APP_INTERVIEW_STATUS = "Interview Scheduled";
const MJM_APP_ASSESSMENT_STATUS = "Assessment/Screening";
const MJM_APP_VIEWED_STATUS = "Application Viewed";
const MJM_APP_DEFAULT_PLATFORM = "Other"; // Default platform if not detected

const MJM_APP_STATUS_HIERARCHY = {
  [MANUAL_REVIEW_NEEDED_TEXT]: -1, // Uses global constant for the text value
  "Update/Other": 0,             // Generic status, often from Gemini parsing
  [MJM_APP_DEFAULT_STATUS]: 1,
  [MJM_APP_VIEWED_STATUS]: 2,
  [MJM_APP_ASSESSMENT_STATUS]: 3,
  [MJM_APP_INTERVIEW_STATUS]: 4,
  [MJM_APP_OFFER_STATUS]: 5,
  [MJM_APP_REJECTED_STATUS]: 5, // Can be same level as offer before acceptance
  [MJM_APP_ACCEPTED_STATUS]: 6
};

// --- "Applications" Sheet (MJM App Tracker): Config for Auto-Reject Stale Applications ---
// For: MJM Core (App Tracker Module only)
const MJM_STALE_WEEKS_THRESHOLD = 7; // Number of weeks after which a non-finalized application is stale
const MJM_STALE_FINAL_STATUSES_FOR_CHECK = new Set([MJM_APP_REJECTED_STATUS, MJM_APP_ACCEPTED_STATUS, "Withdrawn"]); // Statuses exempt from stale check

// --- Email Parsing (MJM Regex Fallback Logic): Keywords & Settings ---
// For: MJM Core (App Tracker Module's Regex Parser only)
const MJM_REJECTION_KEYWORDS = ["unfortunately", "regret to inform", "not moving forward", "decided not to proceed", "other candidates", "filled the position", "thank you for your time but"];
const MJM_OFFER_KEYWORDS = ["pleased to offer", "offer of employment", "job offer", "formally offer you the position"];
const MJM_INTERVIEW_KEYWORDS = ["invitation to interview", "schedule an interview", "interview request", "like to speak with you", "next steps involve an interview", "interview availability"];
const MJM_ASSESSMENT_KEYWORDS = ["assessment", "coding challenge", "online test", "technical screen", "next step is a skill assessment", "take a short test"];
const MJM_APP_VIEWED_KEYWORDS = ["application was viewed", "your application was viewed by", "recruiter viewed your application", "company viewed your application", "viewed your profile for the role"];
const MJM_PLATFORM_DOMAIN_KEYWORDS = { "linkedin.com": "LinkedIn", "indeed.com": "Indeed", "wellfound.com": "Wellfound", "angel.co": "Wellfound" }; // << REVIEW: This might be shareable if RTS also did platform detection, but keeping MJM-specific for now
const MJM_IGNORED_DOMAINS_FOR_COMPANY_PARSE = new Set(['greenhouse.io', 'lever.co', 'myworkday.com', 'icims.com', 'ashbyhq.com', 'smartrecruiters.com', 'bamboohr.com', 'taleo.net', 'gmail.com', 'google.com', 'example.com']); // Specific to MJM's regex company parsing logic

// --- "Potential Job Leads" Sheet (MJM Leads Module): Headers & Settings ---
// For: MJM Core (Leads Module only)
const MJM_LEADS_SHEET_HEADERS = [ // Defines expected order in "Potential Job Leads" sheet
    "Date Added", "Job Title", "Company", "Location", "Source Email Subject",
    "Link to Job Posting", "Status", "Source Email ID", "Processed Timestamp", "Notes"
];
// No column indices needed if MJM_Leads_SheetUtils.gs uses the header array to find columns by name (headerMap approach).
