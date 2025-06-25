// File: MJM_main.gs
// Description: Core orchestration for the Master Job Manager, including full project setup,
// Application Tracker email processing, and stale application rejection.
// Relies on Global_Constants.gs and MJM_Config.gs for configurations.
// UI menus are managed in MJM_UI.gs.

/**
 * Project: Comprehensive AI Job Suite (Integrated)
 */

// --- FULL PROJECT INITIAL SETUP ---
/**
 * Runs the complete initial setup for ALL modules of the integrated application.
 * This is the master setup function to be called once or to reset the project structure.
 */
function runFullProjectInitialSetup() {
  const SCRIPT_APP_NAME = (typeof APP_NAME !== 'undefined' ? APP_NAME : "Comprehensive Job Suite"); // From Global_Constants.gs
  Logger.log(`==== STARTING FULL PROJECT INITIAL SETUP for "${SCRIPT_APP_NAME}" ====`);
  let ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    Logger.log("[WARN] runFullProjectInitialSetup: No UI context available for alerts. Setup will proceed with logs only.");
  }

  let fullSetupMessages = [];
  let overallProjectSuccess = true; // DEFINED HERE - CRITICAL

  // --- Step 0: Ensure the main Spreadsheet exists or is created ---
  const mainSpreadsheet = MJM_getOrCreateSpreadsheet_Core(); // From MJM_SheetUtils.gs
  if (!mainSpreadsheet) {
    const errorMsg = `[FATAL SETUP ERROR] Main application spreadsheet could not be accessed or created. Please check APP_SPREADSHEET_ID or APP_TARGET_FILENAME in Global_Constants.gs. Full setup cannot continue.`;
    Logger.log(errorMsg);
    if (ui) ui.alert("Fatal Setup Error", errorMsg, ui.ButtonSet.OK);
    return; // Critical failure, cannot continue
  }
  fullSetupMessages.push(`Main Spreadsheet ("${mainSpreadsheet.getName()}"): Acquired/Created OK.`);
  Logger.log(`[INFO] FULL SETUP: Using main spreadsheet: "${mainSpreadsheet.getName()}" (ID: ${mainSpreadsheet.getId()})`);

  // --- Step 0.5: Pre-create Dashboard & Helper (ensure they exist early for correct ordering) ---
  Logger.log("\n--- Ensuring MJM Dashboard & Helper Sheets are Present Early ---");
  try {
    const { sheet: dashboardSheetCreated } = MJM_getOrCreateSheet_V2(mainSpreadsheet, DASHBOARD_TAB_NAME, 0); // From MJM_SheetUtils.gs
    if (dashboardSheetCreated) {
      fullSetupMessages.push(`  Dashboard Sheet ("${DASHBOARD_TAB_NAME}"): Existence Confirmed/Created.`);
    } else {
      throw new Error(`Failed to create/ensure Dashboard Sheet ("${DASHBOARD_TAB_NAME}").`);
    }

    const { sheet: helperSheetCreated } = MJM_getOrCreateSheet_V2(mainSpreadsheet, DASHBOARD_HELPER_SHEET_NAME); // From MJM_SheetUtils.gs
    if (helperSheetCreated) {
      fullSetupMessages.push(`  Dashboard Helper Sheet ("${DASHBOARD_HELPER_SHEET_NAME}"): Existence Confirmed/Created.`);
    } else {
      throw new Error(`Failed to create/ensure Dashboard Helper Sheet ("${DASHBOARD_HELPER_SHEET_NAME}").`);
    }
  } catch (e) {
    const errorDetail = `[ERROR] FULL SETUP (Dashboard/Helper Pre-Setup): ${e.toString()}\nStack: ${e.stack || 'No stack'}`;
    Logger.log(errorDetail);
    fullSetupMessages.push(`  Dashboard/Helper Sheets Pre-Setup: FAILED - ${e.message}`);
    overallProjectSuccess = false;
  }

  // --- Step 1: Setup MJM Application Tracker Module components ---
  Logger.log("\n--- Attempting Setup: MJM Application Tracker Module ---");
  if (overallProjectSuccess) {
    try {
      Logger.log(`[DEBUG runFullProjectInitialSetup] BEFORE calling MJM_setupAppTrackerModule. mainSpreadsheet is: ${mainSpreadsheet ? mainSpreadsheet.getName() : 'NULL'}`);
      const appTrackerResult = MJM_setupAppTrackerModule(mainSpreadsheet); // Assumes MJM_setupAppTrackerModule is defined later in this file
      Logger.log(`[DEBUG runFullProjectInitialSetup] AFTER calling MJM_setupAppTrackerModule. appTrackerResult is: ${JSON.stringify(appTrackerResult)}`);

      if (appTrackerResult && typeof appTrackerResult === 'object' && appTrackerResult.hasOwnProperty('messages') && appTrackerResult.hasOwnProperty('success')) {
        fullSetupMessages.push(...appTrackerResult.messages.map(m => `  AppTracker: ${m}`));
        if (!appTrackerResult.success) overallProjectSuccess = false;
      } else {
        Logger.log(`[ERROR runFullProjectInitialSetup] appTrackerResult is not a valid object. Value: ${JSON.stringify(appTrackerResult)}. Marking AppTracker setup as FAILED.`);
        fullSetupMessages.push("  AppTracker: FAILED - MJM_setupAppTrackerModule did not return a valid result object.");
        overallProjectSuccess = false;
      }
    } catch (e) {
      const errorDetail = `[ERROR] FULL SETUP (App Tracker Module): ${e.toString()}\nStack: ${e.stack || 'No stack'}`;
      Logger.log(errorDetail);
      fullSetupMessages.push(`  AppTracker Module Setup: FAILED - ${e.message}`);
      overallProjectSuccess = false;
    }
  } else {
    fullSetupMessages.push("  AppTracker Module Setup: SKIPPED due to previous critical errors.");
    Logger.log("[WARN] FULL SETUP: Skipping AppTracker Module due to previous errors.");
  }

  // --- Step 2: Setup MJM Job Leads Tracker Module components ---
  Logger.log("\n--- Attempting Setup: MJM Job Leads Tracker Module ---");
  if (overallProjectSuccess) {
    try {
      const leadsTrackerResult = MJM_setupLeadsModule(mainSpreadsheet); // From MJM_Leads_Main.gs
      if (leadsTrackerResult && typeof leadsTrackerResult === 'object' && leadsTrackerResult.hasOwnProperty('messages') && leadsTrackerResult.hasOwnProperty('success')) {
        fullSetupMessages.push(...leadsTrackerResult.messages.map(m => `  LeadsTracker: ${m}`));
        if (!leadsTrackerResult.success) overallProjectSuccess = false;
      } else {
        Logger.log(`[ERROR runFullProjectInitialSetup] leadsTrackerResult is not a valid object. Value: ${JSON.stringify(leadsTrackerResult)}. Marking LeadsTracker setup as FAILED.`);
        fullSetupMessages.push("  LeadsTracker: FAILED - MJM_setupLeadsModule did not return a valid result object.");
        overallProjectSuccess = false;
      }
    } catch (e) {
      const errorDetail = `[ERROR] FULL SETUP (Job Leads Module): ${e.toString()}\nStack: ${e.stack || 'No stack'}`;
      Logger.log(errorDetail);
      fullSetupMessages.push(`  LeadsTracker Module Setup: FAILED - ${e.message}`);
      overallProjectSuccess = false;
    }
  } else {
    fullSetupMessages.push("  LeadsTracker Module Setup: SKIPPED due to previous critical errors.");
    Logger.log("[WARN] FULL SETUP: Skipping Job Leads Module due to previous errors.");
  }

  // --- Step 3: Setup Profile Data Sheet (MasterProfile) using RTS logic ---
  Logger.log("\n--- Attempting Setup: Profile Data Sheet (MasterProfile) ---");
  if (overallProjectSuccess) {
    try {
      const profileSheet = RTS_setupMasterResumeSheet(mainSpreadsheet); // From RTS_SheetSetup.gs
      if (profileSheet) {
        fullSetupMessages.push(`  Profile Sheet ("${PROFILE_DATA_SHEET_NAME}"): Setup logic called OK (Outcome depends on user choice/internal logic).`);
        // Note: RTS_setupMasterResumeSheet might return the sheet even if user cancels a clear.
        // Its internal success needs to be robustly handled if that's a condition for overallProjectSuccess.
        // For now, if it returns a sheet object, we consider this step "called".
      } else {
        fullSetupMessages.push(`  Profile Sheet ("${PROFILE_DATA_SHEET_NAME}"): Setup FAILED or was aborted by its internal logic. This is critical for RTS.`);
        Logger.log(`[ERROR] FULL SETUP: Profile Sheet ("${PROFILE_DATA_SHEET_NAME}") setup appears to have failed or was aborted.`);
        overallProjectSuccess = false;
      }
    } catch (e) {
      const errorDetail = `[ERROR] FULL SETUP (Profile Sheet - ${PROFILE_DATA_SHEET_NAME}): ${e.toString()}\nStack: ${e.stack || 'No stack'}`;
      Logger.log(errorDetail);
      fullSetupMessages.push(`  Profile Sheet ("${PROFILE_DATA_SHEET_NAME}") Setup: FAILED - ${e.message}`);
      overallProjectSuccess = false;
    }
  } else {
    fullSetupMessages.push(`  Profile Sheet ("${PROFILE_DATA_SHEET_NAME}") Setup: SKIPPED due to previous critical errors.`);
    Logger.log(`[WARN] FULL SETUP: Skipping Profile Sheet ("${PROFILE_DATA_SHEET_NAME}") due to previous errors.`);
  }

  // --- Step 4: Ensure RTS intermediate processing sheets exist ---
  Logger.log("\n--- Ensuring RTS Intermediate Processing Sheets Exist ---");
  if (overallProjectSuccess) {
    try {
      const { sheet: jdSheet } = MJM_getOrCreateSheet_V2(mainSpreadsheet, JD_ANALYSIS_SHEET_NAME); // From MJM_SheetUtils.gs
      if (jdSheet) {
        fullSetupMessages.push(`  RTS Analysis Sheet ("${JD_ANALYSIS_SHEET_NAME}"): Existence Confirmed/Created.`);
      } else {
        throw new Error(`Failed to create/ensure RTS Analysis Sheet ("${JD_ANALYSIS_SHEET_NAME}").`);
      }

      const { sheet: scoringSheet } = MJM_getOrCreateSheet_V2(mainSpreadsheet, BULLET_SCORING_RESULTS_SHEET_NAME); // From MJM_SheetUtils.gs
      if (scoringSheet) {
        fullSetupMessages.push(`  RTS Scoring Sheet ("${BULLET_SCORING_RESULTS_SHEET_NAME}"): Existence Confirmed/Created.`);
      } else {
        throw new Error(`Failed to create/ensure RTS Scoring Sheet ("${BULLET_SCORING_RESULTS_SHEET_NAME}").`);
      }
    } catch (e) {
      const errorDetail = `[ERROR] FULL SETUP (Ensuring RTS Sheets): ${e.toString()}\nStack: ${e.stack || 'No stack'}`;
      Logger.log(errorDetail);
      fullSetupMessages.push(`  RTS Intermediate Sheets Setup: FAILED - ${e.message}`);
      overallProjectSuccess = false;
    }
  } else {
    fullSetupMessages.push("  RTS Intermediate Sheets Setup: SKIPPED due to previous critical errors.");
    Logger.log("[WARN] FULL SETUP: Skipping RTS Intermediate Sheets due to previous errors.");
  }

  // --- Step 5: Final Cleanup of Default "Sheet1" (if it's a leftover) ---
  Logger.log("\n--- Final Step: Attempt Cleanup of Default 'Sheet1' (if applicable) ---");
  if (mainSpreadsheet) {
    try {
      MJM_cleanupDefaultSheet1(mainSpreadsheet); // From MJM_SheetUtils.gs
      fullSetupMessages.push("  Default 'Sheet1' cleanup: Attempted (see logs for outcome).");
    } catch (e) {
      Logger.log(`[WARN] FULL SETUP (Cleanup 'Sheet1'): Not critical. ${e.toString()}`);
    }
  } else {
     Logger.log(`[WARN] FULL SETUP (Cleanup 'Sheet1'): Skipped as main spreadsheet object was not available.`);
  }

  // --- Final Summary UI Alert ---
  Logger.log("\n--- Finalizing Full Project Setup ---");
  const summaryTitle = `${SCRIPT_APP_NAME} - Full Setup ${overallProjectSuccess ? "Complete" : "Encountered Issues"}`;
  let summaryMessageText = `Full Project Initial Setup Summary for "${SCRIPT_APP_NAME}":\n- ${fullSetupMessages.join('\n- ')}\n\nOverall Status: ${overallProjectSuccess ? "SUCCESSFUL" : "ISSUES ENCOUNTERED (Review logs for critical details!)"}`;
  
  if (!overallProjectSuccess) {
    summaryMessageText += "\n\nIMPORTANT: One or more critical setup steps failed. Please review the execution transcript (View > Execution transcript in script editor) for detailed error messages. Some application features may not work correctly until these are resolved.";
  } else {
    summaryMessageText += "\n\nAll major setup components completed. Please review logs for any non-critical warnings or specific details.";
  }

  Logger.log(summaryMessageText.replace(/- /g, '  - '));
  if (ui) {
    ui.alert(summaryTitle, summaryMessageText, ui.ButtonSet.OK);
  } else {
    Logger.log("No UI context available to display setup summary alert to the user.");
  }

// --- Final Step Before Summary: Apply Sheet Tab Color Coding ---
  Logger.log("\n--- Applying Sheet Tab Color Coding ---");
  if (mainSpreadsheet) { // Ensure mainSpreadsheet object is valid
    try {
      // Define your preferred colors (Hex codes)
      // You can find hex color pickers online easily.
      const COLOR_LEADS = "#6495ED";         // Cornflower Blue
      const COLOR_DASH_APPS = "#FFA500";     // Orange
      const COLOR_PROFILE = "#FFC0CB";       // Pink
      const COLOR_RTS_PROCESS = "#006400";   // Dark Green

      // Helper function to safely set tab color
      const setTabColorSafe = (sheetName, color) => {
        if (!sheetName || typeof sheetName !== 'string') {
            Logger.log(`[WARN] setTabColorSafe: Invalid sheetName provided: ${sheetName}`);
            return;
        }
        const sheet = mainSpreadsheet.getSheetByName(sheetName);
        if (sheet) {
          try {
            sheet.setTabColor(color);
            Logger.log(`    Tab color set for "${sheetName}".`);
          } catch (eColor) {
            Logger.log(`[WARN] Could not set tab color for "${sheetName}": ${eColor.message}`);
          }
        } else {
          Logger.log(`[WARN] Sheet "${sheetName}" not found during tab coloring attempt.`);
        }
      };

      // Apply colors (using sheet name constants from Global_Constants.gs)
      
      // Orange Group
      setTabColorSafe(DASHBOARD_TAB_NAME, COLOR_DASH_APPS);
      setTabColorSafe(DASHBOARD_HELPER_SHEET_NAME, COLOR_DASH_APPS); 
      setTabColorSafe(APP_TRACKER_SHEET_TAB_NAME, COLOR_DASH_APPS); // Applications

      // Cornflower Blue Group
      setTabColorSafe(LEADS_SHEET_TAB_NAME, COLOR_LEADS); // Potential Job Leads

      // Pink Group
      setTabColorSafe(PROFILE_DATA_SHEET_NAME, COLOR_PROFILE); // MasterProfile

      // Dark Green Group
      setTabColorSafe(JD_ANALYSIS_SHEET_NAME, COLOR_RTS_PROCESS);
      setTabColorSafe(BULLET_SCORING_RESULTS_SHEET_NAME, COLOR_RTS_PROCESS);
      
      fullSetupMessages.push("  Sheet tab color coding: Applied (if sheets found).");
    } catch (eTabColors) {
      Logger.log(`[ERROR] Tab color coding process failed: ${eTabColors.message}`);
      fullSetupMessages.push("  Sheet tab color coding: FAILED.");
      // Not typically a critical failure to stop overall success
    }
  } else {
    fullSetupMessages.push("  Sheet tab color coding: SKIPPED (mainSpreadsheet object not available).");
  }

  // Call MJM_cleanupDefaultSheet1 LAST, after all sheets are named and potentially colored
  // This ensures if "Sheet1" was temporarily used and renamed, its final intended color (if any) sticks.
  // Or if "Sheet1" is a truly unwanted leftover, it's removed without affecting other operations.
  Logger.log("\n--- Final Cleanup of Default 'Sheet1' (if applicable) ---");
  if (mainSpreadsheet) {
      try {
          MJM_cleanupDefaultSheet1(mainSpreadsheet); // From MJM_SheetUtils.gs
          fullSetupMessages.push("  Default 'Sheet1' cleanup: Attempted post-coloring.");
      } catch (eCleanup) {
          Logger.log(`[WARN] FULL SETUP (Final Cleanup 'Sheet1'): ${eCleanup.toString()}`);
      }
  }

  Logger.log(`==== FULL PROJECT INITIAL SETUP ${overallProjectSuccess ? "CONCLUDED SUCCESSFULLY" : "CONCLUDED WITH ISSUES"} for "${SCRIPT_APP_NAME}" ====`);
} // End of runFullProjectInitialSetup

/**
 * Sets up ONLY the MJM Application Tracker specific components:
 * Gmail Labels, specific filter for application updates, "Applications" data sheet,
 * Dashboard, Helper sheet, and core triggers for application processing.
 * This is a helper called by runFullProjectInitialSetup.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} mainSpreadsheet The main application spreadsheet object.
 * @return {{success: boolean, messages: string[]}} Object indicating success and log messages.
 */
function MJM_setupAppTrackerModule(mainSpreadsheet) {
  const functionNameForLog = "MJM_setupAppTrackerModule";
  Logger.log(`  -- STARTING ${functionNameForLog} --`);
  let messages = [];
  let success = true;
  let dummyDataAddedToAppSheet = false;
  const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : false);

  if (!mainSpreadsheet) {
    const errorMsg = `[ERROR] ${functionNameForLog}: mainSpreadsheet object not provided. Aborting AppTracker setup.`;
    Logger.log(errorMsg);
    return { success: false, messages: [errorMsg] };
  }
  Logger.log(`  [INFO] ${functionNameForLog}: Operating on spreadsheet "${mainSpreadsheet.getName()}".`);

  // --- 1. Gmail Labels & Filter Setup ---
  Logger.log(`  [INFO] ${functionNameForLog}: Phase 1 - Gmail Setup Starting...`);
  let appTrackerToProcessLabelId = null;
  try {
    getOrCreateLabel(MJM_MASTER_GMAIL_LABEL_PARENT); Utilities.sleep(150); // From MJM_GmailUtils.gs
    getOrCreateLabel(MJM_TRACKER_GMAIL_LABEL_PARENT); Utilities.sleep(150);
    const toProcessLabelObj = getOrCreateLabel(MJM_TRACKER_GMAIL_LABEL_TO_PROCESS); Utilities.sleep(150);
    getOrCreateLabel(MJM_TRACKER_GMAIL_LABEL_PROCESSED); Utilities.sleep(150);
    getOrCreateLabel(MJM_TRACKER_GMAIL_LABEL_MANUAL_REVIEW); Utilities.sleep(150);
    Utilities.sleep(500); // Allow time for labels to propagate if newly created

    const advGmail = Gmail; // Assuming Gmail Advanced Service is enabled
    if (!advGmail?.Users?.Labels) {
      throw new Error("Advanced Gmail Service (Users.Labels API) is not available or not properly enabled for AppTracker setup.");
    }

    const labelsListResponse = advGmail.Users.Labels.list('me');
    if (!labelsListResponse || !labelsListResponse.labels) {
      throw new Error("Advanced Gmail Service: Did not receive a valid labels list response for AppTracker.");
    }
    const targetLabelInfo = labelsListResponse.labels.find(l => l.name === MJM_TRACKER_GMAIL_LABEL_TO_PROCESS);

    if (targetLabelInfo?.id) {
      appTrackerToProcessLabelId = targetLabelInfo.id;
      messages.push("AppTracker Gmail Labels: Verified/Created, 'To Process' ID obtained.");
      if (DEBUG) Logger.log(`    AppTracker Gmail: '${MJM_TRACKER_GMAIL_LABEL_TO_PROCESS}' Label ID: ${appTrackerToProcessLabelId}`);

      if (!advGmail?.Users?.Settings?.Filters) {
        throw new Error("Advanced Gmail Service (Users.Settings.Filters API) is not available/enabled for AppTracker filter setup.");
      }
      const filtersListResponse = advGmail.Users.Settings.Filters.list('me');
      let existingFiltersArray = (filtersListResponse && Array.isArray(filtersListResponse.filter)) ? filtersListResponse.filter : [];
      if (DEBUG && filtersListResponse && !Array.isArray(filtersListResponse.filter) && filtersListResponse.filter === undefined) {
           Logger.log(`    AppTracker Gmail: No filters currently exist for this user (Filters.list response 'filter' property undefined).`);
      } else if (DEBUG) {
           Logger.log(`    AppTracker Gmail: Found ${existingFiltersArray.length} existing Gmail filters. Checking for AppTracker filter...`);
      }

      const filterAlreadyExists = existingFiltersArray.some(f =>
        f.criteria?.query === MJM_TRACKER_GMAIL_FILTER_QUERY_APP_UPDATES && // From MJM_Config.gs
        f.action?.addLabelIds?.includes(appTrackerToProcessLabelId)
      );

      if (!filterAlreadyExists) {
        const filterResource = {
          criteria: { query: MJM_TRACKER_GMAIL_FILTER_QUERY_APP_UPDATES }, // From MJM_Config.gs
          action: { addLabelIds: [appTrackerToProcessLabelId], removeLabelIds: ['INBOX'] }
        };
        advGmail.Users.Settings.Filters.create(filterResource, 'me');
        messages.push("AppTracker Gmail Filter: CREATED successfully.");
        Logger.log(`  [INFO] ${functionNameForLog}: AppTracker Gmail Filter (Query: "${MJM_TRACKER_GMAIL_FILTER_QUERY_APP_UPDATES}"): CREATED for label ID ${appTrackerToProcessLabelId}.`);
      } else {
        messages.push("AppTracker Gmail Filter: Already exists.");
        Logger.log(`  [INFO] ${functionNameForLog}: AppTracker Gmail Filter (Query: "${MJM_TRACKER_GMAIL_FILTER_QUERY_APP_UPDATES}"): Confirmed it already exists.`);
      }
    } else {
      throw new Error(`Could not find or get ID for critical Gmail label "${MJM_TRACKER_GMAIL_LABEL_TO_PROCESS}". Filter creation aborted for AppTracker module.`);
    }
    Logger.log(`  [INFO] ${functionNameForLog}: Phase 1 - Gmail Setup Completed Successfully.`);
  } catch (e) {
    const errorDetail = `[ERROR] ${functionNameForLog} (Gmail Setup Phase): ${e.toString()}\nStack: ${e.stack || 'No stack'}`;
    Logger.log(errorDetail);
    messages.push(`AppTracker Gmail Setup: FAILED - ${e.message}`);
    success = false;
  }

  // --- 2. "Applications" Sheet Setup ---
  Logger.log(`  [INFO] ${functionNameForLog}: Phase 2 - Applications Sheet Setup Starting...`);
  let appDataSheet = null; 
  try {
    // APP_TRACKER_SHEET_TAB_NAME from Global_Constants.gs
    // MJM_getOrCreateSheet_V2 from MJM_SheetUtils.gs
    const { sheet: appSheetFromUtil, newSheetCreatedThisOp: appSheetIsNewlyHandled } = MJM_getOrCreateSheet_V2(mainSpreadsheet, APP_TRACKER_SHEET_TAB_NAME, 1);
    appDataSheet = appSheetFromUtil;

    if (!appDataSheet) {
      messages.push(`Sheet "${APP_TRACKER_SHEET_TAB_NAME}": FAILED to create or retrieve.`);
      success = false; 
      throw new Error(`Applications sheet "${APP_TRACKER_SHEET_TAB_NAME}" could not be established.`);
    }
    messages.push(`Sheet "${APP_TRACKER_SHEET_TAB_NAME}": OK/Exists (Formatting is handled by MJM_SheetUtils).`);

    if (appSheetIsNewlyHandled || appDataSheet.getLastRow() <= 1) {
      if (DEBUG) Logger.log(`    ${functionNameForLog}: AppSheet "${appDataSheet.getName()}" new/empty. LastRow: ${appDataSheet.getLastRow()}. Adding dummy data.`);
      
        const today = new Date();
        const weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
        const twoWeeksAgo = new Date(today.getTime() - 14 * 24 * 60 * 60 * 1000);
        let dummyRowsData = [
          [new Date(), twoWeeksAgo, "LinkedIn", "Alpha Innovations Corp", "Software Engineer I", MJM_APP_DEFAULT_STATUS, MJM_APP_DEFAULT_STATUS, twoWeeksAgo, "Applied to Alpha Innovations Corp", "", ""],
          [new Date(), weekAgo, "Indeed.com", "Beta Technical Solutions", "Data Analyst Professional", MJM_APP_VIEWED_STATUS, MJM_APP_VIEWED_STATUS, weekAgo, "Viewed by Beta Technical Recruiter", "", ""],
          [new Date(), today, "Company Career Site", "Gamma Incorporated", "Product Manager Lead", MJM_APP_INTERVIEW_STATUS, MJM_APP_INTERVIEW_STATUS, today, "Scheduled interview with Gamma Inc.", "", ""]
        ]; // Uses MJM_APP_..._STATUS from MJM_Config.gs
        dummyRowsData = dummyRowsData.map(row => { 
            const fullRow = row.slice(0, MJM_APP_TOTAL_COLUMNS); // MJM_APP_TOTAL_COLUMNS from MJM_Config.gs
            while (fullRow.length < MJM_APP_TOTAL_COLUMNS) fullRow.push("");
            return fullRow;
        });
        const requiredRowsForDummy = 1 + dummyRowsData.length;
        if (appDataSheet.getMaxRows() < requiredRowsForDummy) {
            appDataSheet.insertRowsAfter(appDataSheet.getMaxRows(), requiredRowsForDummy - appDataSheet.getMaxRows());
        }
        appDataSheet.getRange(2, 1, dummyRowsData.length, MJM_APP_TOTAL_COLUMNS).setValues(dummyRowsData);
        dummyDataAddedToAppSheet = true;
        messages.push("AppTracker Sheet: Dummy data added for dashboard initialization.");
        if(DEBUG) Logger.log(`    ${functionNameForLog}: Added ${dummyRowsData.length} dummy rows to "${appDataSheet.getName()}".`);
    }
    Logger.log(`  [INFO] ${functionNameForLog}: Phase 2 - Applications Sheet Setup Completed Successfully.`);
  } catch (e) {
    const errorDetail = `[ERROR] ${functionNameForLog} (Applications Sheet Setup): ${e.toString()}\nStack: ${e.stack || 'No stack'}`;
    Logger.log(errorDetail);
    messages.push(`Applications Sheet ("${APP_TRACKER_SHEET_TAB_NAME}") Setup: FAILED - ${e.message}`);
    success = false;
  }

  // --- 3. Dashboard & Helper Sheet Setup (Includes Dummy Data Handling & Chart Creation) ---
  Logger.log(`  [INFO] ${functionNameForLog}: Phase 3 - Dashboard/Helper, Metrics & Chart Creation Starting...`);
  if (success && mainSpreadsheet && appDataSheet) { // appDataSheet already has dummy data here
    let dashboardSheet = null;    
    let helperSheet = null;       
    let dashboardSetupPhaseSuccess = true; 

    try {
      // ... (Get dashboardSheet and helperSheet as before) ...
      dashboardSheet = mainSpreadsheet.getSheetByName(DASHBOARD_TAB_NAME);
      helperSheet = mainSpreadsheet.getSheetByName(DASHBOARD_HELPER_SHEET_NAME);

      if (dashboardSheet && helperSheet) {
        messages.push(`AppTracker: Sheet "${DASHBOARD_TAB_NAME}" and "${DASHBOARD_HELPER_SHEET_NAME}" references obtained.`);
        Logger.log(`  [INFO] ${functionNameForLog}: Dashboard and Helper sheets retrieved.`);
        
        // 1. Format dashboard & set helper formulas (will use dummy data)
        Logger.log(`    ${functionNameForLog}: Calling MJM_formatDashboardSheet (with dummy data present)...`);
        try {
            MJM_formatDashboardSheet(dashboardSheet, appDataSheet.getName()); 
            messages.push(`AppTracker: Dashboard formatting called.`);
            Logger.log(`    ${functionNameForLog}: MJM_formatDashboardSheet completed.`);
        } catch (eFormatDash) {
            dashboardSetupPhaseSuccess = false; 
            // ... (log error) ...
        }
        
        // 2. Create/Update Charts (will use dummy data for initial creation)
        if (dashboardSetupPhaseSuccess) {
            Logger.log(`    ${functionNameForLog}: Calling MJM_updateDashboardMetrics (with dummy data present)...`);
            try {
                MJM_updateDashboardMetrics(mainSpreadsheet); 
                messages.push("AppTracker Dashboard: Initial metrics & charts creation called.");
                Logger.log(`    ${functionNameForLog}: Initial MJM_updateDashboardMetrics completed.`);
            } catch (eUpdateMetricsInitial) {
                dashboardSetupPhaseSuccess = false;
                // ... (log error) ...
            }
        }

        // 3. NOW Remove Dummy Data
        if (dashboardSetupPhaseSuccess && appDataSheet && dummyDataAddedToAppSheet) { // Check appDataSheet again for safety
            Logger.log(`    ${functionNameForLog}: Attempting to remove dummy data from "${appDataSheet.getName()}".`);
            const numDummyRowsAdded = 3; 
            if (appDataSheet.getLastRow() >= (1 + numDummyRowsAdded)) {
                try {
                    appDataSheet.deleteRows(2, numDummyRowsAdded);
                    Logger.log(`    ${functionNameForLog}: Dummy data removed from Applications sheet.`);
                    messages.push("AppTracker Sheet: Dummy data removed.");
                    // dummyDataAddedToAppSheet = false; // Optional: reset flag if needed elsewhere
                } catch (eDeleteDummy) { /* ... log warning ... */ }
            } // ... else log not enough rows ...
        }

        // 4. Call MJM_updateDashboardMetrics AGAIN to refresh charts based on empty/actual data
        if (dashboardSetupPhaseSuccess) {
            Logger.log(`    ${functionNameForLog}: Calling MJM_updateDashboardMetrics AGAIN (after dummy data removal)...`);
            try {
                MJM_updateDashboardMetrics(mainSpreadsheet); 
                messages.push("AppTracker Dashboard: Final metrics & charts refresh called.");
                Logger.log(`    ${functionNameForLog}: Final MJM_updateDashboardMetrics completed.`);
            } catch (eUpdateMetricsFinal) {
                // This failure is less critical for overall setup success, but should be noted.
                Logger.log(`[WARN] ${functionNameForLog}: Error during final MJM_updateDashboardMetrics: ${eUpdateMetricsFinal.toString()}`);
                messages.push(`AppTracker Dashboard: WARNING - Error during final chart refresh - ${eUpdateMetricsFinal.message}`);
            }
        }
        
      } else { /* ... handle missing dashboard/helper sheets ... */ }
    } catch (eOuter) { /* ... outer catch ... */ }

    if (!dashboardSetupPhaseSuccess) success = false; 
    Logger.log(`  [INFO] ${functionNameForLog}: Phase 3 Concluded with success: ${dashboardSetupPhaseSuccess}.`);
  } else { /* ... handle prior failure or missing main/appData sheets ... */ }

  // --- 4. App Tracker Triggers ---
  Logger.log(`  [INFO] ${functionNameForLog}: Phase 4 - Trigger Setup Starting...`);
  if(success) { 
    try {
      if (MJM_createHourlyTrigger('MJM_processJobApplicationEmails', 1)) { // From MJM_Triggers.gs
        messages.push("Trigger for 'MJM_processJobApplicationEmails' (Hourly): Successfully set up.");
      } else {
        messages.push("Trigger for 'MJM_processJobApplicationEmails' (Hourly): Not newly created (may already exist or error occurred - check logs).");
      }
    } catch (eTrig1) {
      Logger.log(`[ERROR] ${functionNameForLog} (Trigger for Email Processing): ${eTrig1.message}`);
      messages.push(`Trigger for Email Processing: FAILED - ${eTrig1.message}`);
      success = false;
    }
    try {
      if (MJM_createDailyAtHourTrigger('MJM_markStaleApplicationsAsRejected', 2)) { // From MJM_Triggers.gs
        messages.push("Trigger for 'MJM_markStaleApplicationsAsRejected' (Daily): Successfully set up.");
      } else {
        messages.push("Trigger for 'MJM_markStaleApplicationsAsRejected' (Daily): Not newly created (check logs).");
      }
    } catch (eTrig2) {
      Logger.log(`[ERROR] ${functionNameForLog} (Trigger for Stale Apps): ${eTrig2.message}`);
      messages.push(`Trigger for Stale Apps: FAILED - ${eTrig2.message}`);
      success = false;
    }
    Logger.log(`  [INFO] ${functionNameForLog}: Phase 4 - Trigger Setup Completed.`);
  } else {
      messages.push("AppTracker Triggers: SKIPPED due to errors in prior setup phases.");
      Logger.log(`  [WARN] ${functionNameForLog}: Phase 4 (Triggers) SKIPPED due to prior setup failure.`);
  }

  Logger.log(`  -- FINISHED ${functionNameForLog} (${success ? "OK" : "ISSUES ENCOUNTERED"}) --`);
  return { success: success, messages: messages };
} // End of MJM_setupAppTrackerModule

/**
 * Main email processing function for MJM Job Application updates.
 * This function is intended to be triggered automatically (e.g., hourly).
 */
function MJM_processJobApplicationEmails() {
  const SCRIPT_START_TIME = new Date();
  Logger.log(`\n==== STARTING MJM_PROCESS_JOB_APP_EMAILS (${SCRIPT_START_TIME.toLocaleString()}) ====`);
  const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : false); // Global_Constants.gs

  const userProps = PropertiesService.getUserProperties();
  const geminiApiKey = userProps.getProperty(SHARED_GEMINI_API_KEY_PROPERTY); // Global_Constants.gs
  let useGemini = !!(geminiApiKey && geminiApiKey.startsWith("AIza") && geminiApiKey.length > 30);
  if (DEBUG) Logger.log(useGemini ? "[DEBUG] MJM App Emails: Using Gemini." : "[DEBUG] MJM App Emails: Using Regex (Gemini key issue/not set).");

  const mainSS = MJM_getOrCreateSpreadsheet_Core(); // From MJM_SheetUtils.gs
  if (!mainSS) { Logger.log("[FATAL ERROR] MJM App Emails: Main Spreadsheet not found. Aborting."); return; }
  
  // APP_TRACKER_SHEET_TAB_NAME from Global_Constants.gs
  // MJM_getOrCreateSheet_V2 is used here to *ensure* the sheet exists; normally it would, but this is defensive.
  const { sheet: dataSheet } = MJM_getOrCreateSheet_V2(mainSS, APP_TRACKER_SHEET_TAB_NAME); 
  if (!dataSheet) { Logger.log(`[FATAL ERROR] MJM App Emails: Sheet "${APP_TRACKER_SHEET_TAB_NAME}" could not be accessed. Aborting.`); return; }

  // Label names from MJM_Config.gs (e.g., MJM_TRACKER_GMAIL_LABEL_TO_PROCESS)
  let toProcessLabel, processedLabel, manualReviewLabel;
  try {
    toProcessLabel = GmailApp.getUserLabelByName(MJM_TRACKER_GMAIL_LABEL_TO_PROCESS);
    processedLabel = GmailApp.getUserLabelByName(MJM_TRACKER_GMAIL_LABEL_PROCESSED);
    manualReviewLabel = GmailApp.getUserLabelByName(MJM_TRACKER_GMAIL_LABEL_MANUAL_REVIEW);
    if (!toProcessLabel || !processedLabel || !manualReviewLabel) {
      throw new Error(`One or more core App Tracker labels ("${MJM_TRACKER_GMAIL_LABEL_TO_PROCESS}", "${MJM_TRACKER_GMAIL_LABEL_PROCESSED}", "${MJM_TRACKER_GMAIL_LABEL_MANUAL_REVIEW}") not found.`);
    }
  } catch(e) { Logger.log(`[FATAL ERROR] MJM App Emails: Could not retrieve critical Tracker labels. ${e.message}`); return; }

  const lastSheetDataRow = dataSheet.getLastRow();
  const applicationDataCache = {}; 
  const alreadyProcessedMsgIdsThisRun = new Set(); 

  if (lastSheetDataRow >= 2) { 
    try { 
      const colsToLoadForCache = [MJM_APP_COMPANY_COL,MJM_APP_JOB_TITLE_COL,MJM_APP_EMAIL_ID_COL,MJM_APP_STATUS_COL,MJM_APP_PEAK_STATUS_COL]; // From MJM_Config.gs
      const minColCache = Math.min(...colsToLoadForCache); const maxColCache = Math.max(...colsToLoadForCache);
      const numColsForCache = maxColCache - minColCache + 1;
      const cacheRangeValues = dataSheet.getRange(2, minColCache, lastSheetDataRow - 1, numColsForCache).getValues();

      const relCompanyIdx = MJM_APP_COMPANY_COL - minColCache; const relTitleIdx = MJM_APP_JOB_TITLE_COL - minColCache;
      const relEmailIdIdx = MJM_APP_EMAIL_ID_COL - minColCache; const relStatusIdx = MJM_APP_STATUS_COL - minColCache;
      const relPeakStatusIdx = MJM_APP_PEAK_STATUS_COL - minColCache;

      cacheRangeValues.forEach((rowValues, sheetRowIndex) => {
        const emailId = String(rowValues[relEmailIdIdx] || "").trim();
        const company = String(rowValues[relCompanyIdx] || "").trim();
        const companyLCKey = company.toLowerCase();
        // MANUAL_REVIEW_NEEDED_TEXT from Global_Constants.gs
        if (companyLCKey && companyLCKey !== MANUAL_REVIEW_NEEDED_TEXT.toLowerCase() && companyLCKey !== 'n/a') {
          if (!applicationDataCache[companyLCKey]) applicationDataCache[companyLCKey] = [];
          applicationDataCache[companyLCKey].push({
            rowNum: sheetRowIndex + 2, emailId: emailId, company: company, title: String(rowValues[relTitleIdx] || "").trim(),
            status: String(rowValues[relStatusIdx] || "").trim(), peakStatus: String(rowValues[relPeakStatusIdx] || "").trim()
          });
        }
      });
      if (DEBUG) Logger.log(`  MJM App Emails Preload: Cached ${Object.keys(applicationDataCache).length} unique companies.`);
    } catch (e) { Logger.log(`[ERROR] MJM App Emails: Failed during preload cache operation: ${e.toString()}\nStack: ${e.stack}`); }
  } else { if(DEBUG) Logger.log(`  MJM App Emails Preload: Application sheet empty or header only. No cache preloaded.`); }

  const MAX_THREADS_TO_SCAN = 15; 
  const MAX_MESSAGES_TO_PROCESS_THIS_RUN = 20; 
  let processedMessagesCountThisRun = 0;
  let gmailThreadsToScan = [];
  try { gmailThreadsToScan = toProcessLabel.getThreads(0, MAX_THREADS_TO_SCAN); }
  catch (e) { Logger.log(`[ERROR] MJM App Emails: Failed to get threads from label "${toProcessLabel.getName()}": ${e.message}`); return; }

  const newMessagesList = [];
  const emailIdsAlreadyInSheet = new Set();
  Object.values(applicationDataCache).flat().forEach(entry => { if (entry.emailId) emailIdsAlreadyInSheet.add(entry.emailId); });

  gmailThreadsToScan.forEach(thread => {
    try {
      thread.getMessages().forEach(msg => {
        if (!emailIdsAlreadyInSheet.has(msg.getId()) && !alreadyProcessedMsgIdsThisRun.has(msg.getId())) { 
          newMessagesList.push({ messageObj: msg, emailDate: msg.getDate(), gmailThreadId: thread.getId() });
          alreadyProcessedMsgIdsThisRun.add(msg.getId()); 
        }
      });
    } catch (eFetchMsg) { Logger.log(`[WARN] MJM App Emails: Error fetching messages for thread ${thread.getId()}: ${eFetchMsg.message}`); }
  });

  if (newMessagesList.length === 0) {
    Logger.log("[INFO] MJM App Emails: No new unread application email messages to process.");
    if(mainSS) try { MJM_updateDashboardMetrics(mainSS); } catch (eDashboard) { Logger.log(`[WARN] MJM App Emails: Dashboard update failed (no new msgs): ${eDashboard.message}`); } 
    Logger.log(`==== MJM_PROCESS_JOB_APP_EMAILS FINISHED (No new messages) ====`); return;
  }
  newMessagesList.sort((a, b) => new Date(a.emailDate) - new Date(b.emailDate)); 
  Logger.log(`[INFO] MJM App Emails: Found ${newMessagesList.length} new messages to analyze.`);

  let threadProcessingOutcomesMap = {}; 
  let runProcessingStats = { updatedRows: 0, newRowsAdded: 0, errorsEncountered: 0 };

  for (let idx = 0; idx < newMessagesList.length; idx++) {
    if (processedMessagesCountThisRun >= MAX_MESSAGES_TO_PROCESS_THIS_RUN || (new Date().getTime() - SCRIPT_START_TIME.getTime()) / 1000 > 330) {
      Logger.log(`[WARN] MJM App Emails: Reached processing limit. Stopping loop.`); break;
    }
    const { messageObj, emailDate, gmailThreadId } = newMessagesList[idx];
    const messageId = messageObj.getId();
    processedMessagesCountThisRun++;
    if(DEBUG) Logger.log(`  --- Processing App Msg ${processedMessagesCountThisRun}/${newMessagesList.length} (ID: ${messageId}, Thread: ${gmailThreadId}) ---`);

    let extractedCompany = MANUAL_REVIEW_NEEDED_TEXT, extractedTitle = MANUAL_REVIEW_NEEDED_TEXT, extractedStatus = null; // MANUAL_REVIEW_NEEDED_TEXT from Global_Constants.gs
    let emailPlainBody = "", needsManualReviewFlag = false, sheetWriteSuccessful = false;

    try {
      const emailSubject = messageObj.getSubject() || "";
      const emailSender = messageObj.getFrom() || "";
      const emailPermalink = `https://mail.google.com/mail/u/0/#inbox/${messageId}`;
      let detectedPlatform = MJM_APP_DEFAULT_PLATFORM; // MJM_Config.gs
      try { const senderDom = (emailSender.match(/@([^>]+)/) || [])[1]?.toLowerCase(); if(senderDom) for(const k in MJM_PLATFORM_DOMAIN_KEYWORDS) if(senderDom.includes(k)){detectedPlatform=MJM_PLATFORM_DOMAIN_KEYWORDS[k];break;} } catch(eP){Logger.log(`    Platform detection error: ${eP}`);} // MJM_PLATFORM_DOMAIN_KEYWORDS from MJM_Config.gs

      try { emailPlainBody = messageObj.getPlainBody(); } catch (eBody) { Logger.log(`    WARN: Failed to get plain body for Msg ${messageId}: ${eBody.message}`); emailPlainBody = "[Body Fetch Error]";}

      if (useGemini && emailPlainBody.trim() && emailPlainBody !== "[Body Fetch Error]") {
        const geminiResult = MJM_callGemini_forApplicationDetails(emailSubject, emailPlainBody, geminiApiKey); // MJM_GeminiService.gs
        if (geminiResult) {
          extractedCompany = geminiResult.company || MANUAL_REVIEW_NEEDED_TEXT;
          extractedTitle = geminiResult.title || MANUAL_REVIEW_NEEDED_TEXT;
          extractedStatus = geminiResult.status; 
          if (DEBUG) Logger.log(`    Gemini Parsed: C='${extractedCompany}', T='${extractedTitle}', S='${extractedStatus}'`);
          if (!extractedStatus || extractedStatus === MANUAL_REVIEW_NEEDED_TEXT || extractedStatus === "Update/Other") {
            const regexStatusFromBody = MJM_parseBodyForStatus(emailPlainBody); // MJM_ParsingUtils.gs
            if (regexStatusFromBody) extractedStatus = regexStatusFromBody;
          }
        } else { if(DEBUG) Logger.log(`    Gemini call returned null, relying on regex fallback.`); }
      }

      if (extractedCompany === MANUAL_REVIEW_NEEDED_TEXT || extractedTitle === MANUAL_REVIEW_NEEDED_TEXT || !extractedStatus || extractedStatus === MANUAL_REVIEW_NEEDED_TEXT || extractedStatus === "Update/Other") {
        if(DEBUG && useGemini && (extractedCompany === MANUAL_REVIEW_NEEDED_TEXT || extractedTitle === MANUAL_REVIEW_NEEDED_TEXT || !extractedStatus || extractedStatus === MANUAL_REVIEW_NEEDED_TEXT || extractedStatus === "Update/Other")) Logger.log(`    Invoking regex fallback for some fields.`);
        const regexParseResult = MJM_extractCompanyAndTitle(messageObj, detectedPlatform, emailSubject, emailPlainBody); // MJM_ParsingUtils.gs
        if (extractedCompany === MANUAL_REVIEW_NEEDED_TEXT) extractedCompany = regexParseResult.company;
        if (extractedTitle === MANUAL_REVIEW_NEEDED_TEXT) extractedTitle = regexParseResult.title;
        if (!extractedStatus || extractedStatus === MANUAL_REVIEW_NEEDED_TEXT || extractedStatus === "Update/Other") {
          const regexStatusFromBody = MJM_parseBodyForStatus(emailPlainBody); // MJM_ParsingUtils.gs
          if (regexStatusFromBody) extractedStatus = regexStatusFromBody;
        }
        if(DEBUG) Logger.log(`    Regex Fallback/Final Parsed: C='${extractedCompany}', T='${extractedTitle}', S='${extractedStatus}'`);
      }
      
      needsManualReviewFlag = (extractedCompany === MANUAL_REVIEW_NEEDED_TEXT || extractedTitle === MANUAL_REVIEW_NEEDED_TEXT);
      const finalStatusToLog = extractedStatus || MJM_APP_DEFAULT_STATUS; // MJM_Config.gs
      const companyCacheKeyForSearch = (extractedCompany !== MANUAL_REVIEW_NEEDED_TEXT) ? extractedCompany.toLowerCase() : `_mrev_entry_key_${messageId}`; // Needs careful review
      let existingApplicationEntry = null; let sheetRowNumberToUpdate = -1;

      if (extractedCompany !== MANUAL_REVIEW_NEEDED_TEXT && applicationDataCache[companyCacheKeyForSearch]) {
        const matches = applicationDataCache[companyCacheKeyForSearch];
        if (extractedTitle !== MANUAL_REVIEW_NEEDED_TEXT) existingApplicationEntry = matches.find(e => e.title?.toLowerCase() === extractedTitle.toLowerCase());
        if (!existingApplicationEntry && matches.length > 0) existingApplicationEntry = matches.sort((a,b) => b.rowNum - a.rowNum)[0]; // Get most recent if multiple by company
        if (existingApplicationEntry) sheetRowNumberToUpdate = existingApplicationEntry.rowNum;
      }

      let currentRowDataValues; 
      if (sheetRowNumberToUpdate !== -1 && existingApplicationEntry) { 
        if(DEBUG) Logger.log(`    Updating entry at row ${sheetRowNumberToUpdate} for C='${extractedCompany}', T='${extractedTitle}'`);
        currentRowDataValues = dataSheet.getRange(sheetRowNumberToUpdate, 1, 1, MJM_APP_TOTAL_COLUMNS).getValues()[0]; // MJM_Config.gs
        
        currentRowDataValues[MJM_APP_PROCESSED_TIMESTAMP_COL - 1] = new Date(); // MJM_Config.gs
        const currentEmailDateInSheet = currentRowDataValues[MJM_APP_EMAIL_DATE_COL-1] instanceof Date ? currentRowDataValues[MJM_APP_EMAIL_DATE_COL-1] : (currentRowDataValues[MJM_APP_EMAIL_DATE_COL-1] ? new Date(currentRowDataValues[MJM_APP_EMAIL_DATE_COL-1]) : null);
        if (!currentEmailDateInSheet || isNaN(currentEmailDateInSheet.getTime()) || emailDate.getTime() > currentEmailDateInSheet.getTime()) currentRowDataValues[MJM_APP_EMAIL_DATE_COL-1] = emailDate;
        
        const currentLastUpdateInSheet = currentRowDataValues[MJM_APP_LAST_UPDATE_DATE_COL-1] instanceof Date ? currentRowDataValues[MJM_APP_LAST_UPDATE_DATE_COL-1] : (currentRowDataValues[MJM_APP_LAST_UPDATE_DATE_COL-1] ? new Date(currentRowDataValues[MJM_APP_LAST_UPDATE_DATE_COL-1]) : null);
        if (!currentLastUpdateInSheet || isNaN(currentLastUpdateInSheet.getTime()) || emailDate.getTime() > currentLastUpdateInSheet.getTime()) currentRowDataValues[MJM_APP_LAST_UPDATE_DATE_COL-1] = emailDate;
        
        currentRowDataValues[MJM_APP_EMAIL_SUBJECT_COL-1]=emailSubject; currentRowDataValues[MJM_APP_EMAIL_LINK_COL-1]=emailPermalink;
        currentRowDataValues[MJM_APP_EMAIL_ID_COL-1]=messageId; currentRowDataValues[MJM_APP_PLATFORM_COL-1]=detectedPlatform;
        if(extractedCompany!==MANUAL_REVIEW_NEEDED_TEXT && (currentRowDataValues[MJM_APP_COMPANY_COL-1]===MANUAL_REVIEW_NEEDED_TEXT || String(currentRowDataValues[MJM_APP_COMPANY_COL-1]).toLowerCase()!==extractedCompany.toLowerCase())) currentRowDataValues[MJM_APP_COMPANY_COL-1]=extractedCompany;
        if(extractedTitle!==MANUAL_REVIEW_NEEDED_TEXT && (currentRowDataValues[MJM_APP_JOB_TITLE_COL-1]===MANUAL_REVIEW_NEEDED_TEXT || String(currentRowDataValues[MJM_APP_JOB_TITLE_COL-1]).toLowerCase()!==extractedTitle.toLowerCase())) currentRowDataValues[MJM_APP_JOB_TITLE_COL-1]=extractedTitle;

        const statusInSheet = String(currentRowDataValues[MJM_APP_STATUS_COL - 1] || MJM_APP_DEFAULT_STATUS).trim(); // MJM_Config.gs
        if(statusInSheet !== MJM_APP_ACCEPTED_STATUS || finalStatusToLog === MJM_APP_ACCEPTED_STATUS){ // MJM_Config.gs
          const currentRank = MJM_APP_STATUS_HIERARCHY[statusInSheet] ?? 0; // MJM_Config.gs
          const newRank = MJM_APP_STATUS_HIERARCHY[finalStatusToLog] ?? 0;
          if (newRank >= currentRank || finalStatusToLog === MJM_APP_REJECTED_STATUS || finalStatusToLog === MJM_APP_OFFER_STATUS) currentRowDataValues[MJM_APP_STATUS_COL - 1] = finalStatusToLog;
        }
        const updatedStatusInSheet = currentRowDataValues[MJM_APP_STATUS_COL - 1];

        let peakStatus = existingApplicationEntry.peakStatus || String(currentRowDataValues[MJM_APP_PEAK_STATUS_COL-1]||"").trim() || MJM_APP_DEFAULT_STATUS;
        const excludedPeak = new Set([MJM_APP_REJECTED_STATUS, MJM_APP_ACCEPTED_STATUS, MANUAL_REVIEW_NEEDED_TEXT, "Update/Other"]); // Uses Global & MJM_Config constants
        if((MJM_APP_STATUS_HIERARCHY[updatedStatusInSheet]??-2) > (MJM_APP_STATUS_HIERARCHY[peakStatus]??-2) && !excludedPeak.has(updatedStatusInSheet)) peakStatus=updatedStatusInSheet;
        else if(peakStatus === MJM_APP_DEFAULT_STATUS && !excludedPeak.has(updatedStatusInSheet) && (MJM_APP_STATUS_HIERARCHY[updatedStatusInSheet]??0) > (MJM_APP_STATUS_HIERARCHY[MJM_APP_DEFAULT_STATUS]??0)) peakStatus=updatedStatusInSheet;
        currentRowDataValues[MJM_APP_PEAK_STATUS_COL - 1] = peakStatus;
        if(DEBUG) Logger.log(`    Updated Row ${sheetRowNumberToUpdate}. Status: "${updatedStatusInSheet}", Peak: "${peakStatus}"`);
        
        dataSheet.getRange(sheetRowNumberToUpdate, 1, 1, MJM_APP_TOTAL_COLUMNS).setValues([currentRowDataValues]);
        runProcessingStats.updatedRows++; sheetWriteSuccessful = true;
        const cacheItem = applicationDataCache[companyCacheKeyForSearch]?.find(e => e.rowNum === sheetRowNumberToUpdate);
        if(cacheItem) Object.assign(cacheItem, {status: updatedStatusInSheet, peakStatus: peakStatus, emailId: messageId, company:currentRowDataValues[MJM_APP_COMPANY_COL-1], title:currentRowDataValues[MJM_APP_JOB_TITLE_COL-1]});
      } else { 
        if(DEBUG) Logger.log(`    Appending new entry for C='${extractedCompany}', T='${extractedTitle}'`);
        currentRowDataValues = new Array(MJM_APP_TOTAL_COLUMNS).fill("");
        currentRowDataValues[MJM_APP_PROCESSED_TIMESTAMP_COL-1]=new Date(); currentRowDataValues[MJM_APP_EMAIL_DATE_COL-1]=emailDate;
        currentRowDataValues[MJM_APP_PLATFORM_COL-1]=detectedPlatform; currentRowDataValues[MJM_APP_COMPANY_COL-1]=extractedCompany;
        currentRowDataValues[MJM_APP_JOB_TITLE_COL-1]=extractedTitle; currentRowDataValues[MJM_APP_STATUS_COL-1]=finalStatusToLog;
        currentRowDataValues[MJM_APP_LAST_UPDATE_DATE_COL-1]=emailDate; currentRowDataValues[MJM_APP_EMAIL_SUBJECT_COL-1]=emailSubject;
        currentRowDataValues[MJM_APP_EMAIL_LINK_COL-1]=emailPermalink; currentRowDataValues[MJM_APP_EMAIL_ID_COL-1]=messageId;
        const excludedPInit = new Set([MJM_APP_REJECTED_STATUS,MJM_APP_ACCEPTED_STATUS,MANUAL_REVIEW_NEEDED_TEXT,"Update/Other"]);
        const initPeak = !excludedPInit.has(finalStatusToLog)?finalStatusToLog:MJM_APP_DEFAULT_STATUS;
        currentRowDataValues[MJM_APP_PEAK_STATUS_COL-1]=initPeak;
        if(DEBUG) Logger.log(`    Appending. Status: "${finalStatusToLog}", Peak: "${initPeak}"`);
        dataSheet.appendRow(currentRowDataValues);
        runProcessingStats.newRowsAdded++; sheetWriteSuccessful = true;
        const newRowNum = dataSheet.getLastRow();
        const cKeyNew = (extractedCompany !== MANUAL_REVIEW_NEEDED_TEXT)?extractedCompany.toLowerCase():`_mrev_entry_key_${messageId}`;
        if(!applicationDataCache[cKeyNew])applicationDataCache[cKeyNew]=[];
        applicationDataCache[cKeyNew].push({rowNum:newRowNum,emailId:messageId,company:extractedCompany,title:extractedTitle,status:finalStatusToLog,peakStatus:initPeak});
      }

      if(sheetWriteSuccessful) {
          let outcomeLabel = needsManualReviewFlag ? 'manual' : 'done'; 
          if(threadProcessingOutcomesMap[gmailThreadId] !== 'manual') threadProcessingOutcomesMap[gmailThreadId] = outcomeLabel; 
          if(outcomeLabel === 'manual') threadProcessingOutcomesMap[gmailThreadId] = 'manual'; // 'manual' outcome is sticky for the thread
      } else {
          runProcessingStats.errorsEncountered++;
          threadProcessingOutcomesMap[gmailThreadId] = 'manual'; // Assume manual review if sheet write fails
      }
    } catch (errInner) {
        Logger.log(`[ERROR] MJM App Emails Proc Loop: MsgID ${messageId}. ${errInner.message}\nStack: ${errInner.stack}`);
        threadProcessingOutcomesMap[gmailThreadId]='manual';
        runProcessingStats.errorsEncountered++;
    }
    // Note: `alreadyProcessedMsgIdsThisRun.add(messageId)` was done earlier, before the main try-catch in this loop.
    // This means even if processing for a message fails catastrophically within this loop, it won't be re-attempted in *this specific run*.
    // It *will* be re-attempted in the *next hourly run* if it's not moved out of the "To Process" label.
    Utilities.sleep(200 + Math.floor(Math.random() * 100));
  } 

  Logger.log(`  MJM App Emails Loop End. Stats: Updated ${runProcessingStats.updatedRows}, New ${runProcessingStats.newRowsAdded}, Errors ${runProcessingStats.errorsEncountered}.`);
  MJM_applyFinalLabelsToThreads(threadProcessingOutcomesMap, toProcessLabel, processedLabel, manualReviewLabel); // From MJM_GmailUtils.gs
  if(mainSS)try{MJM_updateDashboardMetrics(mainSS);}catch(eDash){Logger.log(`[WARN] MJM App Emails: Final dashboard update fail: ${eDash.message}`);} // From MJM_Dashboard.gs
  Logger.log(`==== MJM_PROCESS_JOB_APP_EMAILS FINISHED (${new Date().toLocaleString()}) === Total time: ${(new Date().getTime() - SCRIPT_START_TIME.getTime())/1000}s ====`);
} // End of MJM_processJobApplicationEmails

/**
 * Marks stale MJM applications (those not in a final state and not updated recently) as rejected.
 * Intended to be run by a daily time-driven trigger.
 */
function MJM_markStaleApplicationsAsRejected() {
  const SCRIPT_START_TIME = new Date();
  Logger.log(`\n==== STARTING MJM_MARK_STALE_APPLICATIONS_AS_REJECTED (${SCRIPT_START_TIME.toLocaleString()}) ====`);
  const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : false); // Global_Constants.gs

  const mainSS = MJM_getOrCreateSpreadsheet_Core(); // From MJM_SheetUtils.gs
  if (!mainSS) { 
    Logger.log("[FATAL ERROR] Stale Apps: Main application spreadsheet could not be accessed. Aborting."); 
    return; 
  }
  
  // APP_TRACKER_SHEET_TAB_NAME from Global_Constants.gs
  // MJM_getOrCreateSheet_V2 for robustness
  const { sheet: dataSheet } = MJM_getOrCreateSheet_V2(mainSS, APP_TRACKER_SHEET_TAB_NAME); 
  if (!dataSheet) { 
    Logger.log(`[FATAL ERROR] Stale Apps: Sheet/Tab "${APP_TRACKER_SHEET_TAB_NAME}" could not be accessed. Aborting.`); 
    return; 
  }

  const allDataRange = dataSheet.getDataRange();
  const allSheetValues = allDataRange.getValues(); 

  if (allSheetValues.length <= 1) { 
    Logger.log("[INFO] Stale Apps: No data rows found in the applications sheet to process for staleness.");
    return;
  }

  const currentDate = new Date();
  const staleDateThreshold = new Date(currentDate);
  staleDateThreshold.setDate(currentDate.getDate() - (MJM_STALE_WEEKS_THRESHOLD * 7)); // MJM_STALE_WEEKS_THRESHOLD from MJM_Config.gs
  
  if(DEBUG) Logger.log(`  Stale Apps: Application considered stale if Last Update Date < ${staleDateThreshold.toLocaleDateString()} (Threshold: ${MJM_STALE_WEEKS_THRESHOLD} weeks)`);

  let updatedStaleApplicationsCount = 0;
  let rowsActuallyProcessedForStaleness = 0; 

  for (let i = 1; i < allSheetValues.length; i++) { // Start from 1 to skip header
    const currentRowArray = allSheetValues[i];
    const sheetRowNumberForLog = i + 1;

    const currentAppStatus = String(currentRowArray[MJM_APP_STATUS_COL - 1] || "").trim(); // From MJM_Config.gs
    const lastUpdateDateValue = currentRowArray[MJM_APP_LAST_UPDATE_DATE_COL - 1]; // From MJM_Config.gs
    let lastUpdateDateObject;

    if (lastUpdateDateValue instanceof Date && !isNaN(lastUpdateDateValue.getTime())) {
      lastUpdateDateObject = lastUpdateDateValue;
    } else if (lastUpdateDateValue && typeof lastUpdateDateValue === 'string' && lastUpdateDateValue.trim() !== "") {
      lastUpdateDateObject = new Date(lastUpdateDateValue);
      if (isNaN(lastUpdateDateObject.getTime())) { 
        if(DEBUG) Logger.log(`    Row ${sheetRowNumberForLog} Skip Stale Check: Invalid date string for Last Update Date: "${lastUpdateDateValue}"`);
        continue; 
      }
    } else {
      if(DEBUG) Logger.log(`    Row ${sheetRowNumberForLog} Skip Stale Check: Missing or unparseable Last Update Date field.`);
      continue; 
    }
    rowsActuallyProcessedForStaleness++; 

    // MJM_STALE_FINAL_STATUSES_FOR_CHECK is a Set from MJM_Config.gs
    // MANUAL_REVIEW_NEEDED_TEXT from Global_Constants.gs
    if (MJM_STALE_FINAL_STATUSES_FOR_CHECK.has(currentAppStatus) || !currentAppStatus || currentAppStatus === MANUAL_REVIEW_NEEDED_TEXT) {
      if(DEBUG && currentAppStatus) Logger.log(`    Row ${sheetRowNumberForLog} Skip Stale: Status "${currentAppStatus}" is final, empty, or requires manual review.`);
      continue;
    }

    if (lastUpdateDateObject.getTime() >= staleDateThreshold.getTime()) { 
      if(DEBUG) Logger.log(`    Row ${sheetRowNumberForLog} Skip Stale: Last Update Date ${lastUpdateDateObject.toLocaleDateString()} is NOT older than threshold date ${staleDateThreshold.toLocaleDateString()}.`);
      continue;
    }

    if(DEBUG) Logger.log(`[INFO] Stale Apps: Row ${sheetRowNumberForLog} - MARKING STALE. Last Update: ${lastUpdateDateObject.toLocaleDateString()}, Old Status: "${currentAppStatus}" -> New Status: "${MJM_APP_REJECTED_STATUS}".`);
    
    allSheetValues[i][MJM_APP_STATUS_COL - 1] = MJM_APP_REJECTED_STATUS; // From MJM_Config.gs
    allSheetValues[i][MJM_APP_LAST_UPDATE_DATE_COL - 1] = currentDate; 
    allSheetValues[i][MJM_APP_PROCESSED_TIMESTAMP_COL - 1] = currentDate; // From MJM_Config.gs
    updatedStaleApplicationsCount++;
  }

  if(DEBUG) Logger.log(`  Stale Apps: Total rows with valid dates considered for staleness check: ${rowsActuallyProcessedForStaleness}.`);
  if (updatedStaleApplicationsCount > 0) {
    try {
      allDataRange.setValues(allSheetValues); 
      Logger.log(`[INFO] Stale Apps: Successfully updated ${updatedStaleApplicationsCount} stale applications to "${MJM_APP_REJECTED_STATUS}".`);
    } catch (eWrite) {
      Logger.log(`[ERROR] Stale Apps: Failed to write updated values back to sheet: ${eWrite.message}\nStack: ${eWrite.stack}`);
    }
  } else {
    Logger.log("[INFO] Stale Apps: No stale applications found meeting all criteria for update.");
  }
  Logger.log(`==== MJM_MARK_STALE_APPLICATIONS_AS_REJECTED END ==== Total Time: ${(new Date().getTime()-SCRIPT_START_TIME.getTime()) / 1000}s ====`);
} // End of MJM_markStaleApplicationsAsRejected
