// File: MJM_Leads_Main.gs (or MRM_Leads_Main.gs)
// Description: Contains the primary functions for the MJM Job Leads Tracker module,
// including initial setup and the ongoing processing of job lead emails.
// Relies on constants from Global_Constants.gs and MJM_Config.gs.

/**
 * Sets up the MJM Job Leads Tracker module components within the provided main spreadsheet.
 * Called by runFullProjectInitialSetup.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} mainSpreadsheet The main application spreadsheet object.
 * @return {{success: boolean, messages: string[]}} An object indicating success and containing setup messages.
 */
function MJM_setupLeadsModule(mainSpreadsheet) { // Name changed to match call from runFullProjectInitialSetup
  const moduleName = "MJM Job Leads Tracker";
  Logger.log(`\n  -- STARTING SETUP for ${moduleName} --`);
  let messages = [];
  let success = true;
  const DEBUG = GLOBAL_DEBUG_MODE; // Global_Constants.gs

  if (!mainSpreadsheet) {
    const errorMsg = `[ERROR] ${moduleName} Setup: Main spreadsheet object NOT PROVIDED. Aborting module setup.`;
    Logger.log(errorMsg);
    return { success: false, messages: [errorMsg] };
  }
  if (DEBUG) Logger.log(`  ${moduleName} Setup: Using spreadsheet "${mainSpreadsheet.getName()}" (ID: ${mainSpreadsheet.getId()})`);

  // --- Step 1: Get/Create and Format Leads Sheet Tab ---
  const leadsSheetName = LEADS_SHEET_TAB_NAME; // From Global_Constants.gs
  let leadsSheet = mainSpreadsheet.getSheetByName(leadsSheetName);

  if (!leadsSheet) {
    if(DEBUG) Logger.log(`  Sheet "${leadsSheetName}" not found. Attempting to create/rename "Sheet1".`);
    const sheets = mainSpreadsheet.getSheets();
    if (sheets.length === 1 && sheets[0].getName() === 'Sheet1' && leadsSheetName !== 'Sheet1') {
        sheets[0].setName(leadsSheetName);
        leadsSheet = sheets[0];
        leadsSheet.clear(); // Clear a renamed "Sheet1"
        Logger.log(`[INFO] ${moduleName} Setup: Renamed existing "Sheet1" to "${leadsSheetName}" and cleared.`);
    } else {
        leadsSheet = mainSpreadsheet.insertSheet(leadsSheetName);
        Logger.log(`[INFO] ${moduleName} Setup: Created new sheet tab: "${leadsSheetName}".`);
    }
    messages.push(`Sheet "${leadsSheetName}": CREATED.`);
  } else {
    Logger.log(`[INFO] ${moduleName} Setup: Found existing sheet tab: "${leadsSheetName}".`);
    messages.push(`Sheet "${leadsSheetName}": Exists.`);
  }

  // Apply formatting using helper from MJM_Leads_SheetUtils.gs
  // MJM_LEADS_SHEET_HEADERS from MJM_Config.gs
  // MJM_Leads_setupSheetFormatting defined in MJM_Leads_SheetUtils.gs
  MJM_Leads_setupSheetFormatting(leadsSheet, MJM_LEADS_SHEET_HEADERS);
  messages.push(`Sheet "${leadsSheetName}": Formatting applied/verified.`);


  // --- Step 2: MJM Leads Gmail Label and Filter Setup ---
  // All MJM_LEADS_GMAIL_... constants are from MJM_Config.gs
  Logger.log(`[INFO] ${moduleName} Setup: Ensuring Leads Gmail labels & filter exist...`);
  try {
    getOrCreateLabel(MJM_MASTER_GMAIL_LABEL_PARENT); Utilities.sleep(100); // General util from MJM_GmailUtils.gs
    getOrCreateLabel(MJM_LEADS_GMAIL_LABEL_PARENT); Utilities.sleep(100);
    getOrCreateLabel(MJM_LEADS_GMAIL_LABEL_NEEDS_PROCESS); Utilities.sleep(100);
    getOrCreateLabel(MJM_LEADS_GMAIL_LABEL_DONE_PROCESS); Utilities.sleep(100);
    messages.push("Leads Gmail Labels: Creation/verification called.");

    let leadsNeedsProcessLabelId = null;
    const advancedGmailService = Gmail;
    if (!advancedGmailService?.Users?.Labels) throw new Error("Advanced Gmail Service (Labels) not available.");
    
    const labelsListResponse = advancedGmailService.Users.Labels.list('me');
    const targetLabelInfo = labelsListResponse.labels?.find(l => l.name === MJM_LEADS_GMAIL_LABEL_NEEDS_PROCESS);

    if (targetLabelInfo?.id) {
      leadsNeedsProcessLabelId = targetLabelInfo.id;
      if(DEBUG) Logger.log(`  Found Label ID for "${MJM_LEADS_GMAIL_LABEL_NEEDS_PROCESS}": ${leadsNeedsProcessLabelId}`);
      
      const filtersListResponse = advancedGmailService.Users.Settings.Filters.list('me');
      const filterExists = (filtersListResponse.filter || []).some(f =>
        f.criteria?.query === MJM_LEADS_GMAIL_FILTER_QUERY &&
        f.action?.addLabelIds?.includes(leadsNeedsProcessLabelId)
      );

      if (!filterExists) {
        const filterResource = {
          criteria: { query: MJM_LEADS_GMAIL_FILTER_QUERY },
          action: { addLabelIds: [leadsNeedsProcessLabelId], removeLabelIds: ['INBOX'] }
        };
        advancedGmailService.Users.Settings.Filters.create(filterResource, 'me');
        messages.push(`Leads Gmail Filter (for query "${MJM_LEADS_GMAIL_FILTER_QUERY}"): CREATED.`);
      } else {
        messages.push(`Leads Gmail Filter (for query "${MJM_LEADS_GMAIL_FILTER_QUERY}"): Exists.`);
      }
    } else {
      throw new Error(`Could not get Gmail Label ID for "${MJM_LEADS_GMAIL_LABEL_NEEDS_PROCESS}". Filter cannot be created.`);
    }
  } catch (e) {
    Logger.log(`[ERROR] ${moduleName} Setup (Gmail Labels/Filter): ${e.message}\nStack: ${e.stack}`);
    messages.push(`Leads Gmail Labels/Filter: FAILED - ${e.message}`);
    success = false;
  }

  // Step 3: UserProperties for Leads Label IDs (Optional - for resilience if label names were to change)
  // Current processing logic (MJM_processJobLeads) gets labels by name. So this is not strictly critical for functionality.
  // But it's good practice to store them if retrieved.
  // MJM_LEADS_USER_PROPERTY_NEEDS_PROCESS_LABEL_ID from MJM_Config.gs
  const userProps = PropertiesService.getUserProperties();
  try {
    const advGmail = Gmail;
    const lblsResp = advGmail.Users.Labels.list('me');
    const needsProcLblInfo = lblsResp.labels?.find(l => l.name === MJM_LEADS_GMAIL_LABEL_NEEDS_PROCESS);
    const doneProcLblInfo = lblsResp.labels?.find(l => l.name === MJM_LEADS_GMAIL_LABEL_DONE_PROCESS);

    if (needsProcLblInfo?.id) userProps.setProperty(MJM_LEADS_USER_PROPERTY_NEEDS_PROCESS_LABEL_ID, needsProcLblInfo.id);
    else Logger.log(`[WARN] ${moduleName} Setup: Could not get/store ID for label "${MJM_LEADS_GMAIL_LABEL_NEEDS_PROCESS}" in UserProperties.`);
    
    if (doneProcLblInfo?.id) userProps.setProperty(MJM_LEADS_USER_PROPERTY_DONE_PROCESS_LABEL_ID, doneProcLblInfo.id);
    else Logger.log(`[WARN] ${moduleName} Setup: Could not get/store ID for label "${MJM_LEADS_GMAIL_LABEL_DONE_PROCESS}" in UserProperties.`);

    if(needsProcLblInfo?.id && doneProcLblInfo?.id) messages.push("Leads Label IDs: Stored/updated in UserProperties.");

  } catch (e) {
     Logger.log(`[WARN] ${moduleName} Setup: Error storing Leads label IDs in UserProperties: ${e.message}`);
     messages.push(`Leads Label IDs UserProperties: Store FAILED - ${e.message}`);
     // Not critical enough to set success = false if primary label/filter setup worked.
  }


  // --- Step 4: Create Time-Driven Trigger for MJM_processJobLeads ---
  const triggerFunctionName = 'MJM_processJobLeads'; // Global name of the function in this file
  try {
    // createDailyAtHourTrigger from MJM_Triggers.gs (or Global_Triggers.gs if you globalized it)
    // This function now DELETES existing clock triggers for this function name before creating.
    if (MJM_createDailyAtHourTrigger(triggerFunctionName, 3)) { // Example: Run daily around 3 AM
      messages.push(`Trigger for "${triggerFunctionName}": CREATED/Re-created daily.`);
    } else {
      // This 'else' might not be hit often if createDailyAtHourTrigger always tries to create
      // unless it itself encounters an error and returns false.
      messages.push(`Trigger for "${triggerFunctionName}": Not newly created (check logs for issues or if an error occurred).`);
      // We might not set success = false here as the trigger function itself logs errors.
    }
  } catch (e) {
    Logger.log(`[ERROR] ${moduleName} Setup (Trigger for ${triggerFunctionName}): ${e.toString()}`);
    messages.push(`Trigger for "${triggerFunctionName}": FAILED - ${e.message}`);
    success = false; // Trigger setup failure is significant
  }

  Logger.log(`  -- FINISHED ${moduleName} SETUP (${success ? "OK" : "ISSUES"}) --`);
  return { success: success, messages: messages };
}

/**
 * Processes emails labeled for job leads for the MJM module.
 * (The rest of the MJM_processJobLeads function from the previous refactor goes here)
 * Key changes: Uses MJM_getOrCreateSpreadsheet_Core(), MJM_Leads_getSheetAndHeaderMap(),
 * MJM-prefixed Gmail labels from MJM_Config, SHARED_GEMINI_API_KEY_PROPERTY from Global,
 * MJM_callGemini_forJobLeads, MJM_parseGeminiResponse_forJobLeads,
 * and MJM_Leads_ specific sheet utils.
 */
function MJM_processJobLeads() {
  const SCRIPT_START_TIME = new Date();
  const moduleName = "MJM Job Leads Processor";
  Logger.log(`\n==== STARTING ${moduleName} (${SCRIPT_START_TIME.toLocaleString()}) ====`);
  const DEBUG = GLOBAL_DEBUG_MODE; // Global_Constants.gs

  // --- Initialization & Config Loading ---
  const userProperties = PropertiesService.getUserProperties();
  const geminiApiKey = userProperties.getProperty(SHARED_GEMINI_API_KEY_PROPERTY); // Global_Constants.gs

  if (!geminiApiKey) {
    Logger.log(`[FATAL] ${moduleName}: Gemini API Key not found (Property: "${SHARED_GEMINI_API_KEY_PROPERTY}"). Aborting.`);
    return;
  }

  // MJM_getOrCreateSpreadsheet_Core from MJM_SheetUtils.gs (uses Global_Constants.gs for APP_SPREADSHEET_ID/FILENAME)
  const mainSpreadsheet = MJM_getOrCreateSpreadsheet_Core();
  if (!mainSpreadsheet) {
    Logger.log(`[FATAL] ${moduleName}: Main application spreadsheet not found. Aborting.`);
    return;
  }

  // LEADS_SHEET_TAB_NAME from Global_Constants.gs
  // MJM_LEADS_SHEET_HEADERS from MJM_Config.gs
  // MJM_Leads_getSheetAndHeaderMap from MJM_Leads_SheetUtils.gs
  const { sheet: dataSheet, headerMap } = MJM_Leads_getSheetAndHeaderMap(mainSpreadsheet, LEADS_SHEET_TAB_NAME, MJM_LEADS_SHEET_HEADERS);
  if (!dataSheet || !headerMap || Object.keys(headerMap).length === 0) {
    Logger.log(`[FATAL] ${moduleName}: Leads sheet "${LEADS_SHEET_TAB_NAME}" or its headers not correctly mapped in "${mainSpreadsheet.getName()}". Aborting.`);
    return;
  }

  // MJM_LEADS_GMAIL_LABEL_NEEDS_PROCESS, MJM_LEADS_GMAIL_LABEL_DONE_PROCESS from MJM_Config.gs
  const needsProcessLabelName = MJM_LEADS_GMAIL_LABEL_NEEDS_PROCESS;
  const doneProcessLabelName = MJM_LEADS_GMAIL_LABEL_DONE_PROCESS;

  const needsProcessLabel = GmailApp.getUserLabelByName(needsProcessLabelName);
  const doneProcessLabel = GmailApp.getUserLabelByName(doneProcessLabelName); // OK if null, just warn later

  if (!needsProcessLabel) {
    Logger.log(`[FATAL] ${moduleName}: Gmail label "${needsProcessLabelName}" not found. Aborting.`);
    return;
  }
  if (!doneProcessLabel) Logger.log(`[WARN] ${moduleName}: Gmail label "${doneProcessLabelName}" not found. Processed threads may not be re-labeled correctly.`);
  if (DEBUG) Logger.log(`[DEBUG] ${moduleName}: Config OK. API Key: ${geminiApiKey.substring(0,5)}... Spreadsheet: "${mainSpreadsheet.getName()}", Data Sheet: "${dataSheet.getName()}"`);

  // --- Email Processing ---
  // MJM_Leads_getProcessedEmailIdsFromSheet from MJM_Leads_SheetUtils.gs
  const processedEmailIds = MJM_Leads_getProcessedEmailIdsFromSheet(dataSheet, headerMap);
  Logger.log(`[INFO] ${moduleName}: Preloaded ${processedEmailIds.size} processed email IDs from sheet "${dataSheet.getName()}".`);

  const LEADS_THREAD_LIMIT_PER_RUN = 10; // Configurable, maybe from Global_Constants.gs or MJM_Config.gs
  const LEADS_MESSAGES_TO_PROCESS_PER_RUN = 15; // Configurable
  let messagesProcessedThisRunCount = 0;

  const threadsToProcess = needsProcessLabel.getThreads(0, LEADS_THREAD_LIMIT_PER_RUN);
  Logger.log(`[INFO] ${moduleName}: Found ${threadsToProcess.length} threads in label "${needsProcessLabelName}".`);

  for (const thread of threadsToProcess) {
    if (messagesProcessedThisRunCount >= LEADS_MESSAGES_TO_PROCESS_PER_RUN) {
      Logger.log(`[INFO] ${moduleName}: Message processing limit (${LEADS_MESSAGES_TO_PROCESS_PER_RUN}) reached for this run.`);
      break;
    }
    const SCRIPT_RUNTIME_SECONDS = (new Date().getTime() - SCRIPT_START_TIME.getTime()) / 1000;
    if (SCRIPT_RUNTIME_SECONDS > 320) { // ~5m 20s, leave buffer for cleanup
      Logger.log(`[WARN] ${moduleName}: Script execution time limit approaching (${SCRIPT_RUNTIME_SECONDS}s). Stopping further thread processing.`);
      break;
    }

    const messagesInThread = thread.getMessages();
    let newMessagesFoundInThread = false;
    let allMessagesInThreadProcessedSuccessfullyThisRun = true;

    for (const message of messagesInThread) {
      if (messagesProcessedThisRunCount >= LEADS_MESSAGES_TO_PROCESS_PER_RUN) break; // Check limit again per message
      const messageId = message.getId();

      if (processedEmailIds.has(messageId)) {
        if(DEBUG) Logger.log(`  Skipping already processed message ID: ${messageId} in thread ${thread.getId()}.`);
        continue; // Skip already processed messages
      }
      
      newMessagesFoundInThread = true;
      messagesProcessedThisRunCount++;
      Logger.log(`\n--- ${moduleName}: Processing new Lead Msg ID: ${messageId}, Subject: "${message.getSubject()}" (${messagesProcessedThisRunCount}/${LEADS_MESSAGES_TO_PROCESS_PER_RUN}) ---`);
      
      let currentMessageOutcomeSuccess = false;
      try {
        const emailBody = message.getPlainBody();
        if (!emailBody || emailBody.trim() === "") {
          Logger.log(`  Msg ${messageId}: Plain body is empty. Marking as 'processed' (nothing to parse).`);
          currentMessageOutcomeSuccess = true; // Successfully determined it's empty
        } else {
          // MJM_callGemini_forJobLeads & MJM_parseGeminiResponse_forJobLeads from MJM_GeminiService.gs
          const geminiApiResponse = MJM_callGemini_forJobLeads(emailBody, geminiApiKey);

          if (geminiApiResponse && geminiApiResponse.success && geminiApiResponse.data) {
            const extractedJobsArray = MJM_parseGeminiResponse_forJobLeads(geminiApiResponse.data);
            if (extractedJobsArray && extractedJobsArray.length > 0) {
              Logger.log(`  Gemini extracted ${extractedJobsArray.length} potential job(s) from msg ${messageId}.`);
              let validJobsWrittenThisMessage = 0;
              for (const jobData of extractedJobsArray) {
                // Basic check: ensure there's a job title, not just "N/A" placeholder from Gemini.
                if (jobData && jobData.jobTitle && String(jobData.jobTitle).trim().toLowerCase() !== 'n/a') {
                  // MJM_Leads_writeJobToSheet from MJM_Leads_SheetUtils.gs
                  MJM_Leads_writeJobToSheet(dataSheet, message, jobData, headerMap);
                  validJobsWrittenThisMessage++;
                } else {
                  if(DEBUG) Logger.log(`    Skipping job item with N/A title from msg ${messageId}: ${JSON.stringify(jobData)}`);
                }
              }
              if (validJobsWrittenThisMessage > 0) currentMessageOutcomeSuccess = true;
              else {
                  Logger.log(`  Msg ${messageId}: Gemini call/parse success, but no valid jobs (all N/A titles or empty array after filtering). Considered processed.`);
                  currentMessageOutcomeSuccess = true; // No actionable jobs, but msg is processed.
              }
            } else {
              Logger.log(`  Msg ${messageId}: Gemini call/parse success, but parsing yielded no job listings array or it was empty.`);
              currentMessageOutcomeSuccess = true; // Message processed, no jobs found.
            }
          } else { // Gemini API call failed or returned no data structure
            Logger.log(`[ERROR] ${moduleName}: Gemini API call FAILED or returned invalid data for msg ${messageId}. Details: ${geminiApiResponse ? geminiApiResponse.error : 'Response object or data was null'}`);
            // MJM_Leads_writeErrorEntryToSheet from MJM_Leads_SheetUtils.gs
            MJM_Leads_writeErrorEntryToSheet(dataSheet, message, "Gemini API Call/Parse Failed for Leads", geminiApiResponse?.error || "Unknown Gemini API error or invalid response", headerMap);
            allMessagesInThreadProcessedSuccessfullyThisRun = false; // Mark thread as having an issue
          }
        }
      } catch (e) {
        Logger.log(`[FATAL SCRIPT ERROR] ${moduleName}: Uncaught exception processing msg ${messageId}: ${e.message}\nStack: ${e.stack}`);
        MJM_Leads_writeErrorEntryToSheet(dataSheet, message, "Critical Script Error during lead processing", e.toString(), headerMap);
        allMessagesInThreadProcessedSuccessfullyThisRun = false; // Mark thread as having an issue
      }

      if (currentMessageOutcomeSuccess) {
        processedEmailIds.add(messageId); // Add to processed set only if successfully handled
      }
      Utilities.sleep(1200 + Math.floor(Math.random() * 800)); // Pause between messages
    } // End loop over messages in a thread

    // After processing all (new) messages in a thread:
    if (newMessagesFoundInThread && allMessagesInThreadProcessedSuccessfullyThisRun) {
      if (doneProcessLabel) thread.removeLabel(needsProcessLabel).addLabel(doneProcessLabel);
      else thread.removeLabel(needsProcessLabel); // At least remove from "NeedsProcess"
      Logger.log(`[INFO] ${moduleName}: Thread ID ${thread.getId()} fully processed for new leads and moved from "${needsProcessLabelName}".`);
    } else if (newMessagesFoundInThread) { // Some new messages existed but one or more failed
      Logger.log(`[WARN] ${moduleName}: Thread ID ${thread.getId()} had processing issues with one or more new messages. NOT moved from "${needsProcessLabelName}". Will be retried next run.`);
    } else if (!newMessagesFoundInThread && messagesInThread.length > 0) { // Thread had messages, but all were already in processedEmailIds
      if(DEBUG) Logger.log(`[DEBUG] ${moduleName}: Thread ID ${thread.getId()} contained only previously processed messages. Ensuring it's labeled 'Done'.`);
      if (doneProcessLabel) thread.removeLabel(needsProcessLabel).addLabel(doneProcessLabel);
      else thread.removeLabel(needsProcessLabel);
    } else { // Thread was empty to begin with or became effectively empty after skipping processed messages
      if(DEBUG) Logger.log(`[DEBUG] ${moduleName}: Thread ID ${thread.getId()} appears empty or had no new messages. Removing from "${needsProcessLabelName}" if still present.`);
      try { if (thread.getLabels().some(l=>l.getName() === needsProcessLabelName)) thread.removeLabel(needsProcessLabel); }
      catch(eLbl) { if(DEBUG) Logger.log(`  Minor error removing label from likely already unlabelled thread ${thread.getId()}: ${eLbl}`);}
    }
    Utilities.sleep(400); // Pause between threads
  } // End loop over threads

  Logger.log(`\n==== ${moduleName} FINISHED (${new Date().toLocaleString()}) === Total Time: ${(new Date().getTime() - SCRIPT_START_TIME.getTime())/1000}s. Messages processed this run: ${messagesProcessedThisRunCount} ====`);
}
