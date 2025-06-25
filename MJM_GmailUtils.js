// File: MJM_GmailUtils.gs
// Description: Contains utility functions for interacting with Gmail, primarily label
// creation and management, specific to MJM module needs.
// Relies on GLOBAL_DEBUG_MODE from Global_Constants.gs.

/**
 * Gets an existing Gmail label by name, or creates it if it doesn't exist.
 * This is a general utility potentially used by multiple MJM setup functions.
 * @param {string} labelName The full desired name of the label (e.g., "Parent/Child").
 * @return {GoogleAppsScript.Gmail.GmailLabel|null} The GmailLabel object or null on failure.
 */
function getOrCreateLabel(labelName) { // This function can remain fairly generic as it's a utility
  if (!labelName || typeof labelName !== 'string' || labelName.trim() === "") {
    Logger.log(`[ERROR] MJM_GmailUtils (getOrCreateLabel): Invalid labelName provided: "${labelName}"`);
    return null;
  }
  let label = null;
  try {
    label = GmailApp.getUserLabelByName(labelName);
  } catch (e) {
    Logger.log(`[ERROR] MJM_GmailUtils (getOrCreateLabel): Error checking for label "${labelName}": ${e.message}`);
    // This can happen if a parent label doesn't exist when checking for "Parent/Child".
    // The createLabel call below often handles creating parent paths.
  }

  if (!label) {
    if (GLOBAL_DEBUG_MODE) Logger.log(`[DEBUG] MJM_GmailUtils (getOrCreateLabel): Label "${labelName}" not found. Attempting to create...`);
    try {
      label = GmailApp.createLabel(labelName);
      Logger.log(`[INFO] MJM_GmailUtils (getOrCreateLabel): Successfully created label: "${labelName}"`);
    } catch (e) {
      Logger.log(`[ERROR] MJM_GmailUtils (getOrCreateLabel): Failed to create label "${labelName}": ${e.message}\nStack: ${e.stack}`);
      return null;
    }
  } else {
    if (GLOBAL_DEBUG_MODE) Logger.log(`[DEBUG] MJM_GmailUtils (getOrCreateLabel): Label "${labelName}" already exists.`);
  }
  return label;
}

/**
 * Applies final processing labels to Gmail threads based on their processing outcomes.
 * Specific to the MJM Application Tracker email processing flow.
 * @param {Object} threadOutcomes An object mapping threadId to outcome ('manual' or 'done').
 * @param {GoogleAppsScript.Gmail.GmailLabel} sourceProcessingLabel The "To Process" label object to be removed.
 * @param {GoogleAppsScript.Gmail.GmailLabel} targetProcessedLabelObj The "Processed" label object to be added for 'done' outcomes.
 * @param {GoogleAppsScript.Gmail.GmailLabel} targetManualReviewLabelObj The "Manual Review" label object to be added for 'manual' outcomes.
 */
function MJM_applyFinalLabelsToThreads(threadOutcomes, sourceProcessingLabel, targetProcessedLabelObj, targetManualReviewLabelObj) {
  const threadIdsToUpdate = Object.keys(threadOutcomes);
  if (threadIdsToUpdate.length === 0) {
    Logger.log("[INFO] MJM_GmailUtils (ApplyLabels): No thread outcomes to process for App Tracker.");
    return;
  }
  Logger.log(`[INFO] MJM_GmailUtils (ApplyLabels): Applying labels for ${threadIdsToUpdate.length} App Tracker threads.`);

  // Validate label objects
  if (!sourceProcessingLabel || typeof sourceProcessingLabel.getName !== 'function' ||
      !targetProcessedLabelObj || typeof targetProcessedLabelObj.getName !== 'function' ||
      !targetManualReviewLabelObj || typeof targetManualReviewLabelObj.getName !== 'function' ) {
    Logger.log(`[ERROR] MJM_GmailUtils (ApplyLabels): One or more invalid label objects provided. Aborting label application.`);
    return;
  }

  const toProcessLabelName = sourceProcessingLabel.getName(); // For checking current labels
  let successfulLabelChanges = 0;
  let labelErrors = 0;

  for (const threadId of threadIdsToUpdate) {
    const outcome = threadOutcomes[threadId];
    const labelToAdd = (outcome === 'manual') ? targetManualReviewLabelObj : targetProcessedLabelObj;
    const labelNameToAdd = labelToAdd.getName();

    try {
      const thread = GmailApp.getThreadById(threadId);
      if (!thread) {
        Logger.log(`[WARN] MJM_GmailUtils (ApplyLabels): Thread ${threadId} not found. Skipping label update.`);
        labelErrors++;
        continue;
      }

      const currentThreadLabels = thread.getLabels().map(l => l.getName());
      let labelsActuallyChangedThisThread = false;

      // Remove "To Process" label
      if (currentThreadLabels.includes(toProcessLabelName)) {
        try {
          thread.removeLabel(sourceProcessingLabel);
          if (GLOBAL_DEBUG_MODE) Logger.log(`[DEBUG] MJM_GmailUtils (ApplyLabels): Removed "${toProcessLabelName}" from thread ${threadId}`);
          labelsActuallyChangedThisThread = true;
        } catch (eRemove) {
          Logger.log(`[WARN] MJM_GmailUtils (ApplyLabels): Failed to remove "${toProcessLabelName}" from thread ${threadId}: ${eRemove.message}`);
          // Continue to attempt adding the new label
        }
      }

      // Add "Processed" or "Manual Review" label
      if (!currentThreadLabels.includes(labelNameToAdd)) {
        try {
          thread.addLabel(labelToAdd);
          Logger.log(`[INFO] MJM_GmailUtils (ApplyLabels): Added "${labelNameToAdd}" to thread ${threadId}`);
          labelsActuallyChangedThisThread = true;
        } catch (eAdd) {
          Logger.log(`[ERROR] MJM_GmailUtils (ApplyLabels): Failed to add "${labelNameToAdd}" to thread ${threadId}: ${eAdd.message}`);
          labelErrors++;
          continue; // Skip to next thread if critical add fails
        }
      } else {
        if (GLOBAL_DEBUG_MODE) Logger.log(`[DEBUG] MJM_GmailUtils (ApplyLabels): Thread ${threadId} already has target label "${labelNameToAdd}".`);
      }

      if (labelsActuallyChangedThisThread) {
        successfulLabelChanges++;
        Utilities.sleep(200 + Math.floor(Math.random() * 100)); // Pause after making changes to a thread
      }
    } catch (eGeneral) {
      Logger.log(`[ERROR] MJM_GmailUtils (ApplyLabels): General error processing thread ${threadId} for labeling: ${eGeneral.message}`);
      labelErrors++;
    }
  }
  Logger.log(`[INFO] MJM_GmailUtils (ApplyLabels): Finished. Label changes/verifications: ${successfulLabelChanges}. Errors: ${labelErrors}.`);
}
