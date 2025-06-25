// File: MJM_Triggers.gs (or MRM_Triggers.gs)
// Description: Contains functions for creating, verifying, and managing
// time-driven triggers specifically for the MJM (Master Job Manager) modules.
// Relies on GLOBAL_DEBUG_MODE from Global_Constants.gs for logging.

/**
 * Creates or verifies a time-based trigger that runs every N hours FOR AN MJM FUNCTION.
 * Checks if a trigger for the specified function name already exists.
 * If one exists, it assumes it's correctly configured and does not delete/recreate.
 *
 * @param {string} mjmFunctionName The name of the MJM function to be triggered (e.g., "MJM_processJobApplicationEmails").
 * @param {number} [everyNHours=1] The interval in hours for the trigger.
 * @return {boolean} True if a new trigger was created in this call, false otherwise (already existed or error).
 */
function MJM_createHourlyTrigger(mjmFunctionName, everyNHours = 1) {
  if (!mjmFunctionName || typeof mjmFunctionName !== 'string' || mjmFunctionName.trim() === "") {
    Logger.log(`[ERROR] MJM_Triggers (MJM_createHourlyTrigger): Invalid or empty mjmFunctionName provided.`);
    return false;
  }
  if (typeof everyNHours !== 'number' || everyNHours < 1 || everyNHours > 23) {
    Logger.log(`[ERROR] MJM_Triggers (MJM_createHourlyTrigger): Invalid 'everyNHours' value: ${everyNHours}. Must be 1-23.`);
    return false;
  }

  let triggerExists = false;
  try {
    const projectTriggers = ScriptApp.getProjectTriggers();
    for (const trigger of projectTriggers) {
      if (trigger.getHandlerFunction() === mjmFunctionName &&
          trigger.getEventType() === ScriptApp.EventType.CLOCK &&
          trigger.getTriggerSource() === ScriptApp.TriggerSource.CLOCK) {
        // Basic check: assumes if one exists for this function & type, it's correctly set up.
        triggerExists = true;
        break;
      }
    }

    if (!triggerExists) {
      ScriptApp.newTrigger(mjmFunctionName)
        .timeBased()
        .everyHours(everyNHours)
        .create();
      Logger.log(`[INFO] MJM_Triggers (MJM_createHourlyTrigger): ${everyNHours}-hourly trigger CREATED for MJM function "${mjmFunctionName}".`);
    } else {
      if (GLOBAL_DEBUG_MODE) Logger.log(`[DEBUG] MJM_Triggers (MJM_createHourlyTrigger): ${everyNHours}-hourly trigger for MJM function "${mjmFunctionName}" already exists.`);
    }
  } catch (e) {
    Logger.log(`[ERROR] MJM_Triggers (MJM_createHourlyTrigger): Failed for "${mjmFunctionName}": ${e.message}\nStack: ${e.stack}`);
    return false;
  }
  return !triggerExists; // True if NEWLY created
}

/**
 * Creates or verifies a daily time-based trigger FOR AN MJM FUNCTION to run at a specific hour.
 * Ensures only one such daily trigger for the given mjmFunctionName is active by deleting existing clock triggers for that function first.
 *
 * @param {string} mjmFunctionName The name of the MJM function to be triggered (e.g., "MJM_markStaleApplicationsAsRejected").
 * @param {number} [hourOfDay=2] The approximate hour (0-23) in the script's timezone.
 * @return {boolean} True if a new trigger was created in this call (after potential deletions), false if an error occurred.
 */
function MJM_createDailyAtHourTrigger(mjmFunctionName, hourOfDay = 2) {
  if (!mjmFunctionName || typeof mjmFunctionName !== 'string' || mjmFunctionName.trim() === "") {
    Logger.log(`[ERROR] MJM_Triggers (MJM_createDailyAtHourTrigger): Invalid or empty mjmFunctionName provided.`);
    return false;
  }
  if (typeof hourOfDay !== 'number' || hourOfDay < 0 || hourOfDay > 23) {
    Logger.log(`[ERROR] MJM_Triggers (MJM_createDailyAtHourTrigger): Invalid 'hourOfDay' value: ${hourOfDay}. Must be 0-23.`);
    return false;
  }

  let newTriggerCreatedThisCall = false;
  try {
    const projectTriggers = ScriptApp.getProjectTriggers();
    for (const trigger of projectTriggers) {
      if (trigger.getHandlerFunction() === mjmFunctionName &&
          trigger.getEventType() === ScriptApp.EventType.CLOCK && // Ensure it's a clock-based trigger
          trigger.getTriggerSource() === ScriptApp.TriggerSource.CLOCK) {
        Logger.log(`[INFO] MJM_Triggers (MJM_createDailyAtHourTrigger): Deleting existing CLOCK trigger for MJM function "${mjmFunctionName}" to ensure clean daily setup.`);
        ScriptApp.deleteTrigger(trigger);
        // Keep searching and deleting any other clock triggers for the same function
      }
    }

    // Create the new daily trigger
    ScriptApp.newTrigger(mjmFunctionName)
      .timeBased()
      .everyDays(1)
      .atHour(hourOfDay)
      .inTimezone(Session.getScriptTimeZone())
      .create();
    newTriggerCreatedThisCall = true;
    Logger.log(`[INFO] MJM_Triggers (MJM_createDailyAtHourTrigger): Daily trigger CREATED for MJM function "${mjmFunctionName}" around ${hourOfDay}:00 ${Session.getScriptTimeZone()}.`);

  } catch (e) {
    Logger.log(`[ERROR] MJM_Triggers (MJM_createDailyAtHourTrigger): Failed for "${mjmFunctionName}": ${e.message}\nStack: ${e.stack}`);
    return false;
  }
  return newTriggerCreatedThisCall;
}

/**
 * Deletes all project triggers for a specific MJM handler function.
 * @param {string} mjmFunctionName The name of the MJM handler function whose triggers should be deleted.
 */
function MJM_deleteAllTriggersForFunction(mjmFunctionName) {
    if (!mjmFunctionName || typeof mjmFunctionName !== 'string' || mjmFunctionName.trim() === "") {
        Logger.log(`[ERROR] MJM_Triggers (MJM_deleteAllTriggersForFunction): Invalid mjmFunctionName.`);
        return;
    }
    try {
        const triggers = ScriptApp.getProjectTriggers();
        let count = 0;
        for (const trigger of triggers) {
            if (trigger.getHandlerFunction() === mjmFunctionName) {
                ScriptApp.deleteTrigger(trigger);
                count++;
            }
        }
        Logger.log(`[INFO] MJM_Triggers (MJM_deleteAllTriggersForFunction): Deleted ${count} trigger(s) for MJM function "${mjmFunctionName}".`);
    } catch (e) {
        Logger.log(`[ERROR] MJM_Triggers (MJM_deleteAllTriggersForFunction): Error deleting triggers for "${mjmFunctionName}": ${e.message}`);
    }
}
