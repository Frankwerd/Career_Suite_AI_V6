// File: MJM_AdminUtils.gs (or MRM_AdminUtils.gs)
// Description: Contains administrative utility functions for project setup and configuration,
// such as managing shared API keys (Gemini & Groq) stored in UserProperties.
// Relies on constants from Global_Constants.gs.

/**
 * Provides a UI prompt to set the SHARED Gemini API Key in UserProperties.
 * Uses SHARED_GEMINI_API_KEY_PROPERTY constant from Global_Constants.gs.
 */
function setSharedGeminiApiKey_UI() {
  let ui;
  const apiKeyProperty = SHARED_GEMINI_API_KEY_PROPERTY; // From Global_Constants.gs
  const serviceName = "Gemini";

  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    Logger.log(`[WARN] MJM_AdminUtils (${serviceName} Key UI): Spreadsheet UI context not available. User needs to run from sheet.`);
    // Using Browser.msgBox as a fallback if ui is not available when this is called.
    try {
      Browser.msgBox(`Cannot Open UI`, `This function must be run from within the Google Sheet to set the ${serviceName} API Key.`, Browser.Buttons.OK);
    } catch (eBrowser) {
      Logger.log(`[ERROR] MJM_AdminUtils (${serviceName} Key UI): Fallback Browser.msgBox also failed. ${eBrowser.message}`);
    }
    return;
  }

  const currentKey = PropertiesService.getUserProperties().getProperty(apiKeyProperty);
  const promptMessage = `Enter the SHARED Google AI ${serviceName} API Key.\nThis will be stored in UserProperties under the key: "${apiKeyProperty}".\n${currentKey ? '(Current key will be OVERWRITTEN)' : '(No key currently set)'}\n\nLeave blank to CLEAR the existing key.`;
  const response = ui.prompt(`Set Shared ${serviceName} API Key`, promptMessage, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const apiKey = response.getResponseText().trim();
    if (apiKey === "") { // User left it blank to clear
      PropertiesService.getUserProperties().deleteProperty(apiKeyProperty);
      ui.alert(`${serviceName} API Key Cleared`, `The ${serviceName} API Key stored under "${apiKeyProperty}" has been cleared.`, ui.ButtonSet.OK); // <<< ADDED ui.ButtonSet.OK
      Logger.log(`[INFO] MJM_AdminUtils: Cleared ${serviceName} API Key for property "${apiKeyProperty}".`);
    } else if (apiKey && apiKey.startsWith("AIza") && apiKey.length > 30) { // Basic validation for Gemini keys
      PropertiesService.getUserProperties().setProperty(apiKeyProperty, apiKey);
      ui.alert(`${serviceName} API Key Saved`, `The ${serviceName} API Key has been saved successfully under property: "${apiKeyProperty}".`, ui.ButtonSet.OK); // <<< ADDED ui.ButtonSet.OK
      Logger.log(`[INFO] MJM_AdminUtils: Updated ${serviceName} API Key for property "${apiKeyProperty}".`);
    } else {
      ui.alert(`${serviceName} API Key Not Saved`, `The entered key does not appear to be a valid ${serviceName} API key (should start with "AIza" and be >30 chars). Please try again.`, ui.ButtonSet.OK); // <<< WAS CORRECT HERE
    }
  } else {
    ui.alert(`${serviceName} API Key Setup Cancelled`, `The ${serviceName} API key setup process was cancelled. No changes made.`, ui.ButtonSet.OK); // <<< WAS CORRECT HERE
  }
}
/**
 * Provides a UI prompt to set the SHARED Groq API Key in UserProperties.
 * Uses SHARED_GROQ_API_KEY_PROPERTY constant from Global_Constants.gs.
 */
function setSharedGroqApiKey_UI() {
  let ui;
  const apiKeyProperty = SHARED_GROQ_API_KEY_PROPERTY; // From Global_Constants.gs
  const serviceName = "Groq";

  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    Logger.log(`[WARN] MJM_AdminUtils (${serviceName} Key UI): Spreadsheet UI context not available.`);
    try {
      Browser.msgBox(`Cannot Open UI`, `This function must be run from within the Google Sheet to set the ${serviceName} API Key.`, Browser.Buttons.OK);
    } catch (eBrowser) {
      Logger.log(`[ERROR] MJM_AdminUtils (${serviceName} Key UI): Fallback Browser.msgBox also failed. ${eBrowser.message}`);
    }
    return;
  }

  const currentKey = PropertiesService.getUserProperties().getProperty(apiKeyProperty);
  const promptMessage = `Enter the SHARED ${serviceName} API Key.\nThis will be stored in UserProperties under the key: "${apiKeyProperty}".\n${currentKey ? '(Current key will be OVERWRITTEN)' : '(No key currently set)'}\n\nLeave blank to CLEAR the existing key. Groq keys usually start with "gsk_".`;
  const response = ui.prompt(`Set Shared ${serviceName} API Key`, promptMessage, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const apiKey = response.getResponseText().trim();
    if (apiKey === "") { // User left it blank to clear
      PropertiesService.getUserProperties().deleteProperty(apiKeyProperty);
      ui.alert(`${serviceName} API Key Cleared`, `The ${serviceName} API Key stored under "${apiKeyProperty}" has been cleared.`, ui.ButtonSet.OK); // <<< CORRECTED: Added ui.ButtonSet.OK
      Logger.log(`[INFO] MJM_AdminUtils: Cleared ${serviceName} API Key for property "${apiKeyProperty}".`);
    } else if (apiKey && apiKey.startsWith("gsk_") && apiKey.length > 40) { // Basic validation for Groq keys
      PropertiesService.getUserProperties().setProperty(apiKeyProperty, apiKey);
      ui.alert(`${serviceName} API Key Saved`, `The ${serviceName} API Key has been saved successfully under property: "${apiKeyProperty}".`, ui.ButtonSet.OK); // <<< CORRECTED: Added ui.ButtonSet.OK
      Logger.log(`[INFO] MJM_AdminUtils: Updated ${serviceName} API Key for property "${apiKeyProperty}".`);
    } else {
      // This one was likely correct already as it's the 'else' for invalid key format
      ui.alert(`${serviceName} API Key Not Saved`, `The entered key does not appear to be a valid ${serviceName} API key (should start with "gsk_" and be quite long). Please try again.`, ui.ButtonSet.OK); 
    }
  } else {
    // This one was likely correct already for cancellation
    ui.alert(`${serviceName} API Key Setup Cancelled`, `The ${serviceName} API key setup process was cancelled. No changes made.`, ui.ButtonSet.OK); 
  }
}

/**
 * TEMPORARY: Manually sets the SHARED Gemini API Key in UserProperties.
 * Edit YOUR_GEMINI_KEY_HERE in the code before running.
 * REMOVE OR CLEAR THE KEY FROM CODE AFTER RUNNING FOR SECURITY.
 */
function TEMPORARY_manualSetSharedGeminiApiKey() {
  const YOUR_GEMINI_KEY_HERE = 'AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'; // <<< EDIT THIS LINE
  const propertyName = SHARED_GEMINI_API_KEY_PROPERTY; // From Global_Constants.gs
  const serviceName = "Gemini";

  if (YOUR_GEMINI_KEY_HERE === 'AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' || YOUR_GEMINI_KEY_HERE.trim() === '') {
    const msg = `ERROR: ${serviceName} API Key not set in TEMPORARY_manualSetShared${serviceName}ApiKey function. Edit the script code first. Target UserProperty: "${propertyName}".`;
    Logger.log(msg);
    try { SpreadsheetApp.getUi().alert('Action Required', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) { Browser.msgBox('Action Required', msg, Browser.Buttons.OK); }
    return;
  }
  PropertiesService.getUserProperties().setProperty(propertyName, YOUR_GEMINI_KEY_HERE);
  const successMsg = `UserProperty "${propertyName}" has been MANUALLY SET with the hardcoded ${serviceName} API Key. IMPORTANT: For security, remove or comment out this function, or at least clear the key variable in the code.`;
  Logger.log(successMsg);
  try { SpreadsheetApp.getUi().alert('API Key Manually Set', successMsg, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) { Browser.msgBox('API Key Manually Set', successMsg, Browser.Buttons.OK); }
}

/**
 * TEMPORARY: Manually sets the SHARED Groq API Key in UserProperties.
 * Edit YOUR_GROQ_KEY_HERE in the code before running.
 */
function TEMPORARY_manualSetSharedGroqApiKey() {
  const YOUR_GROQ_KEY_HERE = 'gsk_XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'; // <<< EDIT THIS LINE
  const propertyName = SHARED_GROQ_API_KEY_PROPERTY; // From Global_Constants.gs
  const serviceName = "Groq";

  if (YOUR_GROQ_KEY_HERE === 'gsk_XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' || YOUR_GROQ_KEY_HERE.trim() === '') {
    const msg = `ERROR: ${serviceName} API Key not set in TEMPORARY_manualSetShared${serviceName}ApiKey function. Edit script code first. Target UserProperty: "${propertyName}".`;
    Logger.log(msg);
    try { SpreadsheetApp.getUi().alert('Action Required', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) { Browser.msgBox('Action Required', msg, Browser.Buttons.OK); }
    return;
  }
  PropertiesService.getUserProperties().setProperty(propertyName, YOUR_GROQ_KEY_HERE);
  const successMsg = `UserProperty "${propertyName}" has been MANUALLY SET with the hardcoded ${serviceName} API Key. IMPORTANT: For security, remove/comment out function or clear key variable.`;
  Logger.log(successMsg);
  try { SpreadsheetApp.getUi().alert('API Key Manually Set', successMsg, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) { Browser.msgBox('API Key Manually Set', successMsg, Browser.Buttons.OK); }
}


/**
 * Displays all UserProperties set for this script project to the logs.
 * Sensitive values like API keys are partially masked.
 */
function showAllUserProperties() {
  const userProps = PropertiesService.getUserProperties().getProperties();
  let logOutput = "[INFO] MJM_AdminUtils: Current UserProperties for this script project:\n";
  let foundKeysCount = 0; // To track if any properties are found

  if (Object.keys(userProps).length === 0) {
    logOutput += "  (No UserProperties are currently set for this project)\n";
  } else {
    foundKeysCount = Object.keys(userProps).length;
    for (const key in userProps) {
      let value = userProps[key];
      // Mask sensitive values
      if (key.toUpperCase().includes('API') || key.toUpperCase().includes('KEY') || key.toUpperCase().includes('SECRET')) {
        if (value && typeof value === 'string' && value.length > 10) {
          value = `${value.substring(0, 4)}...${value.substring(value.length - 4)} (Length: ${value.length})`;
        } else if (value && typeof value === 'string') {
            value = "**** (Short Value, potentially sensitive)";
        }
      }
      logOutput += `  - ${key}: ${value}\n`;
    }
  }
  Logger.log(logOutput); // Log it first, always.

  // --- More Detailed UI Alert ---
  const title = "User Properties Logged";
  let alertMessage = `All UserProperties (${foundKeysCount} found) for this script project have been written to the script's execution logs.\n\n`;
  alertMessage += `To view them:\n`;
  alertMessage += `1. Open the Apps Script editor (Extensions > Apps Script).\n`;
  alertMessage += `2. On the left sidebar, click the "Executions" icon (looks like a play button with lines, or a clock).\n`;
  alertMessage += `3. Find the most recent execution for the function "showAllUserProperties".\n`;
  alertMessage += `4. Click on that execution row to open its details.\n`;
  alertMessage += `5. Look for the log entry starting with "[INFO] MJM_AdminUtils: Current UserProperties...".\n\n`;
  alertMessage += `API Keys and other sensitive values will be partially masked in this log for security.`;

  try {
    SpreadsheetApp.getUi().alert(title, alertMessage, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) { // Fallback if no Spreadsheet UI (e.g., run from editor directly)
    try {
      Browser.msgBox(title, alertMessage, Browser.Buttons.OK);
    } catch (e2) {
      Logger.log("[WARN] showAllUserProperties: Could not display UI alert. Logs are still available in Script Editor Executions.");
      // Silently fail if no UI whatsoever is available (e.g. headless custom function trigger - though this is admin util)
    }
  }
}


// TODO: Add setProfileSheetConfiguration_UI() if needed,
// to set PROFILE_DATA_SHEET_NAME or other sheet config in UserProperties.
// Given we're aiming for a single spreadsheet (APP_SPREADSHEET_ID) and fixed global tab names like
// PROFILE_DATA_SHEET_NAME, this function might not be strictly necessary unless you want users
// to be able to rename those key tabs and store the custom name.
/*
function setProfileSheetConfiguration_UI() {
  // Example: To store the PROFILE_DATA_SHEET_NAME if it were user-configurable
  // For now, it's a global constant, so this is less critical.
  // const propName = 'USER_CONFIG_PROFILE_SHEET_NAME';
  // const currentName = PropertiesService.getUserProperties().getProperty(propName) || PROFILE_DATA_SHEET_NAME; // Default to global const
  // ... UI prompt logic ...
}
*/
