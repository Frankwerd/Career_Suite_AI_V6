// File: RTS_GroqService.gs
// Description: Provides functions for interacting with the Groq API.
// Handles API key retrieval from PropertiesService (using SHARED_GROQ_API_KEY_PROPERTY)
// and makes HTTP requests to the Groq API endpoint.
// Relies on constants from Global_Constants.gs.

/**
 * Calls the Groq API with the provided prompt text using an OpenAI-compatible chat completions endpoint.
 * This function is specifically for RTS module needs.
 *
 * @param {string} promptText The user's content/message for the prompt.
 * @param {string} [modelName=DEFAULT_GROQ_MODEL] The specific Groq model to use.
 *        Defaults to DEFAULT_GROQ_MODEL from Global_Constants.gs.
 * @param {number} [temperature=0.2] Optional. The temperature for generation (0.0 to 2.0).
 * @param {number} [maxTokens=2048] Optional. Maximum number of tokens for the completion.
 * @param {string} [systemContent="You are a helpful..."] Optional. A system message.
 * @return {string|null} The text content of the Groq API response on success,
 *                       or an "ERROR: ..." string on failure or if API key is missing.
 */
function RTS_callGroq( // Renamed for clarity if MJM were to have its own Groq caller for different tasks
  promptText,
  modelName = DEFAULT_GROQ_MODEL, // From Global_Constants.gs
  temperature = 0.2,
  maxTokens = 2048,
  systemContent = "You are a helpful and meticulous AI assistant. Respond ONLY with the requested format (e.g., JSON). Do not include any explanatory text or markdown formatting before or after the JSON output."
) {
  // SHARED_GROQ_API_KEY_PROPERTY from Global_Constants.gs
  // GLOBAL_DEBUG_MODE from Global_Constants.gs
  const DEBUG = GLOBAL_DEBUG_MODE;
  const scriptProperties = PropertiesService.getUserProperties(); // Switched to UserProperties for wider accessibility via Admin UI
  const apiKey = scriptProperties.getProperty(SHARED_GROQ_API_KEY_PROPERTY);

  if (!apiKey) {
    Logger.log(`[ERROR] RTS_GroqService (RTS_callGroq): Groq API Key not found. Please set property: '${SHARED_GROQ_API_KEY_PROPERTY}' in User Properties (via Admin menu).`);
    return `ERROR: Groq API Key Missing (Property: ${SHARED_GROQ_API_KEY_PROPERTY})`;
  }

  const API_ENDPOINT = "https://api.groq.com/openai/v1/chat/completions";
  let messages = [];
  if (systemContent && systemContent.trim() !== "") {
    messages.push({"role": "system", "content": systemContent});
  }
  messages.push({"role": "user", "content": promptText});

  const payload = {
    "messages": messages,
    "model": modelName,
    "temperature": temperature,
    "max_tokens": maxTokens,
    "top_p": 1,
    "stream": false,
    "stop": null
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': { // << CORRECTED HERE
      'Authorization': 'Bearer ' + apiKey
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  if(DEBUG) Logger.log(`  RTS_callGroq: Calling Groq API (Model: ${modelName}, Temp: ${temperature}). User Prompt Len: ${promptText.length}`);
  try {
    const response = UrlFetchApp.fetch(API_ENDPOINT, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);
      if (jsonResponse.choices && jsonResponse.choices[0]?.message?.content) {
        const generatedText = jsonResponse.choices[0].message.content;
        if(DEBUG) Logger.log(`    RTS_callGroq: Success from Groq. Output length: ${generatedText.length}`);
        return generatedText.trim();
      } else if (jsonResponse.error) {
        Logger.log(`[ERROR] RTS_GroqService (RTS_callGroq): Groq API (HTTP 200) returned error object: ${JSON.stringify(jsonResponse.error)}`);
        return `ERROR: Groq API Error (RTS) - ${jsonResponse.error.message || JSON.stringify(jsonResponse.error)}`;
      } else {
        Logger.log(`[ERROR] RTS_GroqService (RTS_callGroq): Groq response structure unexpected (HTTP 200). Full Response (first 1000): ${responseBody.substring(0,1000)}`);
        return "ERROR: Unexpected API Response Structure from Groq (RTS) - HTTP 200";
      }
    } else {
      Logger.log(`[ERROR] RTS_GroqService (RTS_callGroq): Groq API call failed. HTTP Code: ${responseCode}. Body (first 500): ${responseBody.substring(0, 500)}...`);
      let errorMessage = `ERROR: API Call Failed (RTS_Groq) - HTTP ${responseCode}`;
      try { const errorJson = JSON.parse(responseBody); if (errorJson.error?.message) errorMessage += `: ${errorJson.error.message}`; }
      catch (e) { /* Ignore if error body not JSON */ }
      return errorMessage;
    }
  } catch (e) {
    Logger.log(`[EXCEPTION] RTS_GroqService (RTS_callGroq): ${e.toString()}\nStack: ${e.stack}`);
    return `ERROR: Exception during Groq API call (RTS) - ${e.message}`;
  }
}

/**
 * UI-triggered function to allow users to set the SHARED Groq API Key.
 * NOTE: This function is provided here for potential standalone RTS testing.
 * For the integrated MJM application, the primary UI for setting this key
 * should be through MJM_AdminUtils.gs -> setSharedGroqApiKey_UI().
 * This version uses UserProperties to align with centralized storage.
 */
function RTS_SET_SHARED_GROQ_API_KEY_FOR_TESTING_UI() { // Renamed to indicate testing purpose if kept here
  let ui;
  const apiKeyProperty = SHARED_GROQ_API_KEY_PROPERTY; // From Global_Constants.gs
  const serviceName = "Groq (RTS Test Util)";

  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    Logger.log(`[WARN] RTS_GroqService (${serviceName} Key UI): Spreadsheet UI context not available.`);
    // Browser.msgBox might not be available in all contexts where RTS code is eventually run (e.g. Add-on backend)
    // For a simple script-editor run test, Browser.inputBox is better if SpreadsheetApp.getUi() fails
    try {
        const currentKeyInfo = PropertiesService.getUserProperties().getProperty(apiKeyProperty) ? "(Overwrite)" : "(None set)";
        const apiKey = Browser.inputBox(`${serviceName} API Key Setup`, `Enter SHARED Groq API Key ${currentKeyInfo} for property '${apiKeyProperty}':`, Browser.Buttons.OK_CANCEL);
        if (apiKey !== 'cancel' && apiKey !== null) {
            if (apiKey.trim() === "") {
                 PropertiesService.getUserProperties().deleteProperty(apiKeyProperty);
                 Browser.msgBox(`${serviceName} API Key Cleared`, `Property '${apiKeyProperty}' cleared.`, Browser.Buttons.OK);
            } else if (apiKey.trim().startsWith("gsk_")) {
                PropertiesService.getUserProperties().setProperty(apiKeyProperty, apiKey.trim());
                Browser.msgBox(`${serviceName} API Key Saved`, `Property '${apiKeyProperty}' updated.`, Browser.Buttons.OK);
            } else {
                 Browser.msgBox(`${serviceName} API Key NOT Saved`, `Invalid format (should start "gsk_").`, Browser.Buttons.OK);
            }
        } else {
             Browser.msgBox(`${serviceName} API Key Cancelled`, `No changes made.`, Browser.Buttons.OK);
        }
    } catch (e2) {
        Logger.log(`RTS_GroqService (${serviceName} Key UI): Also failed with Browser.inputBox: ${e2.message}. Key cannot be set via this UI utility currently.`);
    }
    return;
  }
  // Fallback to SpreadsheetApp.getUi() logic if it was obtained
  const currentKey = PropertiesService.getUserProperties().getProperty(apiKeyProperty);
  const promptMessage = `Enter SHARED Groq API Key.\nStored as UserProperty: "${apiKeyProperty}".\n${currentKey ? '(Overwrite current)' : '(None set)'}\n\nLeave blank to CLEAR. Groq keys usually start "gsk_". (This is the RTS Test Util)`;
  const response = ui.prompt(`Set ${serviceName} API Key`, promptMessage, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const apiKey = response.getResponseText().trim();
    if (apiKey === "") {
      PropertiesService.getUserProperties().deleteProperty(apiKeyProperty);
      ui.alert(`${serviceName} API Key Cleared`, `Property "${apiKeyProperty}" cleared.`);
    } else if (apiKey && apiKey.startsWith("gsk_") && apiKey.length > 40) {
      PropertiesService.getUserProperties().setProperty(apiKeyProperty, apiKey);
      ui.alert(`${serviceName} API Key Saved`, `Property "${apiKeyProperty}" updated.`);
    } else {
      ui.alert(`${serviceName} API Key NOT Saved`, `Invalid format. Check key.`, ui.ButtonSet.OK);
    }
  } else {
    ui.alert(`${serviceName} API Key Cancelled`, `No changes made.`, ui.ButtonSet.OK);
  }
}
