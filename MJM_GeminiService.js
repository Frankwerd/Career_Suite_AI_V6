// File: MJM_GeminiService.gs (or MRM_GeminiService.gs)
// Description: Handles all interactions with the Google Gemini API for
// AI-powered parsing of email content specifically for MJM module tasks
// (extracting job application details and job leads).
// Relies on constants from Global_Constants.gs (for API key prop, default model, shared text)
// and MJM_Config.gs (for MJM-specific status examples in prompts).

/**
 * Calls the Gemini API to extract company name, job title, and application status
 * from an email subject and body, specifically for MJM Application Tracker.
 *
 * @param {string} emailSubject The subject of the email.
 * @param {string} emailBody The plain text body of the email.
 * @param {string} apiKey The Gemini API key. (Passed directly, usually fetched by caller once per run).
 * @return {Object|null} An object like { company: string, title: string, status: string }
 *                       or null/MANUAL_REVIEW_NEEDED_TEXT as values on failure.
 */
function MJM_callGemini_forApplicationDetails(emailSubject, emailBody, apiKey) {
  // API Key check is usually done by the calling function (e.g., MJM_processJobApplicationEmails) once.
  // If apiKey is not provided to this function, it will fail at UrlFetchApp.
  // GLOBAL_DEBUG_MODE, MANUAL_REVIEW_NEEDED_TEXT from Global_Constants.gs
  // MJM_APP_DEFAULT_STATUS, MJM_APP_REJECTED_STATUS etc. from MJM_Config.gs for prompt examples.
  const DEBUG = GLOBAL_DEBUG_MODE;

  if ((!emailSubject || emailSubject.trim() === "") && (!emailBody || emailBody.trim() === "")) {
    Logger.log("[WARN] MJM_GeminiService (AppDetails): Both email subject and body are empty. Skipping Gemini call.");
    return { company: MANUAL_REVIEW_NEEDED_TEXT, title: MANUAL_REVIEW_NEEDED_TEXT, status: MANUAL_REVIEW_NEEDED_TEXT };
  }

  // DEFAULT_GEMINI_MODEL from Global_Constants.gs
  const modelToUse = DEFAULT_GEMINI_MODEL;
  const API_ENDPOINT = `https://generativelanguage.googleapis.com/v1beta/models/${modelToUse}:generateContent?key=${apiKey}`;
  if(DEBUG) Logger.log(`[DEBUG] MJM_GeminiService (AppDetails): Using API Endpoint: ${API_ENDPOINT.split('key=')[0] + "key=..."}`);

  const bodySnippet = emailBody ? emailBody.substring(0, 12000) : ""; // Max 12k chars for body snippet

  const prompt = `
    Analyze the provided email Subject and Body for a job application tracking system.
    Extract: "company_name", "job_title", and "status".
    Return ONLY a single, valid JSON object: {"company_name": "...", "job_title": "...", "status": "..."}. No markdown.

    **RELEVANCE CHECK (PRIORITY 1):**
    - If the email IS NOT DIRECTLY for a job application submitted by the recipient (e.g., general newsletters, marketing, job alerts not tied to a submission, sales pitches, spam), set ALL three fields to "${MANUAL_REVIEW_NEEDED_TEXT}". Output: {"company_name": "${MANUAL_REVIEW_NEEDED_TEXT}","job_title": "${MANUAL_REVIEW_NEEDED_TEXT}","status": "${MANUAL_REVIEW_NEEDED_TEXT}"}

    **IF APPLICATION-RELATED (PRIORITY 2):**
    1.  "company_name": HIRING COMPANY name. Not the ATS (Greenhouse, Lever), not the job board (LinkedIn, Indeed) unless they are the direct hirer. If unclear, use "${MANUAL_REVIEW_NEEDED_TEXT}".
    2.  "job_title": SPECIFIC job title THE USER APPLIED FOR, as stated in THIS email. If not restated in *this specific email body/subject for this event* (especially common for initial ATS "Application Received" confirmations from Greenhouse/Lever), use "${MANUAL_REVIEW_NEEDED_TEXT}".
    3.  "status": Current application status from THIS email. CHOOSE EXACTLY from this list:
        *   "${MJM_APP_DEFAULT_STATUS}" (Application submitted/received)
        *   "${MJM_APP_REJECTED_STATUS}" (Not moving forward, rejected)
        *   "${MJM_APP_OFFER_STATUS}" (Offer of employment)
        *   "${MJM_APP_INTERVIEW_STATUS}" (Interview invitation/scheduled)
        *   "${MJM_APP_ASSESSMENT_STATUS}" (Assessment, coding challenge, test)
        *   "${MJM_APP_VIEWED_STATUS}" (Application viewed by recruiter/company)
        *   "Update/Other" (General updates, "still reviewing", unclear status)
        If truly ambiguous but application-related, use "${MANUAL_REVIEW_NEEDED_TEXT}" for status as a last resort.

    --- EXAMPLES (using your system's status values) ---
    Subject: Your application was sent to MycoWorks
    Body: LinkedIn. Your application was sent to MycoWorks. Data Architect.
    Output: {"company_name": "MycoWorks","job_title": "Data Architect","status": "${MJM_APP_DEFAULT_STATUS}"}

    Subject: Update on your application for Product Manager at MegaEnterprises
    Body: From: no-reply@greenhouse.io. ...we have decided to move forward with other candidates...
    Output: {"company_name": "MegaEnterprises","job_title": "Product Manager","status": "${MJM_APP_REJECTED_STATUS}"}

    Subject: Thank you for applying to Handshake! (Application received, no title repeated in body)
    Body: no-reply@greenhouse.io. Hi Francis, Thank you for your interest in Handshake! We have received your application...
    Output: {"company_name": "Handshake","job_title": "${MANUAL_REVIEW_NEEDED_TEXT}","status": "${MJM_APP_DEFAULT_STATUS}"}

    Subject: Join our webinar on Future Tech! (Unrelated)
    Output: {"company_name": "${MANUAL_REVIEW_NEEDED_TEXT}","job_title": "${MANUAL_REVIEW_NEEDED_TEXT}","status": "${MANUAL_REVIEW_NEEDED_TEXT}"}
    --- END EXAMPLES ---

    --- EMAIL TO PROCESS ---
    Subject: ${emailSubject}
    Body:
    ${bodySnippet}
    --- END OF EMAIL TO PROCESS ---
    Output JSON:
  `;

  const payload = {
    "contents": [{"parts": [{"text": prompt}]}],
    "generationConfig": { "temperature": 0.2, "maxOutputTokens": 512, "topP": 0.95, "topK": 40 },
    "safetySettings": [ /* ... standard safety settings ... */
      { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" }
    ]
  };
  const options = {'method':'post', 'contentType':'application/json', 'payload':JSON.stringify(payload), 'muteHttpExceptions':true};

  if(DEBUG) Logger.log(`[DEBUG] MJM_GeminiService (AppDetails): Calling API. Prompt len: ${prompt.length}`);
  let response, attempt = 0, maxAttempts = 2;
  let result = { company: MANUAL_REVIEW_NEEDED_TEXT, title: MANUAL_REVIEW_NEEDED_TEXT, status: MANUAL_REVIEW_NEEDED_TEXT }; // Default to this

  while(attempt < maxAttempts){
    attempt++;
    try {
      response = UrlFetchApp.fetch(API_ENDPOINT, options);
      const responseCode = response.getResponseCode(); const responseBody = response.getContentText();
      if(DEBUG) Logger.log(`  (AppDetails Attempt ${attempt}) RC: ${responseCode}. Body(start): ${responseBody.substring(0,150)}`);

      if (responseCode === 200) {
        const jsonResponse = JSON.parse(responseBody);
        if (jsonResponse.candidates && jsonResponse.candidates[0]?.content?.parts?.[0]?.text) {
          let extractedJsonString = jsonResponse.candidates[0].content.parts[0].text.trim();
          if (extractedJsonString.startsWith("```json")) extractedJsonString = extractedJsonString.substring(7).trim();
          if (extractedJsonString.endsWith("```")) extractedJsonString = extractedJsonString.substring(0, extractedJsonString.length - 3).trim();
          
          try {
            const extractedData = JSON.parse(extractedJsonString);
            if (typeof extractedData.company_name !== 'undefined' && typeof extractedData.job_title !== 'undefined' && typeof extractedData.status !== 'undefined') {
              Logger.log(`[INFO] MJM_GeminiService (AppDetails): Success. C:"${extractedData.company_name}", T:"${extractedData.job_title}", S:"${extractedData.status}"`);
              result = { // Use MANUAL_REVIEW_NEEDED_TEXT as fallback if a field is empty but key exists
                  company: extractedData.company_name || MANUAL_REVIEW_NEEDED_TEXT,
                  title: extractedData.job_title || MANUAL_REVIEW_NEEDED_TEXT,
                  status: extractedData.status || MANUAL_REVIEW_NEEDED_TEXT
              };
              return result; // Success, exit loop
            } else { Logger.log(`[WARN] MJM_GeminiService (AppDetails): JSON missing expected fields. Output: ${extractedJsonString}`); }
          } catch (e) { Logger.log(`[ERROR] MJM_GeminiService (AppDetails): Error parsing JSON string from Gemini: ${e}\nString: >>>${extractedJsonString}<<<`); }
        } else if (jsonResponse.promptFeedback?.blockReason) {
          Logger.log(`[ERROR] MJM_GeminiService (AppDetails): Prompt blocked. Reason: ${jsonResponse.promptFeedback.blockReason}`);
          result.status = `Blocked: ${jsonResponse.promptFeedback.blockReason}`; return result; // Exit loop with block reason
        } else { Logger.log(`[ERROR] MJM_GeminiService (AppDetails): API response unexpected (no candidates/text). Body: ${responseBody.substring(0,500)}`);}
      } else if (responseCode === 429 && attempt < maxAttempts) { // Rate limit
        Logger.log(`[WARN] MJM_GeminiService (AppDetails): Rate limit (429). Attempt ${attempt}/${maxAttempts}. Waiting...`);
        Utilities.sleep(5000 + Math.floor(Math.random() * 5000)); continue;
      } else { // Other API errors
        Logger.log(`[ERROR] MJM_GeminiService (AppDetails): API error. Code: ${responseCode}. Body: ${responseBody.substring(0,500)}`);
        // If model not found, that's critical and likely won't resolve with retry.
        if (responseCode === 404 && responseBody.includes("is not found for API version")) {
             Logger.log(`[FATAL] MJM_GeminiService (AppDetails): MODEL "${modelToUse}" NOT FOUND. Check model name.`);
        }
        return result; // Return default on persistent error
      }
    } catch (e) { // Network errors etc.
      Logger.log(`[ERROR] MJM_GeminiService (AppDetails): Exception during API call (Attempt ${attempt}): ${e.toString()}`);
      if (attempt < maxAttempts) { Utilities.sleep(3000); continue; }
    }
    return result; // Return default if all attempts failed for a reason not exiting the loop earlier
  }
  Logger.log(`[ERROR] MJM_GeminiService (AppDetails): Failed to get valid response after ${maxAttempts} attempts.`);
  return result; // Final fallback
}


/**
 * Calls the Gemini API to extract multiple job leads from an email body.
 * This function is specifically for the MJM Job Leads Tracker module.
 * @param {string} emailBody The plain text body of the email.
 * @param {string} apiKey The Gemini API key. (Passed directly by caller).
 * @return {{success: boolean, data: Object|null, error: string|null}} Result object.
 */
function MJM_callGemini_forJobLeads(emailBody, apiKey) {
  // API Key check by caller. GLOBAL_DEBUG_MODE, DEFAULT_GEMINI_MODEL from Global_Constants.gs
  const DEBUG = GLOBAL_DEBUG_MODE;

  if (typeof emailBody !== 'string' || emailBody.trim() === "") {
      Logger.log(`[WARN] MJM_GeminiService (JobLeads): emailBody is not a string or is empty.`);
      return { success: false, data: null, error: "Email body empty or invalid for Job Leads." };
  }

  const modelToUse = DEFAULT_GEMINI_MODEL;
  const API_ENDPOINT = `https://generativelanguage.googleapis.com/v1beta/models/${modelToUse}:generateContent?key=${apiKey}`;

  // --- MOCK RESPONSE LOGIC (FOR OFFLINE TESTING) ---
  // REMOVE OR SECURELY DISABLE FOR PRODUCTION if API key might be placeholder in some dev envs
  /* // << UNCOMMENT FOR MOCKING IF NEEDED, BUT ENSURE API KEY IS OTHERWISE VALID
  if (!apiKey || apiKey.startsWith("AIzaSy_DEV_PLACEHOLDER") ) { // Check for a specific dev placeholder if you use one
      Logger.log("[WARN] MJM_GeminiService (JobLeads): Using MOCK response due to placeholder API Key.");
      // Construct and return a mock response similar to what the actual API would send
      // This mock would be a JS object, not just the text part
      if (emailBody.toLowerCase().includes("multiple job listings inside") || emailBody.toLowerCase().includes("software engineer at google")) {
          return { success: true, data: { candidates: [{ content: { parts: [{ text: JSON.stringify([
              { "jobTitle": "Software Engineer (Mock)", "company": "Tech Alpha (Mock)", "location": "Remote", "linkToJobPosting": "https://example.com/job/alpha" },
              { "jobTitle": "Product Manager (Mock)", "company": "Innovate Beta (Mock)", "location": "New York, NY", "linkToJobPosting": "https://example.com/job/beta" }
          ])}]}}]}, error: null };
      } // Default mock for other cases:
      return { success: true, data: { candidates: [{ content: { parts: [{ text: JSON.stringify([{ "jobTitle": "N/A (Mock Single)", "company": "Some Corp (Mock)", "location": "Remote", "linkToJobPosting": "N/A" }])}]}}]}, error: null };
  }
  */ // << END MOCKING BLOCK

  const promptText = `
    From the following email content, identify each distinct job posting.
    For EACH job posting found, extract: "jobTitle", "company", "location", "linkToJobPosting".
    If a field is not found, use "N/A".
    Format your ENTIRE response as a single, valid JSON array of objects. No markdown.
    Example: [{"jobTitle": "SWE", "company": "Tech Corp", "location": "Remote", "linkToJobPosting": "url"}]
    If no jobs found, return an empty JSON array: [].

    Email Content:
    ---
    ${emailBody.substring(0, 28000)} 
    ---
    JSON Array Output:
  `;

  const payload = {
      contents: [{ parts: [{ "text": promptText }] }],
      generationConfig: { "temperature": 0.2, "maxOutputTokens": 8192 },
      safetySettings: [ /* ... standard safety settings ... */
        { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
        { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
        { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
        { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" }
      ]
  };
  const options = {'method': "post", 'contentType': "application/json", 'payload': JSON.stringify(payload), 'muteHttpExceptions': true};

  if(DEBUG) Logger.log(`[DEBUG] MJM_GeminiService (JobLeads): Calling API. Prompt len: ${promptText.length}`);
  let attempt = 0, maxAttempts = 2;

  while (attempt < maxAttempts) {
      attempt++;
      try {
          const response = UrlFetchApp.fetch(API_ENDPOINT, options);
          const responseCode = response.getResponseCode();
          const responseBody = response.getContentText();

          if (responseCode === 200) {
              if(DEBUG) Logger.log(`  (JobLeads Attempt ${attempt}) RC: ${responseCode}. Body(start): ${responseBody.substring(0,150)}`);
              try {
                  // Gemini should return the candidates array directly in the response body for :generateContent
                  // The actual parseable JSON might be nested within candidates[0].content.parts[0].text.
                  // The *entire responseBody* IS the top-level JSON from Gemini.
                  const fullJsonResponse = JSON.parse(responseBody);
                  // Now, fullJsonResponse should have the "candidates" array as per Gemini API spec.
                  // The actual list of jobs (the array of objects we asked for) is what `MJM_parseGeminiResponse_forJobLeads` expects.
                  return { success: true, data: fullJsonResponse, error: null }; // Pass the whole response
              } catch (jsonParseError) {
                  Logger.log(`[ERROR] MJM_GeminiService (JobLeads): Failed to parse Gemini JSON response: ${jsonParseError}.\nRaw body: ${responseBody.substring(0, 1000)}`);
                  return { success: false, data: null, error: `Failed to parse API JSON: ${jsonParseError}. Response was: ${responseBody.substring(0, 200)}` };
              }
          } else if (responseCode === 429 && attempt < maxAttempts) {
              Logger.log(`[WARN] MJM_GeminiService (JobLeads): Rate limit (429) on attempt ${attempt}. Waiting...`);
              Utilities.sleep(3000 + Math.random() * 2000); continue;
          } else {
              Logger.log(`[ERROR] MJM_GeminiService (JobLeads): API error. Code: ${responseCode}. Body: ${responseBody.substring(0,500)}`);
              if (responseCode === 400 && responseBody.toLowerCase().includes("api key not valid")) Logger.log(`[FATAL] Gemini API Key Invalid. Check property: ${SHARED_GEMINI_API_KEY_PROPERTY}`);
              else if (responseCode === 404 && responseBody.toLowerCase().includes("is not found for api version")) Logger.log(`[FATAL] MODEL "${modelToUse}" NOT FOUND.`);
              return { success: false, data: null, error: `API Error ${responseCode}: ${responseBody.substring(0, 200)}` };
          }
      } catch (e) {
          Logger.log(`[ERROR] MJM_GeminiService (JobLeads): Exception during API call (Attempt ${attempt}): ${e.toString()}`);
          if (attempt < maxAttempts) { Utilities.sleep(2000); continue; }
          return { success: false, data: null, error: `UrlFetchApp Error after ${maxAttempts} attempts: ${e.toString()}` };
      }
  }
  return { success: false, data: null, error: `Exceeded max retry attempts for Gemini JobLeads API call.` };
}

/**
 * Parses the raw JSON data object from the Gemini API response (for job leads)
 * into an array of standardized job lead objects.
 * @param {Object} apiResponseData The full 'data' object from a successful MJM_callGemini_forJobLeads response.
 * @return {Array<Object>} An array of job lead objects, or an empty array if parsing fails or no jobs.
 */
function MJM_parseGeminiResponse_forJobLeads(apiResponseData) {
  // GLOBAL_DEBUG_MODE from Global_Constants.gs
  const DEBUG = GLOBAL_DEBUG_MODE;
  if(DEBUG) Logger.log(`[DEBUG] MJM_GeminiService (ParseLeads): Input data for parsing (first 300 chars): ${JSON.stringify(apiResponseData).substring(0,300)}...`);
  let jobListings = [];
  try {
      let jsonStringFromLLM = "";
      // Extract the text part which contains the JSON array string
      if (apiResponseData && apiResponseData.candidates && apiResponseData.candidates.length > 0 &&
          apiResponseData.candidates[0].content && apiResponseData.candidates[0].content.parts &&
          apiResponseData.candidates[0].content.parts.length > 0 &&
          typeof apiResponseData.candidates[0].content.parts[0].text === 'string') {
          jsonStringFromLLM = apiResponseData.candidates[0].content.parts[0].text.trim();
      } else {
          Logger.log(`[WARN] MJM_GeminiService (ParseLeads): No parsable content string in Gemini response structure for leads.`);
          if (apiResponseData && apiResponseData.promptFeedback?.blockReason) {
              Logger.log(`  Block Reason: ${apiResponseData.promptFeedback.blockReason}. Ratings: ${JSON.stringify(apiResponseData.promptFeedback.safetyRatings)}`);
          }
          return []; // Return empty array
      }

      // Clean potential markdown
      if (jsonStringFromLLM.startsWith("```json")) jsonStringFromLLM = jsonStringFromLLM.substring(7).trim();
      else if (jsonStringFromLLM.startsWith("```")) jsonStringFromLLM = jsonStringFromLLM.substring(3).trim();
      if (jsonStringFromLLM.endsWith("```")) jsonStringFromLLM = jsonStringFromLLM.substring(0, jsonStringFromLLM.length - 3).trim();
      if(DEBUG) Logger.log(`  Cleaned JSON String for Leads (first 500): ${jsonStringFromLLM.substring(0, 500)}...`);

      const parsedData = JSON.parse(jsonStringFromLLM);
      if (Array.isArray(parsedData)) {
          parsedData.forEach(job => {
              if (job && typeof job === 'object' && (job.jobTitle || job.company)) { // Basic validation
                  jobListings.push({
                      jobTitle: job.jobTitle || "N/A", company: job.company || "N/A",
                      location: job.location || "N/A", linkToJobPosting: job.linkToJobPosting || "N/A"
                  });
              } else { Logger.log(`[WARN] MJM_GeminiService (ParseLeads): Skipped invalid item in job listings: ${JSON.stringify(job)}`); }
          });
      } else if (typeof parsedData === 'object' && parsedData !== null && (parsedData.jobTitle || parsedData.company)) { // Handle single object case
          jobListings.push({
              jobTitle: parsedData.jobTitle || "N/A", company: parsedData.company || "N/A",
              location: parsedData.location || "N/A", linkToJobPosting: parsedData.linkToJobPosting || "N/A"
          });
          Logger.log(`[WARN] MJM_GeminiService (ParseLeads): LLM returned a single object, parsed as one job.`);
      } else { Logger.log(`[WARN] MJM_GeminiService (ParseLeads): LLM output not a JSON array or parsable single job object. Output (first 200): ${jsonStringFromLLM.substring(0, 200)}`); }
  } catch (e) {
      Logger.log(`[ERROR] MJM_GeminiService (ParseLeads): Error parsing leads response: ${e.toString()}.\nAPI Data: ${JSON.stringify(apiResponseData).substring(0,300)}`);
  }
  if(DEBUG) Logger.log(`[DEBUG] MJM_GeminiService (ParseLeads): Successfully parsed ${jobListings.length} job objects.`);
  return jobListings;
}
