// File: RTS_TailoringLogic.gs
// Description: Contains functions responsible for constructing prompts and making calls
// to the LLM (via RTS_callGroq from RTS_GroqService.gs) for specific resume tailoring tasks.
// Relies on Global_Constants.gs for default model name and debug mode.

/**
 * Analyzes a job description using the configured LLM (Groq) to extract key information.
 *
 * @param {string} jobDescriptionText The full text of the job description.
 * @return {Object} A JavaScript object with extracted JD info, or an error object.
 */
function RTS_analyzeJobDescription(jobDescriptionText) {
  // GLOBAL_DEBUG_MODE, DEFAULT_GROQ_MODEL from Global_Constants.gs
  const DEBUG = GLOBAL_DEBUG_MODE;

  if (!jobDescriptionText || !jobDescriptionText.trim()) {
    Logger.log("[ERROR] RTS_TailoringLogic (analyzeJobDescription): Input job description text is empty.");
    return { error: "Input job description text is empty." };
  }
  const modelToUse = DEFAULT_GROQ_MODEL; // Using global default

  const prompt = `
    Analyze the following job description text.
    Extract the specified information and return it ONLY as a single, valid JSON object.
    Do not include any explanatory text, markdown, or anything outside the JSON structure.
    The JSON object should have keys: "jobTitle", "companyName", "location", "keyResponsibilities", "requiredTechnicalSkills", "requiredSoftSkills", "experienceLevel", "educationRequirements", "primaryKeywords", "companyCultureClues".
    For keys where information isn't present, use "" for string values or [] for array values. Maintain all keys.

    Job Description Text:
    ---
    ${jobDescriptionText}
    ---

    Output JSON:
  `;

  if(DEBUG) Logger.log(`  RTS_analyzeJobDescription: Calling Groq (model: ${modelToUse}) to analyze JD.`);
  // RTS_callGroq is from RTS_GroqService.gs
  const groqResponseString = RTS_callGroq(prompt, modelToUse, 0.1, 1500);

  if (groqResponseString && !groqResponseString.startsWith("ERROR:")) {
    try {
      let cleanedResponse = groqResponseString.trim();
      if (cleanedResponse.startsWith("```json")) cleanedResponse = cleanedResponse.substring(7).trim();
      else if (cleanedResponse.startsWith("```")) cleanedResponse = cleanedResponse.substring(3).trim();
      if (cleanedResponse.endsWith("```")) cleanedResponse = cleanedResponse.substring(0, cleanedResponse.length - 3).trim();
      
      const parsedJD = JSON.parse(cleanedResponse);
      if (parsedJD && typeof parsedJD.jobTitle !== 'undefined' && Array.isArray(parsedJD.keyResponsibilities)) {
        if(DEBUG) Logger.log("  RTS_analyzeJobDescription: Successfully parsed JD analysis from Groq.");
        return parsedJD;
      } else {
        Logger.log(`[ERROR] RTS_TailoringLogic (analyzeJobDescription): Parsed JSON structure mismatch. Raw: ${cleanedResponse}`);
        return { error: "Groq output structure mismatch for JD analysis.", rawOutput: cleanedResponse };
      }
    } catch (e) {
      Logger.log(`[ERROR] RTS_TailoringLogic (analyzeJobDescription): Failed to parse JSON from Groq. Error: ${e.toString()}. Raw: ${groqResponseString}`);
      return { error: "Failed to parse Groq JSON output for JD analysis.", rawOutput: groqResponseString };
    }
  } else {
    Logger.log(`[ERROR] RTS_TailoringLogic (analyzeJobDescription): Groq API call failed. Response: ${groqResponseString}`);
    return { error: "Groq API call failed for JD analysis.", rawOutput: groqResponseString };
  }
}

/**
 * Evaluates the relevance of a resume section/bullet point to the analyzed job description.
 *
 * @param {string} resumeSectionText The text from a resume section (e.g., a single bullet point).
 * @param {Object} jdAnalysisResults The structured analysis of the JD from RTS_analyzeJobDescription.
 * @return {Object} { relevanceScore: number, matchingKeywords: string[], justification: string }, or error object.
 */
function RTS_matchResumeSection(resumeSectionText, jdAnalysisResults) {
  // GLOBAL_DEBUG_MODE, DEFAULT_GROQ_MODEL from Global_Constants.gs
  const DEBUG = GLOBAL_DEBUG_MODE;

  if (!resumeSectionText?.trim()) return { error: "Input resume section text is empty." };
  if (!jdAnalysisResults?.primaryKeywords || !jdAnalysisResults.jobTitle) return { error: "Invalid JD analysis (missing keywords/jobTitle)." };
  
  const modelToUse = DEFAULT_GROQ_MODEL;
  const jdAnalysisContextString = JSON.stringify(jdAnalysisResults, null, 2);

  const prompt = `
    You are an expert resume analyst. Evaluate the relevance of the "Resume Section Text" 
    to the "Analyzed Job Description Data".
    Return ONLY a single, valid JSON: {"relevanceScore": number (0.0-1.0), "matchingKeywords": string[], "justification": string (1-2 sentences)}.
    If irrelevant, relevanceScore: 0.0.

    --- Analyzed Job Description Data ---
    ${jdAnalysisContextString}

    --- Resume Section Text to Evaluate ---
    ${resumeSectionText}

    --- Output JSON ---
  `;

  if(DEBUG) Logger.log(`  RTS_matchResumeSection: Calling Groq (model: ${modelToUse}) to evaluate relevance.`);
  // RTS_callGroq from RTS_GroqService.gs
  const groqResponseString = RTS_callGroq(prompt, modelToUse, 0.2, 512);

  if (groqResponseString && !groqResponseString.startsWith("ERROR:")) {
    try {
      let cleanedResponse = groqResponseString.trim();
      if (cleanedResponse.startsWith("```json")) cleanedResponse = cleanedResponse.substring(7).trim();
      else if (cleanedResponse.startsWith("```")) cleanedResponse = cleanedResponse.substring(3).trim();
      if (cleanedResponse.endsWith("```")) cleanedResponse = cleanedResponse.substring(0, cleanedResponse.length - 3).trim();
      
      const parsedMatch = JSON.parse(cleanedResponse);
      if (parsedMatch && typeof parsedMatch.relevanceScore === 'number' && Array.isArray(parsedMatch.matchingKeywords) && typeof parsedMatch.justification === 'string') {
        return parsedMatch;
      } else {
        Logger.log(`[ERROR] RTS_TailoringLogic (matchResumeSection): Parsed JSON structure mismatch. Raw: ${cleanedResponse}`);
        return { error: "Groq output structure mismatch for matching.", rawOutput: cleanedResponse };
      }
    } catch (e) {
      Logger.log(`[ERROR] RTS_TailoringLogic (matchResumeSection): Failed to parse JSON. Error: ${e}. Raw: ${groqResponseString}`);
      return { error: "Failed to parse Groq JSON for matching.", rawOutput: groqResponseString };
    }
  } else {
    Logger.log(`[ERROR] RTS_TailoringLogic (matchResumeSection): Groq API call failed. Response: ${groqResponseString}`);
    return { error: "Groq API call failed for matching.", rawOutput: groqResponseString };
  }
}

/**
 * Rewrites/tailors an original resume bullet point to align with the target job.
 *
 * @param {string} originalBullet The original bullet point text.
 * @param {Object} jdAnalysisResults The structured analysis of the JD.
 * @param {string} targetRoleTitle The specific job title being targeted.
 * @return {string} Tailored bullet, "Original bullet not suitable...", or "ERROR: ..." string.
 */
function RTS_tailorBulletPoint(originalBullet, jdAnalysisResults, targetRoleTitle) {
  // GLOBAL_DEBUG_MODE, DEFAULT_GROQ_MODEL from Global_Constants.gs
  const DEBUG = GLOBAL_DEBUG_MODE;

  if (!originalBullet?.trim()) return "ERROR: Original bullet empty.";
  if (!jdAnalysisResults?.primaryKeywords || !jdAnalysisResults.jobTitle) return "ERROR: Invalid JD analysis.";
  if (!targetRoleTitle?.trim()) return "ERROR: Target role empty.";

  const modelToUse = DEFAULT_GROQ_MODEL;
  const jdContextString = `Target Role: ${targetRoleTitle}\nKey Responsibilities (sample): ${(jdAnalysisResults.keyResponsibilities || []).slice(0, 3).join('; ')}\nPrimary Keywords from JD: ${(jdAnalysisResults.primaryKeywords || []).join(', ')}\nRequired Technical Skills: ${(jdAnalysisResults.requiredTechnicalSkills || []).join(', ')}`.trim();

  const prompt = `
    You are an expert resume writer. Rewrite the "Original Resume Bullet" for the "Target Role" using "Job Description Context".
    Guidelines:
    1. Start with a strong action verb.
    2. Quantify achievements (retain/enhance original metrics, do not invent if none).
    3. Naturally incorporate JD keywords/skills IF they genuinely fit. DO NOT FORCE.
    4. Maintain truthfulness. Do not fabricate.
    5. Concise, impactful, 1-2 lines.
    6. If original is already excellent, return it with minor enhancements or unchanged.
    7. If completely irrelevant & untailorable, return EXACTLY: "Original bullet not suitable for significant tailoring towards this role."

    --- Job Description Context ---
    ${jdContextString} 

    --- Original Resume Bullet to Tailor ---
    ${originalBullet}

    --- Rewritten (Tailored) Resume Bullet ---
    (Return ONLY rewritten bullet OR "not suitable" message. Optional JSON: {"rewritten_bullet": "text"} - if used, ONLY JSON.)
  `;

  if(DEBUG) Logger.log(`  RTS_tailorBulletPoint: Calling Groq (model: ${modelToUse}) to tailor bullet.`);
  // RTS_callGroq from RTS_GroqService.gs
  const groqResponseString = RTS_callGroq(prompt, modelToUse, 0.5, 256);

  if (groqResponseString && !groqResponseString.startsWith("ERROR:")) {
    let tailoredText = groqResponseString.trim();
    try {
      let potentialJson = tailoredText;
      if (potentialJson.startsWith("```json")) potentialJson = potentialJson.substring(7).trim();
      else if (potentialJson.startsWith("```")) potentialJson = potentialJson.substring(3).trim();
      if (potentialJson.endsWith("```")) potentialJson = potentialJson.substring(0, potentialJson.length - 3).trim();
      
      const parsedResponse = JSON.parse(potentialJson); // Will throw error if not JSON
      if (parsedResponse && parsedResponse.rewritten_bullet && typeof parsedResponse.rewritten_bullet === 'string') {
        if(DEBUG) Logger.log("    Tailored bullet extracted from Groq JSON response.");
        return parsedResponse.rewritten_bullet.trim();
      }
      // If not the specific JSON, or missing key, it was not intended as JSON by LLM here.
      // This 'else' is implicitly handled by the catch block or direct return of tailoredText.
    } catch (e) {
      // Not JSON, assume it's the direct string output (either tailored bullet or "not suitable" message)
      if(DEBUG) Logger.log("    Groq response for tailorBulletPoint not JSON, using as direct string.");
    }
    return tailoredText; // Return the cleaned original response
  } else {
    Logger.log(`[ERROR] RTS_TailoringLogic (tailorBulletPoint): Groq API call failed. Response: ${groqResponseString}`);
    return `ERROR: Groq API call failed for tailoring: "${originalBullet.substring(0,50)}..."`;
  }
}

/**
 * Generates a tailored professional summary using the LLM (Groq).
 *
 * @param {string} topMatchedExperiencesHighlights Key highlights from the resume.
 * @param {Object} jdAnalysisResults The structured analysis of the JD.
 * @param {string} candidateFullName The full name of the candidate.
 * @return {string} Tailored summary string, or "ERROR: ..." string on failure.
 */
function RTS_generateTailoredSummary(topMatchedExperiencesHighlights, jdAnalysisResults, candidateFullName) {
  // GLOBAL_DEBUG_MODE, DEFAULT_GROQ_MODEL from Global_Constants.gs
  const DEBUG = GLOBAL_DEBUG_MODE;

  if (!topMatchedExperiencesHighlights?.trim()) return "ERROR: Highlights empty.";
  if (!jdAnalysisResults?.jobTitle) return "ERROR: Invalid JD analysis (missing jobTitle).";
  if (!candidateFullName?.trim()) return "ERROR: Candidate name empty.";

  const modelToUse = DEFAULT_GROQ_MODEL;
  const jdContextString = `Target Role: ${jdAnalysisResults.jobTitle}${jdAnalysisResults.companyName ? ' at ' + jdAnalysisResults.companyName : ''}.\nKey Responsibilities (sample): ${(jdAnalysisResults.keyResponsibilities || []).slice(0, 2).join('; ')}\nPrimary Keywords: ${(jdAnalysisResults.primaryKeywords || []).join(', ')}\nRequired Skills: ${(jdAnalysisResults.requiredTechnicalSkills || []).join(', ')}\nExperience Level: ${jdAnalysisResults.experienceLevel || 'Not specified'}.`.trim();

  const prompt = `
    You are an expert resume writer. Craft a compelling professional summary for candidate ${candidateFullName}.
    The candidate is applying for "Target Role" (see "Job Description Context").
    Use "Candidate Highlights" for key skills/achievements.

    Instructions:
    1. Concise, impactful, professional summary: 2-4 sentences, max 80 words.
    2. Directly address "Target Role".
    3. Naturally integrate JD keywords/phrases from "Job Description Context".
    4. Leverage "Candidate Highlights".
    5. Confident, compelling tone.
    6. Avoid clichés (e.g., "results-oriented").
    7. STRICTLY AVOID first-person (I, me, my). Write in third person or impersonally.

    --- Job Description Context ---
    ${jdContextString}

    --- Candidate Highlights ---
    ${topMatchedExperiencesHighlights}

    --- Professional Summary ---
    (Return ONLY summary text. Optional JSON: {"professionalSummary": "text"} - if used, ONLY JSON.)
  `;

  if(DEBUG) Logger.log(`  RTS_generateTailoredSummary: Calling Groq (model: ${modelToUse}) for summary.`);
  // RTS_callGroq from RTS_GroqService.gs
  const groqResponseString = RTS_callGroq(prompt, modelToUse, 0.6, 200);

  if (groqResponseString && !groqResponseString.startsWith("ERROR:")) {
    let summaryText = groqResponseString.trim();
     try {
      let potentialJson = summaryText;
      if (potentialJson.startsWith("```json")) potentialJson = potentialJson.substring(7).trim();
      else if (potentialJson.startsWith("```")) potentialJson = potentialJson.substring(3).trim();
      if (potentialJson.endsWith("```")) potentialJson = potentialJson.substring(0, potentialJson.length - 3).trim();

      const parsedResponse = JSON.parse(potentialJson); // Will throw if not JSON
      if (parsedResponse && parsedResponse.professionalSummary && typeof parsedResponse.professionalSummary === 'string') {
        if(DEBUG) Logger.log("    Summary extracted from Groq JSON response.");
        return parsedResponse.professionalSummary.trim();
      }
    } catch (e) {
      if(DEBUG) Logger.log("    Groq response for summary not JSON, using as direct string.");
    }
    return summaryText; // Return the cleaned original response
  } else {
    Logger.log(`[ERROR] RTS_TailoringLogic (generateTailoredSummary): Groq API call failed. Response: ${groqResponseString}`);
    return `ERROR: Groq API call failed during summary generation.`;
  }
}
