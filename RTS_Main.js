// File: RTS_Main.gs
// Description: Contains the main orchestration logic for the AI Resume Tailoring staged processing
// (Stage 1: Analyze & Score, Stage 2: Tailor Selected, Stage 3: Assemble & Generate).
// Relies on global constants from Global_Constants.gs and functions from other RTS_ prefixed files.

// --- STAGE 1: ANALYZE JD & SCORE MASTER PROFILE BULLETS (WITH QoL ADDITIONS) ---
function RTS_runStage1_AnalyzeAndScore(jobDescriptionText, spreadsheetId) {
  const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : true);
  Logger.log(`--- RTS Stage 1: Starting JD Analysis & Scoring (DEBUG: ${DEBUG}) ---`);

  if (!jobDescriptionText?.trim()) { 
    Logger.log("[ERROR] RTS Stage 1: Job Description text is empty or null."); 
    return { success: false, message: "Job Description text cannot be empty." }; 
  }
  if (!spreadsheetId) { 
    Logger.log("[ERROR] RTS Stage 1: Spreadsheet ID not provided."); 
    return { success: false, message: "Spreadsheet ID is required." }; 
  }

  // Constants from Global_Constants.gs
  const profileDataSheetNameConst = (typeof PROFILE_DATA_SHEET_NAME !== 'undefined' ? PROFILE_DATA_SHEET_NAME : "MasterProfile_Fallback");
  const jdAnalysisSheetNameConst = (typeof JD_ANALYSIS_SHEET_NAME !== 'undefined' ? JD_ANALYSIS_SHEET_NAME : "JDAnalysisData_Fallback");
  const scoringSheetNameConst = (typeof BULLET_SCORING_RESULTS_SHEET_NAME !== 'undefined' ? BULLET_SCORING_RESULTS_SHEET_NAME : "BulletScoringResults_Fallback");
  const apiDelayStage1 = (typeof RTS_API_CALL_DELAY_STAGE1_MS !== 'undefined' ? RTS_API_CALL_DELAY_STAGE1_MS : 1000);
  const userSelectYesConst = (typeof USER_SELECT_YES_VALUE !== 'undefined' ? USER_SELECT_YES_VALUE : "YES");

  const masterProfileObject = RTS_getMasterProfileData(spreadsheetId, profileDataSheetNameConst); // From RTS_MasterResumeData.gs
  if (!masterProfileObject?.personalInfo) {
    Logger.log(`[ERROR] RTS Stage 1: Failed to load master profile from SSID: ${spreadsheetId}, Sheet: "${profileDataSheetNameConst}".`);
    return { success: false, message: `Could not load master profile data. Ensure "${profileDataSheetNameConst}" exists and is structured.` };
  }
  if(DEBUG) Logger.log("  RTS Stage 1: Master Profile loaded successfully.");

  Logger.log("  RTS Stage 1: Analyzing Job Description...");
  Utilities.sleep(apiDelayStage1);
  const jdAnalysis = RTS_analyzeJobDescription(jobDescriptionText); // From RTS_TailoringLogic.gs
  if (!jdAnalysis || jdAnalysis.error) {
    Logger.log(`[ERROR] RTS Stage 1: Job Description analysis failed. Details: ${JSON.stringify(jdAnalysis)}`);
    return { success: false, message: "JD analysis failed.", details: jdAnalysis, jdAnalysis: null };
  }
  if(DEBUG) Logger.log(`  RTS Stage 1: JD Analyzed. Extracted Job Title: ${jdAnalysis.jobTitle || 'N/A (check JD content and parser)'}`);

  let currentSpreadsheet;
  let scoringSheet; 
  const allScoredDataRowsForSheet = []; 
  let itemsScoredCounter = 0;
  const scoringSheetHeaders = ["UniqueID", "Section", "ItemIdentifier", "OriginalBulletText", "RelevanceScore", "MatchingKeywords", "Justification", "SelectToTailor(Manual)", "TailoredBulletText(Stage2)"];

  // ---- Start of MAIN TRY block for sheet operations and scoring ----
  try { 
    currentSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    if (!currentSpreadsheet) throw new Error(`Failed to open spreadsheet ID: ${spreadsheetId}`);

    // JDAnalysisData sheet setup
    let jdSheet = currentSpreadsheet.getSheetByName(jdAnalysisSheetNameConst);
    if (!jdSheet) jdSheet = currentSpreadsheet.insertSheet(jdAnalysisSheetNameConst);
    jdSheet.clearContents(); 
    jdSheet.getRange(1,1).setValue("JD Analysis JSON:").setFontWeight("bold");
    const jdAnalysisString = JSON.stringify(jdAnalysis,null,2); 
    jdSheet.getRange(2,1).setValue(jdAnalysisString.length>49000 ? jdAnalysisString.substring(0,49000)+"...(TRUNCATED)" : jdAnalysisString);
    try { jdSheet.autoResizeColumn(1); } catch (eResizeJD) { if(DEBUG) Logger.log(`    [WARN] Could not auto-resize JDAnalysisSheet column: ${eResizeJD.message}`);}
    if(DEBUG) Logger.log(`  RTS Stage 1: JD Analysis data stored in sheet "${jdAnalysisSheetNameConst}".`);

    // BulletScoringResults sheet setup
    scoringSheet = currentSpreadsheet.getSheetByName(scoringSheetNameConst);
    if (!scoringSheet) scoringSheet = currentSpreadsheet.insertSheet(scoringSheetNameConst);
    scoringSheet.clearContents(); 
    SpreadsheetApp.flush(); 

    const headerRowRange = scoringSheet.getRange(1, 1, 1, scoringSheetHeaders.length);
    headerRowRange.setValues([scoringSheetHeaders]).setFontWeight("bold").setBackground("#EFEFEF")
                   .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    if (scoringSheet.getFrozenRows() !== 1) try {scoringSheet.setFrozenRows(1);}catch(e){/*ignore*/}

    // Data Validation (Dropdown) for "SelectToTailor(Manual)"
    const selectToTailorColName = "SelectToTailor(Manual)";
    const selectToTailorColIdx = scoringSheetHeaders.indexOf(selectToTailorColName) + 1;
    if (selectToTailorColIdx > 0) {
      let maxValRowsCurrent = scoringSheet.getMaxRows();
      if (maxValRowsCurrent < 200) { // Ensure enough rows for validation if sheet is small
          scoringSheet.insertRowsAfter(maxValRowsCurrent, 200 - maxValRowsCurrent);
          maxValRowsCurrent = scoringSheet.getMaxRows();
      }
      if (maxValRowsCurrent > 1) { 
        const validationDataRange = scoringSheet.getRange(2, selectToTailorColIdx, maxValRowsCurrent - 1, 1);
        const validationRule = SpreadsheetApp.newDataValidation().requireValueInList([userSelectYesConst,"NO",""],true).setAllowInvalid(false).setHelpText(`Select an option to include for tailoring.`).build();
        validationDataRange.setDataValidation(validationRule);
        if(DEBUG) Logger.log(`    RTS Stage 1: YES/NO dropdown validation applied to "${selectToTailorColName}" column in "${scoringSheetNameConst}".`);
      }
    } else if(DEBUG) { Logger.log(`    [WARN] RTS Stage 1: Column "${selectToTailorColName}" not found in headers for dropdown.`); }
    if(DEBUG) Logger.log(`  RTS Stage 1: Sheet "${scoringSheetNameConst}" prepared with headers and dropdown validation.`);

    // Helper function for scoring individual bullet points/items
    const scoreSingleProfileItem = (idBase, sectionTitleStr, itemIdentifierStr, textToScore) => {
        const trimmedText = String(textToScore||"").trim();
        if (!trimmedText) return; 
        itemsScoredCounter++; // Use a consistent counter for all scored items
        const fullUniqueId = `${idBase}_Itm${itemsScoredCounter}`; // More unique ID structure
        
        if(DEBUG) Logger.log(`      Scoring (${fullUniqueId}): Section="${sectionTitleStr}", ItemID="${itemIdentifierStr}", Text="${trimmedText.substring(0,60)}..."`);
        Utilities.sleep(apiDelayStage1);
        const matchResult = RTS_matchResumeSection(trimmedText, jdAnalysis); // From RTS_TailoringLogic.gs
        
        let scoreValue = 0.0; 
        let matchingKeywordsArray = []; 
        let justificationText = "LLM match error or no score returned by AI.";
        
        if(matchResult && !matchResult.error && typeof matchResult.relevanceScore === 'number'){ 
            scoreValue = matchResult.relevanceScore; 
            matchingKeywordsArray = matchResult.matchingKeywords || []; 
            justificationText = matchResult.justification || "No justification provided by AI.";
        } else if(matchResult?.error){ 
            justificationText = `LLM Match Error: ${matchResult.error} (Raw LLM Output: ${matchResult.rawOutput||""})`;
        }
        allScoredDataRowsForSheet.push([fullUniqueId, sectionTitleStr, itemIdentifierStr, trimmedText, scoreValue.toFixed(3), matchingKeywordsArray.join("; "), justificationText, "", ""]);
    };

    // --- SCORING LOOPS FOR EACH PROFILE SECTION ---
    // This assumes masterProfileObject.sections contains the parsed data.

    // Score EXPERIENCE section
    if(DEBUG) Logger.log("    RTS Stage 1: Processing EXPERIENCE items for scoring...");
    const experienceProfileSection = masterProfileObject.sections.find(s => s.title === "EXPERIENCE");
    if (experienceProfileSection?.items && Array.isArray(experienceProfileSection.items)) {
      experienceProfileSection.items.forEach((jobEntry, itemIndex) => {
        const itemIdentifierForSheet = jobEntry.company || `ExperienceEntry_${itemIndex}`;
        (jobEntry.responsibilities || []).forEach((bulletTextValue) => {
          scoreSingleProfileItem(`EXP_${itemIndex}`, "EXPERIENCE", itemIdentifierForSheet, bulletTextValue);
        });
      });
    } else if(DEBUG) Logger.log("      No 'EXPERIENCE' items found in profile data to score.");
    
    // Score PROJECTS section
    if(DEBUG) Logger.log("    RTS Stage 1: Processing PROJECTS items for scoring...");
    const projectsProfileSection = masterProfileObject.sections.find(s => s.title === "PROJECTS");
    if (projectsProfileSection?.subsections && Array.isArray(projectsProfileSection.subsections)) {
      projectsProfileSection.subsections.forEach((projectSubSection, subSectionIndex) => {
        (projectSubSection.items || []).forEach((projectEntry, itemIndex) => {
          const itemIdentifierForSheet = projectEntry.projectName || `${projectSubSection.name}_Project_${itemIndex}`;
          // Score descriptionBullets for the project
          (projectEntry.descriptionBullets || []).forEach((bulletTextValue) => {
            scoreSingleProfileItem(`PROJ_S${subSectionIndex}_Item${itemIndex}_Desc`, "PROJECTS", itemIdentifierForSheet, bulletTextValue);
          });
          // If projects also have a flat 'responsibilities' array directly under the item
          (projectEntry.responsibilities || []).forEach((bulletTextValue) => {
            scoreSingleProfileItem(`PROJ_S${subSectionIndex}_Item${itemIndex}_Resp`, "PROJECTS", itemIdentifierForSheet, bulletTextValue);
          });
        });
      });
    } else if(DEBUG) Logger.log("      No 'PROJECTS' subsections/items found in profile data to score.");
    
    // Score LEADERSHIP & UNIVERSITY INVOLVEMENT section
    if(DEBUG) Logger.log("    RTS Stage 1: Processing LEADERSHIP & UNIVERSITY INVOLVEMENT items for scoring...");
    const leadershipProfileSection = masterProfileObject.sections.find(s => s.title === "LEADERSHIP & UNIVERSITY INVOLVEMENT");
    if (leadershipProfileSection?.items && Array.isArray(leadershipProfileSection.items)) {
      leadershipProfileSection.items.forEach((leadershipEntry, itemIndex) => {
        const itemIdentifierForSheet = leadershipEntry.organization || leadershipEntry.role || `LeadershipEntry_${itemIndex}`;
        (leadershipEntry.responsibilities || []).forEach((bulletTextValue) => {
          scoreSingleProfileItem(`LEAD_${itemIndex}`, "LEADERSHIP", itemIdentifierForSheet, bulletTextValue);
        });
      });
    } else if(DEBUG) Logger.log("      No 'LEADERSHIP & UNIVERSITY INVOLVEMENT' items found in profile data.");

    // Score TECHNICAL SKILLS & CERTIFICATES section
    if(DEBUG) Logger.log("    RTS Stage 1: Processing TECHNICAL SKILLS & CERTIFICATES items for scoring...");
    const techSkillsProfileSection = masterProfileObject.sections.find(s => s.title === "TECHNICAL SKILLS & CERTIFICATES");
    if (techSkillsProfileSection?.subsections && Array.isArray(techSkillsProfileSection.subsections)) {
      techSkillsProfileSection.subsections.forEach((techSubSection, subSectionIndex) => {
        const categoryNameForIdentifier = techSubSection.name || `TechCategory_${subSectionIndex}`;
        (techSubSection.items || []).forEach((techItem, itemIndex) => {
          let textToScoreForSkillCert = ""; 
          let skillCertIdentifier = ""; 
          if (techItem.skill && String(techItem.skill).trim()) { 
            skillCertIdentifier = String(techItem.skill).trim();
            textToScoreForSkillCert = skillCertIdentifier;
            if (techItem.details && String(techItem.details).trim()) textToScoreForSkillCert += ` (${String(techItem.details).trim()})`;
          } else if (techItem.name && String(techItem.name).trim()) { // Likely a certificate
            skillCertIdentifier = String(techItem.name).trim();
            textToScoreForSkillCert = skillCertIdentifier;
            if (techItem.issuer && String(techItem.issuer).trim()) textToScoreForSkillCert += ` (Issuer: ${String(techItem.issuer).trim()})`;
            if (techItem.issueDate && String(techItem.issueDate).trim()) textToScoreForSkillCert += `, Issued: ${String(techItem.issueDate).trim()}`;
            if (techItem.details && String(techItem.details).trim()) textToScoreForSkillCert += ` - ${String(techItem.details).trim()}`; // Append details if exists
          }
          if(textToScoreForSkillCert) { // Only score if we have text
            scoreSingleProfileItem(`TECH_Sub${subSectionIndex}_Item${itemIndex}`, "TECHNICAL SKILLS", `${categoryNameForIdentifier}: ${skillCertIdentifier}`, textToScoreForSkillCert);
          }
        });
      });
    } else if(DEBUG) Logger.log("      No 'TECHNICAL SKILLS & CERTIFICATES' subsections found in profile data.");

    // --- Write all scored data rows to the sheet & Apply Formatting ---
    if (allScoredDataRowsForSheet.length > 0) {
      scoringSheet.getRange(2, 1, allScoredDataRowsForSheet.length, scoringSheetHeaders.length).setValues(allScoredDataRowsForSheet);
      if(DEBUG) Logger.log(`  RTS Stage 1: ${itemsScoredCounter} total items/bullets scored and written to "${scoringSheetNameConst}". Applying formatting...`);
      
      // Auto-resize columns, with fixed width for long text columns
      try { 
        scoringSheetHeaders.forEach((header, zeroBasedColIndex) => {
            const oneBasedColIndex = zeroBasedColIndex + 1;
            if (header === "OriginalBulletText" || header === "Justification" || header === "TailoredBulletText(Stage2)") {
                scoringSheet.setColumnWidth(oneBasedColIndex, 350); 
            } else if (header === "MatchingKeywords") {
                scoringSheet.setColumnWidth(oneBasedColIndex, 200);
            } else {
                scoringSheet.autoResizeColumn(oneBasedColIndex);
            }
        });
      } catch (eColResize) { Logger.log(`[WARN] RTS Stage 1: Column resizing error for "${scoringSheetNameConst}": ${eColResize.message}`); }

      // Conditional Formatting (using your specified ranges: 0.7-1 green, 0.4-0.69 yellow, 0.1-0.39 red)
      try {
        if(DEBUG) Logger.log(`    RTS Stage 1: Applying conditional formatting rules to "${scoringSheetNameConst}"...`);
        if (scoringSheet.getLastRow() > 1) { // Check if there's actual data beyond the header
          const scoreColHeaderTitle = "RelevanceScore";
          const scoreColZeroBasedIndex = scoringSheetHeaders.indexOf(scoreColHeaderTitle);
          if (scoreColZeroBasedIndex !== -1) {
            const scoreColSheetLetter = MJM_columnToLetter(scoreColZeroBasedIndex + 1); // MJM_columnToLetter assumed global from MJM_SheetUtils.gs
            // Apply formatting to all data rows, full width
            const dataFormattingA1Range = `A2:${MJM_columnToLetter(scoringSheetHeaders.length)}${scoringSheet.getLastRow()}`;
            const dataFormattingRangeObject = scoringSheet.getRange(dataFormattingA1Range);

            scoringSheet.clearConditionalFormatRules(); // Clear all rules on this sheet before applying new ones
            SpreadsheetApp.flush(); 
            
            let newConditionalFormatRules = [];
            newConditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=$${scoreColSheetLetter}2>=0.7`).setBackground("#c9ead3").setRanges([dataFormattingRangeObject]).build()); // Pale Green for >=0.7
            newConditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND($${scoreColSheetLetter}2>=0.4, $${scoreColSheetLetter}2<0.7)`).setBackground("#fff2cc").setRanges([dataFormattingRangeObject]).build()); // Pale Yellow for 0.4 to <0.7
            newConditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND($${scoreColSheetLetter}2>=0.1, $${scoreColSheetLetter}2<0.4)`).setBackground("#f4cccc").setRanges([dataFormattingRangeObject]).build()); // Pale Red for 0.1 to <0.4
            // Optional: Rule for scores less than 0.1 (but >0) if you want to distinguish very low scores
            // newConditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND($${scoreColSheetLetter}2>0, $${scoreColSheetLetter}2<0.1)`).setBackground("#efefef").setRanges([dataFormattingRangeObject]).build()); // Example: Lightest Grey

            if (newConditionalFormatRules.length > 0) {
                scoringSheet.setConditionalFormatRules(newConditionalFormatRules);
                if(DEBUG) Logger.log(`    RTS Stage 1: Conditional formatting rules applied based on score column ${scoreColSheetLetter} for range ${dataFormattingA1Range}.`);
            }
          } else { if(DEBUG) Logger.log(`    [WARN] RTS Stage 1: Could not find 'RelevanceScore' column header for conditional formatting.`); }
        }
      } catch(eCondFormat) { Logger.log(`[WARN] RTS Stage 1: Error applying conditional formatting: ${eCondFormat.message}\n${eCondFormat.stack ||''}`); }
    } else { if(DEBUG) Logger.log("  RTS Stage 1: No profile items were scored or found, so no data written to sheet."); }

    // Hide extra columns after column I ("TailoredBulletText(Stage2)")
    const lastVisibleHeaderIndexForScoring = scoringSheetHeaders.indexOf("TailoredBulletText(Stage2)"); // Zero-based
    if (lastVisibleHeaderIndexForScoring !== -1) {
        const lastVisibleColumnNumber = lastVisibleHeaderIndexForScoring + 1; // One-based
        const maxColsInScoring = scoringSheet.getMaxColumns();
        if (maxColsInScoring > lastVisibleColumnNumber) {
            try {
                scoringSheet.hideColumns(lastVisibleColumnNumber + 1, maxColsInScoring - lastVisibleColumnNumber);
                if(DEBUG) Logger.log(`    RTS Stage 1: Hid columns in "${scoringSheetNameConst}" from column ${MJM_columnToLetter(lastVisibleColumnNumber + 1)} onwards.`);
            } catch (eHideExtraCols) {
                Logger.log(`[WARN] RTS Stage 1: Could not hide extra columns in scoring sheet: ${eHideExtraCols.message}`);
            }
        }
    }

  // ---- End of MAIN TRY block ----
  } catch (eOuterTry) { 
    Logger.log(`[ERROR] RTS Stage 1 (Outer Try/Catch wrapping Sheet Operations and Main Scoring Logic): ${eOuterTry.toString()}\nStack: ${eOuterTry.stack}`);
    return { success: false, message: `Critical error during Stage 1 execution: ${eOuterTry.toString()}`, jdAnalysis: jdAnalysis /* jdAnalysis might be null if error happened before its assignment */ };
  }

  Logger.log(`--- RTS Stage 1: JD Analysis & Profile Scoring Process Finished ---`);
  return { 
    success: true, 
    message: `Stage 1 Complete. JD Analyzed. ${itemsScoredCounter} items scored and written to "${scoringSheetNameConst}". Please review, make selections with "${userSelectYesConst}" in the sheet, then run Stage 2.`, 
    jdAnalysis: jdAnalysis 
  };
}

// --- STAGE 2: TAILOR SELECTED BULLETS ---
function RTS_runStage2_TailorSelectedBullets(spreadsheetId) {
  const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : true); // Global
  Logger.log(`--- RTS Stage 2: Starting Tailoring Selected Profile Bullets (DEBUG: ${DEBUG}) ---`);

  if (!spreadsheetId) { 
    Logger.log("[ERROR] RTS Stage 2: Spreadsheet ID parameter is missing.");
    return { success: false, message: "Spreadsheet ID is required for Stage 2." }; 
  }

  // Constants from Global_Constants.gs
  const jdAnalysisSheetNameConst = (typeof JD_ANALYSIS_SHEET_NAME !== 'undefined' ? JD_ANALYSIS_SHEET_NAME : "JDAnalysisData_Fallback");
  const scoringSheetNameConst = (typeof BULLET_SCORING_RESULTS_SHEET_NAME !== 'undefined' ? BULLET_SCORING_RESULTS_SHEET_NAME : "BulletScoringResults_Fallback");
  const apiDelayStage2 = (typeof RTS_API_CALL_DELAY_STAGE2_MS !== 'undefined' ? RTS_API_CALL_DELAY_STAGE2_MS : 1000);
  const userSelectYesConst = (typeof USER_SELECT_YES_VALUE !== 'undefined' ? USER_SELECT_YES_VALUE : "YES");

  let currentSpreadsheet, jdAnalysis, scoringSheetDataRows, scoringSheetHeadersFromSheet;
  try {
    currentSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    if (!currentSpreadsheet) throw new Error(`Failed to open spreadsheet with ID: ${spreadsheetId}`);

    const jdSheet = currentSpreadsheet.getSheetByName(jdAnalysisSheetNameConst);
    const scoringSheet = currentSpreadsheet.getSheetByName(scoringSheetNameConst);

    if (!jdSheet) throw new Error(`Required sheet "${jdAnalysisSheetNameConst}" not found in spreadsheet ID ${spreadsheetId}. Run Stage 1 first.`);
    if (!scoringSheet) throw new Error(`Required sheet "${scoringSheetNameConst}" not found in spreadsheet ID ${spreadsheetId}. Run Stage 1 first.`);
    
    let jdAnalysisJsonString = ""; 
    try { jdAnalysisJsonString = jdSheet.getRange(2, 1).getValue(); } 
    catch(e) { throw new Error(`Failed to read JD Analysis from sheet "${jdSheet.getName()}": ${e.message}`); }
    
    try { jdAnalysis = JSON.parse(jdAnalysisJsonString); } 
    catch(e) { throw new Error(`Failed to parse JD Analysis JSON (String starts: "${String(jdAnalysisJsonString).substring(0,100)}..."): ${e.message}`);}
    
    if (!jdAnalysis?.jobTitle) throw new Error("Parsed JD Analysis from sheet is invalid or does not contain 'jobTitle'.");
    if(DEBUG) Logger.log("  RTS Stage 2: JD Analysis loaded successfully from sheet.");

    const allDataFromScoringSheet = scoringSheet.getDataRange().getValues();
    if (allDataFromScoringSheet.length < 2) { // Must have header + at least one data row
      Logger.log("  RTS Stage 2: No scored data rows found in the scoring sheet to process for tailoring.");
      return { success: true, message: "No scored data rows found to tailor.", bulletsTailored: 0 };
    }
    scoringSheetHeadersFromSheet = allDataFromScoringSheet.shift(); // Get actual headers from the sheet
    scoringSheetDataRows = allDataFromScoringSheet;     // Remaining are data rows
    if(DEBUG) Logger.log(`  RTS Stage 2: Loaded ${scoringSheetDataRows.length} data rows from "${scoringSheetNameConst}". Headers: [${scoringSheetHeadersFromSheet.join(', ')}]`);

  } catch (eLoad) {
    Logger.log(`[ERROR] RTS Stage 2 (Prerequisite Data Loading Phase): ${eLoad.toString()}\n${eLoad.stack || 'No stack trace'}`);
    return { success: false, message: `Error loading prerequisite data for Stage 2: ${eLoad.toString()}` };
  }

  // Expected header names
  const originalBulletColHeaderName = "OriginalBulletText";
  const selectToTailorColHeaderName = "SelectToTailor(Manual)";
  const tailoredTextColHeaderName = "TailoredBulletText(Stage2)";

  const originalBulletColIdx = scoringSheetHeadersFromSheet.indexOf(originalBulletColHeaderName);
  const selectColIdx = scoringSheetHeadersFromSheet.indexOf(selectToTailorColHeaderName);
  const tailoredColIdx = scoringSheetHeadersFromSheet.indexOf(tailoredTextColHeaderName);

  if ([originalBulletColIdx, selectColIdx, tailoredColIdx].includes(-1)) {
      const missing = [originalBulletColHeaderName,selectToTailorColHeaderName,tailoredTextColHeaderName].filter(h => scoringSheetHeadersFromSheet.indexOf(h) === -1);
      Logger.log(`[ERROR] RTS Stage 2: One or more required columns missing in "${scoringSheetNameConst}". Missing: [${missing.join(', ')}]. Available: [${scoringSheetHeadersFromSheet.join(', ')}]`);
      return { success: false, message: `Required columns missing in "${scoringSheetNameConst}" for Stage 2. Ensure headers are correct.`};
  }

  let bulletsAttemptedToTailor = 0;
  let bulletsSuccessfullyTailored = 0;
  
  for (let i = 0; i < scoringSheetDataRows.length; i++) {
    const currentRowData = scoringSheetDataRows[i];
    // Check if the selection column indicates "YES" (case-insensitive, trimmed)
    if (String(currentRowData[selectColIdx]).toUpperCase().trim() === userSelectYesConst) {
      bulletsAttemptedToTailor++;
      const originalBulletTextContent = String(currentRowData[originalBulletColIdx]);

      if (!originalBulletTextContent.trim()) { 
          scoringSheetDataRows[i][tailoredColIdx] = "TAILOR_SKIP: Original bullet text was empty."; 
          if(DEBUG) Logger.log(`    RTS Stage 2: Skipped tailoring for sheet row ${i+2} as original bullet was empty.`);
          continue; 
      }
      if(DEBUG) Logger.log(`    RTS Stage 2: Attempting to tailor bullet from sheet row ${i+2}: "${originalBulletTextContent.substring(0,70)}..."`);
      Utilities.sleep(apiDelayStage2); // Global constant
      // RTS_tailorBulletPoint from RTS_TailoringLogic.gs
      const tailoredOutputText = RTS_tailorBulletPoint(originalBulletTextContent, jdAnalysis, jdAnalysis.jobTitle);

      if (tailoredOutputText && !tailoredOutputText.startsWith("ERROR:") && tailoredOutputText.toLowerCase() !== "original bullet not suitable for significant tailoring towards this role.") {
        scoringSheetDataRows[i][tailoredColIdx] = tailoredOutputText.trim(); // Update the array for batch write
        bulletsSuccessfullyTailored++;
        if(DEBUG) Logger.log(`      RTS Stage 2: Successfully tailored bullet for row ${i+2}.`);
      } else {
        scoringSheetDataRows[i][tailoredColIdx] = `TAILOR_FAIL/SKIP: ${tailoredOutputText || 'No output from LLM.'}`;
        if(DEBUG) Logger.log(`    [WARN] RTS Stage 2 - Row ${i+2}: Failed or skipped tailoring for bullet. LLM Message: ${tailoredOutputText}`);
      }
    }
  }

  // Write updated data back to the sheet ONLY if any bullets were marked for tailoring
  if (bulletsAttemptedToTailor > 0) { 
    try { 
      const targetSheetForWriteback = currentSpreadsheet.getSheetByName(scoringSheetNameConst);
      if (targetSheetForWriteback) {
          targetSheetForWriteback.getRange(2, 1, scoringSheetDataRows.length, scoringSheetHeadersFromSheet.length)
                                 .setValues(scoringSheetDataRows);
          if(DEBUG) Logger.log(`  RTS Stage 2: Updated "${scoringSheetNameConst}" sheet with tailoring results/attempts.`);
      } else {
          Logger.log(`[ERROR] RTS Stage 2 (Sheet Writeback): Could not find sheet "${scoringSheetNameConst}" to write back results.`);
      }
    } catch (eSheetWrite) { 
        Logger.log(`[ERROR] RTS Stage 2 (Sheet Writeback): Failed to write tailored bullets back to sheet. ${eSheetWrite.toString()}\n${eSheetWrite.stack || ''}`);
    }
  } else {
      if(DEBUG) Logger.log(`  RTS Stage 2: No bullets were marked with '${userSelectYesConst}' in the selection column for tailoring.`);
  }

  Logger.log(`--- RTS Stage 2: Tailoring Process Finished. Attempted to tailor: ${bulletsAttemptedToTailor}, Successfully tailored: ${bulletsSuccessfullyTailored} ---`);
  return { 
    success: true, 
    message: `Stage 2 Complete. Tailoring Attempted: ${bulletsAttemptedToTailor}, Succeeded: ${bulletsSuccessfullyTailored}. Sheet "${scoringSheetNameConst}" updated.`, 
    bulletsTailored: bulletsSuccessfullyTailored 
  };
}

// --- STAGE 3: ASSEMBLE RESUME OBJECT & GENERATE DOCUMENT ---
function RTS_runStage3_BuildAndGenerateDocument(spreadsheetId) {
  const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : true);
  Logger.log(`--- RTS Stage 3: Starting Final Resume Assembly & Document Generation (DEBUG: ${DEBUG}) ---`);
  // Constants from Global_Constants.gs will be used throughout this function
  // e.g., PROFILE_DATA_SHEET_NAME, JD_ANALYSIS_SHEET_NAME, BULLET_SCORING_RESULTS_SHEET_NAME,
  // PROFILE_SCHEMA_VERSION, RTS_STAGE3_FINAL_INCLUSION_SCORE_THRESHOLD, RTS_STAGE3_MAX_BULLETS_PER_JOB, 
  // USER_SELECT_YES_VALUE, RTS_API_CALL_DELAY_STAGE3_SUMMARY_MS, RESUME_TEMPLATE_DOC_ID

  if (!spreadsheetId) {
    Logger.log("[ERROR] RTS Stage 3: Spreadsheet ID parameter is missing.");
    return { success: false, message: "Spreadsheet ID is required for Stage 3." };
  }

  let masterProfileObject, jdAnalysis, allScoredBulletObjectsFromSheet;
  try {
    // Load Master Profile (PROFILE_DATA_SHEET_NAME from Global)
    masterProfileObject = RTS_getMasterProfileData(spreadsheetId, PROFILE_DATA_SHEET_NAME); // From RTS_MasterResumeData.gs
    if (!masterProfileObject?.personalInfo) { // Check for both object and personalInfo
      throw new Error("Master Profile data could not be loaded or is missing critical 'personalInfo' section.");
    }
    if(DEBUG) Logger.log("  RTS Stage 3: Master Profile loaded successfully.");

    const ss = SpreadsheetApp.openById(spreadsheetId);
    const jdAnalysisSheet = ss.getSheetByName(JD_ANALYSIS_SHEET_NAME); // Global
    const scoringResultSheet = ss.getSheetByName(BULLET_SCORING_RESULTS_SHEET_NAME); // Global
    if(!jdAnalysisSheet || !scoringResultSheet) {
      throw new Error(`Missing prerequisite sheets for Stage 3: JD Analysis ("${JD_ANALYSIS_SHEET_NAME}") or Scoring Results ("${BULLET_SCORING_RESULTS_SHEET_NAME}"). Please run Stage 1 first.`);
    }
    
    let jdAnalysisJsonStringValue = "";
    try { jdAnalysisJsonStringValue = jdAnalysisSheet.getRange(2,1).getValue(); }
    catch(e) { throw new Error(`Error reading JD Analysis JSON string from sheet "${JD_ANALYSIS_SHEET_NAME}": ${e.message}`);}
    
    try { jdAnalysis = JSON.parse(jdAnalysisJsonStringValue); }
    catch(e) { throw new Error(`Error parsing JD Analysis JSON from sheet (length ${jdAnalysisJsonStringValue.length}, starts: "${String(jdAnalysisJsonStringValue).substring(0,100)}..."): ${e.message}`);}
    
    if(!jdAnalysis?.jobTitle) { throw new Error("Parsed JD Analysis from sheet appears invalid or does not contain a 'jobTitle'."); }
    if(DEBUG) Logger.log("  RTS Stage 3: JD Analysis loaded and parsed successfully.");

    const scoringSheetAllValues = scoringResultSheet.getDataRange().getValues();
    if(scoringSheetAllValues.length < 2) { // Needs header + at least one data row
      allScoredBulletObjectsFromSheet = [];
      if(DEBUG) Logger.log(`  RTS Stage 3: No data rows found in "${BULLET_SCORING_RESULTS_SHEET_NAME}".`);
    } else {
      const scoringSheetHeadersArray = scoringSheetAllValues.shift(); // Get headers
      allScoredBulletObjectsFromSheet = scoringSheetAllValues.map(rowDataArray => {
        let obj = {}; 
        scoringSheetHeadersArray.forEach((header, idx) => obj[header] = rowDataArray[idx]); 
        return obj;
      });
    }
    if(DEBUG) Logger.log(`  RTS Stage 3: Loaded ${allScoredBulletObjectsFromSheet.length} scored entries from "${BULLET_SCORING_RESULTS_SHEET_NAME}".`);

  } catch (eLoadData) { 
    Logger.log(`[ERROR] RTS Stage 3 (Prerequisite Data Loading Phase): ${eLoadData.message}\n${eLoadData.stack || ''}`); 
    return { success: false, message: `Error loading prerequisite data for Stage 3: ${eLoadData.message}` };
  }

  // Initialize the final resume object
  const finalTailoredResumeObject = {
    resumeSchemaVersion: (typeof PROFILE_SCHEMA_VERSION !== 'undefined' ? PROFILE_SCHEMA_VERSION : "IntegratedSchema_vX.Y"), // Global
    personalInfo: JSON.parse(JSON.stringify(masterProfileObject.personalInfo)), // Deep copy
    summary: "", 
    sections: []
  };
  let highlightsForAISummaryGeneration = [];

  // Helper to get the text to use (Tailored if available and good, else Original)
  const getEffectiveBulletTextFromScoredEntry = (scoredEntryObject) => {
    const tailoredTextFromStage2 = scoredEntryObject["TailoredBulletText(Stage2)"];
    if (tailoredTextFromStage2 && 
        String(tailoredTextFromStage2).trim() && 
        !String(tailoredTextFromStage2).toUpperCase().startsWith("TAILOR_FAIL") && 
        String(tailoredTextFromStage2).toLowerCase() !== "original bullet not suitable for significant tailoring towards this role.") {
        return String(tailoredTextFromStage2).trim();
    }
    return String(scoredEntryObject.OriginalBulletText || "").trim();
  };
  
  // --- Assemble DYNAMIC Sections based on Scores and User Selections ---
  // Using Global Constants: RTS_STAGE3_FINAL_INCLUSION_SCORE_THRESHOLD, USER_SELECT_YES_VALUE, RTS_STAGE3_MAX_BULLETS_PER_JOB/PROJECT

  // EXPERIENCE Section Assembly
  if(DEBUG) Logger.log("    RTS Stage 3: Assembling EXPERIENCE section...");
  const assembledExperienceItems = [];
  (masterProfileObject.sections.find(s=>s.title==="EXPERIENCE")?.items || []).forEach(masterJobEntry => {
    const selectedFilteredSortedBullets = allScoredBulletObjectsFromSheet
      .filter(scoredEntry => 
        scoredEntry.Section === "EXPERIENCE" && 
        scoredEntry.ItemIdentifier === masterJobEntry.company && // Match by company name as item identifier
        parseFloat(scoredEntry.RelevanceScore) >= RTS_STAGE3_FINAL_INCLUSION_SCORE_THRESHOLD && 
        String(scoredEntry["SelectToTailor(Manual)"]).toUpperCase().trim() === USER_SELECT_YES_VALUE)
      .sort((a,b) => parseFloat(b.RelevanceScore) - parseFloat(a.RelevanceScore)) // Highest score first
      .slice(0, RTS_STAGE3_MAX_BULLETS_PER_JOB) // Limit number of bullets per job
      .map(sb => getEffectiveBulletTextFromScoredEntry(sb)).filter(bullet => bullet); // Get text & remove any empty ones
    if (selectedFilteredSortedBullets.length > 0) {
      assembledExperienceItems.push({ ...masterJobEntry, responsibilities: selectedFilteredSortedBullets });
      highlightsForAISummaryGeneration.push(...selectedFilteredSortedBullets);
    }
  });
  if(assembledExperienceItems.length > 0) finalTailoredResumeObject.sections.push({ title: "EXPERIENCE", items: assembledExperienceItems });

  // PROJECTS Section Assembly
  if(DEBUG) Logger.log("    RTS Stage 3: Assembling PROJECTS section...");
  const assembledProjectSubsections = [];
  (masterProfileObject.sections.find(s=>s.title==="PROJECTS")?.subsections || []).forEach(masterSubSection => {
      const assembledProjectItemsForSubSection = [];
      (masterSubSection.items || []).forEach(masterProjectEntry => {
          const selectedFilteredSortedProjectBullets = allScoredBulletObjectsFromSheet
              .filter(scoredEntry => 
                  scoredEntry.Section === "PROJECTS" && 
                  scoredEntry.ItemIdentifier === masterProjectEntry.projectName && // Match by project name
                  parseFloat(scoredEntry.RelevanceScore) >= RTS_STAGE3_FINAL_INCLUSION_SCORE_THRESHOLD && 
                  String(scoredEntry["SelectToTailor(Manual)"]).toUpperCase().trim() === USER_SELECT_YES_VALUE)
              .sort((a,b)=>parseFloat(b.RelevanceScore)-parseFloat(a.RelevanceScore))
              .slice(0,RTS_STAGE3_MAX_BULLETS_PER_PROJECT) // Global limit for projects
              .map(sb=>getEffectiveBulletTextFromScoredEntry(sb)).filter(bullet=>bullet);
          if(selectedFilteredSortedProjectBullets.length > 0) {
              assembledProjectItemsForSubSection.push({...masterProjectEntry, descriptionBullets: selectedFilteredSortedProjectBullets}); // Assume project items store bullets in 'descriptionBullets'
              highlightsForAISummaryGeneration.push(...selectedFilteredSortedProjectBullets);
          }
      });
      if(assembledProjectItemsForSubSection.length > 0) assembledProjectSubsections.push({name: masterSubSection.name, items: assembledProjectItemsForSubSection});
  });
  if(assembledProjectSubsections.length > 0) finalTailoredResumeObject.sections.push({ title: "PROJECTS", subsections: assembledProjectSubsections });
  
  // LEADERSHIP & UNIVERSITY INVOLVEMENT Section Assembly
  if(DEBUG) Logger.log("    RTS Stage 3: Assembling LEADERSHIP & UNIVERSITY INVOLVEMENT section...");
  const assembledLeadershipItems = [];
  (masterProfileObject.sections.find(s=>s.title==="LEADERSHIP & UNIVERSITY INVOLVEMENT")?.items || []).forEach(masterLeadershipItem => {
    const itemIdentifier = masterLeadershipItem.organization || masterLeadershipItem.role || `LeadItemUndef`; // Identifier for matching
    const selectedFilteredSortedLeadershipBullets = allScoredBulletObjectsFromSheet
      .filter(scoredEntry => 
          scoredEntry.Section === "LEADERSHIP" && // Match "LEADERSHIP" (as set in Stage 1 scoring)
          scoredEntry.ItemIdentifier === itemIdentifier && 
          parseFloat(scoredEntry.RelevanceScore) >= RTS_STAGE3_FINAL_INCLUSION_SCORE_THRESHOLD && 
          String(scoredEntry["SelectToTailor(Manual)"]).toUpperCase().trim() === USER_SELECT_YES_VALUE)
      .sort((a,b)=>parseFloat(b.RelevanceScore)-parseFloat(a.RelevanceScore))
      .slice(0,RTS_STAGE3_MAX_BULLETS_PER_JOB) // Re-using job limit, or define a new global const for leadership bullets
      .map(sb=>getEffectiveBulletTextFromScoredEntry(sb)).filter(bullet=>bullet);
    if(selectedFilteredSortedLeadershipBullets.length > 0) {
      assembledLeadershipItems.push({...masterLeadershipItem, responsibilities: selectedFilteredSortedLeadershipBullets});
      highlightsForAISummaryGeneration.push(...selectedFilteredSortedLeadershipBullets);
    }
  });
  if(assembledLeadershipItems.length > 0) finalTailoredResumeObject.sections.push({ title: "LEADERSHIP & UNIVERSITY INVOLVEMENT", items: assembledLeadershipItems });
  
  // TECHNICAL SKILLS & CERTIFICATES Section Assembly
  if(DEBUG) Logger.log("    RTS Stage 3: Assembling TECHNICAL SKILLS & CERTIFICATES section...");
  const assembledTechSubsections = [];
  (masterProfileObject.sections.find(s=>s.title==="TECHNICAL SKILLS & CERTIFICATES")?.subsections || []).forEach(masterTechSubSection => {
      const selectedTechItemsForSub = [];
      (masterTechSubSection.items || []).forEach(masterTechItem => {
          // The ItemIdentifier for tech skills in scoring sheet was "CategoryName: SkillName" or "CategoryName: CertName"
          const techItemSheetIdentifier = `${masterTechSubSection.name}: ${masterTechItem.skill || masterTechItem.name}`;
          const isSelected = allScoredBulletObjectsFromSheet.some(scoredEntry => 
              scoredEntry.Section === "TECHNICAL SKILLS" && // As set in Stage 1
              scoredEntry.ItemIdentifier === techItemSheetIdentifier &&
              parseFloat(scoredEntry.RelevanceScore) >= RTS_STAGE3_FINAL_INCLUSION_SCORE_THRESHOLD && 
              String(scoredEntry["SelectToTailor(Manual)"]).toUpperCase().trim() === USER_SELECT_YES_VALUE
          );
          if(isSelected) {
              selectedTechItemsForSub.push(masterTechItem); // Add the original master item if selected
          }
      });
      if(selectedTechItemsForSub.length > 0) assembledTechSubsections.push({name: masterTechSubSection.name, items: selectedTechItemsForSub });
  });
  if(assembledTechSubsections.length > 0)finalTailoredResumeObject.sections.push({title:"TECHNICAL SKILLS & CERTIFICATES",subsections:assembledTechSubsections});

  // Generate AI Summary
  if(DEBUG) Logger.log(`    RTS Stage 3: Generating AI summary. Number of collected highlight phrases: ${highlightsForAISummaryGeneration.length}`);
  const highlightsTextForLLM = highlightsForAISummaryGeneration.sort((a,b)=>b.length-a.length).slice(0,RTS_STAGE3_MAX_HIGHLIGHTS_FOR_SUMMARY).join(" \n").trim(); // Global const
  const summaryInputTextToLLM = highlightsTextForLLM ? highlightsTextForLLM : (masterProfileObject.summary || "A highly skilled and motivated professional seeking a challenging role.");
  Utilities.sleep(RTS_API_CALL_DELAY_STAGE3_SUMMARY_MS); // Global const
  // RTS_generateTailoredSummary from RTS_TailoringLogic.gs
  const aiGeneratedSummaryText = RTS_generateTailoredSummary(summaryInputTextToLLM, jdAnalysis, finalTailoredResumeObject.personalInfo.fullName);
  finalTailoredResumeObject.summary = (aiGeneratedSummaryText && !aiGeneratedSummaryText.startsWith("ERROR:")) ? aiGeneratedSummaryText.trim() : (masterProfileObject.summary || "");
  if(DEBUG) Logger.log(`    RTS Stage 3: Summary set. Using AI: ${aiGeneratedSummaryText && !aiGeneratedSummaryText.startsWith("ERROR:")}. Length: ${finalTailoredResumeObject.summary.length}`);

  // Add Static Sections & Ensure Final Order
  if(DEBUG) Logger.log("    RTS Stage 3: Adding static sections (e.g., Education, Honors) and ensuring final document order.");
  const finalResumeSectionOrder = ["EDUCATION", "TECHNICAL SKILLS & CERTIFICATES", "EXPERIENCE", "PROJECTS", "LEADERSHIP & UNIVERSITY INVOLVEMENT", "HONORS & AWARDS"]; // Could be Global Constant
  const titlesCurrentlyInFinalResume = finalTailoredResumeObject.sections.map(s=>s.title.toUpperCase());
  finalResumeSectionOrder.forEach(sectionTitleToEnsure => {
    if(!titlesCurrentlyInFinalResume.includes(sectionTitleToEnsure.toUpperCase())){ // If section not already dynamically added
      const staticSectionDataFromMaster = masterProfileObject.sections.find(s => s.title.toUpperCase() === sectionTitleToEnsure.toUpperCase());
      // Add only if it has content
      if(staticSectionDataFromMaster && ((staticSectionDataFromMaster.items && staticSectionDataFromMaster.items.length > 0) || 
                                   (staticSectionDataFromMaster.subsections && staticSectionDataFromMaster.subsections.some(ss=>ss.items && ss.items.length > 0)))){
        finalTailoredResumeObject.sections.push(JSON.parse(JSON.stringify(staticSectionDataFromMaster))); // Deep copy to avoid reference issues
        if(DEBUG) Logger.log(`      Added static section to output: ${sectionTitleToEnsure}`);
      } else if(DEBUG) Logger.log(`      Static section "${sectionTitleToEnsure}" not found in master or was empty; not adding.`);
    }
  });
  finalTailoredResumeObject.sections.sort((a,b) => finalResumeSectionOrder.indexOf(a.title.toUpperCase()) - finalResumeSectionOrder.indexOf(b.title.toUpperCase()));
  if(DEBUG) Logger.log(`    RTS Stage 3: Final tailored resume object assembled. Section order: ${finalTailoredResumeObject.sections.map(s=>s.title).join(', ')}`);

  // Generate Google Document
  let docTitleForFile = `Tailored Resume - ${finalTailoredResumeObject.personalInfo.fullName || 'Candidate'}`;
  if(jdAnalysis?.jobTitle) docTitleForFile += ` for ${String(jdAnalysis.jobTitle).replace(/[^\w\s-]/g,"").substring(0,35)}`; // Sanitize for filename
  const nowTimestamp = new Date(); 
  const timestampSuffix = ` (${nowTimestamp.getFullYear()}-${("0"+(nowTimestamp.getMonth()+1)).slice(-2)}-${("0"+nowTimestamp.getDate()).slice(-2)}_${("0"+nowTimestamp.getHours()).slice(-2)}-${("0"+nowTimestamp.getMinutes()).slice(-2)})`;
  docTitleForFile += timestampSuffix;

  // RTS_createFormattedResumeDoc from RTS_DocumentService.gs, RESUME_TEMPLATE_DOC_ID from Global_Constants.gs
  const generatedDocUrl = RTS_createFormattedResumeDoc(finalTailoredResumeObject, docTitleForFile, RESUME_TEMPLATE_DOC_ID);

  if (generatedDocUrl) {
    Logger.log(`--- RTS Stage 3: SUCCESS! Tailored Resume Document Generated: ${generatedDocUrl} ---`);
    return { success: true, message: "Stage 3 Complete: Tailored resume document generated successfully.", docUrl: generatedDocUrl, tailoredResumeObjectForDebug: DEBUG ? finalTailoredResumeObject : "Debug object excluded." };
  } else {
    Logger.log("--- RTS Stage 3: FAILED - Document generation error (RTS_createFormattedResumeDoc returned null/false). ---");
    return { success: false, message: "Stage 3 Failed: Document generation process did not return a valid URL.", tailoredResumeObjectForDebug: DEBUG ? finalTailoredResumeObject : "Debug object excluded." };
  }
}

// --- STAGE 3: ASSEMBLE RESUME OBJECT & GENERATE DOCUMENT ---
function RTS_runStage3_BuildAndGenerateDocument(spreadsheetId) {
  const DEBUG = GLOBAL_DEBUG_MODE;
  Logger.log("--- RTS Stage 3: Starting Final Resume Assembly & Document Generation ---");
  // Constants from Global_Constants.gs: PROFILE_DATA_SHEET_NAME, JD_ANALYSIS_SHEET_NAME, BULLET_SCORING_RESULTS_SHEET_NAME,
  // PROFILE_SCHEMA_VERSION, RTS_STAGE3_..._THRESHOLD/MAX_..., USER_SELECT_YES_VALUE, RTS_API_CALL_DELAY_STAGE3_SUMMARY_MS, RESUME_TEMPLATE_DOC_ID

  if (!spreadsheetId) return { success: false, message: "Spreadsheet ID required for Stage 3." };

  let masterProfileObject, jdAnalysis, allScoredBulletObjects;
  try {
    masterProfileObject = RTS_getMasterProfileData(spreadsheetId, PROFILE_DATA_SHEET_NAME); // Global const for sheet name
    if (!masterProfileObject?.personalInfo) throw new Error("Failed to load master profile or personalInfo missing.");
    if(DEBUG) Logger.log("  RTS Stage 3: Master Profile loaded.");

    const currentSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const jdSheet = currentSpreadsheet.getSheetByName(JD_ANALYSIS_SHEET_NAME);
    const scoringSheet = currentSpreadsheet.getSheetByName(BULLET_SCORING_RESULTS_SHEET_NAME);
    if(!jdSheet || !scoringSheet) throw new Error("Missing JD Analysis or Bullet Scoring sheets. Run Stage 1 & 2.");

    const jdStr = jdSheet.getRange(2,1).getValue(); jdAnalysis = JSON.parse(jdStr); // Add try-catch
    if(!jdAnalysis?.jobTitle) throw new Error("JD Analysis data invalid from sheet.");
    if(DEBUG) Logger.log("  RTS Stage 3: JD Analysis loaded.");

    const scoreData = scoringSheet.getDataRange().getValues();
    if(scoreData.length < 2) allScoredBulletObjects = [];
    else {const hdrs=scoreData.shift(); allScoredBulletObjects=scoreData.map(r=>{let o={};hdrs.forEach((h,i)=>o[h]=r[i]);return o;});}
    if(DEBUG) Logger.log(`  RTS Stage 3: Loaded ${allScoredBulletObjects.length} scored entries.`);
  } catch (e) { Logger.log(`[ERROR] RTS Stage 3 (Data Loading): ${e}`); return { success: false, message: `Error loading data: ${e}` };}

  const finalTailoredResumeObject = {
    resumeSchemaVersion: (typeof PROFILE_SCHEMA_VERSION !== 'undefined' ? PROFILE_SCHEMA_VERSION : "unknown_integrated"),
    personalInfo: JSON.parse(JSON.stringify(masterProfileObject.personalInfo)), // Deep copy
    summary: "", sections: []
  };
  let includedContentForSummary = [];

  const getEffectiveBullet = (scoredEntry) => {
    const tailored = scoredEntry["TailoredBulletText(Stage2)"];
    if (tailored && String(tailored).trim() && !String(tailored).toUpperCase().startsWith("TAILOR_FAIL") && String(tailored).toLowerCase() !== "original bullet not suitable for significant tailoring towards this role.") {
        return String(tailored).trim();
    }
    return String(scoredEntry.OriginalBulletText || "").trim();
  };

  // Process EXPERIENCE Section
  if(DEBUG) Logger.log("    RTS Stage 3: Assembling EXPERIENCE section...");
  const tailoredExpSection = { title: "EXPERIENCE", items: [] };
  (masterProfileObject.sections.find(s=>s.title==="EXPERIENCE")?.items || []).forEach(masterJob => {
    const bullets = allScoredBulletObjects
      .filter(sb => sb.Section === "EXPERIENCE" && sb.ItemIdentifier === masterJob.company && parseFloat(sb.RelevanceScore) >= RTS_STAGE3_FINAL_INCLUSION_SCORE_THRESHOLD && String(sb["SelectToTailor(Manual)"]).toUpperCase().trim() === USER_SELECT_YES_VALUE)
      .sort((a,b) => parseFloat(b.RelevanceScore) - parseFloat(a.RelevanceScore))
      .slice(0, RTS_STAGE3_MAX_BULLETS_PER_JOB) // Global const
      .map(sb => getEffectiveBullet(sb)).filter(b => b); // Filter out any empty bullets after processing
    if (bullets.length > 0) {
      tailoredExpSection.items.push({ ...masterJob, responsibilities: bullets });
      includedContentForSummary.push(...bullets);
    }
  });
  if(tailoredExpSection.items.length > 0) finalTailoredResumeObject.sections.push(tailoredExpSection);

  // Process PROJECTS Section
  if(DEBUG) Logger.log("    RTS Stage 3: Assembling PROJECTS section...");
  const tailoredProjSection = { title: "PROJECTS", subsections: [] }; // Assuming projects use subsections as per MasterResumeData
  (masterProfileObject.sections.find(s=>s.title==="PROJECTS")?.subsections || []).forEach(masterSubSection => {
      const tailoredItemsForSub = [];
      (masterSubSection.items || []).forEach(masterProj => {
          const bullets = allScoredBulletObjects
              .filter(sb => sb.Section === "PROJECTS" && sb.ItemIdentifier === masterProj.projectName && parseFloat(sb.RelevanceScore) >= RTS_STAGE3_FINAL_INCLUSION_SCORE_THRESHOLD && String(sb["SelectToTailor(Manual)"]).toUpperCase().trim() === USER_SELECT_YES_VALUE)
              .sort((a,b) => parseFloat(b.RelevanceScore) - parseFloat(a.RelevanceScore))
              .slice(0, RTS_STAGE3_MAX_BULLETS_PER_PROJECT) // Global const
              .map(sb => getEffectiveBullet(sb)).filter(b => b);
          if (bullets.length > 0) {
              // Ensure the project structure matches what RTS_DocumentService expects
              tailoredItemsForSub.push({ 
                  ...masterProj, // Spreads original project details (name, org, role, dates, technologies etc.)
                  descriptionBullets: bullets // Overwrites or sets the bullets
              });
              includedContentForSummary.push(...bullets);
          }
      });
      if(tailoredItemsForSub.length > 0) tailoredProjSection.subsections.push({name: masterSubSection.name, items: tailoredItemsForSub});
  });
  if(tailoredProjSection.subsections.some(ss => ss.items && ss.items.length > 0)) finalTailoredResumeObject.sections.push(tailoredProjSection);

  // Process LEADERSHIP & UNIVERSITY INVOLVEMENT Section
  if(DEBUG) Logger.log("    RTS Stage 3: Assembling LEADERSHIP section...");
  const tailoredLeadSection = { title: "LEADERSHIP & UNIVERSITY INVOLVEMENT", items: [] };
  (masterProfileObject.sections.find(s=>s.title==="LEADERSHIP & UNIVERSITY INVOLVEMENT")?.items || []).forEach(masterItem => {
    const bullets = allScoredBulletObjects
      .filter(sb => sb.Section === "LEADERSHIP" && sb.ItemIdentifier === (masterItem.organization || masterItem.role) && parseFloat(sb.RelevanceScore) >= RTS_STAGE3_FINAL_INCLUSION_SCORE_THRESHOLD && String(sb["SelectToTailor(Manual)"]).toUpperCase().trim() === USER_SELECT_YES_VALUE)
      .sort((a,b) => parseFloat(b.RelevanceScore) - parseFloat(a.RelevanceScore))
      .slice(0, RTS_STAGE3_MAX_BULLETS_PER_JOB) // Re-use job limit or make a new const
      .map(sb => getEffectiveBullet(sb)).filter(b => b);
    if (bullets.length > 0) {
      tailoredLeadSection.items.push({ ...masterItem, responsibilities: bullets });
      includedContentForSummary.push(...bullets);
    }
  });
  if(tailoredLeadSection.items.length > 0) finalTailoredResumeObject.sections.push(tailoredLeadSection);
  
  // Process TECHNICAL SKILLS & CERTIFICATES Section
  if(DEBUG) Logger.log("    RTS Stage 3: Assembling TECHNICAL SKILLS section...");
  const tailoredTechSection = { title: "TECHNICAL SKILLS & CERTIFICATES", subsections: [] };
  const masterTechSubsections = (masterProfileObject.sections.find(s => s.title === "TECHNICAL SKILLS & CERTIFICATES")?.subsections || []);
  masterTechSubsections.forEach(masterSub => {
      const selectedItemsForSub = [];
      (masterSub.items || []).forEach(masterItem => {
          const identifier = masterItem.skill || masterItem.name; // Skill name or Cert name
          const isSelected = allScoredBulletObjects.some(sb => 
              sb.Section === "TECHNICAL SKILLS" && 
              sb.ItemIdentifier === (masterSub.name + ": " + identifier) && // As stored in Stage 1 for skills
              parseFloat(sb.RelevanceScore) >= RTS_STAGE3_FINAL_INCLUSION_SCORE_THRESHOLD && 
              String(sb["SelectToTailor(Manual)"]).toUpperCase().trim() === USER_SELECT_YES_VALUE
          );
          if (isSelected) {
              selectedItemsForSub.push(masterItem); // Add the original master item structure
          }
      });
      if(selectedItemsForSub.length > 0) tailoredTechSection.subsections.push({ name: masterSub.name, items: selectedItemsForSub });
  });
  if(tailoredTechSection.subsections.length > 0) finalTailoredResumeObject.sections.push(tailoredTechSection);

  // Generate AI Summary
  if(DEBUG) Logger.log(`    RTS Stage 3: Generating summary from ${includedContentForSummary.length} highlights.`);
  const highlightsText = includedContentForSummary.sort((a,b)=>b.length-a.length).slice(0,RTS_STAGE3_MAX_HIGHLIGHTS_FOR_SUMMARY).join(" \n "); // Global
  const summaryInput = highlightsText.trim() ? highlightsText : (masterProfileObject.summary || "Highly motivated professional.");
  Utilities.sleep(RTS_API_CALL_DELAY_STAGE3_SUMMARY_MS); // Global
  // RTS_generateTailoredSummary from RTS_TailoringLogic.gs
  const aiSummary = RTS_generateTailoredSummary(summaryInput, jdAnalysis, finalTailoredResumeObject.personalInfo.fullName);
  finalTailoredResumeObject.summary = (aiSummary && !aiSummary.startsWith("ERROR:")) ? aiSummary.trim() : (masterProfileObject.summary || "");

  // Add Static Sections & Final Ordering
  if(DEBUG) Logger.log("    RTS Stage 3: Adding static sections and ordering.");
  const finalSectionOrder = ["EDUCATION", "EXPERIENCE", "PROJECTS", "TECHNICAL SKILLS & CERTIFICATES", "LEADERSHIP & UNIVERSITY INVOLVEMENT", "HONORS & AWARDS"]; // Could be Global Const
  const existingSectionTitles = finalTailoredResumeObject.sections.map(s=>s.title.toUpperCase());
  finalSectionOrder.forEach(titleToEnsure => {
    if(!existingSectionTitles.includes(titleToEnsure.toUpperCase())){
      const staticSection = masterProfileObject.sections.find(s => s.title.toUpperCase() === titleToEnsure.toUpperCase());
      if(staticSection && ((staticSection.items && staticSection.items.length > 0) || (staticSection.subsections && staticSection.subsections.some(ss=>ss.items && ss.items.length > 0)))) {
        finalTailoredResumeObject.sections.push(JSON.parse(JSON.stringify(staticSection))); // Deep copy
      }
    }
  });
  finalTailoredResumeObject.sections.sort((a,b) => finalSectionOrder.indexOf(a.title.toUpperCase()) - finalSectionOrder.indexOf(b.title.toUpperCase()));

  // Generate Document
  let docTitle = `Tailored Resume - ${finalTailoredResumeObject.personalInfo.fullName || 'Candidate'}`;
  if(jdAnalysis?.jobTitle) docTitle += ` for ${jdAnalysis.jobTitle.replace(/[^\w\s-]/g,"").substring(0,30)}`;
  docTitle += ` (${new Date().toISOString().substring(0,10)})`;
  // RTS_createFormattedResumeDoc from RTS_DocumentService.gs, RESUME_TEMPLATE_DOC_ID from Global_Constants.gs
  const docUrl = RTS_createFormattedResumeDoc(finalTailoredResumeObject, docTitle, RESUME_TEMPLATE_DOC_ID);

  if (docUrl) {
    Logger.log(`--- RTS Stage 3: SUCCESS! Document: ${docUrl} ---`);
    return { success: true, message: "Stage 3 Complete: Tailored resume document generated.", docUrl: docUrl, tailoredResumeObjectForDebug: DEBUG ? finalTailoredResumeObject : "Debug data excluded." };
  } else {
    Logger.log("--- RTS Stage 3: FAILED - Document generation error. ---");
    return { success: false, message: "Stage 3 Failed: Document generation error.", tailoredResumeObjectForDebug: DEBUG ? finalTailoredResumeObject : "Debug data excluded." };
  }
}
