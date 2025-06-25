// File: RTS_MasterResumeData.gs
// Description: Responsible for fetching and parsing master resume/profile data from a specified
// Google Sheet tab (now within the main application spreadsheet).
// It transforms the tabular sheet data into a structured JavaScript object.
// Relies on global constants for profile structure, schema version, and debug mode.

// --- HELPER FUNCTIONS for Data Parsing (Specific to Profile Data Structure) ---

/**
 * Formats a date value from a Google Sheet cell for profile sections.
 */
function RTS_Profile_formatDateFromSheet(dateCellValue, isPotentialEndDate = false) {
  if (dateCellValue === null || dateCellValue === undefined || String(dateCellValue).trim() === "") {
    return isPotentialEndDate ? "Present" : "";
  }
  const dateString = String(dateCellValue).trim();
  if (dateString.toUpperCase() === "PRESENT") {
    return "Present";
  }
  let dateObj;
  if (dateCellValue instanceof Date) { dateObj = dateCellValue; }
  else { dateObj = new Date(dateString); }
  if (isNaN(dateObj.getTime())) { return dateString; }
  if (/^\d{4}$/.test(dateString)) { return dateString; }
  if (/^[A-Za-z]{3,9}\s\d{4}$/.test(dateString)) { return dateString; }
  const monthNames = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  return `${monthNames[dateObj.getMonth()]} ${dateObj.getFullYear()}`;
}

/**
 * Splits multi-line text from a sheet cell into an array of strings for profile sections.
 */
function RTS_Profile_smartSplitMultiLine(textCellContent) {
  if (!textCellContent || typeof textCellContent !== 'string' || textCellContent.trim() === "") { return []; }
  const stringWithActualNewlines = textCellContent.replace(/\\n/g, '\n');
  return stringWithActualNewlines.split(/\r\n|\r|\n/g).map(s => s.trim()).filter(s => s && s.length > 0);
}

/**
 * Gets a canonical section title using PROFILE_STRUCTURE from Global_Constants.gs.
 */
function RTS_Profile_getCanonicalSectionTitle(text) {
  if (!text || typeof text !== 'string') return null;
  const upperText = text.trim().toUpperCase();
  const profileStructureConst = (typeof PROFILE_STRUCTURE !== 'undefined' ? PROFILE_STRUCTURE : []);
  const foundSection = profileStructureConst.find(section => section.title && upperText.startsWith(section.title.toUpperCase()));
  return foundSection ? foundSection.title : null;
}

/**
 * Checks if a section is configured as tabular (has a `headers` array in PROFILE_STRUCTURE).
 */
function RTS_Profile_isTabularSection(sectionTitleKey) {
  const profileStructureConst = (typeof PROFILE_STRUCTURE !== 'undefined' ? PROFILE_STRUCTURE : []);
  const sectionConfig = profileStructureConst.find(s => s.title === sectionTitleKey);
  return sectionConfig ? (Array.isArray(sectionConfig.headers)) : false;
}

/**
 * Checks if a row array from sheet.getValues() is effectively blank.
 */
function RTS_Profile_rowIsEffectivelyBlank(rowArray) {
  if (!rowArray || !Array.isArray(rowArray)) return true;
  return rowArray.every(cell => (cell === null || cell === undefined || String(cell).trim() === ""));
}


// --- MAIN DATA FETCHING AND PARSING FUNCTION ---
/**
 * Fetches and parses master profile data.
 * @param {string} spreadsheetId ID of the spreadsheet.
 * @param {string} profileDataSheetName Name of the sheet tab with profile data.
 * @return {Object|null} Structured profile data object or null on error.
 */
function RTS_getMasterProfileData(spreadsheetId, profileDataSheetName) {
  const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : true);
  const SCHEMA_VERSION = (typeof PROFILE_SCHEMA_VERSION !== 'undefined' ? PROFILE_SCHEMA_VERSION : "SchemaVersion_Unavailable");
  const profileStructureConst = (typeof PROFILE_STRUCTURE !== 'undefined' ? PROFILE_STRUCTURE : []);
  
  Logger.log(`--- RTS_MasterProfileData: Loading Profile from Passed SSID: "${spreadsheetId}", Sheet: "${profileDataSheetName}" (DEBUG_MODE: ${DEBUG}) ---`);

  const profileOutput = { profileSchemaVersion: SCHEMA_VERSION, personalInfo: {}, summary: "", sections: [] };

  if (!spreadsheetId || !profileDataSheetName) { Logger.log("[ERROR] RTS_MasterProfileData: Missing spreadsheetId or profileDataSheetName."); return null; }

  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    if (!ss) { Logger.log(`[ERROR] RTS_MasterProfileData: Cannot open spreadsheet ID "${spreadsheetId}".`); return null; }
    const sheet = ss.getSheetByName(profileDataSheetName);
    if (!sheet) { Logger.log(`[ERROR] RTS_MasterProfileData: Sheet "${profileDataSheetName}" not in SSID "${spreadsheetId}".`); return null; }
    if (sheet.getLastRow() === 0) { Logger.log(`[ERROR] RTS_MasterProfileData: Profile sheet "${profileDataSheetName}" is empty.`); return null; }

    const allDataFromSheet = sheet.getDataRange().getValues();
    if (!allDataFromSheet || allDataFromSheet.length === 0) { Logger.log("[ERROR] RTS_MasterProfileData: No data retrieved."); return null; }
    if (DEBUG) Logger.log(`  RTS_MasterProfileData: Retrieved ${allDataFromSheet.length} total rows from "${profileDataSheetName}".`);

    let currentProcessingSectionTitle = null;
    let accumulatedDataRowsForCurrentSection = [];
    let headersForCurrentTabularSection = [];

    for (let rowIndex = 0; rowIndex < allDataFromSheet.length; rowIndex++) {
      const currentRowArray = allDataFromSheet[rowIndex];
      const firstCellTextClean = (currentRowArray[0] || "").toString().trim();
      let identifiedNewSectionTitle = RTS_Profile_getCanonicalSectionTitle(firstCellTextClean);

      if (identifiedNewSectionTitle) {
        if (currentProcessingSectionTitle && (accumulatedDataRowsForCurrentSection.length > 0 || currentProcessingSectionTitle === "SUMMARY")) {
          RTS_Profile_processAccumulatedSection(profileOutput, currentProcessingSectionTitle, headersForCurrentTabularSection, accumulatedDataRowsForCurrentSection);
        }
        currentProcessingSectionTitle = identifiedNewSectionTitle;
        if (DEBUG) Logger.log(`  RTS_MasterProfileData: Found section header: "${currentProcessingSectionTitle}" (from sheet: "${firstCellTextClean}") at row ${rowIndex + 1}.`);
        accumulatedDataRowsForCurrentSection = []; 
        headersForCurrentTabularSection = []; 

        const sectionConfig = profileStructureConst.find(s => s.title === currentProcessingSectionTitle);
        if (sectionConfig && Array.isArray(sectionConfig.headers) && currentProcessingSectionTitle !== "PERSONAL INFO" && currentProcessingSectionTitle !== "SUMMARY") {
          let actualHeaderRowIndex = rowIndex + 1;
          while (actualHeaderRowIndex < allDataFromSheet.length && RTS_Profile_rowIsEffectivelyBlank(allDataFromSheet[actualHeaderRowIndex])) { actualHeaderRowIndex++; }
          if (actualHeaderRowIndex < allDataFromSheet.length && 
              !RTS_Profile_getCanonicalSectionTitle((allDataFromSheet[actualHeaderRowIndex][0] || "").toString().trim()) &&
              !RTS_Profile_rowIsEffectivelyBlank(allDataFromSheet[actualHeaderRowIndex])) {
            headersForCurrentTabularSection = (allDataFromSheet[actualHeaderRowIndex] || []).map(h => String(h || "").trim()).filter(h => h);
            if (DEBUG) Logger.log(`    Headers for "${currentProcessingSectionTitle}": [${headersForCurrentTabularSection.join(" | ")}] (from sheet row ${actualHeaderRowIndex + 1}).`);
            rowIndex = actualHeaderRowIndex; 
          } else {
            if (DEBUG) Logger.log(`  [WARN] No data headers found for tabular section "${currentProcessingSectionTitle}" below row ${rowIndex + 1}. Will use default headers from PROFILE_STRUCTURE if parsing.`);
          }
        } else if (DEBUG) {
          Logger.log(`    Section "${currentProcessingSectionTitle}" is non-tabular or special (e.g., Personal Info, Summary). No sheet sub-headers read.`);
        }
        continue; 
      }

      if (currentProcessingSectionTitle) {
        if (RTS_Profile_rowIsEffectivelyBlank(currentRowArray)) { if(DEBUG) Logger.log(`    Skipping blank data row ${rowIndex + 1} in "${currentProcessingSectionTitle}".`); continue; }
        if (currentProcessingSectionTitle === "SUMMARY") {
            const summaryLine = currentRowArray.map(cell => String(cell || "").trim()).filter(Boolean).join(" ");
            if(summaryLine) accumulatedDataRowsForCurrentSection.push([summaryLine]);
        } else { 
            accumulatedDataRowsForCurrentSection.push(currentRowArray.map(cell => String(cell || ""))); // Keep raw, trimming done in processAccumulated
        }
      }
    } 

    if (currentProcessingSectionTitle && (accumulatedDataRowsForCurrentSection.length > 0 || currentProcessingSectionTitle === "SUMMARY")) {
      RTS_Profile_processAccumulatedSection(profileOutput, currentProcessingSectionTitle, headersForCurrentTabularSection, accumulatedDataRowsForCurrentSection);
    }
    
    if (DEBUG) { 
        Logger.log("  --- RTS_MasterProfileData: Final Parsed Profile Output (End of RTS_getMasterProfileData) ---");
        Logger.log(`  PersonalInfo FullName: ${profileOutput.personalInfo.fullName || 'NOT PARSED! Check MasterProfile Sheet and parser logic.'}`);
        if(Object.keys(profileOutput.personalInfo).length > 0) Logger.log(`    Full PersonalInfo Object: ${JSON.stringify(profileOutput.personalInfo)}`); else Logger.log(`    PersonalInfo Object: Is Empty.`);
        Logger.log(`  Summary (Starts with): "${profileOutput.summary.substring(0, 100)}..." (Total Summary Length: ${profileOutput.summary.length})`);
        Logger.log(`  Total Structured Sections in profileOutput.sections: ${profileOutput.sections.length}`);
        profileOutput.sections.forEach(s => {
            let itemsCount = 0; if (s.items) itemsCount = s.items.length; else if (s.subsections) itemsCount = s.subsections.reduce((sum, sub) => sum + (sub.items ? sub.items.length : 0), 0);
            Logger.log(`    Output Section: ${s.title}, Effective Items Found: ${itemsCount}`);
        });
        Logger.log("  --- End of Parsed Profile Output Preview ---");
    }
    Logger.log(`[SUCCESS] RTS_MasterProfileData: Finished loading and parsing from "${profileDataSheetName}". Candidate: ${profileOutput.personalInfo.fullName || "FullName NOT PARSED!"}`);
    return profileOutput;

  } catch (e) {
    Logger.log(`[CRITICAL EXCEPTION] RTS_MasterProfileData: While processing "${profileDataSheetName}". Error: ${e.message}\nStack: ${e.stack}`);
    return null;
  }
}

/**
 * Processes accumulated raw data for a specific profile section.
 */
function RTS_Profile_processAccumulatedSection(profileOutput, sectionTitleKey, sectionHeadersFromSheet, dataRows) {
  const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : true);
  const NUM_BULLETS = (typeof NUM_DEDICATED_PROFILE_BULLET_COLUMNS !== 'undefined' ? NUM_DEDICATED_PROFILE_BULLET_COLUMNS : 3);
  const profileStructureConst = (typeof PROFILE_STRUCTURE !== 'undefined' ? PROFILE_STRUCTURE : []);

  if(DEBUG) Logger.log(`    RTS_Profile_processAccumulatedSection: START - Section: "${sectionTitleKey}". Headers from Sheet: [${sectionHeadersFromSheet ? sectionHeadersFromSheet.join(" | ") : 'NONE'}]. Data Rows: ${dataRows ? dataRows.length : '0'}`);

  if (sectionTitleKey === "PERSONAL INFO") {
    if(DEBUG) Logger.log("      Processing PERSONAL INFO as Key-Value pairs from dataRows...");
    profileOutput.personalInfo = {}; // Initialize
    dataRows.forEach((rowArray, rIdx) => {
      let keyFromCell = String(rowArray[0] || "").trim();
      const valueFromCell = String(rowArray[1] || "").trim();
      if(DEBUG) Logger.log(`        PI DataRow ${rIdx + 1}: KeyOnSheet="${keyFromCell}", ValueOnSheet="${valueFromCell}"`);
      if (!keyFromCell || !valueFromCell) { if(DEBUG) Logger.log(`          Skipping PI Row ${rIdx + 1}: empty key or value.`); return; }

      let normalizedKey = keyFromCell.toLowerCase().replace(/\s+|url/gi, ''); // Normalize: lowercase, no spaces, no "url"
      let mapped = false;
      if (normalizedKey.includes("fullname") || normalizedKey.includes("name")) { profileOutput.personalInfo.fullName = valueFromCell; mapped=true;}
      else if (normalizedKey.includes("linkedin")) { profileOutput.personalInfo.linkedin = valueFromCell; mapped=true;}
      else if (normalizedKey.includes("github")) { profileOutput.personalInfo.github = valueFromCell; mapped=true;}
      else if (normalizedKey.includes("portfolio")||normalizedKey.includes("website")) { profileOutput.personalInfo.portfolio = valueFromCell; mapped=true;}
      else if (normalizedKey.includes("phone") || normalizedKey.includes("contactnumber")) { profileOutput.personalInfo.phone = valueFromCell; mapped=true;}
      else if (normalizedKey.includes("email") || normalizedKey.includes("e-mail")) { profileOutput.personalInfo.email = valueFromCell; mapped=true;}
      else if (normalizedKey.includes("location") || normalizedKey.includes("address") || normalizedKey.includes("city")) { profileOutput.personalInfo.location = valueFromCell; mapped=true;}
      else { 
        let camelKey = keyFromCell.replace(/[^a-zA-Z0-9\s_]+/g, '').replace(/[\s_]+(.)/g, (_match, char) => char.toUpperCase());
        if (camelKey && camelKey.trim() !== "") {
          camelKey = camelKey.charAt(0).toLowerCase() + camelKey.slice(1);
          profileOutput.personalInfo[camelKey] = valueFromCell; mapped=true;
          if(DEBUG) Logger.log(`          Fallback PI mapping: .personalInfo.${camelKey} = "${valueFromCell}" (Original Key: "${keyFromCell}")`);
        } else { if(DEBUG) Logger.log(`          PI Key "${keyFromCell}" (Row ${rIdx+1}) resulted in invalid camelCase key, SKIPPED.`);}
      }
       if(DEBUG && mapped) Logger.log(`          Processed PI Row ${rIdx+1}: ("${keyFromCell}") successfully mapped.`);
    });
    if(DEBUG) Logger.log(`      FINAL profileOutput.personalInfo after processing all rows: ${JSON.stringify(profileOutput.personalInfo)}`);
    return;
  }

  if (sectionTitleKey === "SUMMARY") {
    profileOutput.summary = dataRows.map(rowArray => String(rowArray[0] || "").trim()).filter(Boolean).join("\n\n").trim();
    if(DEBUG) Logger.log(`      SUMMARY processed. Final Length: ${profileOutput.summary.length}. Content starts: "${profileOutput.summary.substring(0,70)}..."`);
    return;
  }
  
  const sectionConfig = profileStructureConst.find(s => s.title === sectionTitleKey);
  if (!sectionConfig || !Array.isArray(sectionConfig.headers) || dataRows.length === 0) {
    if(DEBUG) Logger.log(`    Section "${sectionTitleKey}" has no data rows or no header config in PROFILE_STRUCTURE. Will ensure empty section structure in output object.`);
    const existingS = profileOutput.sections.find(s => s.title === sectionTitleKey);
    if (!existingS) {
      if(sectionTitleKey === "TECHNICAL SKILLS & CERTIFICATES" || sectionTitleKey === "PROJECTS") profileOutput.sections.push({title:sectionTitleKey, subsections:[]});
      else profileOutput.sections.push({title:sectionTitleKey, items:[]});
    }
    return;
  }

  const headersToUse = (sectionHeadersFromSheet && sectionHeadersFromSheet.length > 0 && sectionHeadersFromSheet.length >= sectionConfig.headers.length) ? 
                       sectionHeadersFromSheet : sectionConfig.headers;
  if (!headersToUse || headersToUse.length === 0) {
      if(DEBUG) Logger.log(`    [CRITICAL WARN] For section "${sectionTitleKey}", no effective headers could be determined (neither from sheet nor from PROFILE_STRUCTURE). Cannot accurately map data items.`);
      return;
  }
  if(DEBUG) Logger.log(`    Processing "${sectionTitleKey}" tabular data using headers: [${headersToUse.join(" | ")}]`);

  const parsedItems = dataRows.map(rowArray => {
    let item = {}; headersToUse.forEach((header,idx)=>{if(header?.trim()!=="" && idx<rowArray.length)item[header.trim()]=String(rowArray[idx]||"").trim();}); return item;
  });

  const finalProcessedItems = parsedItems.map(itemFS => {
    const stdItem = {};
    for(const sheetHdr in itemFS){let jsK=sheetHdr.replace(/[^a-zA-Z0-9\s_]+/g,'').replace(/[\s_]+(.)/g, (m,c)=>c.toUpperCase());jsK=jsK.charAt(0).toLowerCase()+jsK.slice(1);
      const map={"jobtitle":"jobTitle","company":"company","startdate":"startDate","enddate":"endDate","institution":"institution","degree":"degree","relevantcoursework":"relevantCoursework","categoryname":"categoryName","skillitem":"skill","issuedate":"issueDate","projectname":"projectName","githubname1":"githubName1","githuburl1":"githubUrl1","awardname":"awardName","rolename":"role","technologies":"technologies"};
      if(jsK.match(/^(responsibility|descriptionBullet)\d+$/i))jsK=jsK.charAt(0).toLowerCase()+jsK.slice(1); else if(map[jsK.toLowerCase()])jsK=map[jsK.toLowerCase()]; else if (map[sheetHdr]) jsKey = map[sheetHdr]; // Check original header if camelCase didn't match map
      stdItem[jsK]=itemFS[sheetHdr];
    }
    if(stdItem.startDate)stdItem.startDate=RTS_Profile_formatDateFromSheet(stdItem.startDate); if(stdItem.endDate)stdItem.endDate=RTS_Profile_formatDateFromSheet(stdItem.endDate,true);
    if(stdItem.date && sectionTitleKey === "HONORS & AWARDS") stdItem.date = RTS_Profile_formatDateFromSheet(stdItem.date); 
    if(stdItem.issueDate && sectionTitleKey === "TECHNICAL SKILLS & CERTIFICATES") stdItem.issueDate = RTS_Profile_formatDateFromSheet(stdItem.issueDate);
    if(sectionTitleKey==="EXPERIENCE"||sectionTitleKey==="LEADERSHIP & UNIVERSITY INVOLVEMENT"){stdItem.responsibilities=[];for(let k=1;k<=NUM_BULLETS;k++){const bk=`responsibility${k}`;if(stdItem[bk])stdItem.responsibilities.push(stdItem[bk]);delete stdItem[bk];}}
    else if(sectionTitleKey==="PROJECTS"){stdItem.descriptionBullets=[];for(let k=1;k<=NUM_BULLETS;k++){const bk=`descriptionBullet${k}`;if(stdItem[bk])stdItem.descriptionBullets.push(stdItem[bk]);delete stdItem[bk];} if(typeof stdItem.technologies==='string')stdItem.technologies=RTS_Profile_smartSplitMultiLine(stdItem.technologies);else if(!Array.isArray(stdItem.technologies))stdItem.technologies=[]; stdItem.githubLinks=[];if(stdItem.githubName1&&stdItem.githubUrl1)stdItem.githubLinks.push({name:stdItem.githubName1,url:stdItem.githubUrl1});delete stdItem.githubName1;delete stdItem.githubUrl1;}
    if(sectionTitleKey==="EDUCATION" && typeof stdItem.relevantCoursework==='string')stdItem.relevantCoursework=RTS_Profile_smartSplitMultiLine(stdItem.relevantCoursework); else if(sectionTitleKey==="EDUCATION" && !Array.isArray(stdItem.relevantCoursework))stdItem.relevantCoursework=[];
    return stdItem;
  });

  let targetSect=profileOutput.sections.find(s=>s.title===sectionTitleKey);
  if(!targetSect){if(sectionTitleKey==="TECHNICAL SKILLS & CERTIFICATES"||sectionTitleKey==="PROJECTS")targetSect={title:sectionTitleKey,subsections:[]};else targetSect={title:sectionTitleKey,items:[]};profileOutput.sections.push(targetSect);}
  if(sectionTitleKey==="TECHNICAL SKILLS & CERTIFICATES"){const subsM=new Map();finalProcessedItems.forEach(it=>{const catN=it.categoryName||"Uncategorized";if(!subsM.has(catN))subsM.set(catN,{name:catN,items:[]});const skN=it.skill||it.name||it.skillItem;const det=it.details||"";if(it.issuer||it.issueDate)subsM.get(catN).items.push({name:skN,issuer:it.issuer||"",issueDate:it.issueDate||"",details:det});else if(skN)subsM.get(catN).items.push({skill:skN,details:det});});targetSect.subsections=Array.from(subsM.values());}
  else if(sectionTitleKey==="PROJECTS"){if(targetSect.subsections.length===0)targetSect.subsections.push({name:"General Projects",items:[]});targetSect.subsections[0].items.push(...finalProcessedItems);}
  else{targetSect.items.push(...finalProcessedItems);}
  if(DEBUG&&finalProcessedItems.length>0)Logger.log(`    RTS_Profile_processAccumulatedSection: "${sectionTitleKey}" processed. Effective items added/mapped: ${finalProcessedItems.length}.`);
  else if(DEBUG)Logger.log(`    RTS_Profile_processAccumulatedSection: "${sectionTitleKey}" processed. No valid items were mapped from data or headers issue.`);
}
