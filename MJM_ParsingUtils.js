// File: MJM_ParsingUtils.gs (or MRM_ParsingUtils.gs)
// Description: Contains functions dedicated to parsing email content (subject, body, sender)
// for the MJM Application Tracker module using regular expressions and keyword matching
// as a fallback to AI parsing.
// Relies on constants from MJM_Config.gs (for MJM-specific keywords/settings)
// and Global_Constants.gs (for shared values like MANUAL_REVIEW_NEEDED_TEXT, GLOBAL_DEBUG_MODE).

/**
 * Attempts to parse a company name from the sender's email domain.
 * Ignores common ATS and generic domains defined in MJM_Config.gs.
 * @param {string} sender The raw "From" string of the email.
 * @return {string|null} The parsed company name or null if not determinable.
 */
function MJM_parseCompanyFromDomain(sender) {
  const emailMatch = sender.match(/<([^>]+)>/);
  if (!emailMatch || !emailMatch[1]) return null;

  const emailAddress = emailMatch[1];
  const domainParts = emailAddress.split('@');
  if (domainParts.length !== 2) return null;

  let domain = domainParts[1].toLowerCase();

  // MJM_IGNORED_DOMAINS_FOR_COMPANY_PARSE from MJM_Config.gs
  if (MJM_IGNORED_DOMAINS_FOR_COMPANY_PARSE.has(domain) && !domain.includes('wellfound.com')) {
    // Exception for wellfound which can sometimes be the actual company for direct job posts.
    return null;
  }

  // Remove common prefixes and TLDs
  domain = domain.replace(/^(?:careers|jobs|recruiting|apply|hr|talent|notification|notifications|team|hello|no-reply|noreply)[.-]?/i, '');
  domain = domain.replace(/\.(com|org|net|io|co|ai|dev|xyz|tech|ca|uk|de|fr|app|eu|us|info|biz|work|agency|careers|招聘|group|global|inc|llc|ltd|corp|gmbh)$/i, '');
  domain = domain.replace(/[^a-z0-9]+/gi, ' '); // Replace non-alphanumeric with space
  domain = domain.split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' '); // Capitalize each word

  return domain.trim() || null;
}

/**
 * Attempts to parse a company name from the sender's display name part of the "From" string.
 * Cleans common ATS/platform noise and generic terms.
 * @param {string} sender The raw "From" string of the email.
 * @return {string|null} The parsed company name or null.
 */
function MJM_parseCompanyFromSenderName(sender) {
  const nameMatch = sender.match(/^"?(.*?)"?\s*</);
  let name = nameMatch ? nameMatch[1].trim() : sender.split('<')[0].trim();

  if (!name || name.includes('@') || name.length < 2) return null;

  // Remove common ATS/platform noise - could expand this list in MJM_Config.gs if needed
  name = name.replace(/\|\s*(?:greenhouse|lever|wellfound|workday|ashby|icims|smartrecruiters|taleo|bamboohr|recruiterbox|jazzhr|workable|breezyhr|notion)\b/i, '');
  name = name.replace(/\s*(?:via Wellfound|via LinkedIn|via Indeed|from Greenhouse|from Lever|Careers at|Hiring at)\b/gi, '');
  // Remove common generic terms
  name = name.replace(/\s*(?:Careers|Recruiting|Recruitment|Hiring Team|Hiring|Talent Acquisition|Talent|HR|Team|Notifications?|Jobs?|Updates?|Apply|Notification|Hello|No-?Reply|Support|Info|Admin|Department)\b/gi, '');
  // Remove trailing legal entities, punctuation, and trim
  name = name.replace(/[|,_.\s]+(?:Inc\.?|LLC\.?|Ltd\.?|Corp\.?|GmbH|Solutions|Services|Group|Global|Technologies|Labs|Studio|Ventures)?$/i, '').trim();
  name = name.replace(/^(?:The|A)\s+/i, '').trim(); // Remove leading "The", "A"

  if (name.length > 1 && !/^(?:noreply|no-reply|jobs|careers|support|info|admin|hr|talent|recruiting|team|hello)$/i.test(name.toLowerCase())) {
    return name;
  }
  return null;
}

/**
 * Extracts company and job title from an email message using regex patterns and sender info.
 * This is the fallback when AI (Gemini) parsing is not used or fails for MJM Application Tracker.
 *
 * @param {GoogleAppsScript.Gmail.GmailMessage} message The Gmail message object.
 * @param {string} detectedPlatform The platform detected (e.g., "LinkedIn", "Indeed"). Uses MJM_APP_DEFAULT_PLATFORM from MJM_Config.gs.
 * @param {string} emailSubject The subject of the email.
 * @param {string} plainBody The plain text body of the email.
 * @return {{company: string, title: string}} An object containing the extracted company and title.
 *         Defaults to MANUAL_REVIEW_NEEDED_TEXT (global const) if extraction fails.
 */
function MJM_extractCompanyAndTitle(message, detectedPlatform, emailSubject, plainBody) {
  // MANUAL_REVIEW_NEEDED_TEXT from Global_Constants.gs
  // MJM_APP_DEFAULT_PLATFORM from MJM_Config.gs
  let company = MANUAL_REVIEW_NEEDED_TEXT;
  let title = MANUAL_REVIEW_NEEDED_TEXT;
  const sender = message.getFrom();

  if (GLOBAL_DEBUG_MODE) Logger.log(`[DEBUG] MJM_ParsingUtils (Regex C/T Fallback) for subj: "${emailSubject.substring(0,100)}"`);

  let tempCompanyFromDomain = MJM_parseCompanyFromDomain(sender);
  let tempCompanyFromName = MJM_parseCompanyFromSenderName(sender);
  if (GLOBAL_DEBUG_MODE) Logger.log(`  Sender Parse -> Name Extracted: "${tempCompanyFromName}", Domain Extracted: "${tempCompanyFromDomain}"`);

  // Platform-specific logic (e.g., for Wellfound)
  // MJM_PLATFORM_DOMAIN_KEYWORDS could be used here if "Wellfound" is a key.
  if (detectedPlatform === "Wellfound" && plainBody) { // "Wellfound" is a string literal here, assuming MJM_PLATFORM_DOMAIN_KEYWORDS values match
    let wfCoSub = emailSubject.match(/update from (.*?)(?: \|| at |$)/i) ||
                  emailSubject.match(/application to (.*?)(?: successfully| at |$)/i) ||
                  emailSubject.match(/New introduction from (.*?)(?: for |$)/i);
    if (wfCoSub && wfCoSub[1]) company = wfCoSub[1].trim();

    if (title === MANUAL_REVIEW_NEEDED_TEXT && plainBody && sender.toLowerCase().includes("team@hi.wellfound.com")) {
        const markerPhrase = "if there's a match, we will make an email introduction."; // Specific to some Wellfound emails
        const markerIndex = plainBody.toLowerCase().indexOf(markerPhrase);
        if (markerIndex !== -1) {
            const relevantText = plainBody.substring(markerIndex + markerPhrase.length);
            const titleMatch = relevantText.match(/^\s*\*\s*([A-Za-z\s.,:&'\/-]+?)(?:\s*\(| at | \n|$)/m);
            if (titleMatch && titleMatch[1]) title = titleMatch[1].trim();
        }
    }
  }

  // General Regex patterns for subject line parsing
  const subjectParsePatterns = [
    { r: /Application for(?: the)?\s+(.+?)\s+at\s+([^-:|–—]+)/i, tI: 1, cI: 2 },
    { r: /Invite(?:.*?)(?:to|for)(?: an)? interview(?:.*?)\sfor\s+(?:the\s)?(.+?)(?:\s+at\s+([^-:|–—]+))?/i, tI: 1, cI: 2 },
    { r: /Your application for(?: the)?\s+(.+?)\s+at\s+([^-:|–—]+)/i, tI: 1, cI: 2 },
    { r: /Regarding your application for\s+(.+?)(?:\s-\s(.*?))?(?:\s@\s(.*?))?$/i, tI: 1, cI: 3, cI2: 2}, // Greenhouse
    { r: /^(?:Update on|Your Application to|Thank you for applying to)\s+([^-:|–—]+?)(?:\s*-\s*([^-:|–—]+))?$/i, cI: 1, tI: 2 }, // Lever style
    { r: /applying to\s+(.+?)\s+at\s+([^-:|–—]+)/i, tI: 1, cI: 2 },
    { r: /interest in the\s+(.+?)\s+role(?:\s+at\s+([^-:|–—]+))?/i, tI: 1, cI: 2 },
    { r: /update on your\s+(.+?)\s+app(?:lication)?(?:\s+at\s+([^-:|–—]+))?/i, tI: 1, cI: 2 }
  ];

  for (const patternInfo of subjectParsePatterns) {
    let match = emailSubject.match(patternInfo.r);
    if (match) {
      if (GLOBAL_DEBUG_MODE) Logger.log(`  Matched subject pattern: ${patternInfo.r}`);
      let extractedTitle = patternInfo.tI > 0 && match[patternInfo.tI] ? match[patternInfo.tI].trim() : null;
      let extractedCompany = patternInfo.cI > 0 && match[patternInfo.cI] ? match[patternInfo.cI].trim() : null;
      if (!extractedCompany && patternInfo.cI2 > 0 && match[patternInfo.cI2]) extractedCompany = match[patternInfo.cI2].trim();

      // Heuristic for Lever-style ambiguous subjects (Company - Title or Title - Company)
      if (patternInfo.cI === 1 && patternInfo.tI === 2 && extractedCompany && extractedTitle) {
          if (/\b(engineer|manager|analyst|developer|specialist|lead|director|coordinator|architect|consultant|designer|recruiter|associate|intern)\b/i.test(extractedCompany) &&
             !/\b(engineer|manager|analyst|developer|specialist|lead|director|coordinator|architect|consultant|designer|recruiter|associate|intern)\b/i.test(extractedTitle)) {
              [extractedCompany, extractedTitle] = [extractedTitle, extractedCompany]; // Swap
              if (GLOBAL_DEBUG_MODE) Logger.log(`    Swapped Company/Title based on keywords. New C: ${extractedCompany}, New T: ${extractedTitle}`);
          }
      }
      
      if (extractedTitle && (title === MANUAL_REVIEW_NEEDED_TEXT || title === MJM_APP_DEFAULT_STATUS)) title = extractedTitle; // MJM_APP_DEFAULT_STATUS from MJM_Config.gs
      if (extractedCompany && (company === MANUAL_REVIEW_NEEDED_TEXT || company === MJM_APP_DEFAULT_PLATFORM)) company = extractedCompany; // MJM_APP_DEFAULT_PLATFORM from MJM_Config.gs
      
      // If both found satisfactory values, break early
      if (company !== MANUAL_REVIEW_NEEDED_TEXT && title !== MANUAL_REVIEW_NEEDED_TEXT &&
          company !== MJM_APP_DEFAULT_PLATFORM && title !== MJM_APP_DEFAULT_STATUS) break;
    }
  }

  // If subject parsing failed for company, use sender name parse
  if (company === MANUAL_REVIEW_NEEDED_TEXT && tempCompanyFromName) company = tempCompanyFromName;

  // Body Scan Fallback if still needed
  if ((company === MANUAL_REVIEW_NEEDED_TEXT || title === MANUAL_REVIEW_NEEDED_TEXT || company === MJM_APP_DEFAULT_PLATFORM || title === MJM_APP_DEFAULT_STATUS) && plainBody) {
    const bodyFirstKChars = plainBody.substring(0, 1500).replace(/<[^>]+>/g, ' '); // Look in first 1500 chars, strip HTML
    if (company === MANUAL_REVIEW_NEEDED_TEXT || company === MJM_APP_DEFAULT_PLATFORM) {
      let bodyCompanyMatch = bodyFirstKChars.match(/(?:applying to|application with|interview with|position at|role at|opportunity at|Thank you for your interest in working at)\s+([A-Z][A-Za-z\s.&'-]+(?:LLC|Inc\.?|Ltd\.?|Corp\.?|GmbH|Group|Solutions|Technologies)?)(?:[.,\s\n\(]|$)/i);
      if (bodyCompanyMatch && bodyCompanyMatch[1]) company = bodyCompanyMatch[1].trim();
    }
    if (title === MANUAL_REVIEW_NEEDED_TEXT || title === MJM_APP_DEFAULT_STATUS) {
      let bodyTitleMatch = bodyFirstKChars.match(/(?:application for the|position of|role of|applying for the|interview for the|title:)\s+([A-Za-z][A-Za-z0-9\s.,:&'\/\(\)-]+?)(?:\s\(| at | with |[\s.,\n\(]|$)/i);
      if (bodyTitleMatch && bodyTitleMatch[1]) title = bodyTitleMatch[1].trim();
    }
  }

  // Last resort for company: use domain parse if other methods failed
  if (company === MANUAL_REVIEW_NEEDED_TEXT && tempCompanyFromDomain) company = tempCompanyFromDomain;

  // Cleaning function for extracted entities
  const cleanEntity = (entityText, isTitleField = false) => {
    if (!entityText || entityText === MANUAL_REVIEW_NEEDED_TEXT || entityText === MJM_APP_DEFAULT_STATUS || entityText === MJM_APP_DEFAULT_PLATFORM || entityText.toLowerCase() === "n/a") return MANUAL_REVIEW_NEEDED_TEXT;
    let cleaned = entityText.split(/[\n\r#(]| - /)[0]; // Take first line, remove text after # or some " - " patterns
    cleaned = cleaned.replace(/ (?:inc|llc|ltd|corp|gmbh)[\.,]?$/i, '').replace(/[,"']?$/, ''); // Remove legal suffixes, trailing punctuation
    cleaned = cleaned.replace(/^(?:The|A)\s+/i, ''); // Remove leading "The", "A"
    if (isTitleField) { // Specific cleaning for job titles
        cleaned = cleaned.replace(/JR\d+\s*[-–—]?\s*/i, ''); // Remove JR requisition numbers often in titles
        cleaned = cleaned.replace(/\(Senior\)/i, 'Senior'); // Hoist (Senior) out of parentheses
        // Remove common parenthetical additions that aren't part of the core title
        cleaned = cleaned.replace(/\(.*?(?:remote|hybrid|onsite|contract|part-time|full-time|intern|co-op|stipend|urgent|hiring|opening|various locations).*?\)/gi, '');
        cleaned = cleaned.replace(/[-–—:]\s*(?:remote|hybrid|onsite|contract|part-time|full-time|intern|co-op|various locations)\s*$/gi, '');
    }
    cleaned = cleaned.replace(/[\u2018\u2019\u201A\u201B\u2032\u2035]/g, "'").replace(/[\u201C\u201D\u201E\u201F\u2033\u2036]/g, '"'); // Normalize smart quotes
    cleaned = cleaned.replace(/&/gi, '&').replace(/ /gi, ' '); // Normalize & and non-breaking space
    cleaned = cleaned.replace(/\s+/g, ' ').trim(); // Normalize whitespace
    cleaned = cleaned.replace(/^[-\s#*.,]+|[,\s]+$/g, ''); // Clean leading/trailing common list/punctuation chars and spaces again
    return cleaned.length < 2 ? MANUAL_REVIEW_NEEDED_TEXT : cleaned; // If cleaning results in too short a string
  };

  company = cleanEntity(company);
  title = cleanEntity(title, true); // Pass true for title-specific cleaning

  if (GLOBAL_DEBUG_MODE) Logger.log(`  MJM_ParsingUtils Final Fallback Regex Result -> Company:"${company}", Title:"${title}"`);
  return {company: company, title: title};
}

/**
 * Parses the email body for status keywords specific to MJM Application Tracker.
 * Uses keyword arrays from MJM_Config.gs.
 *
 * @param {string} plainBody The plain text body of the email.
 * @return {string|null} The detected status (e.g., MJM_APP_REJECTED_STATUS) or null if no specific keywords found.
 */
function MJM_parseBodyForStatus(plainBody) {
  if (!plainBody || plainBody.length < 10) {
    if (GLOBAL_DEBUG_MODE) Logger.log("[DEBUG] MJM_ParsingUtils (Regex Status): Body too short/missing for status parse.");
    return null;
  }
  // Normalize body text for keyword matching
  let bodyLower = plainBody.toLowerCase().replace(/[.,!?;:()\[\]{}'"“”‘’\-–—]/g, ' ').replace(/\s+/g, ' ').trim();

  // Keyword arrays from MJM_Config.gs (e.g., MJM_OFFER_KEYWORDS, MJM_INTERVIEW_KEYWORDS, etc.)
  if (MJM_OFFER_KEYWORDS.some(k => bodyLower.includes(k))) {
    if (GLOBAL_DEBUG_MODE) Logger.log(`  Regex Status: Matched OFFER.`); return MJM_APP_OFFER_STATUS;
  }
  if (MJM_INTERVIEW_KEYWORDS.some(k => bodyLower.includes(k))) {
    if (GLOBAL_DEBUG_MODE) Logger.log(`  Regex Status: Matched INTERVIEW.`); return MJM_APP_INTERVIEW_STATUS;
  }
  if (MJM_ASSESSMENT_KEYWORDS.some(k => bodyLower.includes(k))) {
    if (GLOBAL_DEBUG_MODE) Logger.log(`  Regex Status: Matched ASSESSMENT.`); return MJM_APP_ASSESSMENT_STATUS;
  }
  if (MJM_APP_VIEWED_KEYWORDS.some(k => bodyLower.includes(k))) {
    if (GLOBAL_DEBUG_MODE) Logger.log(`  Regex Status: Matched APP_VIEWED.`); return MJM_APP_VIEWED_STATUS;
  }
  if (MJM_REJECTION_KEYWORDS.some(k => bodyLower.includes(k))) {
    if (GLOBAL_DEBUG_MODE) Logger.log(`  Regex Status: Matched REJECTION.`); return MJM_APP_REJECTED_STATUS;
  }

  if (GLOBAL_DEBUG_MODE) Logger.log("[DEBUG] MJM_ParsingUtils (Regex Status): No specific MJM status keywords found by regex.");
  return null; // Return null if no strong match, implying MJM_APP_DEFAULT_STATUS might be used by caller
}
