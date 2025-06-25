// File: RTS_DocumentService.gs
// Description: Responsible for creating and formatting the final resume Google Document
// for the RTS module. Uses a template document and populates it with data.
// Relies on RESUME_TEMPLATE_DOC_ID and GLOBAL_DEBUG_MODE from Global_Constants.gs.

// --- UTILITY FUNCTIONS (Specific to this Document Service) ---
if (typeof RegExp.escape !== 'function') {RegExp.escape = s => typeof s==='string'?s.replace(/[-\/\\^$*+?.()|[\]{}]/g,'\\$&'):'';}

function RTS_Doc_findAndReplaceText(body, placeholder, textToInsert) {
  const searchPattern = RegExp.escape(placeholder);
  const effectiveText = (textToInsert === null || textToInsert === undefined) ? "" : String(textToInsert);
  // GLOBAL_DEBUG_MODE from Global_Constants.gs
  const DEBUG = (typeof GLOBAL_DEBUG_MODE !== 'undefined' ? GLOBAL_DEBUG_MODE : false);
  if(DEBUG && placeholder !== "{SUMMARY_CONTENT}") Logger.log(`    RTS_Doc_findAndReplaceText: Replacing "${placeholder}" with text of len ${effectiveText.length}.`);
  body.replaceText(searchPattern, effectiveText);
}

function RTS_Doc_sanitizeBaseAttributes(rawAttrs) {
  if (!rawAttrs) return {}; const sanitized = {};
  const allowed = [DocumentApp.Attribute.FONT_FAMILY, DocumentApp.Attribute.FONT_SIZE, DocumentApp.Attribute.FOREGROUND_COLOR,
    DocumentApp.Attribute.BACKGROUND_COLOR, DocumentApp.Attribute.BOLD, DocumentApp.Attribute.ITALIC, DocumentApp.Attribute.UNDERLINE,
    DocumentApp.Attribute.STRIKETHROUGH, DocumentApp.Attribute.LINE_SPACING, DocumentApp.Attribute.SPACING_BEFORE,
    DocumentApp.Attribute.SPACING_AFTER, DocumentApp.Attribute.INDENT_START, DocumentApp.Attribute.INDENT_END,
    DocumentApp.Attribute.HORIZONTAL_ALIGNMENT];
  for (const k in rawAttrs) { const eK = allowed.find(a=>a.toString()===k); if(eK){let v=rawAttrs[k]; if(eK===DocumentApp.Attribute.HORIZONTAL_ALIGNMENT&&typeof v==='string'){v=DocumentApp.HorizontalAlignment[v.toUpperCase()]||DocumentApp.HorizontalAlignment.LEFT;} if(v!==null)sanitized[eK]=v;}}
  return sanitized;
}

function RTS_Doc_copyAttributesPreservingEnums(originalAttrs) {
  const newAttrs = {}; if (!originalAttrs) return newAttrs; for (const key in originalAttrs) newAttrs[key] = originalAttrs[key]; return newAttrs;
}

function RTS_Doc_populateBlockPlaceholder(body, placeholder, itemsData, renderItemFunction) {
  const DEBUG = GLOBAL_DEBUG_MODE;
  if(DEBUG) Logger.log(`  RTS_Doc_populateBlock: Placeholder "${placeholder}", items: ${itemsData?.length||0}`);
  const range = body.findText(RegExp.escape(placeholder)); if(!range){Logger.log(`[WARN] Placeholder "${placeholder}" NOT FOUND.`); return;}
  let el=range.getElement(); if(el.getType()!==DocumentApp.ElementType.PARAGRAPH){el=el.getParent(); if(el.getType()!==DocumentApp.ElementType.PARAGRAPH){Logger.log(`[ERROR] Placeholder "${placeholder}" not in Para.`); return;}}
  const para=el.asParagraph(); const baseAttrs=RTS_Doc_sanitizeBaseAttributes(para.getAttributes()); para.clear().setAttributes({[DocumentApp.Attribute.SPACING_AFTER]:0, [DocumentApp.Attribute.SPACING_BEFORE]:0}); // Reset spacing of placeholder para
  let currentEl = para; // Start inserting after the (now empty) placeholder paragraph
  if(itemsData?.length>0){ if(DEBUG)Logger.log(`    Rendering ${itemsData.length} items for "${placeholder}"...`);
    for(let i=0;i<itemsData.length;i++){currentEl=renderItemFunction(itemsData[i],body,currentEl,baseAttrs,i); if(!currentEl)break;}}
  else if(DEBUG)Logger.log(`    No items for "${placeholder}".`);
}

// --- INDIVIDUAL SECTION/ITEM RENDER FUNCTIONS (RTS_Doc_ prefixed) ---
function RTS_Doc_renderEducationItem(edu, body, insertAfterElement, baseAttrs, index) {
  const DEBUG = GLOBAL_DEBUG_MODE; 
  if(DEBUG) Logger.log(`      renderEducationItem: Index ${index}, Inst: "${edu.institution || 'N/A'}"`);
  let currentLastElement = insertAfterElement; 
  let insertionIndex = body.getChildIndex(currentLastElement) + 1;
  const itemAttrs = RTS_Doc_copyAttributesPreservingEnums(baseAttrs);

  if (index > 0) {
    const spacingPara = body.insertParagraph(insertionIndex++, "");
    spacingPara.setAttributes({[DocumentApp.Attribute.SPACING_BEFORE]: 6, [DocumentApp.Attribute.SPACING_AFTER]: 0, [DocumentApp.Attribute.LINE_SPACING]: 0.8});
    currentLastElement = spacingPara;
  }
  
  const lineAttrs = RTS_Doc_copyAttributesPreservingEnums(itemAttrs);
  lineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = (index === 0 && !(itemAttrs[DocumentApp.Attribute.SPACING_BEFORE] > 0) ) ? 0 : 1; 
  lineAttrs[DocumentApp.Attribute.SPACING_AFTER] = 1;
  lineAttrs[DocumentApp.Attribute.LINE_SPACING] = itemAttrs[DocumentApp.Attribute.LINE_SPACING] || 1.0; 

  if (edu.institution && String(edu.institution).trim()) {
    const instPara = body.insertParagraph(insertionIndex++, String(edu.institution).trim());
    instPara.setAttributes(RTS_Doc_copyAttributesPreservingEnums(lineAttrs)); 
    instPara.setBold(true);
    currentLastElement = instPara;
  }

  let dateText = String(edu.endDate || "").trim();
  const startDateText = String(edu.startDate || "").trim();
  if (dateText.toUpperCase() !== "PRESENT" && startDateText) dateText = `${startDateText} – ${dateText}`;
  else if (startDateText && !String(edu.endDate || "").trim()) dateText = `${startDateText} – Present`;
  else if (!startDateText && dateText.toUpperCase() === "PRESENT") dateText = `Dates N/A → Present`; 
  else if (!startDateText && !String(edu.endDate || "").trim()) dateText = ""; 
  
  if (dateText) {
    const datePara = body.insertParagraph(insertionIndex++, dateText);
    const dateLineAttrs = RTS_Doc_copyAttributesPreservingEnums(lineAttrs);
    dateLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 0; 
    datePara.setAttributes(dateLineAttrs).setBold(false).setItalic(true);
    currentLastElement = datePara;
  }

  // Line 3: Degree [, Location (Location Italic)] - CORRECTED LOGIC
  const degreeTextClean = String(edu.degree || "").trim();
  const locationTextClean = String(edu.location || "").trim();
  if (degreeTextClean || locationTextClean) {
    const degreeLocPara = body.insertParagraph(insertionIndex++, ""); // Create empty paragraph
    const degreeLineAttrs = RTS_Doc_copyAttributesPreservingEnums(lineAttrs);
    degreeLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = (edu.institution || dateText) ? 1 : 0;
    degreeLocPara.setAttributes(degreeLineAttrs).setBold(false).setItalic(false);

    if (degreeTextClean) {
      degreeLocPara.appendText(degreeTextClean);
    }
    if (locationTextClean) {
      const prefix = degreeTextClean ? ", " : "";
      degreeLocPara.appendText(prefix); // Append separator first
      const startItalicIndex = degreeLocPara.getText().length; // Get length BEFORE appending location
      degreeLocPara.appendText(locationTextClean);             // Append location
      degreeLocPara.editAsText().setItalic(startItalicIndex, degreeLocPara.getText().length - 1, true); // Italicize only the location part
    }
    currentLastElement = degreeLocPara;
  }
  
  if (edu.gpa && String(edu.gpa).trim()) {
    const gpaPara = body.insertParagraph(insertionIndex++, `GPA: ${String(edu.gpa).trim()}`);
    let gpaAttrs = RTS_Doc_copyAttributesPreservingEnums(lineAttrs);
    gpaAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 1;
    gpaAttrs[DocumentApp.Attribute.FONT_SIZE] = Math.max(8, (Number(lineAttrs[DocumentApp.Attribute.FONT_SIZE]) || 10) - 1);
    gpaPara.setAttributes(gpaAttrs);
    currentLastElement = gpaPara;
  }

  if (Array.isArray(edu.relevantCoursework) && edu.relevantCoursework.length > 0) {
    const coursesText = edu.relevantCoursework.map(c => String(c).trim()).filter(Boolean).join(", ");
    if (coursesText) {
      const courseworkPara = body.insertParagraph(insertionIndex++, "");
      let courseworkAttrs = RTS_Doc_copyAttributesPreservingEnums(lineAttrs);
      courseworkAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 1;
      courseworkPara.setAttributes(courseworkAttrs);
      courseworkPara.appendText("Relevant Coursework: ").setItalic(true); 
      courseworkPara.appendText(coursesText).setItalic(false); 
      currentLastElement = courseworkPara;
    }
  }
  return currentLastElement;
}

function RTS_Doc_renderExperienceJob(job, body, insertAfterElement, baseAttrs, index) {
  const DEBUG = GLOBAL_DEBUG_MODE; 
  if(DEBUG) Logger.log(`      renderExperienceJob: Index ${index}, Company: "${job.company || 'N/A'}"`);
  let currentLastEl = insertAfterElement; 
  let insertionIndex = body.getChildIndex(currentLastEl) + 1;
  const itemLineAttrs = RTS_Doc_copyAttributesPreservingEnums(baseAttrs);

  if (index > 0) {
    const spacingPara = body.insertParagraph(insertionIndex++, "");
    spacingPara.setAttributes({[DocumentApp.Attribute.SPACING_BEFORE]: 6, [DocumentApp.Attribute.SPACING_AFTER]: 0, [DocumentApp.Attribute.LINE_SPACING]: 0.8});
    currentLastEl = spacingPara;
  }
  
  const firstLineAttrs = RTS_Doc_copyAttributesPreservingEnums(itemLineAttrs);
  firstLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = (index === 0 && !(itemLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] > 0) ) ? 0 : 1;
  firstLineAttrs[DocumentApp.Attribute.SPACING_AFTER] = 1;
  firstLineAttrs[DocumentApp.Attribute.LINE_SPACING] = itemLineAttrs[DocumentApp.Attribute.LINE_SPACING] || 1.0;

  // Line 1: Company Name (Bold)
  if (job.company && String(job.company).trim()) {
    const companyPara = body.insertParagraph(insertionIndex++, String(job.company).trim());
    companyPara.setAttributes(RTS_Doc_copyAttributesPreservingEnums(firstLineAttrs));
    companyPara.setBold(true);
    currentLastEl = companyPara;
  }
  
  // Line 2: Job Title, Location (Location Italic) - CORRECTED
  const titleTextClean = String(job.jobTitle || "").trim();
  const locationTextClean = String(job.location || "").trim();
  if (titleTextClean || locationTextClean) {
    const titleLocPara = body.insertParagraph(insertionIndex++, ""); // Create empty paragraph
    const titleLineAttrs = RTS_Doc_copyAttributesPreservingEnums(firstLineAttrs);
    titleLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 0; 
    titleLocPara.setAttributes(titleLineAttrs).setBold(false).setItalic(false); // Set base style for the line

    if (titleTextClean) {
      titleLocPara.appendText(titleTextClean);
    }

    if (locationTextClean) {
      const separator = titleTextClean ? ", " : "";
      titleLocPara.appendText(separator); // Append separator
      const startItalicIndex = titleLocPara.getText().length; // Get length *before* appending location
      titleLocPara.appendText(locationTextClean);             // Append location
      // Italicize only the location part. End index is length - 1.
      titleLocPara.editAsText().setItalic(startItalicIndex, titleLocPara.getText().length - 1, true);
    }
    currentLastEl = titleLocPara;
  }
  
  // Line 3: Dates (Italic)
  let dateStringExp = String(job.endDate || "").trim();
  const startDateExpClean = String(job.startDate || "").trim();
  if (dateStringExp.toUpperCase() !== "PRESENT" && startDateExpClean) dateStringExp = `${startDateExpClean} – ${dateStringExp}`;
  else if (startDateExpClean && !String(job.endDate || "").trim()) dateStringExp = `${startDateExpClean} – Present`;
  else if (!startDateExpClean && dateStringExp.toUpperCase() === "PRESENT") dateStringExp = `Start Date N/A → Present`;
  else if (!startDateExpClean && !String(job.endDate || "").trim()) dateStringExp = ""; 
  
  if (dateStringExp) {
    const datePara = body.insertParagraph(insertionIndex++, dateStringExp);
    const dateLineAttrs = RTS_Doc_copyAttributesPreservingEnums(firstLineAttrs);
    dateLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 0;
    datePara.setAttributes(dateLineAttrs).setBold(false).setItalic(true);
    currentLastEl = datePara;
  }

  // Responsibilities (Bullet points)
  if (job.responsibilities && Array.isArray(job.responsibilities) && job.responsibilities.length > 0) {
    const bulletParaStyle = RTS_Doc_copyAttributesPreservingEnums(firstLineAttrs);
    bulletParaStyle[DocumentApp.Attribute.INDENT_START] = 36; 
    bulletParaStyle[DocumentApp.Attribute.INDENT_FIRST_LINE] = 18; 
    bulletParaStyle[DocumentApp.Attribute.SPACING_BEFORE] = (job.company || titleTextClean || dateStringExp) ? 3 : 0;
    bulletParaStyle[DocumentApp.Attribute.SPACING_AFTER] = 1; 

    job.responsibilities.forEach((bulletText, bulletIdx) => {
      const trimmedBullet = String(bulletText || "").trim();
      if (trimmedBullet) {
        const bulletListItem = body.insertListItem(insertionIndex++, trimmedBullet);
        const currentBulletStyle = RTS_Doc_copyAttributesPreservingEnums(bulletParaStyle);
        if (bulletIdx > 0) currentBulletStyle[DocumentApp.Attribute.SPACING_BEFORE] = 1; 
        bulletListItem.setAttributes(currentBulletStyle);
        bulletListItem.setGlyphType(DocumentApp.GlyphType.BULLET); 
        currentLastEl = bulletListItem;
      }
    });
  }
  return currentLastEl;
}

function RTS_Doc_renderProjectItem(project, body, insertAfterElement, baseAttrs, index) {
  const DEBUG = GLOBAL_DEBUG_MODE; 
  if(DEBUG) Logger.log(`      renderProjectItem: Index ${index}, Project Name: "${project.projectName || 'N/A'}"`);
  let currentLastElement = insertAfterElement; 
  let insertionIndex = body.getChildIndex(currentLastElement) + 1;
  const itemLineAttrs = RTS_Doc_copyAttributesPreservingEnums(baseAttrs);

  // Spacing between project entries
  if (index > 0) {
    const spacingPara = body.insertParagraph(insertionIndex++, "");
    spacingPara.setAttributes({[DocumentApp.Attribute.SPACING_BEFORE]: 6, [DocumentApp.Attribute.SPACING_AFTER]: 0, [DocumentApp.Attribute.LINE_SPACING]: 0.8});
    currentLastElement = spacingPara;
  }
  
  const lineAttrs = RTS_Doc_copyAttributesPreservingEnums(itemLineAttrs);
  lineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = (index === 0 && !(itemLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] > 0) ) ? 0 : 1;
  lineAttrs[DocumentApp.Attribute.SPACING_AFTER] = 1;
  lineAttrs[DocumentApp.Attribute.LINE_SPACING] = itemLineAttrs[DocumentApp.Attribute.LINE_SPACING] || 1.15; // Slightly more for projects

  // Line 1: Project Name (Bold)
  if (project.projectName && String(project.projectName).trim()) {
    const namePara = body.insertParagraph(insertionIndex++, String(project.projectName).trim());
    namePara.setAttributes(RTS_Doc_copyAttributesPreservingEnums(lineAttrs));
    namePara.setBold(true);
    currentLastElement = namePara;
  }

  // Line 2: Role and/or Organization (Organization Italic if present with Role)
  const roleTextClean = String(project.role || "").trim();
  const orgTextClean = String(project.organization || "").trim();
  if (roleTextClean || (orgTextClean && orgTextClean.toLowerCase() !== String(project.projectName||"").toLowerCase()) ) { // Only show org if different from project name
    const roleOrgPara = body.insertParagraph(insertionIndex++, "");
    const roleOrgLineAttrs = RTS_Doc_copyAttributesPreservingEnums(lineAttrs);
    roleOrgLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 0; // Tight to project name
    roleOrgPara.setAttributes(roleOrgLineAttrs).setBold(false).setItalic(false);

    if (roleTextClean) {
      roleOrgPara.appendText(roleTextClean);
    }
    if (orgTextClean && orgTextClean.toLowerCase() !== String(project.projectName||"").toLowerCase()) {
      const separator = roleTextClean ? " (" : ""; // e.g. "Lead Developer (Personal Project)"
      const suffix = roleTextClean ? ")" : "";
      roleOrgPara.appendText(separator);
      const startItalicHere = roleOrgPara.getText().length;
      roleOrgPara.appendText(orgTextClean);
      if (orgTextClean.length > 0) {
        roleOrgPara.editAsText().setItalic(startItalicHere, roleOrgPara.getText().length - 1, true);
      }
      roleOrgPara.appendText(suffix);
    }
    currentLastElement = roleOrgPara;
  }
  
  // Line 3: Dates (Italic, tight to previous)
  let dateStringProj = String(project.endDate || "").trim();
  const startDateProjClean = String(project.startDate || "").trim();
  if (dateStringProj.toUpperCase() !== "PRESENT" && startDateProjClean) dateStringProj = `${startDateProjClean} – ${dateStringProj}`;
  else if (startDateProjClean && !String(project.endDate || "").trim()) dateStringProj = `${startDateProjClean} – Present`;
  else if (!startDateProjClean && dateStringProj.toUpperCase() === "PRESENT") dateStringProj = "Ongoing Project";
  else if (!startDateProjClean && !String(project.endDate || "").trim()) dateStringProj = "";
  
  if (dateStringProj) {
    const datePara = body.insertParagraph(insertionIndex++, dateStringProj);
    const dateLineAttrs = RTS_Doc_copyAttributesPreservingEnums(lineAttrs);
    dateLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 0;
    datePara.setAttributes(dateLineAttrs).setBold(false).setItalic(true);
    currentLastElement = datePara;
  }

  // Description Bullets
  if (project.descriptionBullets && Array.isArray(project.descriptionBullets) && project.descriptionBullets.length > 0) {
    const bulletParaStyle = RTS_Doc_copyAttributesPreservingEnums(lineAttrs);
    bulletParaStyle[DocumentApp.Attribute.INDENT_START] = 36; 
    bulletParaStyle[DocumentApp.Attribute.INDENT_FIRST_LINE] = 18; 
    bulletParaStyle[DocumentApp.Attribute.SPACING_BEFORE] = (project.projectName || roleTextClean || dateStringProj) ? 3 : 0; 
    bulletParaStyle[DocumentApp.Attribute.SPACING_AFTER] = 1; 

    project.descriptionBullets.forEach((bulletText, bulletIdx) => {
      const trimmedBullet = String(bulletText || "").trim();
      if (trimmedBullet) {
        const bulletListItem = body.insertListItem(insertionIndex++, trimmedBullet);
        const currentBulletStyle = RTS_Doc_copyAttributesPreservingEnums(bulletParaStyle);
        if (bulletIdx > 0) currentBulletStyle[DocumentApp.Attribute.SPACING_BEFORE] = 1; 
        bulletListItem.setAttributes(currentBulletStyle);
        bulletListItem.setGlyphType(DocumentApp.GlyphType.BULLET); 
        currentLastElement = bulletListItem;
      }
    });
  }
  
  // Technologies Used (Label bold & italic, list normal)
  const technologiesList = Array.isArray(project.technologies) ? project.technologies.map(t=>String(t||"").trim()).filter(Boolean) : [];
  if (technologiesList.length > 0) {
    const techPara = body.insertParagraph(insertionIndex++, "");
    const techLineAttrs = RTS_Doc_copyAttributesPreservingEnums(lineAttrs);
    techLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = (project.descriptionBullets?.length > 0 || project.projectName || roleTextClean || dateStringProj) ? 3 : 1;
    techPara.setAttributes(techLineAttrs);
    techPara.appendText("Technologies: ").setBold(true).setItalic(true);
    techPara.appendText(technologiesList.join(", ")).setBold(false).setItalic(false);
    currentLastElement = techPara;
  }

  // GitHub Links (and other links like Impact, FutureDevelopment if they are URLs)
  (project.githubLinks || []).forEach((linkObj, linkIdx) => {
    if (linkObj.url && String(linkObj.url).trim()) {
      const linkPara = body.insertParagraph(insertionIndex++, "");
      const linkLineAttrs = RTS_Doc_copyAttributesPreservingEnums(lineAttrs);
      linkLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = (linkIdx === 0 && (technologiesList.length > 0 || project.descriptionBullets?.length > 0)) ? 3 : 1;
      linkPara.setAttributes(linkLineAttrs);
      const linkName = String(linkObj.name || "Repository").trim();
      linkPara.appendText(`GitHub (${linkName}): `).setBold(true); // Label for GitHub
      linkPara.appendText(String(linkObj.url).trim()).setLinkUrl(String(linkObj.url).trim()).setUnderline(true).setForegroundColor("#1155CC").setBold(false);
      currentLastElement = linkPara;
    }
  });
  // Simple text lines for Impact and Future Development
  if (project.impact && String(project.impact).trim()) {
    const impactPara = body.insertParagraph(insertionIndex++, "");
    const impactLineAttrs = RTS_Doc_copyAttributesPreservingEnums(lineAttrs);
    impactLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 3;
    impactPara.setAttributes(impactLineAttrs);
    impactPara.appendText("Impact: ").setBold(true);
    impactPara.appendText(String(project.impact).trim()).setBold(false);
    currentLastElement = impactPara;
  }
  if (project.futureDevelopment && String(project.futureDevelopment).trim()) {
    const futurePara = body.insertParagraph(insertionIndex++, "");
    const futureLineAttrs = RTS_Doc_copyAttributesPreservingEnums(lineAttrs);
    futureLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 1;
    futurePara.setAttributes(futureLineAttrs);
    futurePara.appendText("Future Development: ").setBold(true);
    futurePara.appendText(String(project.futureDevelopment).trim()).setBold(false);
    currentLastElement = futurePara;
  }

  return currentLastElement;
}

function RTS_Doc_renderTechnicalSkillsList(skillsSectionData, body, insertAfterElement, baseAttrs) {
  const DEBUG = GLOBAL_DEBUG_MODE;
  if(DEBUG) Logger.log(`      renderTechnicalSkillsList: Processing skills/certificates.`);
  let currentLastElement = insertAfterElement;
  let insertionIndex = body.getChildIndex(currentLastElement) + 1;
  
  const subsections = skillsSectionData.subsections || [];
  if (subsections.length === 0 && DEBUG) {
    Logger.log("        No subsections found in Technical Skills data.");
    // Insert a placeholder paragraph if you want the section title to still appear if no skills are listed later
    // const p = body.insertParagraph(insertionIndex, "(No technical skills or certificates listed for this tailoring)");
    // p.setAttributes(baseAttrs).setItalic(true);
    // return p; // Or simply return insertAfterElement if section should be blank
    return currentLastElement; 
  }

  subsections.forEach((subsection, subIdx) => {
    const categoryName = String(subsection.name || "Uncategorized Technical Skills").trim();
    const itemsInCategory = subsection.items || [];

    if (categoryName && Array.isArray(itemsInCategory) && itemsInCategory.length > 0) {
      const lineAttrs = RTS_Doc_copyAttributesPreservingEnums(baseAttrs);
      // Spacing before the first category vs. between categories
      lineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = (subIdx === 0 && !(baseAttrs[DocumentApp.Attribute.SPACING_BEFORE] > 0) ) ? 0 : 4;
      lineAttrs[DocumentApp.Attribute.SPACING_AFTER] = 1; 
      lineAttrs[DocumentApp.Attribute.LINE_SPACING] = baseAttrs[DocumentApp.Attribute.LINE_SPACING] || 1.0;
      
      const skillCategoryLinePara = body.insertParagraph(insertionIndex++, ""); // Create empty paragraph
      skillCategoryLinePara.setAttributes(lineAttrs);
      
      // Category Name (e.g., "Programming Languages:") - Bold
      skillCategoryLinePara.appendText(categoryName + ": ").setBold(true);
      
      let skillsAndCertsTextArray = [];
      itemsInCategory.forEach(item => {
        let entryText = "";
        if (item.skill && String(item.skill).trim()) { // It's a skill item
          entryText = String(item.skill).trim();
          if (item.details && String(item.details).trim()) {
            entryText += ` (${String(item.details).trim()})`; 
          }
        } else if (item.name && String(item.name).trim()) { // It's a certificate/license item
          entryText = String(item.name).trim();
          let extras = [];
          if (item.issuer && String(item.issuer).trim()) extras.push(String(item.issuer).trim());
          if (item.issueDate && String(item.issueDate).trim()) extras.push(String(item.issueDate).trim());
          // Could also add item.details for certs if that field exists for them.
          if (extras.length > 0) entryText += ` (${extras.join(", ")})`;
        } else if (typeof item === 'string' && String(item).trim()) { // Fallback if item is just a string
          entryText = String(item).trim();
        }
        if (entryText) skillsAndCertsTextArray.push(entryText);
      });
      
      const skillsAndCertsJoinedString = skillsAndCertsTextArray.join(", ").trim();
      if (skillsAndCertsJoinedString) {
        skillCategoryLinePara.appendText(skillsAndCertsJoinedString).setBold(false).setItalic(false);
      }
      currentLastElement = skillCategoryLinePara;
    }
  });
  return currentLastElement;
}

function RTS_Doc_renderLeadershipItem(item, body, insertAfterElement, baseAttrs, index) {
  const DEBUG = GLOBAL_DEBUG_MODE; 
  if(DEBUG) Logger.log(`      renderLeadershipItem: Index ${index}, Org: "${item.organization || 'N/A'}"`);
  let currentLastElement = insertAfterElement; 
  let insertionIndex = body.getChildIndex(currentLastElement) + 1;
  const itemLineAttrs = RTS_Doc_copyAttributesPreservingEnums(baseAttrs);

  if (index > 0) {
    const spacingPara = body.insertParagraph(insertionIndex++, "");
    spacingPara.setAttributes({[DocumentApp.Attribute.SPACING_BEFORE]: 6, [DocumentApp.Attribute.SPACING_AFTER]: 0, [DocumentApp.Attribute.LINE_SPACING]: 0.8});
    currentLastElement = spacingPara;
  }
  
  const lineAttrs = RTS_Doc_copyAttributesPreservingEnums(itemLineAttrs);
  lineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = (index === 0 && !(itemLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] > 0)) ? 0 : 1;
  lineAttrs[DocumentApp.Attribute.SPACING_AFTER] = 1;
  lineAttrs[DocumentApp.Attribute.LINE_SPACING] = itemLineAttrs[DocumentApp.Attribute.LINE_SPACING] || 1.0;

  // Line 1: Organization (Bold)
  if (item.organization && String(item.organization).trim()) {
    const orgPara = body.insertParagraph(insertionIndex++, String(item.organization).trim());
    orgPara.setAttributes(RTS_Doc_copyAttributesPreservingEnums(lineAttrs));
    orgPara.setBold(true);
    currentLastElement = orgPara;
  }

  // Line 2: Dates (Italic, tight to Organization)
  let dateTextLead = String(item.endDate || "").trim();
  const startDateLeadClean = String(item.startDate || "").trim();
  if (dateTextLead.toUpperCase() !== "PRESENT" && startDateLeadClean) dateTextLead = `${startDateLeadClean} – ${dateTextLead}`;
  else if (startDateLeadClean && !String(item.endDate || "").trim()) dateTextLead = `${startDateLeadClean} – Present`;
  else if (!startDateLeadClean && dateTextLead.toUpperCase() === "PRESENT") dateTextLead = "Start Date N/A → Present";
  else if (!startDateLeadClean && !String(item.endDate || "").trim()) dateTextLead = "";

  if (dateTextLead) {
    const datePara = body.insertParagraph(insertionIndex++, dateTextLead);
    const dateLineAttrs = RTS_Doc_copyAttributesPreservingEnums(lineAttrs);
    dateLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 0; 
    datePara.setAttributes(dateLineAttrs).setBold(false).setItalic(true);
    currentLastElement = datePara;
  }

  // Line 3: Role [, Location (Location Italic)] - CORRECTED
  const roleTextClean = String(item.role || "").trim();
  const locationTextClean = String(item.location || "").trim();
  if (roleTextClean || locationTextClean) {
    const roleLocPara = body.insertParagraph(insertionIndex++, "");
    const roleLineAttrs = RTS_Doc_copyAttributesPreservingEnums(lineAttrs);
    // Add a bit of space if organization or dates were present above
    roleLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = (item.organization || dateTextLead) ? 1 : 0; 
    roleLocPara.setAttributes(roleLineAttrs).setBold(false).setItalic(false);

    if (roleTextClean) {
      roleLocPara.appendText(roleTextClean);
    }
    if (locationTextClean) {
      const separator = roleTextClean ? ", " : "";
      roleLocPara.appendText(separator);
      const startItalicHere = roleLocPara.getText().length;
      roleLocPara.appendText(locationTextClean);
      if (locationTextClean.length > 0) {
        roleLocPara.editAsText().setItalic(startItalicHere, roleLocPara.getText().length - 1, true);
      }
    }
    currentLastElement = roleLocPara;
  }
  
  // Line 4: Description (Optional plain text)
  const descriptionTextClean = String(item.description || "").trim();
  if (descriptionTextClean) {
    const descPara = body.insertParagraph(insertionIndex++, descriptionTextClean);
    const descLineAttrs = RTS_Doc_copyAttributesPreservingEnums(lineAttrs);
    descLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 1;
    descPara.setAttributes(descLineAttrs);
    currentLastElement = descPara;
  }

  // Line 5+: Responsibilities (Bullet points)
  if (item.responsibilities && Array.isArray(item.responsibilities) && item.responsibilities.length > 0) {
    const bulletParaStyle = RTS_Doc_copyAttributesPreservingEnums(lineAttrs);
    bulletParaStyle[DocumentApp.Attribute.INDENT_START] = 36; 
    bulletParaStyle[DocumentApp.Attribute.INDENT_FIRST_LINE] = 18; 
    bulletParaStyle[DocumentApp.Attribute.SPACING_BEFORE] = (item.organization || titleTextClean || dateTextLead || descriptionTextClean) ? 3 : 0; 
    bulletParaStyle[DocumentApp.Attribute.SPACING_AFTER] = 1; 

    item.responsibilities.forEach((bulletText, bulletIdx) => {
      const trimmedBullet = String(bulletText || "").trim();
      if (trimmedBullet) {
        const bulletListItem = body.insertListItem(insertionIndex++, trimmedBullet);
        const currentBulletStyle = RTS_Doc_copyAttributesPreservingEnums(bulletParaStyle);
        if (bulletIdx === 0 && descriptionTextClean) currentBulletStyle[DocumentApp.Attribute.SPACING_BEFORE] = 1; // Tighter if after description
        else if (bulletIdx > 0) currentBulletStyle[DocumentApp.Attribute.SPACING_BEFORE] = 1; 
        bulletListItem.setAttributes(currentBulletStyle);
        bulletListItem.setGlyphType(DocumentApp.GlyphType.BULLET); 
        currentLastElement = bulletListItem;
      }
    });
  }
  return currentLastElement;
}

function RTS_Doc_renderHonorItem(award, body, insertAfterElement, baseAttrs, index) {
  const DEBUG = GLOBAL_DEBUG_MODE; 
  if(DEBUG) Logger.log(`      renderHonorItem: Index ${index}, Award: "${award.awardName || 'N/A'}"`);
  let currentLastElement = insertAfterElement; 
  let insertionIndex = body.getChildIndex(currentLastElement) + 1;

  let awardText = String(award.awardName || "Unnamed Award").trim();
  const detailsText = String(award.details || "").trim();
  const dateText = String(award.date || "").trim(); // Assumes date is already formatted string by MasterResumeData

  if (detailsText) awardText += ` (${detailsText})`;
  if (dateText) awardText += (detailsText || String(award.awardName || "").trim()) ? ` - ${dateText}` : dateText; // Add separator if other text exists
  
  if (awardText.trim()) {
      const honorBulletStyle = RTS_Doc_copyAttributesPreservingEnums(baseAttrs);
      // If it's the first honor, use placeholder's before-spacing or 0. Subsequent honors get small spacing.
      honorBulletStyle[DocumentApp.Attribute.SPACING_BEFORE] = (index === 0 && !(baseAttrs[DocumentApp.Attribute.SPACING_BEFORE] > 0) ) ? 0 : 2; 
      honorBulletStyle[DocumentApp.Attribute.SPACING_AFTER] = baseAttrs[DocumentApp.Attribute.SPACING_AFTER] || 1; // Minimal space after
      honorBulletStyle[DocumentApp.Attribute.INDENT_START] = 36;         // Standard bullet indent
      honorBulletStyle[DocumentApp.Attribute.INDENT_FIRST_LINE] = 18;   // Standard hanging indent for bullets

      const honorListItem = body.insertListItem(insertionIndex++, awardText.trim());
      honorListItem.setAttributes(honorBulletStyle);
      honorListItem.setGlyphType(DocumentApp.GlyphType.BULLET);
      currentLastElement = honorListItem;
  }
  return currentLastElement;
}

// --- DOCUMENT CLEANUP ---
function RTS_Doc_cleanupAllEmptyLines(body) {
  const DEBUG=GLOBAL_DEBUG_MODE; if(DEBUG) Logger.log("  RTS_Doc_cleanup: Starting aggressive empty line removal.");
  let removed=0; for(let i=body.getNumChildren()-1;i>=0;i--){const el=body.getChild(i); if(el.getType()===DocumentApp.ElementType.PARAGRAPH){const p=el.asParagraph();if(p.getNumChildren()===0||p.getText().replace(/\s/g,"")===""){try{p.removeFromParent();removed++;}catch(e){}}}}
  if(DEBUG) Logger.log(`    RTS_Doc_cleanup: Removed ${removed} visually blank paragraphs.`);
}

// --- MAIN DOCUMENT CREATION FUNCTION ---
/**
 * Creates a formatted resume Google Document from a resume data object.
 * Uses RESUME_TEMPLATE_DOC_ID from Global_Constants.gs.
 */
function RTS_createFormattedResumeDoc(resumeDataObject, documentTitle, resumeTemplateId = (typeof RESUME_TEMPLATE_DOC_ID !== 'undefined' ? RESUME_TEMPLATE_DOC_ID : "")) {
  const DEBUG = GLOBAL_DEBUG_MODE; // From Global_Constants.gs
  let currentTask = "Function Entry & Validation";
  Logger.log(`--- RTS_DocumentService: Starting 'RTS_createFormattedResumeDoc' for title: "${documentTitle}" ---`);
  if(DEBUG) Logger.log(`    Using Template ID (from param or global): "${resumeTemplateId}"`);

  if (!resumeDataObject?.personalInfo) {Logger.log("[ERROR] RTS_DocService: Invalid resumeDataObject or missing personalInfo."); return null;}
  if (!resumeTemplateId || String(resumeTemplateId).toUpperCase().includes("YOUR_") || resumeTemplateId.trim() === "") {
    Logger.log("[ERROR] RTS_DocService: RESUME_TEMPLATE_DOC_ID is placeholder, missing, or invalid in Global_Constants.gs or parameter.");
    return null;
  }

  const pi = resumeDataObject.personalInfo;
  let finalDocTitle = String(documentTitle || `Resume - ${pi.fullName || 'Candidate'}`).trim();
  // Add timestamp to ensure unique document names and provide versioning context
  if (!finalDocTitle.match(/\(\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\)/)) { // Check for specific timestamp format
    const now = new Date();
    const timestamp = `${now.getFullYear()}-${("0"+(now.getMonth()+1)).slice(-2)}-${("0"+now.getDate()).slice(-2)}_${("0"+now.getHours()).slice(-2)}-${("0"+now.getMinutes()).slice(-2)}-${("0"+now.getSeconds()).slice(-2)}`;
    finalDocTitle += ` (Generated ${timestamp})`;
  }
  if(DEBUG) Logger.log(`    Effective document title: "${finalDocTitle}"`);

  let newDocFile, doc, body;
  try {
    currentTask = "Accessing Template & Copying";
    if(DEBUG) Logger.log(`  RTS_DocService: Task - ${currentTask}`);
    const templateFile = DriveApp.getFileById(resumeTemplateId);
    newDocFile = templateFile.makeCopy(finalDocTitle, DriveApp.getRootFolder()); // Or specify a target folder
    doc = DocumentApp.openById(newDocFile.getId());
    body = doc.getBody(); // ASSIGN BODY IMMEDIATELY AFTER OPENING DOC
    if(DEBUG) Logger.log(`    Template copied to "${doc.getName()}", body retrieved.`);

    currentTask = "Populating Personal Info & Summary";
    if(DEBUG) Logger.log(`  RTS_DocService: Task - ${currentTask}`);
    if (pi) {
      if(DEBUG) Logger.log(`    PI Data: ${JSON.stringify(pi)}`);
      RTS_Doc_findAndReplaceText(body, "{FULL_NAME}", pi.fullName);
      // Construct contact line 1 (Location, Phone, Email)
      const contactParts1 = [pi.location, pi.phone, pi.email].filter(Boolean).map(s=>String(s).trim());
      RTS_Doc_findAndReplaceText(body, "{CONTACT_LINE_1}", contactParts1.join("  •  "));
      // Handle {CONTACT_LINKS} placeholder: LinkedIn, Portfolio, GitHub
      const linksPlaceholderPattern = RegExp.escape("{CONTACT_LINKS}");
      const linksRange = body.findText(linksPlaceholderPattern);
      if (linksRange) {
          const linksElement = linksRange.getElement();
          if (linksElement?.getParent()?.getType() === DocumentApp.ElementType.PARAGRAPH) {
              const linksPara = linksElement.getParent().asParagraph();
              linksPara.clear(); // Clear the placeholder text
              let firstLinkAdded = true;
              const appendLinkWithSeparator = (text, url) => {
                  if(!url || !String(url).trim()) return;
                  if (!firstLinkAdded) linksPara.appendText("  •  ");
                  linksPara.appendText(text).setLinkUrl(url).setUnderline(true).setForegroundColor("#1155CC"); // Standard link blue
                  firstLinkAdded = false;
              };
              appendLinkWithSeparator("LinkedIn", pi.linkedin);
              appendLinkWithSeparator("Portfolio", pi.portfolio);
              appendLinkWithSeparator("GitHub", pi.github);
              if(DEBUG && firstLinkAdded) Logger.log(`      No links found in personalInfo to populate {CONTACT_LINKS}.`);
          } else if(DEBUG) Logger.log(`    [WARN] Parent of {CONTACT_LINKS} not a Paragraph.`);
      } else if(DEBUG) Logger.log(`    [WARN] Placeholder {CONTACT_LINKS} not found.`);
    } else if(DEBUG) Logger.log(`    [WARN] Personal Info (pi) object is null/undefined for document header.`);
    RTS_Doc_findAndReplaceText(body, "{SUMMARY_CONTENT}", resumeDataObject.summary);
    SpreadsheetApp.flush(); // Try to apply text changes

    currentTask = "Populating Dynamic Resume Sections";
    if(DEBUG) Logger.log(`  RTS_DocService: Task - ${currentTask}`);
    if (resumeDataObject.sections?.length > 0) {
      const sectionRenderOrder = [ // This order determines how sections appear in the doc
        "EDUCATION", "TECHNICAL SKILLS & CERTIFICATES", "EXPERIENCE", 
        "PROJECTS", "LEADERSHIP & UNIVERSITY INVOLVEMENT", "HONORS & AWARDS"
      ];
      sectionRenderOrder.forEach(sectionTitleKey => {
        const sectionData = resumeDataObject.sections.find(s => s?.title?.toUpperCase() === sectionTitleKey.toUpperCase());
        if (sectionData) {
          if(DEBUG) Logger.log(`    Found section data for: "${sectionTitleKey}"`);
          let placeholder, renderFunc, items, type = "items";
          switch (sectionTitleKey) {
            case "EDUCATION": placeholder="{EDUCATION_ITEMS}"; renderFunc=RTS_Doc_renderEducationItem; items=sectionData.items; break;
            case "EXPERIENCE": placeholder="{EXPERIENCE_JOBS}"; renderFunc=RTS_Doc_renderExperienceJob; items=sectionData.items; break;
            case "PROJECTS": placeholder="{PROJECT_ITEMS}"; renderFunc=RTS_Doc_renderProjectItem; type="subsections"; items=sectionData.subsections; break; // Projects rendered via subsections of items
            case "TECHNICAL SKILLS & CERTIFICATES": placeholder="{TECHNICAL_SKILLS_LIST}"; renderFunc=RTS_Doc_renderTechnicalSkillsList; type="fullSection"; items=[sectionData]; break; // Pass whole section obj
            case "LEADERSHIP & UNIVERSITY INVOLVEMENT": placeholder="{LEADERSHIP_ITEMS}"; renderFunc=RTS_Doc_renderLeadershipItem; items=sectionData.items; break;
            case "HONORS & AWARDS": placeholder="{HONORS_LIST}"; renderFunc=RTS_Doc_renderHonorItem; items=sectionData.items; break;
            default: Logger.log(`    [WARN] Unknown section title "${sectionTitleKey}" in render order.`); return;
          }
          
          let itemsToRender = [];
          if (type === "items" && Array.isArray(items)) itemsToRender = items;
          else if (type === "subsections" && Array.isArray(items)) { // For projects
            items.forEach(sub => { if(Array.isArray(sub.items)) itemsToRender.push(...sub.items); });
          } else if (type === "fullSection" && Array.isArray(items)) itemsToRender = items; // For skills list

          if(itemsToRender.length > 0 && renderFunc){
            RTS_Doc_populateBlockPlaceholder(body, placeholder, itemsToRender, renderFunc);
          } else { RTS_Doc_findAndReplaceText(body, placeholder, ""); if(DEBUG) Logger.log(`      No items or render func for "${sectionTitleKey}", cleared placeholder.`);}
        } else { // Section data not found in resumeDataObject
            const phMap = {"EDUCATION":"{EDUCATION_ITEMS}", "EXPERIENCE":"{EXPERIENCE_JOBS}", "PROJECTS":"{PROJECT_ITEMS}", "TECHNICAL SKILLS & CERTIFICATES":"{TECHNICAL_SKILLS_LIST}", "LEADERSHIP & UNIVERSITY INVOLVEMENT":"{LEADERSHIP_ITEMS}", "HONORS & AWARDS":"{HONORS_LIST}"};
            if(phMap[sectionTitleKey]) RTS_Doc_findAndReplaceText(body, phMap[sectionTitleKey], ""); // Clear placeholder if section is missing
            if(DEBUG) Logger.log(`    Section "${sectionTitleKey}" not found in resumeDataObject. Placeholder (if any) cleared.`);
        }
      });
    } else if(DEBUG) Logger.log("    No 'sections' array in resumeDataObject or it's empty.");

    currentTask = "Final Document Cleanup";
    if(DEBUG) Logger.log(`  RTS_DocService: Task - ${currentTask}`);
    RTS_Doc_cleanupAllEmptyLines(body);
    SpreadsheetApp.flush(); // Final flush

    currentTask = "Save and Close";
    if(DEBUG) Logger.log(`  RTS_DocService: Task - ${currentTask}`);
    doc.saveAndClose();

    Logger.log(`[SUCCESS] RTS_DocumentService: Document "${doc.getName()}" generated successfully. URL: ${doc.getUrl()}`);
    return doc.getUrl();

  } catch (e) {
    Logger.log(`[CRITICAL ERROR] RTS_DocumentService: During task "${currentTask}": ${e.message}\nStack: ${e.stack || 'No stack trace'}`);
    if (newDocFile?.getId()) { // Check if newDocFile is defined before trying to get ID
      try { DriveApp.getFileById(newDocFile.getId()).setTrashed(true); Logger.log(`  Attempted to trash incomplete document copy: ${newDocFile.getName()}`); } 
      catch (delErr) { Logger.log(`  Error trashing incomplete copy (ID: ${newDocFile.getId()}): ${delErr.message}`);}
    } else if (doc?.getId()){ // Fallback if newDocFile not defined but doc was
        try { DriveApp.getFileById(doc.getId()).setTrashed(true); Logger.log(`  Attempted to trash doc (ID: ${doc.getId()}).`); } 
        catch (delErr) { Logger.log(`  Error trashing doc by ID: ${delErr.message}`);}
    }
    return null; 
  }
}
