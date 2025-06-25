# Career_Suite_AI_v6
# Comprehensive AI-Powered Job Search & Application Suite **REVISED NAME**

**Turn your job search into an efficient, data-driven, and AI-powered operation. This all-in-one Google Apps Script suite, built directly into Google Sheets, transforms the chaotic process of finding, tracking, and applying for jobs into a streamlined, intelligent workflow.**

This is not just one script, but a **cohesive suite of interconnected modules**:

1.  **🤖 The Master Job Manager (MJM):** An automated system that tracks your job applications by parsing your emails, providing a live dashboard with key metrics and a funnel analysis.
2.  **📝 The AI Resume Tailor (RTS):** A sophisticated, three-stage AI engine that analyzes job descriptions, scores your master profile against them, and generates a perfectly tailored resume document on command.

Easy access setup link: https://docs.google.com/spreadsheets/d/12jj5lTyu_MzA6KBkfD-30mj-KYHaX-BjouFMtPIIzFc/edit?usp=sharing

---

## 🚀 Core Features

### Master Job Manager (MJM) Module

*   **⚡ Automated Application Tracking:** Set up a Gmail filter with one click. The script automatically processes application update emails (e.g., "Application Received," "Invitation to Interview," "Rejected"), updating their status in your tracker sheet.
*   **🧠 Intelligent Dual-Engine Email Parsing:**
    *   **AI-Powered (Gemini):** Leverages Google's Gemini API for high-accuracy extraction of company name, job title, and application status from unstructured emails.
    *   **RegEx-Fortified:** Includes a robust, "battle-tested" regular expression engine as a fallback, ensuring reliability even if the AI is unavailable. This parser is smart enough to ignore common ATS noise (Greenhouse, Lever) and identify key status-change keywords.
*   **📈 Live KPI Dashboard:** A beautifully designed and fully automated dashboard gives you a real-time overview of your job search with key metrics like:
    *   Total & Active Applications
    *   Interview & Offer Rates
    *   Application Funnel Analysis (Applied -> Viewed -> Interview -> Offer)
    *   Platform Distribution (LinkedIn vs. Indeed, etc.)
*   **🔍 Proactive Job Lead Sourcing:** A parallel AI engine that processes "job alert" emails, performing complex **one-to-many extraction** to pull multiple distinct job opportunities into a clean, actionable database of potential leads.
*   **⏲️ Failsafe Trigger Management:** Scripts run automatically on hourly and daily triggers. The setup is idempotent and robust, preventing the creation of duplicate triggers.

### AI Resume Tailor (RTS) Module

*   **📄 Centralized Master Profile:** Define your entire professional history—work experience, projects, skills, education—in one structured "MasterProfile" sheet. This serves as the single source of truth for all tailored resumes, based on a flexible, configuration-driven schema.
*   **🤖 Sophisticated Human-in-the-Loop AI Workflow:** A powerful, three-stage process that combines AI's analytical power with your final judgment.
    1.  **Stage 1: AI Diagnostics & Scoring:** The AI deconstructs a job description and then scores *every single bullet point* in your master profile for relevance, providing a score and a written justification. A color-coded interface lets you easily review the AI's analysis.
    2.  **Stage 2: Human Curation:** You decide which high-scoring or promising bullets to include by simply marking them "YES" in the sheet. The AI only proceeds with your explicit selections.
    3.  **Stage 3: AI-Enhanced Content Generation:** For each selected bullet, the AI copywriter rewrites it to align perfectly with the job description, incorporating keywords and optimizing for impact. It even **synthesizes** an entirely new, tailored professional summary based on your best-matched experience.
*   **✨ One-Click, High-Fidelity Document Generation:** The final stage uses a **custom-built templating engine** to take the structured, AI-tailored data and render a pixel-perfect, professionally formatted Google Doc resume. You can change the entire look and feel of your resume simply by editing the template document—no code changes needed.

---

## 🛠️ Tech Stack & Architecture

This project is a showcase of advanced Google Apps Script engineering, applied AI, and robust architectural patterns.

*   **Primary Language:** Google Apps Script (JavaScript)
*   **AI & LLMs:**
    *   **Google Gemini Pro:** Used for its powerful reasoning and data extraction capabilities in both the Application Tracker and Leads modules.
    *   **Groq (Gemma, Llama, Mixtral):** The core engine of the Resume Tailor, leveraged for its speed and precision in analysis, scoring, and creative content generation.
*   **Platform:** Google Workspace (Sheets, Docs, Gmail, Drive)
*   **Key Architectural Concepts:**
    *   **Modular Design:** The project is cleanly separated into distinct modules (MJM, RTS) and utility libraries (Admin, Triggers, Parsers, Services), demonstrating the Separation of Concerns principle.
    *   **Configuration-Driven Systems:** Core logic is driven by centralized configuration objects (`Global_Constants.gs`), making the entire suite highly extensible and easy to maintain. The RTS `MasterProfile` structure is a prime example, capable of handling new resume sections without code changes.
    *   **Resilient, Fault-Tolerant Logic:** Includes automated retries for API calls, graceful fallbacks from AI to RegEx parsing, and "leave no trace" error handling that automatically trashes incomplete artifacts.
    *   **Dynamic "Document as a Template" Engine:** The resume generation service separates content from presentation, allowing for full visual customization by the user via a simple Google Doc template.
*   **Quality Assurance & Testing:**
    *   The project includes a **comprehensive, multi-level testing suite** accessible from the UI, featuring **unit tests** for core AI functions, **component tests** for the document renderer, and full **end-to-end integration tests** for each stage of the resume tailoring workflow.

---

## ⚙️ Setup & Installation

**Prerequisites:**
1.  A Google Account.
2.  API Keys for **Google AI (Gemini)** and **Groq**.

### Initial 5-Minute Setup

1.  **Create the Spreadsheet:**
    *   Create a new Google Sheet. This will be the home for your entire job suite.
    *   **Crucial:** Note the **ID** of your new spreadsheet from the URL (`.../spreadsheets/d/SPREADSHEET_ID_IS_HERE/edit...`).
2.  **Deploy the Code:**
    *   In your new sheet, go to `Extensions > Apps Script`.
    *   Delete any code in the `Code.gs` file.
    *   Create new `.gs` script files for each of the files in this repository (e.g., `Global_Constants.gs`, `MJM_UI.gs`, `RTS_Main.gs`, etc.) and copy-paste the content of each file into its corresponding script file in the editor.
3.  **Configure Core Constants:**
    *   Open `Global_Constants.gs`.
    *   In the `APP_SPREADSHEET_ID` constant, paste the spreadsheet ID from step 1.
    *   In the `RESUME_TEMPLATE_DOC_ID` constant, paste the ID of a Google Doc you want to use as your resume template. A sample template with the required placeholders (`{FULL_NAME}`, `{EXPERIENCE_JOBS}`, etc.) is recommended.
4.  **Save & Refresh:**
    *   Save all script files in the Apps Script editor.
    *   Return to your spreadsheet and **refresh the browser window**. A new menu `⚙️ Comprehensive AI Job Suite` should appear.
5.  **Set API Keys:**
    *   Go to the new menu: `⚙️ ... > 🔧 Admin & Configuration > Set SHARED Gemini API Key`. Paste your key.
    *   Go to `⚙️ ... > 🔧 Admin & Configuration > Set SHARED Groq API Key`. Paste your key.
6.  **Run the Full Project Setup:**
    *   Go to `⚙️ ... > ▶️ Initial Project Setup > RUN FULL PROJECT SETUP (All Modules)`.
    *   The script will ask for permissions. **Grant them.** This is required for the script to manage your sheet, create Gmail labels/filters, and create documents.
    *   The script will now automatically create all the necessary sheets, triggers, and Gmail labels. This may take a minute.

**Your AI Job Suite is now ready to use!**

---

## 📖 How to Use

### Tracking Applications (MJM)

*   **It's Automatic!** As you receive application updates to your Gmail, the filter you created will label them. The hourly trigger will process them, and you will see the updates appear in your `Applications` sheet and on the `Dashboard`.
*   **Manual Processing:** You can also manually trigger the email processing via the menu `MJM: Manual Processing > Process Application Update Emails`.

### Generating a Tailored Resume (RTS)

This is a three-step process guided by the UI menu.

1.  **Populate your Master Profile:**
    *   Go to the `MasterProfile` sheet and fill it out completely, following the existing structure. This is your professional database.
2.  **Step 1: Analyze & Score:**
    *   Go to `RTS: Resume Tailoring > STEP 1: Analyze JD & Score Profile Bullets`.
    *   Paste the full job description you are targeting into the prompt.
    *   The script will create/update the `BulletScoringResults` sheet, complete with relevance scores and color-coding.
3.  **Step 2: Curate Your Bullets:**
    *   Review the `BulletScoringResults` sheet.
    *   For every bullet point you want to include (and potentially enhance), select **"YES"** from the dropdown in the `SelectToTailor(Manual)` column.
    *   Once you've made your selections, go to `RTS: Resume Tailoring > STEP 2: Tailor Selected Bullets...`.
    *   The script will use the AI to rewrite only the bullets you selected, placing the improved text in the `TailoredBulletText(Stage2)` column.
4.  **Step 3: Generate the Document:**
    *   With your tailored bullets ready, go to `RTS: Resume Tailoring > STEP 3: Generate Tailored Resume Document`.
    *   The script will intelligently assemble the best content, generate a new AI summary, and create a perfectly formatted Google Doc. A link to the new resume will appear in a pop-up.
