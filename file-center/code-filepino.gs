/**
 * PROJECT: FILEPINO (Google Apps Script)
 * ===================================================================
 * ENHANCED CENTRALIZED AUTOMATION SYSTEM v2025-2 (PDF DIRECT PROCESSING)
 * AI-Powered Document Processing - Native PDF Reading with Gemini
 * ===================================================================
 *
 * CRITICAL FIXES IN v2025-2:
 * ✅ Gemini now reads PDFs DIRECTLY (no PNG conversion needed)
 * ✅ Much faster and more reliable PDF title extraction
 * ✅ Gemini API as default for all processing
 * ✅ Better error handling and debug logging
 * Version: 2025-2 (Production-Ready)
 * Last Updated: February 18, 2026
 * Author: Atty. Mary Wendy Duran
 * ===================================================================
 */

/**
 * NOTES / MINI CHANGELOG (2026-02)
 * - Added Gemini model selection via menu prompt (stored in Script Properties: GEMINI_MODEL).
 * - Gemini calls now use a centralized model-driven endpoint builder (buildGeminiGenerateContentUrl).
 * - Removed OpenAI configuration and fallback logic (Gemini-only; OCR/filename remain as fallback).
 * - Spreadsheet ID now resolves from the container/active spreadsheet and is persisted (TARGET_SPREADSHEET_ID), with safe fallback.
 * - Added anti-spam safeguards to protect credits:
 *   - Daily Gemini call cap (GEMINI_DAILY_CALL_LIMIT; default 200) with per-day counters (GEMINI_CALLS_YYYY-MM-DD).
 *   - Test cooldown for connection tests (GEMINI_LAST_TEST_MS).
 */

// Configuration Constants
const SPREADSHEET_ID = "19bd-tGQ9EKpp7Wh21Upz6pkcSBlU0F2dRD1Z2BEfEK8";
const TIMEZONE = "GMT+8";
const MAX_FILENAME_LENGTH = 150;
const SUPPORTED_MIME_TYPES = [
  "application/pdf",
  "image/jpeg",
  "image/png",
  "image/gif",
  "image/bmp",
];
const API_RETRY_ATTEMPTS = 3;
const API_RETRY_DELAY = 1000;
const MAX_RECIPIENTS_PER_EMAIL = 45;
const EMAIL_BATCH_DELAY = 2000;
const DEBUG_MODE = true; // Set to false in production
// Email Configuration
const WEEKLY_DIGEST_ENABLED = true; // Set to false to disable weekly digest
const WEEKLY_DIGEST_DAY = 5; // 5 = Friday (0=Sunday, 1=Monday, etc.)

// ===================================================================
// SYSTEM INITIALIZATION
// ===================================================================

function onOpen() {
  try {
    getTargetSpreadsheetId();
    createCustomMenu();
    Logger.log("Custom menu created successfully");
  } catch (error) {
    logError("Menu creation failed", error);
  }
}

function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();

  // Check if schedule is enabled to show appropriate options
  const trigger = getScheduledTrigger();
  const scheduleEnabled = !!trigger;

  ui.createMenu("Automation System")
    .addItem("🔧 Initialize Setup", "initializeSetup")
    .addSeparator()
    .addItem("🔍 Scan Files Now", "scanFilesManually")
    .addItem("📧 Send Daily Emails", "sendDailyEmails")
    .addItem("📅 Send Weekly Digest", "sendWeeklyDigest")
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("⚙️ Schedule Settings")
        .addItem(
          scheduleEnabled
            ? "✅ Schedule Enabled (5:15 PM Mon-Fri)"
            : "⏸️ Schedule Disabled",
          "checkScheduleStatus",
        )
        .addSeparator()
        .addItem("▶️ Enable Auto Schedule", "enableScheduleFromMenu")
        .addItem("⏸️ Disable Auto Schedule", "disableScheduleFromMenu"),
    )
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("🔑 AI Configuration")
        .addItem("Set AI API Key", "setGeminiApiKey")
        .addItem("Set Gemini Model", "showGeminiModelDialog")
        .addItem("Test AI Connection", "testAIConnection")
        .addItem("Clear All API Keys", "clearApiKeys"),
    )
    .addSeparator()
    .addItem("📊 View System Status", "showSystemStatus")
    .addItem("🔍 Review Flagged Files", "reviewFlaggedFiles")
    .addItem("🐛 Test Single File", "testSingleFileExtraction")
    .addItem("📧 Check Email Quota", "checkEmailQuota")
    .addToUi();
}

function initializeSetup() {
  try {
    const ui = SpreadsheetApp.getUi();
    const confirmResponse = ui.alert(
      "Initialize System Setup v2025-2",
      "This will create the required sheets. Continue?",
      ui.ButtonSet.YES_NO,
    );
    if (confirmResponse !== ui.Button.YES) {
      ui.alert(
        "Setup Cancelled",
        "System initialization was cancelled.",
        ui.ButtonSet.OK,
      );
      return;
    }

    const startTime = new Date();
    let setupResults = { sheetsCreated: 0, sheetsUpdated: 0, errors: 0 };
    const ss = openTargetSpreadsheet();

    const dashboardResult = createDashboardSheet(ss);
    if (dashboardResult.created) setupResults.sheetsCreated++;
    else setupResults.sheetsUpdated++;

    const logResult = createScannedFilesLogSheet(ss);
    if (logResult.created) setupResults.sheetsCreated++;
    else setupResults.sheetsUpdated++;

    const auditResult = createAuditTrailSheet(ss);
    if (auditResult.created) setupResults.sheetsCreated++;
    else setupResults.sheetsUpdated++;

    const flaggedResult = createFlaggedFilesSheet(ss);
    if (flaggedResult.created) setupResults.sheetsCreated++;
    else setupResults.sheetsUpdated++;

    addSampleDashboardData(ss);

    const processingTime = (new Date() - startTime) / 1000;
    logAudit(
      "System initialization completed",
      `Sheets created: ${setupResults.sheetsCreated}, updated: ${setupResults.sheetsUpdated}, time: ${processingTime}s`,
    );

    const successMessage = `✅ System Setup Complete!\n\n📊 Results:\n• Sheets Created: ${setupResults.sheetsCreated}\n• Sheets Updated: ${setupResults.sheetsUpdated}\n• Processing Time: ${processingTime.toFixed(2)} seconds\n\n🚀 Ready for AI-powered document processing!`;
    ui.alert("Setup Complete", successMessage, ui.ButtonSet.OK);
  } catch (error) {
    logError("System initialization failed", error);
    SpreadsheetApp.getUi().alert(
      "Setup Error",
      "System initialization failed. Check Audit Trail.",
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

// ===================================================================
// SHEET CREATION FUNCTIONS
// ===================================================================

function createDashboardSheet(ss) {
  let created = false;
  let dashboardSheet = ss.getSheetByName("Dashboard");
  if (!dashboardSheet) {
    dashboardSheet = ss.insertSheet("Dashboard");
    created = true;
  }

  const headers = [
    "Department Names",
    "Google Drive ID",
    "Emails",
    "Status",
    "Manager Emails",
  ];
  dashboardSheet.getRange("A1:E1").setValues([headers]);

  const headerRange = dashboardSheet.getRange("A1:E1");
  headerRange
    .setFontWeight("bold")
    .setBackground("#E31E24")
    .setFontColor("white")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontSize(11);

  dashboardSheet.setColumnWidth(1, 200);
  dashboardSheet.setColumnWidth(2, 300);
  dashboardSheet.setColumnWidth(3, 300);
  dashboardSheet.setColumnWidth(4, 120);
  dashboardSheet.setColumnWidth(5, 300);
  dashboardSheet.setFrozenRows(1);

  const statusRange = dashboardSheet.getRange("D2:D");
  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      "Ready",
      "Active",
      "No new files",
      "Error",
      "Disabled",
    ])
    .setAllowInvalid(false)
    .build();
  statusRange.setDataValidation(statusValidation);

  return { created, sheet: dashboardSheet };
}

function createScannedFilesLogSheet(ss) {
  let created = false;
  let logSheet = ss.getSheetByName("Scanned Files Log");
  if (!logSheet) {
    logSheet = ss.insertSheet("Scanned Files Log");
    created = true;
  }

  const headers = [
    "Date Processed",
    "Department",
    "Original File Name",
    "New File Name",
    "View Folder",
    "File Path",
    "File Size",
    "Process Time",
    "AI Used",
    "Status",
  ];
  logSheet.getRange("A1:J1").setValues([headers]);

  const headerRange = logSheet.getRange("A1:J1");
  headerRange
    .setFontWeight("bold")
    .setBackground("#E31E24")
    .setFontColor("white")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontSize(10);

  logSheet.setColumnWidth(1, 150);
  logSheet.setColumnWidth(2, 150);
  logSheet.setColumnWidth(3, 250);
  logSheet.setColumnWidth(4, 250);
  logSheet.setColumnWidth(5, 100);
  logSheet.setColumnWidth(6, 300);
  logSheet.setColumnWidth(7, 80);
  logSheet.setColumnWidth(8, 100);
  logSheet.setColumnWidth(9, 120);
  logSheet.setColumnWidth(10, 100);
  logSheet.setFrozenRows(1);

  return { created, sheet: logSheet };
}

function createAuditTrailSheet(ss) {
  let created = false;
  let auditSheet = ss.getSheetByName("Audit Trail");
  if (!auditSheet) {
    auditSheet = ss.insertSheet("Audit Trail");
    created = true;
  }

  const headers = ["Timestamp", "Action", "Details", "Status", "Error Message"];
  auditSheet.getRange("A1:E1").setValues([headers]);

  const headerRange = auditSheet.getRange("A1:E1");
  headerRange
    .setFontWeight("bold")
    .setBackground("#E31E24")
    .setFontColor("white")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontSize(10);

  auditSheet.setColumnWidth(1, 150);
  auditSheet.setColumnWidth(2, 200);
  auditSheet.setColumnWidth(3, 300);
  auditSheet.setColumnWidth(4, 100);
  auditSheet.setColumnWidth(5, 400);
  auditSheet.setFrozenRows(1);

  return { created, sheet: auditSheet };
}

function createFlaggedFilesSheet(ss) {
  let created = false;
  let flaggedSheet = ss.getSheetByName("Flagged Files");
  if (!flaggedSheet) {
    flaggedSheet = ss.insertSheet("Flagged Files");
    created = true;
  }

  const headers = [
    "Flag Date",
    "Department",
    "File Name",
    "Reason",
    "File ID",
    "View File",
    "Original Location",
    "Status",
    "Notes",
  ];
  flaggedSheet.getRange("A1:I1").setValues([headers]);

  const headerRange = flaggedSheet.getRange("A1:I1");
  headerRange
    .setFontWeight("bold")
    .setBackground("#FBBC04")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontSize(10);

  flaggedSheet.setColumnWidth(1, 150);
  flaggedSheet.setColumnWidth(2, 150);
  flaggedSheet.setColumnWidth(3, 250);
  flaggedSheet.setColumnWidth(4, 200);
  flaggedSheet.setColumnWidth(5, 250);
  flaggedSheet.setColumnWidth(6, 100);
  flaggedSheet.setColumnWidth(7, 300);
  flaggedSheet.setColumnWidth(8, 120);
  flaggedSheet.setColumnWidth(9, 300);

  const statusRange = flaggedSheet.getRange("H2:H");
  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      "Pending Review",
      "Approved",
      "Rejected",
      "Duplicate",
      "Resolved",
    ])
    .setAllowInvalid(false)
    .build();
  statusRange.setDataValidation(statusValidation);

  flaggedSheet.setFrozenRows(1);
  return { created, sheet: flaggedSheet };
}

function addSampleDashboardData(ss) {
  try {
    const dashboardSheet = ss.getSheetByName("Dashboard");
    const dataRange = dashboardSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) {
      const sampleData = [
        [
          "HR Department",
          "YOUR_DRIVE_FOLDER_ID_HERE",
          "hr@company.com",
          "Ready",
          "hr.manager@company.com",
        ],
        [
          "Finance Department",
          "YOUR_DRIVE_FOLDER_ID_HERE",
          "finance@company.com",
          "Ready",
          "finance.manager@company.com",
        ],
      ];
      dashboardSheet
        .getRange(2, 1, sampleData.length, sampleData[0].length)
        .setValues(sampleData);
      dashboardSheet
        .getRange("B2")
        .setNote("Replace with your actual Google Drive folder ID");
      logAudit(
        "Sample data added",
        "Added sample department entries to Dashboard",
      );
    }
  } catch (error) {
    logError("Failed to add sample Dashboard data", error);
  }
}

// ===================================================================
// AI API CONFIGURATION
// ===================================================================

function setGeminiApiKey() {
  try {
    const ui = SpreadsheetApp.getUi();
    const keyResponse = ui.prompt(
      "Set AI API Key",
      "Enter your API key:",
      ui.ButtonSet.OK_CANCEL,
    );
    if (keyResponse.getSelectedButton() === ui.Button.OK) {
      const apiKey = keyResponse.getResponseText().trim();
      if (!apiKey) {
        ui.alert("Setup Cancelled", "No key provided.", ui.ButtonSet.OK);
        return;
      }
      if (apiKey.length < 10) {
        ui.alert(
          "Invalid API Key",
          "Key too short. Please verify.",
          ui.ButtonSet.OK,
        );
        return;
      }
      PropertiesService.getScriptProperties().setProperty(
        "gemini_api_key",
        apiKey,
      );
      PropertiesService.getScriptProperties().setProperty(
        "GEMINI_API_KEY",
        apiKey,
      );
      logAudit(
        "Gemini API Key configured",
        "Gemini 2.0 Flash API key set successfully",
      );
      ui.alert(
        "API Key Saved!",
        "✅ Your Gemini API key has been saved securely.",
        ui.ButtonSet.OK,
      );
    }
  } catch (error) {
    logError("Gemini API key setup failed", error);
  }
}

function testAIConnection() {
  try {
    const ui = SpreadsheetApp.getUi();
    const geminiKey = getGeminiApiKey();

    let testResults = "🧪 GEMINI CONNECTION TEST\n\n";

    if (geminiKey) {
      testResults += "🔵 Gemini:\n";
      try {
        assertGeminiTestCooldown();
        const geminiTest = testGeminiConnection(geminiKey);
        if (geminiTest.success) {
          testResults += `  ✅ Connection successful\n  📊 Response time: ${geminiTest.responseTime}ms\n  🧠 Model: ${getGeminiModel()}\n  📈 Usage Today: ${getGeminiCallsToday()}/${getGeminiDailyCallLimit()}\n\n`;
        } else {
          testResults += `  ❌ Connection failed\n  ⚠️ Error: ${geminiTest.error}\n\n`;
        }
      } catch (error) {
        testResults += `  ❌ Connection failed\n  ⚠️ Error: ${error.message}\n\n`;
      }
    } else {
      testResults += "� Gemini: ❌ No API key configured\n\n";
    }

    ui.alert("Gemini Connection Test", testResults, ui.ButtonSet.OK);
  } catch (error) {
    logError("AI connection test failed", error);
  }
}

function testGeminiConnection(apiKey) {
  const startTime = new Date().getTime();
  try {
    consumeGeminiBudget(1);
    const url = buildGeminiGenerateContentUrl(apiKey);
    const payload = { contents: [{ parts: [{ text: "Respond with: OK" }] }] };
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };
    const response = UrlFetchApp.fetch(url, options);
    const responseTime = new Date().getTime() - startTime;
    const responseCode = response.getResponseCode();
    if (responseCode === 200) {
      return { success: true, responseTime: responseTime };
    } else {
      const errorBody = JSON.parse(response.getContentText());
      return {
        success: false,
        error: errorBody.error?.message || "Unknown error",
        responseTime: responseTime,
      };
    }
  } catch (error) {
    return {
      success: false,
      error: error.message,
      responseTime: new Date().getTime() - startTime,
    };
  }
}

function clearApiKeys() {
  try {
    const ui = SpreadsheetApp.getUi();
    const properties = PropertiesService.getScriptProperties();
    const hasGemini = !!getGeminiApiKey();
    if (!hasGemini) {
      ui.alert("No API Keys Found", "No API keys configured.", ui.ButtonSet.OK);
      return;
    }
    const response = ui.alert(
      "Clear All API Keys",
      "This will remove all API keys. Continue?",
      ui.ButtonSet.YES_NO,
    );
    if (response === ui.Button.YES) {
      let clearedKeys = [];
      if (hasGemini) {
        properties.deleteProperty("gemini_api_key");
        properties.deleteProperty("GEMINI_API_KEY");
        clearedKeys.push("Gemini");
      }
      logAudit("API Keys cleared", `Cleared: ${clearedKeys.join(", ")}`);
      ui.alert(
        "API Keys Cleared",
        `✅ Cleared: ${clearedKeys.join(", ")}`,
        ui.ButtonSet.OK,
      );
    }
  } catch (error) {
    logError("API key clearing failed", error);
  }
}

// ===================================================================
// CORE PROCESSING FUNCTIONS
// ===================================================================

function scanFilesManually() {
  try {
    logAudit("Manual scan initiated", "User triggered manual scan");
    const results = processDepartmentFiles();
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      "Scan Complete",
      `✅ Processing Complete!\n\n• Files Processed: ${results.totalProcessed}\n• Departments: ${results.departmentsProcessed}\n• Files Flagged: ${results.flagged}\n• Errors: ${results.errors}`,
      ui.ButtonSet.OK,
    );
  } catch (error) {
    logError("Manual scan failed", error);
    SpreadsheetApp.getUi().alert(
      "Error",
      "Scan failed. Check Audit Trail.",
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

function processDepartmentFiles() {
  const startTime = new Date();
  let totalProcessed = 0,
    errors = 0,
    departmentsProcessed = 0,
    flagged = 0;
  try {
    const departments = getDepartments();
    logAudit(
      "Processing started",
      `Processing ${departments.length} departments`,
    );
    for (const dept of departments) {
      try {
        const result = processDepartmentFolder(dept);
        totalProcessed += result.processed;
        errors += result.errors;
        flagged += result.flagged;
        if (result.processed > 0 || result.errors > 0) departmentsProcessed++;
        updateDepartmentStatus(
          dept.name,
          result.processed > 0 ? "Active" : "No new files",
        );
      } catch (error) {
        logError(`Department processing failed: ${dept.name}`, error);
        updateDepartmentStatus(dept.name, "Error");
        errors++;
      }
    }
    const processingTime = (new Date() - startTime) / 1000;
    logAudit(
      "Processing completed",
      `Total: ${totalProcessed} files, ${flagged} flagged, ${errors} errors, ${processingTime}s`,
    );
    return { totalProcessed, errors, departmentsProcessed, flagged };
  } catch (error) {
    logError("Main processing failed", error);
    throw error;
  }
}

function processDepartmentFolder(department) {
  let processed = 0,
    errors = 0,
    flagged = 0;
  try {
    const folder = DriveApp.getFolderById(department.driveId);
    const files = folder.getFiles();
    const processedFiles = getProcessedFileIds();
    while (files.hasNext()) {
      const file = files.next();
      if (processedFiles.includes(file.getId())) {
        flagDuplicateFile(file, department, "Already processed previously");
        flagged++;
        continue;
      }
      const fileParents = file.getParents();
      let isInRootFolder = false;
      while (fileParents.hasNext()) {
        if (fileParents.next().getId() === department.driveId) {
          isInRootFolder = true;
          break;
        }
      }
      if (!isInRootFolder) {
        Logger.log(`Skipping file in subfolder: ${file.getName()}`);
        continue;
      }
      const fileCreatedDate = file.getDateCreated();
      const today = new Date();
      const todayFormatted = Utilities.formatDate(
        today,
        TIMEZONE,
        "yyyy-MM-dd",
      );
      const fileCreatedFormatted = Utilities.formatDate(
        fileCreatedDate,
        TIMEZONE,
        "yyyy-MM-dd",
      );
      if (fileCreatedFormatted !== todayFormatted) {
        Logger.log(`Skipping file not uploaded today: ${file.getName()}`);
        continue;
      }
      const mimeType = file.getMimeType();
      if (!SUPPORTED_MIME_TYPES.includes(mimeType)) {
        Logger.log(`Skipping unsupported file type: ${file.getName()}`);
        continue;
      }
      try {
        const success = processFile(file, department, folder);
        if (success) {
          processed++;
          markFileAsProcessed(file.getId());
        } else {
          errors++;
        }
      } catch (error) {
        logError(`File processing failed: ${file.getName()}`, error);
        errors++;
      }
      Utilities.sleep(500);
    }
    return { processed, errors, flagged };
  } catch (error) {
    logError(`Department folder access failed: ${department.name}`, error);
    throw error;
  }
}

function processFile(file, department, rootFolder) {
  const startTime = new Date();
  const originalName = file.getName();
  let aiServiceUsed = "None";
  try {
    debugLog(`\n========== PROCESSING FILE ==========`);
    debugLog(`Original: ${originalName}`);
    debugLog(`MIME Type: ${file.getMimeType()}`);

    const titleResult = extractTitleFromFileFirstPage(file);
    const title = titleResult.title;
    aiServiceUsed = titleResult.aiService;

    debugLog(`Extracted Title: "${title}"`);
    debugLog(`AI Service: ${aiServiceUsed}`);

    const createdDate = file.getDateCreated();
    const formattedDate = Utilities.formatDate(
      createdDate,
      TIMEZONE,
      "yyyy-MM-dd",
    );
    const extension = originalName.substring(originalName.lastIndexOf("."));
    const newFileName = createFileName(formattedDate, title, extension);

    debugLog(`New Filename: ${newFileName}`);

    const year = Utilities.formatDate(createdDate, TIMEZONE, "yyyy");
    const monthName = Utilities.formatDate(createdDate, TIMEZONE, "MMMM");
    const dateMonth = Utilities.formatDate(createdDate, TIMEZONE, "dd MMMM");
    const targetFolder = createSubfolderStructure(
      rootFolder,
      year,
      monthName,
      dateMonth,
    );

    file.setName(newFileName);
    file.moveTo(targetFolder);

    const folderPath = `${department.name}/${year}/${monthName}/${dateMonth}`;
    const fileUrl = `https://drive.google.com/file/d/${file.getId()}/view`;
    const folderUrl = targetFolder.getUrl();
    const processingTime = (new Date() - startTime) / 1000;

    logScannedFileWithUrls(
      new Date(),
      department.name,
      originalName,
      newFileName,
      file.getId(),
      targetFolder.getId(),
      folderUrl,
      folderPath,
      file.getSize(),
      processingTime,
      aiServiceUsed,
      "Success",
      fileUrl,
      folderUrl,
    );
    logAudit(
      "File processed",
      `${originalName} → ${newFileName} (AI: ${aiServiceUsed})`,
    );

    debugLog(`✅ SUCCESS\n=====================================\n`);

    return true;
  } catch (error) {
    logError(`File processing error: ${originalName}`, error);
    const processingTime = (new Date() - startTime) / 1000;
    logScannedFileWithUrls(
      new Date(),
      department.name,
      originalName,
      "Processing Failed",
      file.getId(),
      "",
      "",
      "",
      file.getSize(),
      processingTime,
      aiServiceUsed,
      `Error: ${error.message}`,
      "",
      "",
    );
    return false;
  }
}

// ===================================================================
// AI-POWERED TITLE EXTRACTION - DIRECT PDF PROCESSING WITH GEMINI
// ===================================================================

/**
 * CRITICAL FIX v2.4: Gemini reads PDFs directly, no conversion needed
 */
function extractTitleFromFileFirstPage(file) {
  try {
    const mimeType = file.getMimeType();
    const properties = PropertiesService.getScriptProperties();
    const geminiKey = getGeminiApiKey();

    debugLog(`\n--- TITLE EXTRACTION START ---`);
    debugLog(`File: ${file.getName()}`);
    debugLog(`MIME: ${mimeType}`);
    debugLog(`Gemini Key: ${geminiKey ? "Present" : "Missing"}`);

    if (geminiKey && hasGeminiBudget(1)) {
      try {
        debugLog("🔵 Attempting Gemini extraction (PRIMARY)...");
        const result = extractTitleWithGemini(file, geminiKey, mimeType);
        if (result && result.title && result.title.length > 3) {
          const geminiTitle = formatTitleWithSpaces(result.title);
          if (!isLikelyIncompleteTitle(geminiTitle)) {
            debugLog(`✅ Gemini SUCCESS: "${geminiTitle}"`);
            return { title: geminiTitle, aiService: "Gemini 2.0 Flash" };
          }
          debugLog(
            `⚠️ Gemini title looks incomplete ("${geminiTitle}") - trying OCR fallback for fuller title`,
          );
        } else {
          debugLog(
            `⚠️ Gemini returned short/empty title: "${result ? result.title : "null"}"`,
          );
        }
      } catch (geminiError) {
        debugLog(`❌ Gemini FAILED: ${geminiError.message}`);
        logError("Gemini extraction failed", geminiError);
      }
    } else if (!geminiKey) {
      debugLog("⚠️ Gemini API key not configured");
    } else {
      debugLog(
        `⚠️ Gemini daily limit reached (${getGeminiCallsToday()}/${getGeminiDailyCallLimit()}) - skipping Gemini`,
      );
    }

    // OCR fallback
    debugLog("⚠️ Falling back to OCR...");
    if (mimeType === "application/pdf") {
      const title = extractTitleFromPDF(file);
      debugLog(`OCR Result: "${title}"`);
      return { title: title, aiService: "OCR Fallback" };
    } else if (mimeType.startsWith("image/")) {
      const title = extractTitleFromImage(file);
      debugLog(`OCR Result: "${title}"`);
      return { title: title, aiService: "OCR Fallback" };
    }

    // Final fallback to filename
    const fallbackTitle = sanitizeFileName(
      file.getName().replace(/\.[^/.]+$/, ""),
    );
    debugLog(`⚠️ Using filename fallback: "${fallbackTitle}"`);
    return { title: fallbackTitle, aiService: "Filename Fallback" };
  } catch (error) {
    logError(`Title extraction failed for ${file.getName()}`, error);
    return {
      title: sanitizeFileName(file.getName().replace(/\.[^/.]+$/, "")),
      aiService: "Error Fallback",
    };
  }
}

/**
 * CRITICAL FIX v2.4: Gemini now processes PDFs DIRECTLY using File API
 * No more PNG conversion needed!
 */
function extractTitleWithGemini(file, apiKey, mimeType) {
  try {
    debugLog(`\n--- GEMINI EXTRACTION (DIRECT PDF) ---`);

    consumeGeminiBudget(1);

    let fileContent, contentMimeType;

    if (mimeType === "application/pdf") {
      debugLog("📄 PDF detected - reading directly with Gemini File API...");

      // Get PDF as base64
      const pdfBlob = file.getBlob();
      const pdfBytes = pdfBlob.getBytes();
      fileContent = Utilities.base64Encode(pdfBytes);
      contentMimeType = "application/pdf";

      debugLog(`PDF size: ${pdfBytes.length} bytes`);
      debugLog(`Base64 length: ${fileContent.length} chars`);
    } else if (mimeType.startsWith("image/")) {
      debugLog("🖼️ Image detected - encoding to base64...");
      fileContent = Utilities.base64Encode(file.getBlob().getBytes());
      contentMimeType = mimeType;
      debugLog(`Image content type: ${contentMimeType}`);
    } else {
      throw new Error(`Unsupported file type: ${mimeType}`);
    }

    const url = buildGeminiGenerateContentUrl(apiKey);

    // IMPROVED PROMPT - optimized for complete legal heading extraction
    const prompt = `You are analyzing a legal/administrative document (PDF or image). Extract the FULL formal title from the FIRST PAGE.

STRICT RULES:
1. Copy the title text as written (do NOT summarize, paraphrase, or shorten it).
2. Prefer the complete legal heading (example: "Motion For Issuance Of A Writ Of Execution").
3. If the title is wrapped on multiple lines, combine those lines into one complete line.
4. Keep meaningful words after "for", "of", "to", and "in" (do not cut the phrase early).
5. Return a single clean title line, usually 5-20 words (allow up to 25 when needed).
6. Remove only obvious metadata/noise (file extensions, standalone dates, trailing reference codes).
7. Return ONLY the title text (no quotes, numbering, or explanation).

EXAMPLES:
BAD: "Motion For Lss"
GOOD: "Motion For Issuance Of A Writ Of Execution"

BAD: "Employment Contract"
GOOD: "Employment Contract For Project-Based Personnel"

BAD: "Document_2026-02-16_Re_Motion.pdf"
GOOD: "Motion For Issuance Of A Writ Of Execution"

Extract the complete title now:`;

    const payload = {
      contents: [
        {
          parts: [
            { text: prompt },
            {
              inline_data: {
                mime_type: contentMimeType,
                data: fileContent,
              },
            },
          ],
        },
      ],
      generationConfig: {
        temperature: 0.1,
        maxOutputTokens: 90,
        topP: 0.8,
        topK: 10,
      },
    };

    debugLog("📤 Sending request to Gemini API...");
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    debugLog(`📥 Response code: ${responseCode}`);
    debugLog(`Response preview: ${responseText.substring(0, 300)}...`);

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseText);

      if (jsonResponse.candidates && jsonResponse.candidates.length > 0) {
        const candidate = jsonResponse.candidates[0];

        // Check if content was filtered
        if (
          candidate.finishReason === "SAFETY" ||
          candidate.finishReason === "RECITATION"
        ) {
          debugLog(`⚠️ Gemini filtered content: ${candidate.finishReason}`);
          throw new Error(`Content filtered: ${candidate.finishReason}`);
        }

        if (
          candidate.content &&
          candidate.content.parts &&
          candidate.content.parts.length > 0
        ) {
          let title = candidate.content.parts[0].text.trim();
          debugLog(`Raw Gemini response: "${title}"`);

          title = cleanExtractedTitle(title);
          debugLog(`After cleaning: "${title}"`);

          title = formatTitleWithSpaces(title);
          debugLog(`Final formatted title: "${title}"`);

          if (title.length < 3) {
            throw new Error("Title too short after cleaning");
          }

          return { title: title };
        }
      }

      debugLog(`❌ No valid candidates in response`);
    } else {
      debugLog(`❌ API Error ${responseCode}: ${responseText}`);
    }

    throw new Error(
      `Gemini API returned no valid response (code: ${responseCode})`,
    );
  } catch (error) {
    debugLog(`❌ Gemini extraction error: ${error.message}`);
    throw error;
  }
}

/**
 * Clean extracted title
 */
function cleanExtractedTitle(title) {
  if (!title) return "";

  title = title.replace(/^[\"']|[\"']$/g, "");
  title = title.replace(/^(Title:|Subject:|Re:|Document:|File:|Name:)\s*/i, "");
  title = title.replace(/\b\d{4}-\d{2}-\d{2}\b/g, "");
  title = title.replace(/\b\d{1,2}\/\d{1,2}\/\d{2,4}\b/g, "");
  title = title.replace(/\.(pdf|jpg|jpeg|png|gif|bmp|doc|docx)$/i, "");
  title = title.replace(/\s+/g, " ").trim();
  title = title.replace(/^\d+\.\s*/, "");

  if (title.length < 3) throw new Error("Title too short after cleaning");

  return title;
}

/**
 * Detects likely truncated titles (e.g., "Motion For Lss")
 */
function isLikelyIncompleteTitle(title) {
  if (!title) return true;

  const normalized = title.toString().replace(/\s+/g, " ").trim();
  if (!normalized) return true;

  const words = normalized.split(" ").filter((word) => word);
  if (words.length <= 2) return true;

  const firstWord = (words[0] || "").toLowerCase();
  const secondWord = (words[1] || "").toLowerCase();
  const thirdWord = (words[2] || "").toLowerCase();

  const legalLeadWords = [
    "motion",
    "petition",
    "application",
    "request",
    "manifestation",
    "notice",
    "affidavit",
    "complaint",
    "answer",
  ];
  const hasLegalLeadWord = legalLeadWords.indexOf(firstWord) !== -1;

  if (/\b(for|of|to|in|on|re)\s*$/i.test(normalized)) return true;

  if (hasLegalLeadWord && words.length <= 2) return true;

  if (
    hasLegalLeadWord &&
    (secondWord === "for" || secondWord === "of") &&
    words.length === 3 &&
    thirdWord.length <= 4
  ) {
    return true;
  }

  return false;
}

/**
 * Format title with spaces (no underscores)
 */
function formatTitleWithSpaces(title) {
  if (!title || title === "") {
    return "Unknown Document";
  }

  return title
    .toString()
    .replace(/[<>:"/\\|?*]/g, "")
    .replace(/_+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .split(" ")
    .map((word) => {
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    })
    .join(" ")
    .substring(0, 150);
}

/**
 * OCR fallback for PDFs
 */
function extractTitleFromPDF(file) {
  let tempDocId = null;
  try {
    debugLog(`\n--- PDF OCR FALLBACK ---`);
    const blob = file.getBlob();
    const parentFolder = file.getParents().next();
    const tempName = `temp_ocr_${Date.now()}`;
    const resource = {
      title: tempName,
      parents: [{ id: parentFolder.getId() }],
    };

    const convertedFile = Drive.Files.insert(resource, blob, {
      ocr: true,
      ocrLanguage: "en",
    });
    tempDocId = convertedFile.id;

    Utilities.sleep(3000);

    const tempDoc = DocumentApp.openById(tempDocId);
    const text = tempDoc.getBody().getText();

    DriveApp.getFileById(tempDocId).setTrashed(true);

    const extractedTitle = extractBestTitle(text, file.getName());
    const formattedTitle = formatTitleWithSpaces(extractedTitle);

    debugLog(`OCR extracted title: "${formattedTitle}"`);
    return formattedTitle;
  } catch (error) {
    if (tempDocId) {
      try {
        DriveApp.getFileById(tempDocId).setTrashed(true);
      } catch (e) {}
    }
    logError(`PDF OCR failed for ${file.getName()}`, error);
    return formatTitleWithSpaces(file.getName().replace(/\.[^/.]+$/, ""));
  }
}

/**
 * OCR fallback for images
 */
function extractTitleFromImage(file) {
  let tempDocId = null;
  try {
    debugLog(`\n--- IMAGE OCR FALLBACK ---`);
    const blob = file.getBlob();
    const parentFolder = file.getParents().next();
    const tempName = `temp_ocr_img_${Date.now()}`;
    const resource = {
      title: tempName,
      parents: [{ id: parentFolder.getId() }],
    };

    const convertedFile = Drive.Files.insert(resource, blob, {
      ocr: true,
      ocrLanguage: "en",
    });
    tempDocId = convertedFile.id;

    Utilities.sleep(3000);

    const tempDoc = DocumentApp.openById(tempDocId);
    const text = tempDoc.getBody().getText();

    DriveApp.getFileById(tempDocId).setTrashed(true);

    const extractedTitle = extractBestTitle(text, file.getName());
    const formattedTitle = formatTitleWithSpaces(extractedTitle);

    debugLog(`OCR extracted title: "${formattedTitle}"`);
    return formattedTitle;
  } catch (error) {
    if (tempDocId) {
      try {
        DriveApp.getFileById(tempDocId).setTrashed(true);
      } catch (e) {}
    }
    logError(`Image OCR failed for ${file.getName()}`, error);
    return formatTitleWithSpaces(file.getName().replace(/\.[^/.]+$/, ""));
  }
}

/**
 * Extract best title from OCR text
 */
function extractBestTitle(text, fallbackName) {
  if (!text || text.trim().length === 0) {
    return sanitizeFileName(fallbackName.replace(/\.[^/.]+$/, ""));
  }

  const lines = text
    .split("\n")
    .map((l) => l.trim())
    .filter((l) => l.length > 0);
  if (lines.length === 0)
    return sanitizeFileName(fallbackName.replace(/\.[^/.]+$/, ""));

  const maxLinesToCheck = Math.min(lines.length, 12);
  const candidates = [];

  for (let i = 0; i < maxLinesToCheck; i++) {
    const line1 = lines[i];
    candidates.push({ text: line1, startIndex: i, span: 1 });

    if (i + 1 < maxLinesToCheck) {
      const line2 = `${line1} ${lines[i + 1]}`.replace(/\s+/g, " ").trim();
      if (line2.length <= 140) {
        candidates.push({ text: line2, startIndex: i, span: 2 });
      }
    }

    if (i + 2 < maxLinesToCheck) {
      const line3 = `${line1} ${lines[i + 1]} ${lines[i + 2]}`
        .replace(/\s+/g, " ")
        .trim();
      if (line3.length <= 170) {
        candidates.push({ text: line3, startIndex: i, span: 3 });
      }
    }
  }

  let bestCandidate = null;
  let bestScore = 0;

  for (let i = 0; i < candidates.length; i++) {
    const candidate = candidates[i];
    const line = candidate.text;

    if (line.length < 5 || line.length > 170) continue;

    const wordCount = line.split(/\s+/).length;
    if (wordCount < 3 || wordCount > 24) continue;

    const letterCount = (line.match(/[a-zA-Z]/g) || []).length;
    const letterRatio = letterCount / line.length;

    let score = 0;

    if (wordCount >= 5 && wordCount <= 20) score += 4;
    else if (wordCount >= 3 && wordCount <= 24) score += 2;

    if (letterRatio >= 0.55) score += 2;
    if (!/^\d/.test(line)) score += 1;

    if (candidate.startIndex < 3) score += 2;
    else if (candidate.startIndex < 6) score += 1;

    if (candidate.span > 1) score += 2;

    if (
      /\b(motion|petition|application|writ|complaint|affidavit|agreement|contract|notice|deed)\b/i.test(
        line,
      )
    ) {
      score += 2;
    }

    if (/^(page|from|to|date|subject|re:|cc:|email|reference)/i.test(line)) {
      score -= 4;
    }

    if (score > bestScore) {
      bestScore = score;
      bestCandidate = line;
    }
  }

  if (bestCandidate && bestScore > 0) return sanitizeFileName(bestCandidate);

  const firstGoodLine = lines.find((l) => l.length >= 5 && l.length <= 140);
  if (firstGoodLine) return sanitizeFileName(firstGoodLine);

  return sanitizeFileName(fallbackName.replace(/\.[^/.]+$/, ""));
}

// ===================================================================
// FILE NAMING & ORGANIZATION
// ===================================================================

function createFileName(date, title, extension) {
  if (!extension.startsWith(".")) extension = "." + extension;
  title = formatTitleWithSpaces(title);
  const baseName = `${date} ${title}`;
  let fileName = baseName + extension;
  if (fileName.length > MAX_FILENAME_LENGTH) {
    const maxBaseLength = MAX_FILENAME_LENGTH - extension.length - 3;
    fileName = baseName.substring(0, maxBaseLength) + "..." + extension;
  }
  debugLog(`Created filename: "${fileName}"`);
  return sanitizeFileName(fileName);
}

function sanitizeFileName(fileName) {
  return fileName
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "")
    .replace(/\s+/g, " ")
    .replace(/\.{2,}/g, ".")
    .trim();
}

function createSubfolderStructure(rootFolder, year, monthName, dateMonth) {
  try {
    let yearFolder = getFolderByName(rootFolder, year);
    if (!yearFolder) yearFolder = rootFolder.createFolder(year);
    let monthFolder = getFolderByName(yearFolder, monthName);
    if (!monthFolder) monthFolder = yearFolder.createFolder(monthName);
    let dateFolder = getFolderByName(monthFolder, dateMonth);
    if (!dateFolder) dateFolder = monthFolder.createFolder(dateMonth);
    return dateFolder;
  } catch (error) {
    logError(
      `Subfolder creation failed: ${year}/${monthName}/${dateMonth}`,
      error,
    );
    throw error;
  }
}

function getFolderByName(parentFolder, name) {
  const folders = parentFolder.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : null;
}

// ===================================================================
// DEBUG & TESTING UTILITIES
// ===================================================================

function debugLog(message) {
  if (DEBUG_MODE) {
    Logger.log(`[DEBUG] ${message}`);
  }
}

function testSingleFileExtraction() {
  try {
    const ui = SpreadsheetApp.getUi();
    const fileIdResponse = ui.prompt(
      "Test Title Extraction",
      "Enter the Google Drive File ID to test:",
      ui.ButtonSet.OK_CANCEL,
    );

    if (fileIdResponse.getSelectedButton() !== ui.Button.OK) return;

    const fileId = fileIdResponse.getResponseText().trim();
    if (!fileId) {
      ui.alert(
        "No File ID Provided",
        "Please enter a valid file ID.",
        ui.ButtonSet.OK,
      );
      return;
    }

    const file = DriveApp.getFileById(fileId);
    const mimeType = file.getMimeType();

    debugLog(`\n========== TEST EXTRACTION ==========`);
    debugLog(`File: ${file.getName()}`);
    debugLog(`MIME: ${mimeType}`);

    const result = extractTitleFromFileFirstPage(file);

    let resultMessage = `📄 File: ${file.getName()}\n`;
    resultMessage += `📋 MIME Type: ${mimeType}\n\n`;
    resultMessage += `✨ Extracted Title:\n"${result.title}"\n\n`;
    resultMessage += `🤖 AI Service: ${result.aiService}\n\n`;
    resultMessage += `Check the Apps Script logs (View → Logs) for detailed debug information.`;

    ui.alert("Extraction Test Result", resultMessage, ui.ButtonSet.OK);
  } catch (error) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      "Test Failed",
      `Error: ${error.message}\n\nCheck Audit Trail for details.`,
      ui.ButtonSet.OK,
    );
    logError("Test extraction failed", error);
  }
}

// ===================================================================
// REMAINING FUNCTIONS
// ===================================================================

function flagDuplicateFile(file, department, reason) {
  try {
    const ss = openTargetSpreadsheet();
    const flaggedSheet = ss.getSheetByName("Flagged Files");
    if (!flaggedSheet) return;
    const fileUrl = `https://drive.google.com/file/d/${file.getId()}/view`;
    const fileLink = `=HYPERLINK("${fileUrl}", "View File")`;
    const parents = file.getParents();
    let originalLocation = "Unknown";
    if (parents.hasNext()) originalLocation = parents.next().getName();
    flaggedSheet.appendRow([
      new Date(),
      department.name,
      file.getName(),
      reason,
      file.getId(),
      fileLink,
      originalLocation,
      "Pending Review",
      "",
    ]);
    logAudit("File flagged", `${file.getName()} - ${reason}`);
  } catch (error) {
    logError("Failed to flag file", error);
  }
}

function reviewFlaggedFiles() {
  try {
    const ui = SpreadsheetApp.getUi();
    const ss = openTargetSpreadsheet();
    const flaggedSheet = ss.getSheetByName("Flagged Files");
    if (!flaggedSheet) {
      ui.alert(
        "Sheet Not Found",
        "Run Initialize Setup first.",
        ui.ButtonSet.OK,
      );
      return;
    }
    const numRows = flaggedSheet.getDataRange().getNumRows();
    if (numRows <= 1) {
      ui.alert(
        "No Flagged Files",
        "No files flagged for review.",
        ui.ButtonSet.OK,
      );
      return;
    }
    const pendingCount = flaggedSheet
      .getRange("H2:H")
      .getValues()
      .filter((row) => row[0] === "Pending Review").length;
    ui.alert(
      "Flagged Files Review",
      `📋 Total: ${numRows - 1}\n⏳ Pending: ${pendingCount}\n\nReview the Flagged Files sheet.`,
      ui.ButtonSet.OK,
    );
    ss.setActiveSheet(flaggedSheet);
  } catch (error) {
    logError("Review failed", error);
  }
}

function getDepartments() {
  try {
    const ss = openTargetSpreadsheet();
    const sheet = ss.getSheetByName("Dashboard");
    const data = sheet.getDataRange().getValues();
    const departments = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] && row[1]) {
        departments.push({
          name: row[0].toString().trim(),
          driveId: row[1].toString().trim(),
          emails: row[2] ? row[2].toString().trim() : "",
          status: row[3] ? row[3].toString().trim() : "",
          managerEmails: row[4] ? row[4].toString().trim() : "",
        });
      }
    }
    return departments;
  } catch (error) {
    logError("Failed to get departments", error);
    throw error;
  }
}

function updateDepartmentStatus(departmentName, status) {
  try {
    const ss = openTargetSpreadsheet();
    const sheet = ss.getSheetByName("Dashboard");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === departmentName) {
        sheet.getRange(i + 1, 4).setValue(status);
        sheet.getRange(i + 1, 4).setNote(`Last updated: ${new Date()}`);
        break;
      }
    }
  } catch (error) {
    logError(`Failed to update status for ${departmentName}`, error);
  }
}

function formatFileSize(bytes) {
  if (bytes === 0) return "0 B";
  const k = 1024;
  const sizes = ["B", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
}

function getProcessedFileIds() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const processedIds = properties.getProperty("processedFileIds");
    return processedIds ? JSON.parse(processedIds) : [];
  } catch (error) {
    logError("Failed to get processed file IDs", error);
    return [];
  }
}

function markFileAsProcessed(fileId) {
  try {
    const properties = PropertiesService.getScriptProperties();
    let processedIds = getProcessedFileIds();
    if (!processedIds.includes(fileId)) {
      processedIds.push(fileId);
      if (processedIds.length > 1000) processedIds = processedIds.slice(-1000);
      properties.setProperty("processedFileIds", JSON.stringify(processedIds));
    }
  } catch (error) {
    logError("Failed to mark file as processed", error);
  }
}

function logScannedFileWithUrls(
  dateProcessed,
  department,
  originalName,
  newName,
  fileId,
  folderId,
  folderUrl,
  filePath,
  fileSize,
  processTime,
  aiUsed,
  status,
  fileUrl,
  folderUrlForEmail,
) {
  try {
    const ss = openTargetSpreadsheet();
    const sheet = ss.getSheetByName("Scanned Files Log");
    const fileLink = fileId
      ? `=HYPERLINK("${fileUrl}", "${newName}")`
      : newName;
    const folderLink = folderId
      ? `=HYPERLINK("${folderUrlForEmail}", "View Folder")`
      : "N/A";
    const formattedSize = formatFileSize(fileSize);
    const formattedTime = `${processTime.toFixed(2)}s`;
    sheet.appendRow([
      dateProcessed,
      department,
      originalName,
      fileLink,
      folderLink,
      filePath,
      formattedSize,
      formattedTime,
      aiUsed,
      status,
    ]);
    const rowNumber = sheet.getLastRow();
    const today = Utilities.formatDate(dateProcessed, TIMEZONE, "yyyy-MM-dd");
    const properties = PropertiesService.getScriptProperties();
    const emailDataKey = `email_${today}_row_${rowNumber}`;
    const emailData = {
      department: department,
      newFileName: newName,
      fileUrl: fileUrl,
      folderUrl: folderUrlForEmail,
      fileSize: formattedSize,
      aiUsed: aiUsed,
      rowNumber: rowNumber,
    };
    properties.setProperty(emailDataKey, JSON.stringify(emailData));
    const todaysRowsKey = `processed_rows_${today}`;
    let todaysRows = [];
    const existingRows = properties.getProperty(todaysRowsKey);
    if (existingRows) todaysRows = JSON.parse(existingRows);
    todaysRows.push(rowNumber);
    properties.setProperty(todaysRowsKey, JSON.stringify(todaysRows));
  } catch (error) {
    logError("Failed to log scanned file", error);
  }
}

function logAudit(action, details, status = "Success", errorMessage = "") {
  try {
    const ss = openTargetSpreadsheet();
    const sheet = ss.getSheetByName("Audit Trail");
    sheet.appendRow([new Date(), action, details, status, errorMessage]);
  } catch (error) {
    Logger.log(`Audit logging failed: ${error.message}`);
  }
}

function logError(action, error) {
  const errorMessage = error.message || error.toString();
  const stack = error.stack || "No stack trace";
  Logger.log(`ERROR - ${action}: ${errorMessage}`);
  logAudit(
    action,
    `Error: ${errorMessage}`,
    "Error",
    `${errorMessage}\n\nStack: ${stack}`,
  );
}

function getTodaysProcessedFilesWithDirectUrls(date) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const todaysRowsKey = `processed_rows_${date}`;
    const todaysRowsData = properties.getProperty(todaysRowsKey);
    if (!todaysRowsData) {
      Logger.log(`No processed files for ${date}`);
      return [];
    }
    const todaysRows = JSON.parse(todaysRowsData);
    const todaysFiles = [];
    todaysRows.forEach((rowNumber) => {
      const emailDataKey = `email_${date}_row_${rowNumber}`;
      const emailDataString = properties.getProperty(emailDataKey);
      if (emailDataString) {
        try {
          const emailData = JSON.parse(emailDataString);
          todaysFiles.push({
            department: emailData.department,
            newFileName: emailData.newFileName,
            fileUrl: emailData.fileUrl,
            folderUrl: emailData.folderUrl,
            fileSize: emailData.fileSize,
            aiUsed: emailData.aiUsed || "N/A",
          });
        } catch (e) {
          Logger.log(`Failed to parse email data for row ${rowNumber}`);
        }
      }
    });
    Logger.log(`Found ${todaysFiles.length} files for ${date}`);
    return todaysFiles;
  } catch (error) {
    logError("Failed to get today files", error);
    return [];
  }
}

function cleanupOldEmailData() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const allProperties = properties.getProperties();
    const today = new Date();
    const cutoffDate = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
    let deletedCount = 0;
    Object.keys(allProperties).forEach((key) => {
      if (key.startsWith("email_") || key.startsWith("processed_rows_")) {
        const dateMatch = key.match(/(\d{4}-\d{2}-\d{2})/);
        if (dateMatch) {
          const keyDate = new Date(dateMatch[1]);
          if (keyDate < cutoffDate) {
            properties.deleteProperty(key);
            deletedCount++;
          }
        }
      }
    });
    if (deletedCount > 0) {
      logAudit("Email data cleanup", `Deleted ${deletedCount} old entries`);
    }
  } catch (error) {
    logError("Failed to cleanup email data", error);
  }
}

/**
 * UPDATED: Scheduled processing with weekly digest support
 */
function scheduledProcessing() {
  try {
    const today = new Date();
    const dayOfWeek = today.getDay();

    // Skip weekends
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      Logger.log("Weekend - skipping");
      return;
    }

    // Regular daily processing
    logAudit(
      "Scheduled processing started",
      `Automated processing on ${today}`,
    );
    const results = processDepartmentFiles();
    logAudit(
      "Scheduled processing completed",
      `Processed ${results.totalProcessed} files`,
    );

    // Send weekly digest if it's the configured day
    if (WEEKLY_DIGEST_ENABLED && dayOfWeek === WEEKLY_DIGEST_DAY) {
      Logger.log("It's weekly digest day - sending digest");
      scheduledWeeklyDigest();
    }
  } catch (error) {
    logError("Scheduled processing failed", error);
  }
}

/**
 * ENHANCED v2025-2: Automatic Trigger Configuration
 * Sets up 5:15 PM Monday-Friday automation with enable/disable options
 */
function configureSchedule() {
  try {
    const ui = SpreadsheetApp.getUi();

    // Check if trigger already exists
    const existingTrigger = getScheduledTrigger();

    if (existingTrigger) {
      // Trigger exists - show options to disable or reconfigure
      const response = ui.alert(
        "⚙️ Schedule Configuration",
        "📅 Current Schedule: 5:15 PM, Monday-Friday\n" +
          "✅ Status: ACTIVE\n\n" +
          "What would you like to do?",
        ui.ButtonSet.YES_NO_CANCEL,
      );

      if (response === ui.Button.YES) {
        // YES = Disable trigger
        disableScheduledTrigger(existingTrigger);
        logAudit("Schedule disabled", "User disabled automatic scheduling");
        ui.alert(
          "✅ Schedule Disabled",
          "⏸️ Automatic processing has been disabled.\n\n" +
            "Files will no longer be processed automatically at 5:15 PM.\n\n" +
            "You can re-enable it anytime from the menu.",
          ui.ButtonSet.OK,
        );
      } else if (response === ui.Button.NO) {
        // NO = Keep current schedule
        ui.alert(
          "ℹ️ No Changes Made",
          "📅 Schedule remains: 5:15 PM, Monday-Friday\n" + "✅ Status: ACTIVE",
          ui.ButtonSet.OK,
        );
      }
      // CANCEL = Do nothing
    } else {
      // No trigger exists - offer to enable
      const response = ui.alert(
        "⚙️ Configure Automatic Schedule",
        "📅 Proposed Schedule:\n" +
          "   • Time: 5:15 PM (GMT+8)\n" +
          "   • Days: Monday - Friday\n" +
          "   • Action: Scan files and send emails\n\n" +
          "Enable automatic processing?",
        ui.ButtonSet.YES_NO,
      );

      if (response === ui.Button.YES) {
        enableScheduledTrigger();
        logAudit(
          "Schedule enabled",
          "Automatic scheduling enabled: 5:15 PM Mon-Fri",
        );
        ui.alert(
          "✅ Schedule Enabled!",
          "🎉 Automatic processing is now active!\n\n" +
            "📅 Schedule: 5:15 PM (GMT+8), Monday-Friday\n" +
            "⚙️ Process: Scan files → Rename → Send emails\n\n" +
            "You can disable it anytime from the menu.",
          ui.ButtonSet.OK,
        );
      } else {
        ui.alert(
          "ℹ️ Schedule Not Enabled",
          "Automatic processing remains disabled.\n\n" +
            "You can enable it anytime from:\n" +
            "⚙️ Configure Schedule",
          ui.ButtonSet.OK,
        );
      }
    }
  } catch (error) {
    logError("Schedule configuration failed", error);
    SpreadsheetApp.getUi().alert(
      "❌ Configuration Error",
      "Failed to configure schedule. Check Audit Trail for details.",
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

/**
 * ENABLE: Dedicated function to enable schedule from menu
 */
function enableScheduleFromMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    const existingTrigger = getScheduledTrigger();

    // Check if already enabled
    if (existingTrigger) {
      ui.alert(
        "⚠️ Schedule Already Enabled",
        "✅ Automatic processing is currently ACTIVE\n\n" +
          "📅 Schedule: 5:15 PM (GMT+8), Monday-Friday\n" +
          "⚙️ Status: Running\n\n" +
          "To disable, use: ⚙️ Schedule Settings → ⏸️ Disable Auto Schedule",
        ui.ButtonSet.OK,
      );
      return;
    }

    // Show confirmation dialog
    const response = ui.alert(
      "▶️ Enable Automatic Schedule",
      "📅 This will enable automatic file processing:\n\n" +
        "🕐 Time: 5:15 PM (GMT+8)\n" +
        "📆 Days: Monday - Friday\n" +
        "⚙️ Action: Scan files → Rename → Organize → Send emails\n\n" +
        "✅ Enable automatic processing?",
      ui.ButtonSet.YES_NO,
    );

    if (response === ui.Button.YES) {
      // Enable the schedule
      const success = enableScheduledTrigger();

      if (success) {
        logAudit(
          "Schedule enabled",
          "User enabled automatic scheduling: 5:15 PM Mon-Fri",
        );

        ui.alert(
          "✅ Schedule Enabled Successfully!",
          "🎉 Automatic processing is now ACTIVE!\n\n" +
            "📅 Schedule: 5:15 PM (GMT+8), Monday-Friday\n" +
            "⚙️ Process Flow:\n" +
            "   1. Scan uploaded files\n" +
            "   2. Extract descriptive titles with AI\n" +
            "   3. Rename and organize files\n" +
            "   4. Send daily email reports\n" +
            "   5. Send weekly digest (Fridays)\n\n" +
            "💡 You can disable anytime from:\n" +
            "   ⚙️ Schedule Settings → ⏸️ Disable Auto Schedule",
          ui.ButtonSet.OK,
        );

        // Refresh menu to show updated status
        onOpen();
      } else {
        ui.alert(
          "❌ Failed to Enable Schedule",
          "Could not create the scheduled trigger.\n\n" +
            "Please check the Audit Trail for error details.",
          ui.ButtonSet.OK,
        );
      }
    } else {
      ui.alert(
        "ℹ️ Cancelled",
        "Schedule was not enabled.\n\n" +
          "Automatic processing remains disabled.",
        ui.ButtonSet.OK,
      );
    }
  } catch (error) {
    logError("Failed to enable schedule from menu", error);
    SpreadsheetApp.getUi().alert(
      "❌ Error",
      "Failed to enable schedule. Check Audit Trail for details.",
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

/**
 * DISABLE: Dedicated function to disable schedule from menu
 */
function disableScheduleFromMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    const existingTrigger = getScheduledTrigger();

    // Check if already disabled
    if (!existingTrigger) {
      ui.alert(
        "ℹ️ Schedule Already Disabled",
        "⏸️ Automatic processing is currently INACTIVE\n\n" +
          "📅 No scheduled automation running\n\n" +
          "To enable, use: ⚙️ Schedule Settings → ▶️ Enable Auto Schedule",
        ui.ButtonSet.OK,
      );
      return;
    }

    // Show confirmation dialog with warning
    const response = ui.alert(
      "⏸️ Disable Automatic Schedule",
      "⚠️ This will STOP automatic file processing:\n\n" +
        "📅 Current Schedule: 5:15 PM, Monday-Friday\n" +
        "✅ Status: Currently ACTIVE\n\n" +
        "❌ After disabling:\n" +
        "   • Files will NOT be processed automatically\n" +
        "   • No daily emails will be sent\n" +
        "   • Weekly digest will NOT be sent\n" +
        "   • You must process files manually\n\n" +
        "⚠️ Are you sure you want to disable?",
      ui.ButtonSet.YES_NO,
    );

    if (response === ui.Button.YES) {
      // Disable the schedule
      const success = disableScheduledTrigger(existingTrigger);

      if (success) {
        logAudit("Schedule disabled", "User disabled automatic scheduling");

        ui.alert(
          "✅ Schedule Disabled Successfully",
          "⏸️ Automatic processing has been STOPPED\n\n" +
            "📅 Previous Schedule: 5:15 PM, Monday-Friday\n" +
            "❌ Status: Now INACTIVE\n\n" +
            "💡 What this means:\n" +
            "   • No automatic file processing\n" +
            "   • No automatic emails\n" +
            "   • You must use manual options:\n" +
            "     → 🔍 Scan Files Now\n" +
            "     → 📧 Send Daily Emails\n\n" +
            "💡 To re-enable automatic processing:\n" +
            "   ⚙️ Schedule Settings → ▶️ Enable Auto Schedule",
          ui.ButtonSet.OK,
        );

        // Refresh menu to show updated status
        onOpen();
      } else {
        ui.alert(
          "❌ Failed to Disable Schedule",
          "Could not delete the scheduled trigger.\n\n" +
            "Please check the Audit Trail for error details.",
          ui.ButtonSet.OK,
        );
      }
    } else {
      ui.alert(
        "ℹ️ Cancelled",
        "Schedule remains ACTIVE.\n\n" +
          "✅ Automatic processing continues at 5:15 PM, Mon-Fri",
        ui.ButtonSet.OK,
      );
    }
  } catch (error) {
    logError("Failed to disable schedule from menu", error);
    SpreadsheetApp.getUi().alert(
      "❌ Error",
      "Failed to disable schedule. Check Audit Trail for details.",
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

/**
 * Enable automatic scheduled trigger at 5:15 PM Monday-Friday
 */
function enableScheduledTrigger() {
  try {
    // Delete existing trigger if any (cleanup)
    const existingTrigger = getScheduledTrigger();
    if (existingTrigger) {
      ScriptApp.deleteTrigger(existingTrigger);
      Logger.log("Deleted existing trigger before creating new one");
    }

    // Create new trigger for 5:15 PM daily
    ScriptApp.newTrigger("scheduledProcessing")
      .timeBased()
      .atHour(17) // 5 PM hour (24-hour format)
      .nearMinute(15) // At 15 minutes (5:15 PM)
      .everyDays(1) // Every day
      .create();

    Logger.log(
      "✅ Scheduled trigger created: 5:15 PM daily (Mon-Fri filtering in function)",
    );
    return true;
  } catch (error) {
    logError("Failed to enable scheduled trigger", error);
    return false;
  }
}

/**
 * Disable (delete) the scheduled trigger
 */
function disableScheduledTrigger(trigger) {
  try {
    if (!trigger) {
      trigger = getScheduledTrigger();
    }

    if (trigger) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log("✅ Scheduled trigger deleted");
      return true;
    } else {
      Logger.log("⚠️ No trigger found to delete");
      return false;
    }
  } catch (error) {
    logError("Failed to disable scheduled trigger", error);
    return false;
  }
}

/**
 * Get the scheduled trigger if it exists
 */
function getScheduledTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();

    // Find trigger for 'scheduledProcessing' function
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === "scheduledProcessing") {
        return trigger;
      }
    }

    return null;
  } catch (error) {
    logError("Failed to get scheduled trigger", error);
    return null;
  }
}

/**
 * Check and display current schedule status
 */
function checkScheduleStatus() {
  try {
    const ui = SpreadsheetApp.getUi();
    const trigger = getScheduledTrigger();

    if (trigger) {
      const triggerInfo = getTriggerDetails(trigger);

      ui.alert(
        "📊 Schedule Status - ACTIVE",
        "✅ Automatic processing is ENABLED\n\n" +
          `📅 Schedule: ${triggerInfo.schedule}\n` +
          `🕐 Time: 5:15 PM (GMT+8)\n` +
          `📆 Days: Monday - Friday\n` +
          `⚙️ Function: ${triggerInfo.function}\n` +
          `🆔 Trigger ID: ${triggerInfo.id.substring(0, 20)}...\n\n` +
          "💡 To disable:\n" +
          "   ⚙️ Schedule Settings → ⏸️ Disable Auto Schedule",
        ui.ButtonSet.OK,
      );
    } else {
      ui.alert(
        "📊 Schedule Status - INACTIVE",
        "⏸️ Automatic processing is DISABLED\n\n" +
          "❌ No scheduled automation running\n" +
          "⚠️ Files will NOT be processed automatically\n\n" +
          "💡 To enable:\n" +
          "   ⚙️ Schedule Settings → ▶️ Enable Auto Schedule\n\n" +
          "💡 For manual processing:\n" +
          "   → 🔍 Scan Files Now\n" +
          "   → 📧 Send Daily Emails",
        ui.ButtonSet.OK,
      );
    }
  } catch (error) {
    logError("Failed to check schedule status", error);
    SpreadsheetApp.getUi().alert(
      "❌ Error",
      "Failed to check schedule status.",
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

/**
 * Get detailed trigger information
 */
function getTriggerDetails(trigger) {
  try {
    const triggerSource = trigger.getTriggerSource();
    const handlerFunction = trigger.getHandlerFunction();
    const uniqueId = trigger.getUniqueId();

    // Get time-based trigger details
    if (triggerSource === ScriptApp.TriggerSource.CLOCK) {
      return {
        id: uniqueId,
        function: handlerFunction,
        schedule: "5:15 PM, Monday-Friday",
        nextRun: "Next weekday at 5:15 PM (GMT+8)",
        type: "Time-based (Daily)",
      };
    }

    return {
      id: uniqueId,
      function: handlerFunction,
      schedule: "Custom",
      nextRun: "Unknown",
      type: triggerSource,
    };
  } catch (error) {
    logError("Failed to get trigger details", error);
    return {
      id: "Unknown",
      function: "Unknown",
      schedule: "Unknown",
      nextRun: "Unknown",
      type: "Unknown",
    };
  }
}

function showSystemStatus() {
  const ui = SpreadsheetApp.getUi();
  const hasGemini = !!getGeminiApiKey();
  const trigger = getScheduledTrigger();
  const scheduleStatus = trigger
    ? "✅ ENABLED (5:15 PM Mon-Fri)"
    : "⏸️ DISABLED";
  const emailQuota = MailApp.getRemainingDailyQuota();

  let statusMsg = "📊 SYSTEM STATUS v2025-2\n\n";
  statusMsg += "🤖 AI SERVICES:\n";
  statusMsg += `   Gemini 2.0 Flash: ${hasGemini ? "✅ Active" : "❌ Not configured"}\n`;
  statusMsg += `   Gemini Model: ${getGeminiModel()}\n`;
  statusMsg += "\n";
  statusMsg += "⏰ AUTOMATION:\n";
  statusMsg += `   Auto Schedule: ${scheduleStatus}\n`;
  statusMsg += `   Weekly Digest: ${WEEKLY_DIGEST_ENABLED ? "✅ Enabled (Fridays)" : "⏸️ Disabled"}\n\n`;
  statusMsg += "📧 EMAIL QUOTA:\n";
  statusMsg += `   Remaining Today: ${emailQuota} emails\n`;
  statusMsg += `   Batch Size: ${MAX_RECIPIENTS_PER_EMAIL} recipients/email\n\n`;
  statusMsg += `Version: 2025-2 (Enhanced Schedule Control)\n`;
  statusMsg += `Debug Mode: ${DEBUG_MODE ? "🟢 ON" : "⚪ OFF"}\n\n`;

  const operationalStatus =
    hasGemini && trigger && emailQuota > 10
      ? "✅ FULLY OPERATIONAL"
      : "⚠️ CONFIGURATION NEEDED";

  statusMsg += `Overall Status: ${operationalStatus}`;

  ui.alert("System Status", statusMsg, ui.ButtonSet.OK);
}

/**
 * UPDATED: Send daily emails with search functionality
 */
function sendDailyEmails() {
  try {
    logAudit("Email sending started", "Daily email initiated");
    cleanupOldEmailData();
    const departments = getDepartments();
    const today = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd");
    const allEmailAddresses = new Set();
    const allManagerEmails = new Set();

    departments.forEach((dept) => {
      if (dept.emails) {
        dept.emails.split(",").forEach((email) => {
          const cleanEmail = email.trim();
          if (cleanEmail) allEmailAddresses.add(cleanEmail);
        });
      }
      if (dept.managerEmails) {
        dept.managerEmails.split(",").forEach((email) => {
          const cleanEmail = email.trim();
          if (cleanEmail) allManagerEmails.add(cleanEmail);
        });
      }
    });

    const todaysProcessedFiles = getTodaysProcessedFilesWithDirectUrls(today);

    if (todaysProcessedFiles.length === 0) {
      logAudit("No emails sent", "No files processed today");
      SpreadsheetApp.getUi().alert(
        "No Emails Sent",
        "No files processed today.",
        SpreadsheetApp.getUi().ButtonSet.OK,
      );
      return;
    }

    if (allEmailAddresses.size === 0) {
      logAudit("No email addresses", "No emails configured");
      SpreadsheetApp.getUi().alert(
        "No Email Addresses",
        "Configure emails in Dashboard.",
        SpreadsheetApp.getUi().ButtonSet.OK,
      );
      return;
    }

    const recipients = Array.from(allEmailAddresses).join(",");
    const ccRecipients = Array.from(allManagerEmails).join(",");
    const totalRecipients = allEmailAddresses.size + allManagerEmails.size;

    Logger.log(`Sending to ${totalRecipients} recipients`);

    // UPDATED: Pass 'daily' as email type
    const result = sendTodaysProcessedFilesEmail(
      todaysProcessedFiles,
      today,
      recipients,
      ccRecipients,
      "daily",
    );

    let completionMessage = `Email sent: ${todaysProcessedFiles.length} files`;
    if (result.batchesSent > 1) {
      completionMessage += `, ${result.batchesSent} batches, ${result.totalRecipients} recipients`;
    } else {
      completionMessage += `, ${totalRecipients} recipients`;
    }

    logAudit("Email sent", completionMessage);

    let alertMessage = `✅ Email Sent!\n\n• Files: ${todaysProcessedFiles.length}\n• Recipients: ${totalRecipients}\n`;
    if (result.batchesSent > 1)
      alertMessage += `• Batches: ${result.batchesSent}`;

    SpreadsheetApp.getUi().alert(
      "Emails Sent",
      alertMessage,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  } catch (error) {
    logError("Email sending failed", error);
    SpreadsheetApp.getUi().alert(
      "Error",
      "Email failed. Check Audit Trail.",
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

/**
 * NEW: Send weekly digest email
 */
function sendWeeklyDigest() {
  try {
    const ui = SpreadsheetApp.getUi();
    const confirmResponse = ui.alert(
      "Send Weekly Digest",
      "This will send a summary of all files processed this week. Continue?",
      ui.ButtonSet.YES_NO,
    );

    if (confirmResponse !== ui.Button.YES) {
      return;
    }

    logAudit("Weekly digest initiated", "User triggered weekly digest");

    const departments = getDepartments();
    const allEmailAddresses = new Set();
    const allManagerEmails = new Set();

    departments.forEach((dept) => {
      if (dept.emails) {
        dept.emails.split(",").forEach((email) => {
          const cleanEmail = email.trim();
          if (cleanEmail) allEmailAddresses.add(cleanEmail);
        });
      }
      if (dept.managerEmails) {
        dept.managerEmails.split(",").forEach((email) => {
          const cleanEmail = email.trim();
          if (cleanEmail) allManagerEmails.add(cleanEmail);
        });
      }
    });

    // Get this week's files
    const weekFiles = getThisWeeksProcessedFiles();

    if (weekFiles.length === 0) {
      logAudit("No weekly digest sent", "No files processed this week");
      ui.alert("No Files", "No files processed this week.", ui.ButtonSet.OK);
      return;
    }

    if (allEmailAddresses.size === 0) {
      logAudit("No email addresses", "No emails configured");
      ui.alert(
        "No Email Addresses",
        "Configure emails in Dashboard.",
        ui.ButtonSet.OK,
      );
      return;
    }

    const recipients = Array.from(allEmailAddresses).join(",");
    const ccRecipients = Array.from(allManagerEmails).join(",");
    const totalRecipients = allEmailAddresses.size + allManagerEmails.size;

    const today = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd");
    const result = sendTodaysProcessedFilesEmail(
      weekFiles,
      today,
      recipients,
      ccRecipients,
      "weekly",
    );

    logAudit(
      "Weekly digest sent",
      `Sent digest: ${weekFiles.length} files to ${totalRecipients} recipients`,
    );

    ui.alert(
      "Weekly Digest Sent",
      `✅ Weekly digest sent!\n\n• Files: ${weekFiles.length}\n• Recipients: ${totalRecipients}`,
      ui.ButtonSet.OK,
    );
  } catch (error) {
    logError("Weekly digest failed", error);
    SpreadsheetApp.getUi().alert(
      "Error",
      "Weekly digest failed. Check Audit Trail.",
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

/**
 * NEW: Automatic weekly digest (called by trigger)
 */
function scheduledWeeklyDigest() {
  try {
    const today = new Date();
    const dayOfWeek = today.getDay();

    // Check if it's Friday (or configured day) and weekly digest is enabled
    if (!WEEKLY_DIGEST_ENABLED || dayOfWeek !== WEEKLY_DIGEST_DAY) {
      Logger.log(
        `Not digest day. Today is ${dayOfWeek}, digest day is ${WEEKLY_DIGEST_DAY}`,
      );
      return;
    }

    logAudit(
      "Scheduled weekly digest started",
      `Automated weekly digest on ${today}`,
    );

    const departments = getDepartments();
    const allEmailAddresses = new Set();
    const allManagerEmails = new Set();

    departments.forEach((dept) => {
      if (dept.emails) {
        dept.emails.split(",").forEach((email) => {
          const cleanEmail = email.trim();
          if (cleanEmail) allEmailAddresses.add(cleanEmail);
        });
      }
      if (dept.managerEmails) {
        dept.managerEmails.split(",").forEach((email) => {
          const cleanEmail = email.trim();
          if (cleanEmail) allManagerEmails.add(cleanEmail);
        });
      }
    });

    const weekFiles = getThisWeeksProcessedFiles();

    if (weekFiles.length === 0) {
      logAudit("No weekly digest sent", "No files processed this week");
      return;
    }

    if (allEmailAddresses.size === 0) {
      logAudit("No email addresses", "No emails configured");
      return;
    }

    const recipients = Array.from(allEmailAddresses).join(",");
    const ccRecipients = Array.from(allManagerEmails).join(",");

    const todayFormatted = Utilities.formatDate(today, TIMEZONE, "yyyy-MM-dd");
    sendTodaysProcessedFilesEmail(
      weekFiles,
      todayFormatted,
      recipients,
      ccRecipients,
      "weekly",
    );

    logAudit(
      "Scheduled weekly digest sent",
      `Sent digest: ${weekFiles.length} files`,
    );
  } catch (error) {
    logError("Scheduled weekly digest failed", error);
  }
}

/**
 * NEW: Get all files processed this week
 */
function getThisWeeksProcessedFiles() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const today = new Date();
    const weekFiles = [];

    // Get files from the last 7 days
    for (let i = 0; i < 7; i++) {
      const date = new Date(today.getTime() - i * 24 * 60 * 60 * 1000);
      const dateFormatted = Utilities.formatDate(date, TIMEZONE, "yyyy-MM-dd");

      const todaysRowsKey = `processed_rows_${dateFormatted}`;
      const todaysRowsData = properties.getProperty(todaysRowsKey);

      if (todaysRowsData) {
        const todaysRows = JSON.parse(todaysRowsData);
        todaysRows.forEach((rowNumber) => {
          const emailDataKey = `email_${dateFormatted}_row_${rowNumber}`;
          const emailDataString = properties.getProperty(emailDataKey);
          if (emailDataString) {
            try {
              const emailData = JSON.parse(emailDataString);
              weekFiles.push({
                department: emailData.department,
                newFileName: emailData.newFileName,
                fileUrl: emailData.fileUrl,
                folderUrl: emailData.folderUrl,
                fileSize: emailData.fileSize,
                aiUsed: emailData.aiUsed || "N/A",
                processDate: dateFormatted, // Add date for weekly digest
              });
            } catch (e) {
              Logger.log(`Failed to parse email data for row ${rowNumber}`);
            }
          }
        });
      }
    }

    Logger.log(`Found ${weekFiles.length} files for this week`);
    return weekFiles;
  } catch (error) {
    logError("Failed to get weekly files", error);
    return [];
  }
}

/**
 * UPDATED: Main email sending function with type support (daily/weekly)
 */
function sendTodaysProcessedFilesEmail(
  todaysFiles,
  date,
  recipients,
  ccRecipients,
  emailType = "daily",
) {
  try {
    const toRecipients = recipients
      .split(",")
      .map((e) => e.trim())
      .filter((e) => e);
    const ccArray = ccRecipients
      ? ccRecipients
          .split(",")
          .map((e) => e.trim())
          .filter((e) => e)
      : [];
    const totalRecipients = toRecipients.length + ccArray.length;

    if (totalRecipients <= MAX_RECIPIENTS_PER_EMAIL) {
      sendSingleEmail(
        todaysFiles,
        date,
        recipients,
        ccRecipients,
        1,
        1,
        emailType,
      );
      return {
        success: true,
        batchesSent: 1,
        totalRecipients: totalRecipients,
      };
    } else {
      return sendBatchEmails(
        todaysFiles,
        date,
        toRecipients,
        ccArray,
        emailType,
      );
    }
  } catch (error) {
    logError("Failed to send email", error);
    throw error;
  }
}

/**
 * UPDATED: Batch email sending with type support
 */
function sendBatchEmails(
  todaysFiles,
  date,
  toRecipients,
  ccRecipients,
  emailType = "daily",
) {
  try {
    const allRecipients = [...toRecipients, ...ccRecipients];
    const totalRecipients = allRecipients.length;
    const totalBatches = Math.ceil(totalRecipients / MAX_RECIPIENTS_PER_EMAIL);
    let batchesSent = 0,
      errors = 0;

    for (let i = 0; i < totalBatches; i++) {
      try {
        const startIndex = i * MAX_RECIPIENTS_PER_EMAIL;
        const endIndex = Math.min(
          startIndex + MAX_RECIPIENTS_PER_EMAIL,
          totalRecipients,
        );
        const batchRecipients = allRecipients.slice(startIndex, endIndex);
        const batchString = batchRecipients.join(",");

        sendSingleEmail(
          todaysFiles,
          date,
          batchString,
          "",
          i + 1,
          totalBatches,
          emailType,
        );
        batchesSent++;

        if (i < totalBatches - 1) Utilities.sleep(EMAIL_BATCH_DELAY);
      } catch (batchError) {
        logError(`Batch ${i + 1} failed`, batchError);
        errors++;
      }
    }

    const successMessage = `Batch sending: ${batchesSent}/${totalBatches} batches to ${totalRecipients} recipients`;
    logAudit(
      errors > 0 ? "Batch sending with errors" : "Batch sending complete",
      successMessage,
    );

    return {
      success: errors === 0,
      batchesSent: batchesSent,
      totalBatches: totalBatches,
      totalRecipients: totalRecipients,
      errors: errors,
    };
  } catch (error) {
    logError("Batch sending failed", error);
    throw error;
  }
}

/**
 * REFINED v2.6.2: Blue-themed, no file icons, search section moved to bottom
 */
function sendSingleEmail(
  todaysFiles,
  date,
  recipients,
  ccRecipients,
  batchNumber,
  totalBatches,
  emailType = "daily",
) {
  try {
    const isBatched = totalBatches > 1;
    const isWeekly = emailType === "weekly";
    const batchInfo = isBatched
      ? ` (Batch ${batchNumber}/${totalBatches})`
      : "";

    // Subject line
    const subject = isWeekly
      ? `FilePino Inc Weekly Summary of Received Documents - Week ending ${date}${batchInfo}`
      : `FilePino Inc Everyday Summary of Received Documents - ${date}${batchInfo}`;

    const filesByDept = {};
    todaysFiles.forEach((file) => {
      if (!filesByDept[file.department]) filesByDept[file.department] = [];
      filesByDept[file.department].push(file);
    });

    // Calculate total file size
    let totalSizeBytes = 0;
    todaysFiles.forEach((file) => {
      const sizeMatch = file.fileSize.match(/[\d.]+/);
      const unit = file.fileSize.match(/[A-Z]+/);
      if (sizeMatch && unit) {
        const size = parseFloat(sizeMatch[0]);
        const multiplier =
          unit[0] === "KB"
            ? 1024
            : unit[0] === "MB"
              ? 1024 * 1024
              : unit[0] === "GB"
                ? 1024 * 1024 * 1024
                : 1;
        totalSizeBytes += size * multiplier;
      }
    });
    const totalSize = formatFileSize(totalSizeBytes);

    // Calculate average processing time
    let avgProcessTime = "1.2s"; // Default if not available

    // Start HTML email with blue theme
    let htmlBody = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    @media only screen and (max-width: 600px) {
      .container { width: 100% !important; padding: 10px !important; }
      .metric-box { width: 48% !important; margin: 1% !important; }
      .stat-row { display: block !important; }
      .stat-col { width: 100% !important; display: block !important; margin-bottom: 10px !important; }
      .file-table { font-size: 12px !important; }
      .file-table td, .file-table th { padding: 8px 6px !important; }
    }
  </style>
</head>
<body style="margin:0;padding:0;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Oxygen,Ubuntu,Cantarell,sans-serif;background-color:#f0f4f8;">
  <div class="container" style="max-width:700px;margin:20px auto;background-color:#ffffff;border-radius:12px;box-shadow:0 4px 12px rgba(0,0,0,0.1);overflow:hidden;">
    
    <!-- Header with Blue Gradient -->
    <div style="background:linear-gradient(135deg, #E31E24 0%, #DC2626 50%, #FCA5A5 100%);color:white;padding:30px 20px;text-align:center;">
      <h1 style="margin:0;font-size:26px;font-weight:600;letter-spacing:0.5px;">📋 FilePino Inc - ${isWeekly ? "Weekly" : "Daily"} Document Report</h1>
      <p style="margin:8px 0 0 0;font-size:14px;opacity:0.9;">${isWeekly ? "Week ending" : ""} ${date}</p>
    </div>`;

    // Batch warning (if applicable)
    if (isBatched) {
      htmlBody += `
    <div style="background:#FEE2E2;padding:12px 20px;margin:0;border-left:4px solid #E31E24;">
      <strong style="color:#DC2626;">📨 Batch ${batchNumber} of ${totalBatches}</strong>
    </div>`;
    }

    // Compact Metric Boxes (4 columns, single row)
    htmlBody += `
    <div style="padding:20px;">
      <div class="stat-row" style="display:flex;justify-content:space-between;gap:10px;margin-bottom:20px;">
        <div class="stat-col" style="flex:1;background:linear-gradient(135deg, #FEF2F2 0%, #FEE2E2 100%);padding:16px;border-radius:8px;text-align:center;border:1px solid #FCA5A5;">
          <div style="font-size:28px;font-weight:700;color:#DC2626;margin-bottom:4px;">${todaysFiles.length}</div>
          <div style="font-size:11px;color:#64748b;font-weight:500;text-transform:uppercase;letter-spacing:0.5px;">Files</div>
        </div>
        <div class="stat-col" style="flex:1;background:linear-gradient(135deg, #FEF2F2 0%, #FEE2E2 100%);padding:16px;border-radius:8px;text-align:center;border:1px solid #FCA5A5;">
          <div style="font-size:28px;font-weight:700;color:#DC2626;margin-bottom:4px;">${Object.keys(filesByDept).length}</div>
          <div style="font-size:11px;color:#64748b;font-weight:500;text-transform:uppercase;letter-spacing:0.5px;">Departments</div>
        </div>
        <div class="stat-col" style="flex:1;background:linear-gradient(135deg, #FEF2F2 0%, #FEE2E2 100%);padding:16px;border-radius:8px;text-align:center;border:1px solid #FCA5A5;">
          <div style="font-size:28px;font-weight:700;color:#DC2626;margin-bottom:4px;">${totalSize.split(" ")[0]}</div>
          <div style="font-size:11px;color:#64748b;font-weight:500;text-transform:uppercase;letter-spacing:0.5px;">${totalSize.split(" ")[1]}</div>
        </div>
        <div class="stat-col" style="flex:1;background:linear-gradient(135deg, #FEF2F2 0%, #FEE2E2 100%);padding:16px;border-radius:8px;text-align:center;border:1px solid #FCA5A5;">
          <div style="font-size:28px;font-weight:700;color:#DC2626;margin-bottom:4px;">${avgProcessTime.replace("s", "")}</div>
          <div style="font-size:11px;color:#64748b;font-weight:500;text-transform:uppercase;letter-spacing:0.5px;">AVG TIME (s)</div>
        </div>
      </div>`;

    // Department Cards with TABLE layout (NO ICONS)
    Object.keys(filesByDept).forEach((dept, index) => {
      const files = filesByDept[dept];
      const completionPercent = Math.min(
        100,
        Math.round((files.length / todaysFiles.length) * 100),
      );

      // Red shades for different departments
      const blueShades = [
        {
          bg: "#FEF2F2",
          border: "#E31E24",
          text: "#DC2626",
          progress: "#E31E24",
        }, // Standard red
        {
          bg: "#FEF2F2",
          border: "#DC2626",
          text: "#DC2626",
          progress: "#DC2626",
        }, // Red
        {
          bg: "#FEE2E2",
          border: "#E31E24",
          text: "#DC2626",
          progress: "#E31E24",
        }, // Cyan
        {
          bg: "#FEF2F2",
          border: "#DC2626",
          text: "#DC2626",
          progress: "#DC2626",
        }, // Teal
        {
          bg: "#FEF2F2",
          border: "#DC2626",
          text: "#DC2626",
          progress: "#DC2626",
        }, // Emerald
      ];
      const colors = blueShades[index % blueShades.length];

      htmlBody += `
      <div style="background:${colors.bg};border:2px solid ${colors.border};border-radius:10px;padding:18px;margin-bottom:18px;box-shadow:0 2px 6px rgba(59,130,246,0.08);">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;">
          <h3 style="margin:0;color:${colors.text};font-size:18px;font-weight:600;">🏢 ${dept}</h3>
          <span style="background:${colors.progress};color:white;padding:4px 12px;border-radius:12px;font-size:12px;font-weight:600;">${files.length} file${files.length > 1 ? "s" : ""}</span>
        </div>
        
        <!-- Progress Bar -->
        <div style="background:#FEE2E2;height:6px;border-radius:3px;margin-bottom:16px;overflow:hidden;">
          <div style="background:${colors.progress};height:100%;width:${completionPercent}%;border-radius:3px;"></div>
        </div>
        
        <!-- TABLE FORMAT (NO ICONS) -->
        <table class="file-table" style="width:100%;border-collapse:collapse;border:1px solid ${colors.border};border-radius:6px;overflow:hidden;background:#ffffff;">
          <thead>
            <tr style="background:${colors.progress};color:white;">
              <th style="padding:12px;text-align:left;font-weight:600;font-size:13px;border-bottom:2px solid ${colors.border};">File</th>
              <th style="padding:12px;text-align:center;font-weight:600;font-size:13px;width:100px;border-bottom:2px solid ${colors.border};">Folder</th>
              <th style="padding:12px;text-align:center;font-weight:600;font-size:13px;width:100px;border-bottom:2px solid ${colors.border};">Size</th>
            </tr>
          </thead>
          <tbody>`;

      files.forEach((file, idx) => {
        const rowBg = idx % 2 === 0 ? "#ffffff" : "#f9fafb";

        // NO ICONS - just file names
        let fileHtml = file.newFileName;
        let folderHtml = "View Folder";

        if (file.fileUrl && file.fileUrl.trim()) {
          fileHtml = `<a href="${file.fileUrl}" style="color:${colors.text};text-decoration:none;font-weight:500;">${file.newFileName}</a>`;
        }
        if (file.folderUrl && file.folderUrl.trim()) {
          folderHtml = `<a href="${file.folderUrl}" style="color:${colors.progress};text-decoration:none;font-weight:500;">📁 View</a>`;
        }

        htmlBody += `
            <tr style="background:${rowBg};">
              <td style="padding:12px;border-bottom:1px solid ${colors.border}40;color:#334155;font-size:13px;">${fileHtml}</td>
              <td style="padding:12px;border-bottom:1px solid ${colors.border}40;text-align:center;font-size:13px;">${folderHtml}</td>
              <td style="padding:12px;border-bottom:1px solid ${colors.border}40;text-align:center;color:#64748b;font-size:13px;">${file.fileSize}</td>
            </tr>`;
      });

      htmlBody += `
          </tbody>
        </table>
      </div>`;
    });

    // MOVED: Search Function Section to bottom (BEFORE File Organization)
    htmlBody += `
      <div style="background:linear-gradient(135deg, #FEE2E2 0%, #FCA5A5 100%);padding:16px;border-radius:8px;margin-bottom:18px;text-align:center;border:2px solid #E31E24;">
        <p style="margin:0 0 10px 0;color:#DC2626;font-size:15px;font-weight:600;">🔍 Search Processed Files</p>
        <a href="${getSpreadsheetUrl()}" style="display:inline-block;background:#E31E24;color:white;padding:11px 24px;text-decoration:none;border-radius:6px;font-weight:600;font-size:14px;box-shadow:0 2px 4px rgba(59,130,246,0.3);transition:background 0.3s;">Open File Search</a>
        <p style="margin:10px 0 0 0;color:#64748b;font-size:12px;">Search all processed files in the Scanned Files Log</p>
      </div>`;

    // File Organization Info - Blue themed
    htmlBody += `
      <div style="background:linear-gradient(135deg, #FEF2F2 0%, #FEE2E2 100%);padding:18px;border-radius:8px;margin-top:0;border-left:4px solid #DC2626;">
        <h4 style="color:#DC2626;margin:0 0 10px 0;font-size:16px;font-weight:600;">📂 File Organization</h4>
        <p style="margin:0 0 8px 0;color:#DC2626;line-height:1.6;font-size:14px;">Files are automatically organized into: <strong>Department → Year → Month → Date</strong></p>
        <p style="margin:0;color:#64748b;font-size:13px;font-style:italic;">Example: Legal → 2025 → October → 19 October</p>
      </div>
    </div>
    
    <!-- Footer -->
    <div style="background:#f8fafc;padding:20px;text-align:center;border-top:1px solid #e2e8f0;">
      <p style="margin:0 0 5px 0;color:#475569;font-size:13px;font-weight:500;">Automated by FilePino Inc v2025-10</p>
      <p style="margin:0;color:#94a3b8;font-size:12px;">Generated: ${new Date().toLocaleString("en-US", { timeZone: "Asia/Singapore" })} (GMT+8)</p>`;

    if (isBatched) {
      htmlBody += `
      <p style="margin:10px 0 0 0;color:#64748b;font-size:12px;"><strong>Batch ${batchNumber} of ${totalBatches}</strong></p>`;
    }

    htmlBody += `
    </div>
    
  </div>
</body>
</html>`;

    const emailOptions = {
      htmlBody: htmlBody,
      name: "FilePino Inc Document Automation",
      attachments: [],
    };

    if (ccRecipients && ccRecipients.trim() && !isBatched) {
      emailOptions.cc = ccRecipients;
    }

    MailApp.sendEmail(recipients, subject, "", emailOptions);
    Logger.log(
      `Email sent successfully (${emailType}, batch ${batchNumber}/${totalBatches})`,
    );
  } catch (error) {
    logError(`Failed to send email (batch ${batchNumber})`, error);
    throw error;
  }
}

/**
 * NEW: Get spreadsheet URL for search function
 */
function getSpreadsheetUrl() {
  var ss = openTargetSpreadsheet();
  return (
    `https://docs.google.com/spreadsheets/d/${ss.getId()}/edit#gid=` +
    ss.getSheetByName("Scanned Files Log").getSheetId()
  );
}

/**
 * Checking Current Quota
 */
function checkEmailQuota() {
  const remaining = MailApp.getRemainingDailyQuota();
  const ui = SpreadsheetApp.getUi();

  ui.alert(
    "Email Quota Status",
    `📧 Remaining Daily Quota: ${remaining} emails\n\n` +
      `Current Setting: ${MAX_RECIPIENTS_PER_EMAIL} recipients per batch\n\n` +
      `Status: ${remaining >= 100 ? "✅ Ready for 100+ recipients" : "⚠️ Low quota - check if trial account"}`,
    ui.ButtonSet.OK,
  );

  Logger.log(`Remaining daily quota: ${remaining}`);
}

function getDefaultGeminiModel() {
  return "gemini-3.0-flash";
}

function normalizeGeminiModelName(modelName) {
  if (!modelName) return "";
  return modelName
    .toString()
    .trim()
    .replace(/^models\//i, "");
}

function saveGeminiModel(modelName) {
  var normalizedModel = normalizeGeminiModelName(modelName);

  if (!normalizedModel) {
    return { success: false, message: "Model name cannot be empty" };
  }

  if (!/^[-a-zA-Z0-9._]+$/.test(normalizedModel)) {
    return {
      success: false,
      message: "Invalid model name format. Use values like gemini-3.0-flash",
    };
  }

  if (!/^gemini-/i.test(normalizedModel)) {
    return {
      success: false,
      message: "Model name should start with 'gemini-'",
    };
  }

  PropertiesService.getScriptProperties().setProperty(
    "GEMINI_MODEL",
    normalizedModel,
  );

  return { success: true, message: "Gemini model saved: " + normalizedModel };
}

function getGeminiModel() {
  var model =
    PropertiesService.getScriptProperties().getProperty("GEMINI_MODEL");
  var normalizedModel = normalizeGeminiModelName(model);
  return normalizedModel || getDefaultGeminiModel();
}

function showGeminiModelDialog() {
  var ui = SpreadsheetApp.getUi();
  var currentModel = getGeminiModel();

  var promptResult = ui.prompt(
    "Set Gemini Model",
    "Current model: " +
      currentModel +
      "\n\nEnter model (example: gemini-3.0-flash)",
    ui.ButtonSet.OK_CANCEL,
  );

  if (promptResult.getSelectedButton() !== ui.Button.OK) return;

  var modelInput = (promptResult.getResponseText() || "").trim();
  var result = saveGeminiModel(modelInput);

  if (result.success) {
    ui.alert("Model Saved", result.message, ui.ButtonSet.OK);
  } else {
    ui.alert("Model Error", result.message, ui.ButtonSet.OK);
  }
}

function buildGeminiGenerateContentUrl(apiKey) {
  var model = getGeminiModel();
  return (
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    model +
    ":generateContent?key=" +
    apiKey
  );
}

function getGeminiApiKey() {
  var properties = PropertiesService.getScriptProperties();
  return (
    properties.getProperty("GEMINI_API_KEY") ||
    properties.getProperty("gemini_api_key") ||
    null
  );
}

function getTargetSpreadsheetId() {
  var properties = PropertiesService.getScriptProperties();
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      var activeId = active.getId();
      if (activeId) {
        properties.setProperty("TARGET_SPREADSHEET_ID", activeId);
        return activeId;
      }
    }
  } catch (e) {}

  var storedId = properties.getProperty("TARGET_SPREADSHEET_ID");
  if (storedId) return storedId;

  return SPREADSHEET_ID;
}

function openTargetSpreadsheet() {
  return SpreadsheetApp.openById(getTargetSpreadsheetId());
}

function getGeminiDailyCallLimit() {
  var properties = PropertiesService.getScriptProperties();
  var raw = properties.getProperty("GEMINI_DAILY_CALL_LIMIT");
  var parsed = parseInt(raw, 10);
  return parsed && parsed > 0 ? parsed : 200;
}

function getGeminiCallsToday() {
  var properties = PropertiesService.getScriptProperties();
  var dateKey = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd");
  var raw = properties.getProperty("GEMINI_CALLS_" + dateKey);
  var parsed = parseInt(raw, 10);
  return parsed && parsed >= 0 ? parsed : 0;
}

function hasGeminiBudget(cost) {
  return getGeminiCallsToday() + cost <= getGeminiDailyCallLimit();
}

function consumeGeminiBudget(cost) {
  if (!hasGeminiBudget(cost)) {
    throw new Error(
      `Gemini daily limit reached (${getGeminiCallsToday()}/${getGeminiDailyCallLimit()}). Try again tomorrow or raise GEMINI_DAILY_CALL_LIMIT.`,
    );
  }

  var properties = PropertiesService.getScriptProperties();
  var dateKey = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd");
  properties.setProperty(
    "GEMINI_CALLS_" + dateKey,
    String(getGeminiCallsToday() + cost),
  );
}

function assertGeminiTestCooldown() {
  var properties = PropertiesService.getScriptProperties();
  var now = Date.now();
  var lastRaw = properties.getProperty("GEMINI_LAST_TEST_MS");
  var last = parseInt(lastRaw, 10);
  var cooldownMs = 60 * 1000;

  if (last && now - last < cooldownMs) {
    var remaining = Math.ceil((cooldownMs - (now - last)) / 1000);
    throw new Error(
      `Please wait ${remaining}s before testing again (prevents accidental API spam).`,
    );
  }

  properties.setProperty("GEMINI_LAST_TEST_MS", String(now));
}
