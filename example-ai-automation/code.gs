/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * AI Document Analyzer - Example Apps Script Automation
 * ═══════════════════════════════════════════════════════════════════════════════
 *
 * PURPOSE:
 * Example automation showing how to use the common-functions library to build
 * a Google Sheets-based AI document analysis system using Gemini.
 *
 * WHAT THIS DOES:
 * 1. Scans a Google Drive folder for documents
 * 2. Uses Gemini AI to analyze each document and extract key information
 * 3. Logs results to a Google Sheet
 * 4. Sends email notifications with analysis summaries
 * 5. Tracks processing status and handles errors
 *
 * SETUP:
 * 1. Copy this file + common-functions/GeminiAPI.gs + common-functions/DriveUtils.gs
 *    + common-functions/SheetUtils.gs + common-functions/EmailUtils.gs
 *    + common-functions/ConfigUtils.gs + common-functions/LoggingUtils.gs into your project
 * 2. Run setupAutomation() from the menu
 * 3. Configure your Gemini API key and Drive folder
 * 4. Run processDocuments() to start analysis
 *
 * VERSION: 1.0.0
 * ═══════════════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════════════
// CONFIGURATION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Edit these constants to customize the automation
 */
var CONFIG = {
  // Sheet configuration
  SHEET_NAME: "Document Analysis", // Main sheet name
  LOG_SHEET_NAME: "Processing Log", // Activity log sheet

  // Gemini configuration (will use stored key from common functions)
  DEFAULT_PROMPT:
    "Analyze this document and provide:\n" +
    "1. Document type (Contract, Invoice, Letter, etc.)\n" +
    "2. Main subject/topic\n" +
    "3. Key dates mentioned\n" +
    "4. Important parties/organizations\n" +
    "5. Overall sentiment (positive, neutral, negative)\n\n" +
    "Format as: TYPE | SUBJECT | DATES | PARTIES | SENTIMENT",

  // File filtering
  SUPPORTED_MIME_TYPES: [
    "application/pdf",
    "image/jpeg",
    "image/png",
    "image/gif",
  ],
  MAX_FILE_SIZE_MB: 20, // Gemini API limit consideration

  // Email notifications
  NOTIFICATION_EMAIL: null, // Set to email address for notifications
  EMAIL_SUBJECT: "AI Document Analysis Complete",

  // Processing limits
  MAX_FILES_PER_RUN: 10, // Prevent timeout/quota issues
  RATE_LIMIT_DELAY_MS: 2000, // Delay between API calls (ms)
};

// ═══════════════════════════════════════════════════════════════════════════════
// MENU & SETUP
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * onOpen - Creates the custom menu when spreadsheet opens
 * Run this automatically or manually to see the menu
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu("🤖 AI Document Analyzer")
    .addItem("📋 Setup Wizard", "setupAutomation")
    .addSeparator()
    .addItem("🔑 Set Gemini API Key", "setGeminiApiKey")
    .addItem("📁 Set Drive Folder", "setDriveFolder")
    .addItem("⚙️ Configuration", "showConfiguration")
    .addSeparator()
    .addItem("▶️ Process Documents", "processDocuments")
    .addItem("📊 View Results", "switchToResultsSheet")
    .addSeparator()
    .addItem("🔍 Test Connection", "testGeminiConnection")
    .addItem("📝 View Logs", "switchToLogSheet")
    .addToUi();
}

/**
 * setupAutomation - Step-by-step setup wizard
 * Guides user through initial configuration
 */
function setupAutomation() {
  var ui = SpreadsheetApp.getUi();

  // Step 1: Welcome
  var welcome = ui.alert(
    "🤖 AI Document Analyzer Setup",
    "Welcome! This wizard will help you configure the automation.\n\n" +
      "You'll need:\n" +
      "1. A Gemini API key (free from aistudio.google.com)\n" +
      "2. A Google Drive folder with documents to analyze\n\n" +
      "Ready to start?",
    ui.ButtonSet.YES_NO,
  );

  if (welcome !== ui.Button.YES) return;

  // Step 2: API Key
  var keyResult = setGeminiApiKey();
  if (!keyResult) {
    ui.alert(
      "Setup cancelled. Please set your API key manually from the menu.",
    );
    return;
  }

  // Step 3: Drive Folder
  var folderResponse = ui.prompt(
    "Step 2: Drive Folder",
    "Enter your Google Drive folder URL or ID:\n" +
      "(This is where your documents to analyze are stored)",
    ui.ButtonSet.OK_CANCEL,
  );

  if (folderResponse.getSelectedButton() !== ui.Button.OK) return;

  var folderId = folderResponse.getResponseText().trim();
  var folder = getFolderByIdOrUrl(folderId);

  if (!folder) {
    ui.alert("❌ Invalid folder. Please check the URL/ID and try again.");
    return;
  }

  // Store folder ID
  PropertiesService.getScriptProperties().setProperty(
    "DRIVE_FOLDER_ID",
    folder.getId(),
  );

  // Step 4: Initialize sheets
  initializeSheets();

  // Step 5: Completion
  ui.alert(
    "✅ Setup Complete!",
    "Your automation is ready!\n\n" +
      "Drive Folder: " +
      folder.getName() +
      "\n" +
      "API Key: Configured\n\n" +
      "Next steps:\n" +
      "1. Add documents to your Drive folder\n" +
      "2. Click '🤖 AI Document Analyzer' → '▶️ Process Documents'\n\n" +
      "The system will analyze documents and log results to this sheet.",
    ui.ButtonSet.OK,
  );
}

/**
 * setDriveFolder - Allow user to update Drive folder later
 */
function setDriveFolder() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    "Set Drive Folder",
    "Enter Drive folder URL or ID:",
    ui.ButtonSet.OK_CANCEL,
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  var folderId = response.getResponseText().trim();
  var folder = getFolderByIdOrUrl(folderId);

  if (folder) {
    PropertiesService.getScriptProperties().setProperty(
      "DRIVE_FOLDER_ID",
      folder.getId(),
    );
    ui.alert("✅ Folder set to: " + folder.getName());
  } else {
    ui.alert("❌ Could not access folder. Check permissions and ID.");
  }
}

/**
 * showConfiguration - Display current configuration status
 */
function showConfiguration() {
  var folderId =
    PropertiesService.getScriptProperties().getProperty("DRIVE_FOLDER_ID");
  var folder = folderId ? getFolderByIdOrUrl(folderId) : null;

  var apiKey =
    PropertiesService.getUserProperties().getProperty("GEMINI_API_KEY");

  var statusItems = [
    {
      label: "Gemini API Key",
      value: apiKey ? "✅ Configured" : "❌ Not set",
      status: apiKey ? "ok" : "error",
    },
    {
      label: "Drive Folder",
      value: folder
        ? folder.getName()
        : folderId
          ? "❌ Invalid ID"
          : "❌ Not set",
      status: folder ? "ok" : "error",
    },
    {
      label: "Max Files/Run",
      value: CONFIG.MAX_FILES_PER_RUN.toString(),
      status: "info",
    },
    {
      label: "Rate Limit",
      value: CONFIG.RATE_LIMIT_DELAY_MS / 1000 + "s",
      status: "info",
    },
  ];

  showStatusPanel(statusItems, "AI Document Analyzer - Configuration");
}

// ═══════════════════════════════════════════════════════════════════════════════
// MAIN AUTOMATION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * processDocuments - Main processing function
 * Scans Drive folder, analyzes documents with Gemini, logs to sheet
 */
function processDocuments() {
  var logSheet = initializeLogSheet(
    SpreadsheetApp.getActiveSpreadsheet(),
    CONFIG.LOG_SHEET_NAME,
  );

  // Check prerequisites
  var folderId =
    PropertiesService.getScriptProperties().getProperty("DRIVE_FOLDER_ID");
  if (!folderId) {
    logActivity(logSheet, "SETUP_ERROR", "No Drive folder configured", {
      status: "FAILED",
    });
    SpreadsheetApp.getUi().alert(
      "❌ Please run Setup Wizard first (menu: 🤖 AI Document Analyzer → 📋 Setup Wizard)",
    );
    return;
  }

  var apiKey =
    PropertiesService.getUserProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) {
    logActivity(logSheet, "SETUP_ERROR", "No Gemini API key configured", {
      status: "FAILED",
    });
    SpreadsheetApp.getUi().alert(
      "❌ Please set your Gemini API key first (menu: 🤖 AI Document Analyzer → 🔑 Set Gemini API Key)",
    );
    return;
  }

  // Get or create main sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = initializeSheets();
  }

  // Scan Drive folder
  var folder = getFolderByIdOrUrl(folderId);
  if (!folder) {
    logError(
      logSheet,
      "Drive Access",
      new Error("Cannot access folder: " + folderId),
    );
    SpreadsheetApp.getUi().alert(
      "❌ Cannot access Drive folder. Check permissions.",
    );
    return;
  }

  logActivity(logSheet, "SCAN_START", "Scanning folder: " + folder.getName(), {
    status: "INFO",
  });

  // List files
  var filesResult = listFilesInFolder(folder, {
    mimeTypes: CONFIG.SUPPORTED_MIME_TYPES,
    maxResults: CONFIG.MAX_FILES_PER_RUN,
  });

  if (!filesResult.success || filesResult.files.length === 0) {
    logActivity(logSheet, "SCAN_COMPLETE", "No files found to process", {
      status: "INFO",
    });
    SpreadsheetApp.getUi().alert("ℹ️ No documents found in folder to analyze.");
    return;
  }

  logActivity(
    logSheet,
    "SCAN_COMPLETE",
    "Found " + filesResult.files.length + " files",
    { status: "INFO" },
  );

  // Process each file
  var processedCount = 0;
  var errorCount = 0;
  var results = [];

  for (var i = 0; i < filesResult.files.length; i++) {
    var fileInfo = filesResult.files[i];

    try {
      // Show progress toast every 2 files
      if (i % 2 === 0) {
        showProgressDialog(
          "Analyzing Documents",
          "Processing " + (i + 1) + " of " + filesResult.files.length + "...",
          Math.round((i / filesResult.files.length) * 100),
          "toast",
        );
      }

      // Check if already processed (by file ID in column A)
      if (isFileAlreadyProcessed(sheet, fileInfo.id)) {
        logActivity(logSheet, "SKIP", "Already processed: " + fileInfo.name, {
          status: "INFO",
        });
        continue;
      }

      // Get file object for Gemini
      var file = DriveApp.getFileById(fileInfo.id);

      // Check file size
      var sizeMB = fileInfo.size / (1024 * 1024);
      if (sizeMB > CONFIG.MAX_FILE_SIZE_MB) {
        logActivity(
          logSheet,
          "SKIP",
          "File too large: " +
            fileInfo.name +
            " (" +
            sizeMB.toFixed(1) +
            " MB)",
          { status: "WARNING" },
        );
        appendResultToSheet(
          sheet,
          fileInfo,
          "SKIPPED",
          "File exceeds " + CONFIG.MAX_FILE_SIZE_MB + " MB limit",
        );
        continue;
      }

      // Analyze with Gemini
      var analysisResult = callGeminiWithFile(file, CONFIG.DEFAULT_PROMPT, {
        temperature: 0.2,
        maxOutputTokens: 500,
      });

      if (analysisResult.success) {
        // Parse the structured response
        var parsed = parseGeminiResponse(analysisResult.text);

        appendResultToSheet(
          sheet,
          fileInfo,
          "COMPLETED",
          analysisResult.text,
          parsed,
        );
        logActivity(logSheet, "ANALYZE_SUCCESS", fileInfo.name, {
          status: "SUCCESS",
        });
        results.push({
          name: fileInfo.name,
          status: "success",
          summary: parsed,
        });
        processedCount++;
      } else {
        appendResultToSheet(sheet, fileInfo, "ERROR", analysisResult.error);
        logActivity(
          logSheet,
          "ANALYZE_FAILED",
          fileInfo.name + ": " + analysisResult.error,
          { status: "FAILED" },
        );
        errorCount++;
      }

      // Rate limiting delay
      if (i < filesResult.files.length - 1) {
        Utilities.sleep(CONFIG.RATE_LIMIT_DELAY_MS);
      }
    } catch (error) {
      logError(logSheet, "Processing " + fileInfo.name, error);
      appendResultToSheet(sheet, fileInfo, "ERROR", error.toString());
      errorCount++;
    }
  }

  // Completion summary
  var summaryMsg =
    "✅ Processing Complete!\n\n" +
    "Files analyzed: " +
    processedCount +
    "\n" +
    "Errors: " +
    errorCount +
    "\n" +
    "Skipped: " +
    (filesResult.files.length - processedCount - errorCount);

  logActivity(logSheet, "BATCH_COMPLETE", summaryMsg, { status: "INFO" });

  // Send email notification if configured
  if (CONFIG.NOTIFICATION_EMAIL && processedCount > 0) {
    sendNotificationEmail(results, processedCount, errorCount);
  }

  // Show completion alert
  SpreadsheetApp.getUi().alert(summaryMsg);

  // Switch to results sheet
  switchToResultsSheet();
}

// ═══════════════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * initializeSheets - Create and format the main and log sheets
 */
function initializeSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create/ensure main analysis sheet
  var mainSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!mainSheet) {
    mainSheet = ss.insertSheet(CONFIG.SHEET_NAME);

    // Set up headers
    var headers = [
      "File ID",
      "File Name",
      "File Type",
      "Size (MB)",
      "Modified Date",
      "Status",
      "Raw Analysis",
      "Document Type",
      "Subject",
      "Key Dates",
      "Parties",
      "Sentiment",
      "Processed At",
    ];

    mainSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    mainSheet
      .getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#4285f4")
      .setFontColor("white");
    mainSheet.setFrozenRows(1);

    // Set column widths
    mainSheet.setColumnWidth(1, 150); // File ID
    mainSheet.setColumnWidth(2, 200); // File Name
    mainSheet.setColumnWidth(3, 100); // File Type
    mainSheet.setColumnWidth(4, 80); // Size
    mainSheet.setColumnWidth(5, 120); // Modified
    mainSheet.setColumnWidth(6, 100); // Status
    mainSheet.setColumnWidth(7, 300); // Raw Analysis
    mainSheet.setColumnWidth(8, 120); // Doc Type
    mainSheet.setColumnWidth(9, 150); // Subject
    mainSheet.setColumnWidth(10, 150); // Key Dates
    mainSheet.setColumnWidth(11, 150); // Parties
    mainSheet.setColumnWidth(12, 80); // Sentiment
    mainSheet.setColumnWidth(13, 120); // Processed At
  }

  // Create log sheet via common function
  initializeLogSheet(ss, CONFIG.LOG_SHEET_NAME);

  return mainSheet;
}

/**
 * isFileAlreadyProcessed - Check if file ID exists in sheet
 */
function isFileAlreadyProcessed(sheet, fileId) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === fileId) return true; // Column A = File ID
  }
  return false;
}

/**
 * appendResultToSheet - Add analysis result to main sheet
 */
function appendResultToSheet(sheet, fileInfo, status, rawText, parsed) {
  parsed = parsed || {};

  var row = [
    fileInfo.id,
    fileInfo.name,
    fileInfo.mimeType,
    (fileInfo.size / (1024 * 1024)).toFixed(2),
    Utilities.formatDate(
      fileInfo.modified,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH:mm",
    ),
    status,
    rawText,
    parsed.type || "",
    parsed.subject || "",
    parsed.dates || "",
    parsed.parties || "",
    parsed.sentiment || "",
    new Date(),
  ];

  sheet.appendRow(row);

  // Color code the status cell
  var lastRow = sheet.getLastRow();
  var statusCell = sheet.getRange(lastRow, 6);

  if (status === "COMPLETED") {
    statusCell.setBackground("#d9ead3"); // Green
  } else if (status === "ERROR") {
    statusCell.setBackground("#f4cccc"); // Red
  } else if (status === "SKIPPED") {
    statusCell.setBackground("#fff2cc"); // Yellow
  }
}

/**
 * parseGeminiResponse - Parse structured output from Gemini
 * Expected format: TYPE | SUBJECT | DATES | PARTIES | SENTIMENT
 */
function parseGeminiResponse(text) {
  var result = {
    type: "",
    subject: "",
    dates: "",
    parties: "",
    sentiment: "",
  };

  try {
    // Try to split by pipe character first
    var parts = text.split("|").map(function (p) {
      return p.trim();
    });

    if (parts.length >= 5) {
      result.type = parts[0];
      result.subject = parts[1];
      result.dates = parts[2];
      result.parties = parts[3];
      result.sentiment = parts[4];
    } else {
      // Fallback: try to extract using line-by-line parsing
      var lines = text.split("\n");
      lines.forEach(function (line) {
        if (line.toLowerCase().indexOf("type:") !== -1)
          result.type = line.replace(/.*type:/i, "").trim();
        if (line.toLowerCase().indexOf("subject:") !== -1)
          result.subject = line.replace(/.*subject:/i, "").trim();
        if (line.toLowerCase().indexOf("date:") !== -1)
          result.dates = line.replace(/.*date:/i, "").trim();
        if (line.toLowerCase().indexOf("parties:") !== -1)
          result.parties = line.replace(/.*parties:/i, "").trim();
        if (line.toLowerCase().indexOf("sentiment:") !== -1)
          result.sentiment = line.replace(/.*sentiment:/i, "").trim();
      });
    }
  } catch (e) {
    // If parsing fails, return empty but don't fail
  }

  return result;
}

/**
 * sendNotificationEmail - Send summary email
 */
function sendNotificationEmail(results, successCount, errorCount) {
  try {
    var successItems = results.filter(function (r) {
      return r.status === "success";
    });

    var htmlBody =
      "<h2>AI Document Analysis Complete</h2>" +
      "<p><strong>Files analyzed:</strong> " +
      successCount +
      "<br>" +
      "<strong>Errors:</strong> " +
      errorCount +
      "</p>" +
      "<h3>Summary of Analyzed Documents:</h3><ul>";

    successItems.forEach(function (item) {
      htmlBody +=
        "<li><strong>" +
        item.name +
        "</strong><br>" +
        "Type: " +
        (item.summary.type || "N/A") +
        " | " +
        "Subject: " +
        (item.summary.subject || "N/A") +
        "</li>";
    });

    htmlBody += "</ul><p>View full results in the Google Sheet.</p>";

    sendHtmlEmail({
      to: CONFIG.NOTIFICATION_EMAIL,
      subject: CONFIG.EMAIL_SUBJECT + " (" + successCount + " files)",
      htmlBody: htmlBody,
      fromName: "AI Document Analyzer",
    });
  } catch (e) {
    // Silent fail on email - not critical
  }
}

/**
 * switchToResultsSheet - Navigate to results sheet
 */
function switchToResultsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (sheet) {
    ss.setActiveSheet(sheet);
  }
}

/**
 * switchToLogSheet - Navigate to log sheet
 */
function switchToLogSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (sheet) {
    ss.setActiveSheet(sheet);
  }
}
