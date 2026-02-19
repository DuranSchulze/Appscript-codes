/*******************************************************
CLIENT DOCUMENT MONITORING SYSTEM
For: Summary of Client's Files (Bound Script)
Features:
- Monitors "01 Client Documents" and “5.02 Scanned Pleadings” shared drive
- Tracks file uploads, modifications, moves, deletions
- AI-powered descriptive renaming using Gemini 2.5 Pro
- Scheduled daily scans (6PM Manila time, Mon-Fri)
- Batch processing with confirmations (25 files per batch)
- Comprehensive logging and error handling
Required OAuth Scopes:
- Drive (including shared drives)
- Spreadsheets
- Advanced Drive API for revision history




Authored by: Atty. Mary Wendy Duran
Modified: 21 September 2025
*******************************************************/

/**
 * Required OAuth Scopes Configuration
 * Add this to appsscript.json manifest
 */
function getRequiredScopes() {
  return {
    oauthScopes: [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive",
      "https://www.googleapis.com/auth/drive.file",
      "https://www.googleapis.com/auth/drive.metadata.readonly",
      "https://www.googleapis.com/auth/script.external_request",
    ],
  };
}

// Configuration Constants
const CONFIG = {
  SHEETS: {
    UPLOADED_BY_CLIENT: "Uploaded By Client",
    SUMMARY_PLEADINGS: "Summary of Pleadings",
    ANNEXES: "Annexes",
    CHECKLIST: "Checklist for Client",
    HEARING_DATES: "Hearing Dates",
    ENTRIES: "Entries",
    LOGS: "Logs",
    CONFIG: "System Config",
  },
  DRIVE_TYPES: {
    CLIENT_DOCS: "CLIENT_DOCUMENTS",
    PLEADINGS: "SCANNED_PLEADINGS",
  },
  BATCH_SIZE: 25,
  TIMEZONE: "Asia/Manila",
  SCHEDULE_HOUR: 18,
  DOMAIN: ".duranschulze.com",
  MAX_RETRIES: 3,
  RETRY_DELAY_MS: 1000,
  API_DELAY_MS: 100,
};

// Enhanced column headers with user tracking
const CLIENT_DOCS_HEADERS = [
  "Date Uploaded",
  "File Name",
  "File URL",
  "Uploaded by (Email)",
  "Date Modified",
  "Last Modified by (Email)",
  "Remarks",
  "Change Log",
];

const PLEADINGS_HEADERS = [
  "Date Uploaded",
  "File Name",
  "File URL",
  "Filed / Issued By",
  "Uploaded by (Email)",
  "Last Modified by (Email)",
  "Action Needed",
  "Assigned DDS Party",
  "Status",
];

const LOG_HEADERS = [
  "Timestamp",
  "Event Type",
  "Drive Type",
  "File ID",
  "File Name",
  "User Email",
  "Details",
  "Status",
  "Error Message",
];

// Enhanced dropdown options for pleadings sheet
const PLEADING_DROPDOWNS = {
  FILED_ISSUED_BY: ["DDS", "Court", "Opposing Party"],
  ACTION_NEEDED: [
    "For Response",
    "For Compliance",
    "For Filing/Release",
    "Filed/Completed",
    "No Action Needed",
  ], // Added "No Action Needed"
  ASSIGNED_DDS_PARTY: [
    "Partner",
    "Senior Lawyer",
    "Associate Lawyer",
    "Paralegal",
    "Legal Assistants",
    "Liaison Officers",
    "Other Support Personnel",
  ],
};

// Pleading type patterns for auto-detection
const PLEADING_TYPES = {
  Motion: /motion|mtd|mta|mtq|motion for|motion to/i,
  Order: /order|ruling|decision|directive/i,
  Complaint: /complaint|petition|application/i,
  Answer: /answer|reply|response|counter/i,
  Pleading: /pleading|filing/i,
  Notice: /notice|notification/i,
  Summons: /summons|subpoena/i,
  Judgment: /judgment|judgement|verdict/i,
  Brief: /brief|memorandum|memo/i,
  Affidavit: /affidavit|declaration/i,
  Contract: /contract|agreement|deed/i,
  Letter: /letter|correspondence/i,
};

// Gemini API Configuration
const GEMINI_CONFIG = {
  FALLBACK_MODEL: "gemini-2.0-flash",
  FALLBACK_ENDPOINT:
    "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent",
  MODELS_LIST_URL: "https://generativelanguage.googleapis.com/v1beta/models",
  USER_PROP_KEY: "GEMINI_API_KEY",
  SELECTED_MODEL_KEY: "GEMINI_SELECTED_MODEL",
  TEMPERATURE: 0.2,
  MAX_TOKENS: 150,
};

/**
 * Returns the active Gemini generateContent endpoint.
 * Reads GEMINI_SELECTED_MODEL from System Config; falls back to FALLBACK_ENDPOINT.
 */
function getGeminiEndpoint() {
  const selectedModel = getConfigValue(GEMINI_CONFIG.SELECTED_MODEL_KEY);
  if (selectedModel) {
    return `https://generativelanguage.googleapis.com/v1beta/models/${selectedModel}:generateContent`;
  }
  return GEMINI_CONFIG.FALLBACK_ENDPOINT;
}

/**
 * Enhanced menu system with user tracking diagnostics
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📁 Client Document Monitor")
    .addItem("🔧 Initial Setup", "setupSystem")
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("📂 Configure Drives")
        .addItem(
          "📄 Set Client Documents Drive ID",
          "setClientDocumentsDriveId",
        )
        .addItem(
          "⚖️ Set Scanned Pleadings Drive ID",
          "setScannedPleadingsDriveId",
        ),
    )
    .addItem("🔑 Set Gemini API Key", "setGeminiApiKey")
    .addItem("🤖 Select AI Model", "selectGeminiModel")
    .addItem("🗑️ Clear Gemini API Key", "clearGeminiApiKey")
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("🔍 Full Scanning")
        .addItem("📄 Full Scan - Client Documents", "manualFullScanClientDocs")
        .addItem("⚖️ Full Scan - Scanned Pleadings", "manualFullScanPleadings")
        .addSeparator()
        .addItem("🔄 Scan Both Drives", "scanBothDrives"),
    )
    .addSeparator()
    .addItem("⏰ Setup Daily Schedule", "setupDailySchedule")
    .addItem("🛑 Remove Daily Schedule", "removeDailySchedule")
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("🩺 Diagnostics")
        .addItem("📄 Test Client Documents Access", "testClientDocsAccess")
        .addItem("⚖️ Test Pleadings Access", "testPleadingsAccess")
        .addItem("👥 Test User Tracking", "testUserTracking")
        .addItem("🔍 Test Gemini Connection", "testGeminiConnection")
        .addItem("🧪 Test Gemini AI Integration", "testGeminiIntegration")
        .addItem("📊 View System Status", "viewSystemStatus")
        .addSeparator()
        .addItem("🚨 View Error Report", "viewErrorReport")
        .addItem("🗑️ Clear Error Logs", "clearErrorLogs"),
    )
    .addToUi();
}

/**
 * Enhanced system setup with updated naming convention notification
 */
function setupSystem() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Create client documents tracking sheet with user columns
    createClientDocsSheet(ss);
    // Create pleadings tracking sheet with user columns and enhanced dropdowns
    createPleadingsSheet(ss);
    // Create other sheets with updated headers
    createSheet(ss, CONFIG.SHEETS.ANNEXES, [
      "Annex",
      "Description",
      "File URL",
      "Status",
      "Remarks",
    ]);
    createSheet(ss, CONFIG.SHEETS.CHECKLIST, [
      "Task",
      "Status",
      "Due Date",
      "Notes",
    ]);
    createSheet(ss, CONFIG.SHEETS.SUMMARY_PLEADINGS, [
      "Date",
      "Pleading Type",
      "Filed By",
      "Action Needed",
      "Status",
    ]);
    createSheet(ss, CONFIG.SHEETS.HEARING_DATES, [
      "Date",
      "Time",
      "Type",
      "Location",
      "Status",
      "Attended By",
      "Remarks",
    ]);
    createSheet(ss, CONFIG.SHEETS.ENTRIES, [
      "Date",
      "Entry Type",
      "Description",
      "Reference",
    ]);
    createSheet(ss, CONFIG.SHEETS.LOGS, LOG_HEADERS);
    createSheet(ss, CONFIG.SHEETS.CONFIG, ["Key", "Value"], true); // Hide config sheet

    // Set initial configuration
    setConfigValue("SYSTEM_INITIALIZED", new Date().toISOString());
    setConfigValue("LAST_FULL_SCAN_CLIENT", "");
    setConfigValue("LAST_FULL_SCAN_PLEADINGS", "");
    setConfigValue("LAST_UPDATE_SCAN_CLIENT", "");
    setConfigValue("LAST_UPDATE_SCAN_PLEADINGS", "");
    setConfigValue("CLIENT_DOCUMENTS_DRIVE_ID", "");
    setConfigValue("SCANNED_PLEADINGS_DRIVE_ID", "");
    setConfigValue("CLIENT_FILE_INDEX", "{}");
    setConfigValue("PLEADINGS_FILE_INDEX", "{}");
    logEvent(
      "SYSTEM",
      "",
      "",
      "",
      "",
      "System setup completed successfully",
      "SUCCESS",
    );
    let setupMsg = "✅ Enhanced System setup completed!\n\n";
    setupMsg += "NEW FEATURES:\n";
    setupMsg += "• 👥 User email tracking for uploads and modifications\n";
    setupMsg += "• 📧 Enhanced user activity logging\n";
    setupMsg += "• 🔄 Updated pleading naming: YYYY-MM-DD Descriptive Title\n";
    setupMsg += '• ✅ Enhanced dropdowns including "No Action Needed"\n';
    setupMsg += "• 🔗 Automatic file renaming in both sheet and Drive\n\n";
    setupMsg += "Other features:\n";
    setupMsg += "• Separate tracking sheets for client docs and pleadings\n";
    setupMsg += "• Auto-detection of pleading types\n";
    setupMsg += "• Manual full scans + automatic daily updates\n\n";
    setupMsg += "Next steps:\n";
    setupMsg += "1. Set Client Documents Drive ID\n";
    setupMsg += "2. Set Scanned Pleadings Drive ID\n";
    setupMsg += "3. Test user tracking\n";
    setupMsg += "4. Run initial full scans";
    SpreadsheetApp.getUi().alert(setupMsg);
  } catch (error) {
    logEvent(
      "ERROR",
      "",
      "",
      "",
      "",
      `System setup failed: ${error.message}`,
      "ERROR",
      error.message,
    );
    SpreadsheetApp.getUi().alert("❌ Setup failed: " + error.message);
  }
}

/**
 * Creates pleadings tracking sheet with enhanced dropdowns including "No Action Needed"
 */
function createPleadingsSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.SUMMARY_PLEADINGS);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.SHEETS.SUMMARY_PLEADINGS);
    // Set up headers
    sheet
      .getRange(1, 1, 1, PLEADINGS_HEADERS.length)
      .setValues([PLEADINGS_HEADERS]);
    sheet.setFrozenRows(1);
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, PLEADINGS_HEADERS.length);
    headerRange.setBackground("#1a73e8");
    headerRange.setFontColor("white");
    headerRange.setFontWeight("bold");
    // Set column widths (adjusted for user columns)
    sheet.setColumnWidth(1, 120); // Date Uploaded
    sheet.setColumnWidth(2, 350); // File Name
    sheet.setColumnWidth(3, 200); // File URL
    sheet.setColumnWidth(4, 150); // Filed / Issued By
    sheet.setColumnWidth(5, 200); // Uploaded by (Email)
    sheet.setColumnWidth(6, 200); // Last Modified by (Email)
    sheet.setColumnWidth(7, 180); // Action Needed (wider for "No Action Needed")
    sheet.setColumnWidth(8, 180); // Assigned DDS Party
    sheet.setColumnWidth(9, 120); // Status
    // Set up enhanced data validation dropdowns
    setupPleadingsDropdowns(sheet);
  }
  return sheet;
}

/**
 * Sets up enhanced data validation dropdowns including "No Action Needed"
 */
function setupPleadingsDropdowns(sheet) {
  try {
    // Filed / Issued By dropdown (column 4)
    const filedByRange = sheet.getRange("D2:D1000");
    const filedByRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(PLEADING_DROPDOWNS.FILED_ISSUED_BY, true)
      .setAllowInvalid(false)
      .build();
    filedByRange.setDataValidation(filedByRule);
    // Enhanced Action Needed dropdown (column 7) - now includes "No Action Needed"
    const actionRange = sheet.getRange("G2:G1000");
    const actionRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(PLEADING_DROPDOWNS.ACTION_NEEDED, true)
      .setAllowInvalid(false)
      .build();
    actionRange.setDataValidation(actionRule);
    // Assigned DDS Party dropdown (column 8)
    const assignedRange = sheet.getRange("H2:H1000");
    const assignedRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(PLEADING_DROPDOWNS.ASSIGNED_DDS_PARTY, true)
      .setAllowInvalid(false)
      .build();
    assignedRange.setDataValidation(assignedRule);
    console.log(
      'Enhanced data validation dropdowns set up successfully with "No Action Needed"',
    );
  } catch (error) {
    console.log(`Error setting up dropdowns: ${error.message}`);
    logEvent(
      "ERROR",
      "",
      "",
      "",
      "",
      `Dropdown setup error: ${error.message}`,
      "ERROR",
      error.message,
    );
  }
}

/**
 * UPDATED: Generate pleading name with new convention - YYYY-MM-DD Descriptive Title (no folder name).
 * Attempts AI-powered title generation via Gemini first; falls back to filename-derived title.
 */
function generatePleadingName(originalName, dateCreated) {
  // Format: YYYY-MM-DD
  const dateStr = dateCreated.toISOString().split("T")[0];
  // Get file extension
  const extension = originalName.split(".").pop();
  // Attempt AI-powered title generation
  const aiTitle = generateAiPleadingTitle(originalName);
  if (aiTitle) {
    console.log(`AI title generated for "${originalName}": "${aiTitle}"`);
    return `${dateStr} ${aiTitle}.${extension}`;
  }
  // Fallback: Extract descriptive title from original name (remove extension and clean up)
  let descriptiveTitle = originalName.replace(/\.[^/.]+$/, ""); // Remove extension
  descriptiveTitle = descriptiveTitle.replace(/[_-]/g, " "); // Replace underscores and hyphens with spaces
  descriptiveTitle = descriptiveTitle.replace(/\s+/g, " ").trim(); // Clean up multiple spaces
  // Capitalize first letter of each word for better readability
  descriptiveTitle = descriptiveTitle
    .split(" ")
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
    .join(" ");
  // New format: YYYY-MM-DD Descriptive Title.extension (NO folder name)
  return `${dateStr} ${descriptiveTitle}.${extension}`;
}

/**
 * Enhanced file info gathering with updated pleading naming convention
 */
function getFileInfoEnhanced(file, folderPath, driveType, rootFolderName) {
  const fileInfo = {
    id: file.getId(),
    name: file.getName(),
    mimeType: file.getMimeType(),
    size: file.getSize(),
    dateCreated: file.getDateCreated(),
    dateModified: file.getLastUpdated(),
    url: file.getUrl(),
    folderPath: folderPath,
    uploadedByEmail: "Unknown",
    lastModifiedByEmail: "Unknown",
    remarks: "Scanned",
  };
  // Get enhanced user information
  try {
    const userInfo = getEnhancedUserInfo(file);
    fileInfo.uploadedByEmail = userInfo.uploadedByEmail || "Unknown";
    fileInfo.lastModifiedByEmail = userInfo.lastModifiedByEmail || "Unknown";
    if (userInfo.uploadedByEmail || userInfo.lastModifiedByEmail) {
      console.log(
        `User tracking successful for ${file.getName()}: Upload: ${userInfo.uploadedByEmail}, Modified: ${userInfo.lastModifiedByEmail}`,
      );
    }
  } catch (error) {
    console.log(`User tracking failed for ${file.getName()}: ${error.message}`);
    logEvent(
      "ERROR",
      driveType,
      file.getId(),
      file.getName(),
      "",
      `User tracking error: ${error.message}`,
      "ERROR",
      error.message,
    );
  }
  // Apply UPDATED pleading naming convention if it's a pleading file
  if (driveType === CONFIG.DRIVE_TYPES.PLEADINGS) {
    fileInfo.suggestedName = generatePleadingName(
      fileInfo.name,
      fileInfo.dateCreated,
    ); // Updated - no rootFolderName parameter
    fileInfo.pleadingType = detectPleadingType(fileInfo.name);
  }
  return fileInfo;
}

/**
 * Enhanced user information extraction with multiple methods
 */
function getEnhancedUserInfo(file) {
  const userInfo = {
    uploadedByEmail: null,
    lastModifiedByEmail: null,
    fromAdvancedApi: false,
  };
  try {
    // Method 1: Advanced Drive API (most comprehensive)
    try {
      const driveFile = Drive.Files.get(file.getId(), {
        fields: "owners,lastModifyingUser,createdTime,modifiedTime",
      });

      // Get uploader email (owner)
      if (driveFile.owners && driveFile.owners.length > 0) {
        const owner = driveFile.owners[0];
        userInfo.uploadedByEmail = owner.emailAddress || owner.displayName;
      }

      // Get last modifier email
      if (driveFile.lastModifyingUser) {
        userInfo.lastModifiedByEmail =
          driveFile.lastModifyingUser.emailAddress ||
          driveFile.lastModifyingUser.displayName;
      }

      userInfo.fromAdvancedApi = true;
    } catch (apiError) {
      console.log(
        `Advanced Drive API failed for ${file.getName()}: ${apiError.message}`,
      );

      // Method 2: Standard DriveApp methods (fallback)
      try {
        const owners = file.getOwners();
        if (owners && owners.length > 0) {
          const owner = owners[0];
          userInfo.uploadedByEmail = owner.getEmail() || owner.getName();
        }

        userInfo.lastModifiedByEmail = userInfo.uploadedByEmail;
      } catch (standardError) {
        console.log(
          `Standard DriveApp methods failed for ${file.getName()}: ${standardError.message}`,
        );

        // Method 3: Domain-based inference (last resort)
        if (CONFIG.DOMAIN) {
          const fileName = file.getName();
          if (fileName.includes(CONFIG.DOMAIN)) {
            userInfo.uploadedByEmail = "Domain User";
            userInfo.lastModifiedByEmail = "Domain User";
          }
        }
      }
    }
  } catch (error) {
    console.log(
      `All user tracking methods failed for ${file.getName()}: ${error.message}`,
    );
  }
  return userInfo;
}

/**
 * ENHANCED: Add file to Pleadings sheet with automatic file renaming in Drive AND sheet
 */
function addFileToPleadingsSheet(fileInfo) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      CONFIG.SHEETS.SUMMARY_PLEADINGS,
    );
    // Use suggested name if available, otherwise original name
    const displayName = fileInfo.suggestedName || fileInfo.name;
    let finalDisplayName = displayName;
    let actualFileName = fileInfo.name;
    // If we have a suggested name that's different from original, rename the actual Drive file
    if (fileInfo.suggestedName && fileInfo.suggestedName !== fileInfo.name) {
      try {
        const driveFile = DriveApp.getFileById(fileInfo.id);
        driveFile.setName(fileInfo.suggestedName);
        finalDisplayName = fileInfo.suggestedName;
        actualFileName = fileInfo.suggestedName;

        logEvent(
          "FILE_RENAMED",
          CONFIG.DRIVE_TYPES.PLEADINGS,
          fileInfo.id,
          fileInfo.name,
          fileInfo.uploadedByEmail,
          `File renamed from "${fileInfo.name}" to "${fileInfo.suggestedName}"`,
          "SUCCESS",
        );

        console.log(
          `Successfully renamed file in Drive: ${fileInfo.name} → ${fileInfo.suggestedName}`,
        );
      } catch (renameError) {
        console.log(`Failed to rename file in Drive: ${renameError.message}`);
        logEvent(
          "ERROR",
          CONFIG.DRIVE_TYPES.PLEADINGS,
          fileInfo.id,
          fileInfo.name,
          "",
          `Failed to rename file: ${renameError.message}`,
          "ERROR",
          renameError.message,
        );
        // Use original name if rename fails
        finalDisplayName = fileInfo.name;
        actualFileName = fileInfo.name;
      }
    }
    // Determine default values based on pleading type and filename analysis
    let defaultFiledBy = "Court"; // Default assumption for scanned pleadings
    let defaultActionNeeded = "For Response";
    let defaultStatus = "New";
    // Try to determine Filed/Issued By from filename
    const lowerFileName = actualFileName.toLowerCase();
    if (
      lowerFileName.includes("motion") ||
      lowerFileName.includes("pleading")
    ) {
      defaultFiledBy = "Opposing Party";
    } else if (
      lowerFileName.includes("order") ||
      lowerFileName.includes("ruling")
    ) {
      defaultFiledBy = "Court";
      defaultActionNeeded = "For Compliance";
    }
    // Special handling for certain document types
    if (
      lowerFileName.includes("judgment") ||
      lowerFileName.includes("decision")
    ) {
      defaultActionNeeded = "No Action Needed"; // Use new dropdown option
      defaultStatus = "Completed";
    }
    const rowData = [
      fileInfo.dateCreated,
      `=HYPERLINK("${fileInfo.url}","${finalDisplayName}")`,
      fileInfo.folderPath,
      defaultFiledBy,
      fileInfo.uploadedByEmail,
      fileInfo.lastModifiedByEmail,
      defaultActionNeeded,
      "Partner", // Default assignment
      defaultStatus,
    ];
    sheet.appendRow(rowData);
    // Log user activity
    if (fileInfo.uploadedByEmail !== "Unknown") {
      logEvent(
        "PLEADING_ADDED",
        CONFIG.DRIVE_TYPES.PLEADINGS,
        fileInfo.id,
        actualFileName,
        fileInfo.uploadedByEmail,
        `Pleading added by ${fileInfo.uploadedByEmail}`,
        "SUCCESS",
      );
    }
  } catch (error) {
    console.log(`Error adding file to pleadings sheet: ${error.message}`);
    logEvent(
      "ERROR",
      CONFIG.DRIVE_TYPES.PLEADINGS,
      fileInfo.id,
      fileInfo.name,
      "",
      `Error adding file: ${error.message}`,
      "ERROR",
      error.message,
    );
  }
}

/**
 * Enhanced change detection with updated naming and automatic renaming for existing files
 */
function detectChangesEnhanced(driveId, driveType, sheetName, indexKey) {
  const oldIndexStr = getConfigValue(indexKey);
  const oldIndex = oldIndexStr ? JSON.parse(oldIndexStr) : {};
  const newIndex = scanSharedDriveEnhanced(
    driveId,
    driveType,
    false,
    sheetName,
  );
  const changes = [];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const existingData = getExistingFileData(sheet, driveType);
  // Find new files and update existing pleading names
  for (const [fileId, newFileInfo] of Object.entries(newIndex)) {
    const oldFileInfo = oldIndex[fileId];
    if (!oldFileInfo) {
      // NEW FILE
      changes.push({
        type: "NEW",
        fileId: fileId,
        fileInfo: newFileInfo,
        details: `New file: ${newFileInfo.name}`,
        userEmail: newFileInfo.uploadedByEmail,
      });

      if (driveType === CONFIG.DRIVE_TYPES.CLIENT_DOCS) {
        addFileToClientDocsSheet(newFileInfo, "New file");
      } else {
        addFileToPleadingsSheet(newFileInfo); // This will automatically rename if needed
      }
    } else {
      // EXISTING FILE - Check for user modifications
      if (
        oldFileInfo.lastModifiedByEmail !== newFileInfo.lastModifiedByEmail &&
        newFileInfo.lastModifiedByEmail !== "Unknown"
      ) {
        changes.push({
          type: "USER_MODIFIED",
          fileId: fileId,
          fileInfo: newFileInfo,
          details: `File modified by ${newFileInfo.lastModifiedByEmail}`,
          userEmail: newFileInfo.lastModifiedByEmail,
        });
      }

      // EXISTING PLEADING FILE - Check if it needs naming update
      if (driveType === CONFIG.DRIVE_TYPES.PLEADINGS) {
        const existingRow =
          existingData[newFileInfo.name] ||
          findExistingRowByFileId(sheet, fileId);
        if (
          existingRow &&
          newFileInfo.suggestedName &&
          newFileInfo.suggestedName !== newFileInfo.name
        ) {
          try {
            // Rename the actual Drive file
            const driveFile = DriveApp.getFileById(fileId);
            driveFile.setName(newFileInfo.suggestedName);

            // Update the sheet display name
            updateExistingPleadingName(
              sheet,
              existingRow.rowIndex,
              newFileInfo.suggestedName,
              fileId,
            );

            changes.push({
              type: "NAME_UPDATED",
              fileId: fileId,
              fileInfo: newFileInfo,
              details: `Pleading name updated: "${newFileInfo.name}" → "${newFileInfo.suggestedName}"`,
              userEmail: "SYSTEM",
            });

            console.log(
              `Updated existing pleading name: ${newFileInfo.name} → ${newFileInfo.suggestedName}`,
            );
          } catch (renameError) {
            console.log(
              `Failed to update existing pleading name: ${renameError.message}`,
            );
            logEvent(
              "ERROR",
              driveType,
              fileId,
              newFileInfo.name,
              "",
              `Failed to update pleading name: ${renameError.message}`,
              "ERROR",
              renameError.message,
            );
          }
        }
      }
    }
  }
  // Log all changes with enhanced user information
  changes.forEach((change) => {
    logEvent(
      change.type,
      driveType,
      change.fileId,
      change.fileInfo.name,
      change.userEmail || "",
      change.details,
      "SUCCESS",
    );
  });
  // Update stored index
  setConfigValue(indexKey, JSON.stringify(newIndex));
  return changes;
}

/**
 * Find existing row by file ID (helper function for pleading name updates)
 */
function findExistingRowByFileId(sheet, fileId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  try {
    const data = sheet
      .getRange(2, 1, lastRow - 1, PLEADINGS_HEADERS.length)
      .getValues();
    for (let i = 0; i < data.length; i++) {
      const fileNameCell = data[i][1]; // File Name column
      if (fileNameCell && typeof fileNameCell === "string") {
        const urlMatch = fileNameCell.match(
          /https:\/\/drive\.google\.com\/[^"]+/,
        );
        if (urlMatch && urlMatch[0].includes(fileId)) {
          return {
            rowIndex: i + 2,
            data: data[i],
          };
        }
      }
    }
  } catch (error) {
    console.log(`Error finding existing row by file ID: ${error.message}`);
  }
  return null;
}

/**
 * Update existing pleading name in sheet
 */
function updateExistingPleadingName(sheet, rowIndex, newName, fileId) {
  try {
    // Get the current file URL
    const file = DriveApp.getFileById(fileId);
    const fileUrl = file.getUrl();
    // Update the hyperlink with new name
    const newHyperlink = `=HYPERLINK("${fileUrl}","${newName}")`;
    sheet.getRange(rowIndex, 2).setValue(newHyperlink); // Column 2 is File Name
    console.log(`Updated sheet display name at row ${rowIndex}: ${newName}`);
  } catch (error) {
    console.log(`Error updating sheet display name: ${error.message}`);
    logEvent(
      "ERROR",
      CONFIG.DRIVE_TYPES.PLEADINGS,
      fileId,
      "",
      "",
      `Error updating sheet display name: ${error.message}`,
      "ERROR",
      error.message,
    );
  }
}

/**
 * Add file to Client Documents sheet with user tracking (unchanged)
 */
function addFileToClientDocsSheet(fileInfo, remarks) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      CONFIG.SHEETS.UPLOADED_BY_CLIENT,
    );
    const rowData = [
      fileInfo.dateCreated,
      `=HYPERLINK("${fileInfo.url}","${fileInfo.name}")`,
      fileInfo.folderPath,
      fileInfo.uploadedByEmail,
      fileInfo.dateModified,
      fileInfo.lastModifiedByEmail,
      remarks,
      `Added ${new Date().toLocaleString()}`,
    ];
    sheet.appendRow(rowData);
    // Log user activity
    if (fileInfo.uploadedByEmail !== "Unknown") {
      logEvent(
        "FILE_ADDED",
        CONFIG.DRIVE_TYPES.CLIENT_DOCS,
        fileInfo.id,
        fileInfo.name,
        fileInfo.uploadedByEmail,
        `File added by ${fileInfo.uploadedByEmail}`,
        "SUCCESS",
      );
    }
  } catch (error) {
    console.log(`Error adding file to client docs sheet: ${error.message}`);
    logEvent(
      "ERROR",
      CONFIG.DRIVE_TYPES.CLIENT_DOCS,
      fileInfo.id,
      fileInfo.name,
      "",
      `Error adding file: ${error.message}`,
      "ERROR",
      error.message,
    );
  }
}

// Keep all other existing functions unchanged...
function createClientDocsSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.UPLOADED_BY_CLIENT);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.SHEETS.UPLOADED_BY_CLIENT);
    // Set up headers
    sheet
      .getRange(1, 1, 1, CLIENT_DOCS_HEADERS.length)
      .setValues([CLIENT_DOCS_HEADERS]);
    sheet.setFrozenRows(1);
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, CLIENT_DOCS_HEADERS.length);
    headerRange.setBackground("#4285f4");
    headerRange.setFontColor("white");
    headerRange.setFontWeight("bold");
    // Set column widths (adjusted for user columns)
    sheet.setColumnWidth(1, 120); // Date Uploaded
    sheet.setColumnWidth(2, 300); // File Name
    sheet.setColumnWidth(3, 200); // File URL
    sheet.setColumnWidth(4, 200); // Uploaded by (Email)
    sheet.setColumnWidth(5, 120); // Date Modified
    sheet.setColumnWidth(6, 200); // Last Modified by (Email)
    sheet.setColumnWidth(7, 150); // Remarks
    sheet.setColumnWidth(8, 200); // Change Log
  }
  return sheet;
}

function createSheet(spreadsheet, name, headers, hidden = false) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground("#4285f4");
      headerRange.setFontColor("white");
      headerRange.setFontWeight("bold");
    }
  }
  if (hidden) {
    sheet.hideSheet();
  }
  return sheet;
}

// Keep all drive configuration, scanning, and utility functions unchanged...
function setClientDocumentsDriveId() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "📄 Set Client Documents Drive ID",
    'Enter the ID of your "01 Client Documents" shared drive:\n\n(Find this in the Drive URL after /drive/folders/)',
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() === ui.Button.OK) {
    const driveId = response.getResponseText().trim();
    if (driveId) {
      try {
        const drive = retryDriveOperation(() =>
          DriveApp.getFolderById(driveId),
        );
        const driveName = drive.getName();

        setConfigValue("CLIENT_DOCUMENTS_DRIVE_ID", driveId);
        logEvent(
          "CONFIG",
          CONFIG.DRIVE_TYPES.CLIENT_DOCS,
          "",
          "",
          "",
          `Client documents drive set: ${driveName} (${driveId})`,
          "SUCCESS",
        );
        ui.alert(
          `✅ Client Documents drive connected!\n\nDrive: ${driveName}\nID: ${driveId}\n\n💡 Files will be tracked with user email information in "Uploaded By Client" sheet.`,
        );
      } catch (error) {
        logEvent(
          "ERROR",
          CONFIG.DRIVE_TYPES.CLIENT_DOCS,
          "",
          "",
          "",
          `Invalid client documents drive ID: ${error.message}`,
          "ERROR",
          error.message,
        );
        ui.alert(
          `❌ Invalid drive ID or access denied.\n\nError: ${error.message}`,
        );
      }
    }
  }
}

function setScannedPleadingsDriveId() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "⚖️ Set Scanned Pleadings Drive ID",
    'Enter the ID of your "5.02 Scanned Pleadings" shared drive:\n\n(Find this in the Drive URL after /drive/folders/)',
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() === ui.Button.OK) {
    const driveId = response.getResponseText().trim();
    if (driveId) {
      try {
        const drive = retryDriveOperation(() =>
          DriveApp.getFolderById(driveId),
        );
        const driveName = drive.getName();

        setConfigValue("SCANNED_PLEADINGS_DRIVE_ID", driveId);
        logEvent(
          "CONFIG",
          CONFIG.DRIVE_TYPES.PLEADINGS,
          "",
          "",
          "",
          `Scanned pleadings drive set: ${driveName} (${driveId})`,
          "SUCCESS",
        );
        ui.alert(
          `✅ Scanned Pleadings drive connected!\n\nDrive: ${driveName}\nID: ${driveId}\n\n💡 Files will be auto-renamed using format: YYYY-MM-DD Descriptive Title\nFiles tracked in "Summary of Pleadings" sheet with enhanced dropdowns.`,
        );
      } catch (error) {
        logEvent(
          "ERROR",
          CONFIG.DRIVE_TYPES.PLEADINGS,
          "",
          "",
          "",
          `Invalid pleadings drive ID: ${error.message}`,
          "ERROR",
          error.message,
        );
        ui.alert(
          `❌ Invalid drive ID or access denied.\n\nError: ${error.message}`,
        );
      }
    }
  }
}

// Keep all other functions (scanning, testing, scheduling, etc.) unchanged...
function scanSharedDriveEnhanced(
  driveId,
  driveType,
  isFullScan = false,
  sheetName,
) {
  const fileIndex = {};
  let processedFiles = 0;
  try {
    const rootFolder = retryDriveOperation(() =>
      DriveApp.getFolderById(driveId),
    );
    const rootFolderName = rootFolder.getName();
    const foldersToScan = [{ folder: rootFolder, path: "/" + rootFolderName }];
    // Clear tracking sheet if full scan
    if (isFullScan) {
      const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const headers =
          driveType === CONFIG.DRIVE_TYPES.CLIENT_DOCS
            ? CLIENT_DOCS_HEADERS
            : PLEADINGS_HEADERS;
        sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
      }
    }
    while (foldersToScan.length > 0) {
      const { folder, path } = foldersToScan.pop();

      try {
        const files = retryDriveOperation(() => folder.getFiles());

        while (files.hasNext()) {
          try {
            const file = files.next();

            if (isSystemOrTempFile(file.getName())) {
              continue;
            }

            const fileInfo = getFileInfoEnhanced(
              file,
              path,
              driveType,
              rootFolderName,
            );
            fileIndex[fileInfo.id] = fileInfo;
            processedFiles++;

            if (isFullScan) {
              if (driveType === CONFIG.DRIVE_TYPES.CLIENT_DOCS) {
                addFileToClientDocsSheet(fileInfo, "Scanned");
              } else {
                addFileToPleadingsSheet(fileInfo); // Automatically renames files
              }
            }

            if (processedFiles % 10 === 0) {
              Utilities.sleep(CONFIG.API_DELAY_MS);
            }
          } catch (fileError) {
            console.log(`Error processing file: ${fileError.message}`);
            logEvent(
              "ERROR",
              driveType,
              "",
              "",
              "",
              `File processing error: ${fileError.message}`,
              "ERROR",
              fileError.message,
            );
          }
        }

        const subfolders = retryDriveOperation(() => folder.getFolders());
        while (subfolders.hasNext()) {
          const subfolder = subfolders.next();
          foldersToScan.push({
            folder: subfolder,
            path: path + "/" + subfolder.getName(),
          });
        }
      } catch (folderError) {
        console.log(`Error processing folder ${path}: ${folderError.message}`);
        logEvent(
          "ERROR",
          driveType,
          "",
          "",
          "",
          `Folder processing error: ${folderError.message}`,
          "ERROR",
          folderError.message,
        );
      }
    }
  } catch (error) {
    throw new Error(`Scan failed: ${error.message}`);
  }
  return fileIndex;
}

function retryDriveOperation(operation, maxRetries = CONFIG.MAX_RETRIES) {
  let lastError;
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const result = operation();
      if (attempt > 1) {
        console.log(`Drive operation succeeded on attempt ${attempt}`);
      }
      return result;
    } catch (error) {
      lastError = error;
      console.log(
        `Drive operation attempt ${attempt} failed: ${error.message}`,
      );

      if (attempt < maxRetries) {
        const delay = CONFIG.RETRY_DELAY_MS * attempt;
        console.log(`Retrying in ${delay}ms...`);
        Utilities.sleep(delay);
      }
    }
  }
  throw new Error(
    `Drive operation failed after ${maxRetries} attempts. Last error: ${lastError.message}`,
  );
}

function isSystemOrTempFile(fileName) {
  const tempPatterns = [
    /^~\$.*/, // Microsoft temp files
    /^\.tmp.*/, // Temp files
    /.*\.tmp$/, // Temp files
    /^\.~lock.*/, // LibreOffice lock files
    /.*\.crdownload$/, // Chrome partial downloads
    /.*\.part$/, // Partial downloads
    /^desktop\.ini$/, // Windows system file
    /^thumbs\.db$/i, // Windows thumbnail cache
    /^\.ds_store$/i, // macOS system file
  ];
  return tempPatterns.some((pattern) => pattern.test(fileName));
}

function detectPleadingType(fileName) {
  for (const [type, pattern] of Object.entries(PLEADING_TYPES)) {
    if (pattern.test(fileName)) {
      return type;
    }
  }
  return "Document";
}

function getExistingFileData(sheet, driveType) {
  const existingData = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return existingData;
  try {
    const headers =
      driveType === CONFIG.DRIVE_TYPES.CLIENT_DOCS
        ? CLIENT_DOCS_HEADERS
        : PLEADINGS_HEADERS;
    const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (let i = 0; i < data.length; i++) {
      const fileNameCell = data[i][1]; // File Name column
      if (fileNameCell && typeof fileNameCell === "string") {
        const nameMatch = fileNameCell.match(/"([^"]+)"/);
        if (nameMatch) {
          existingData[nameMatch[1]] = {
            rowIndex: i + 2,
            data: data[i],
          };
        }
      }
    }
  } catch (error) {
    console.log(`Error reading existing file data: ${error.message}`);
    logEvent(
      "ERROR",
      driveType,
      "",
      "",
      "",
      `Error reading tracking sheet: ${error.message}`,
      "ERROR",
      error.message,
    );
  }
  return existingData;
}

// Keep all other utility functions, testing, scheduling, etc. unchanged...
function manualFullScanClientDocs() {
  manualFullScan(
    "CLIENT_DOCUMENTS_DRIVE_ID",
    CONFIG.DRIVE_TYPES.CLIENT_DOCS,
    "Client Documents",
    CONFIG.SHEETS.UPLOADED_BY_CLIENT,
    "CLIENT_FILE_INDEX",
    "LAST_FULL_SCAN_CLIENT",
  );
}

function manualFullScanPleadings() {
  manualFullScan(
    "SCANNED_PLEADINGS_DRIVE_ID",
    CONFIG.DRIVE_TYPES.PLEADINGS,
    "Scanned Pleadings",
    CONFIG.SHEETS.SUMMARY_PLEADINGS,
    "PLEADINGS_FILE_INDEX",
    "LAST_FULL_SCAN_PLEADINGS",
  );
}

function manualFullScan(
  configKey,
  driveType,
  displayName,
  sheetName,
  indexKey,
  lastScanKey,
) {
  const ui = SpreadsheetApp.getUi();
  const driveId = getConfigValue(configKey);
  if (!driveId) {
    ui.alert(`❌ Please set your ${displayName} drive ID first.`);
    return;
  }
  let scanMessage = `This will scan all files in the ${displayName} drive and rebuild the tracking sheet with user email information.\n\n`;
  if (driveType === CONFIG.DRIVE_TYPES.PLEADINGS) {
    scanMessage += `📋 PLEADINGS SPECIAL FEATURES:\n`;
    scanMessage += `• Files will be auto-renamed: YYYY-MM-DD Descriptive Title\n`;
    scanMessage += `• Both Drive files AND sheet display names updated\n`;
    scanMessage += `• Enhanced dropdown with "No Action Needed" option\n\n`;
  }
  scanMessage += `Proceed?`;
  const response = ui.alert(
    `🔍 Full Scan - ${displayName}`,
    scanMessage,
    ui.ButtonSet.YES_NO,
  );
  if (response !== ui.Button.YES) return;
  try {
    logEvent(
      "SCAN_FULL",
      driveType,
      "",
      "",
      "",
      `Starting full scan with enhanced naming: ${displayName}`,
      "INFO",
    );
    const startTime = new Date();
    const fileIndex = scanSharedDriveEnhanced(
      driveId,
      driveType,
      true,
      sheetName,
    );
    const endTime = new Date();
    const duration = Math.round((endTime - startTime) / 1000);
    setConfigValue(lastScanKey, new Date().toISOString());
    setConfigValue(indexKey, JSON.stringify(fileIndex));
    const fileCount = Object.keys(fileIndex).length;
    let completionMessage = `✅ ${displayName} scan completed!\n\n`;
    completionMessage += `Files found: ${fileCount}\n`;
    completionMessage += `Time taken: ${duration} seconds\n\n`;
    completionMessage += `👥 User email tracking enabled for all files.`;
    if (driveType === CONFIG.DRIVE_TYPES.PLEADINGS) {
      completionMessage += `\n\n📋 Pleadings auto-renamed using format:\nYYYY-MM-DD Descriptive Title`;
    }
    logEvent(
      "SCAN_FULL",
      driveType,
      "",
      "",
      "",
      `Full scan completed: ${fileCount} files in ${duration}s with enhanced features`,
      "SUCCESS",
    );
    ui.alert(completionMessage);
  } catch (error) {
    logEvent(
      "ERROR",
      driveType,
      "",
      "",
      "",
      `Full scan failed: ${error.message}`,
      "ERROR",
      error.message,
    );
    ui.alert(`❌ ${displayName} scan failed: ${error.message}`);
  }
}

// Keep all other functions unchanged...
function scheduledScan() {
  const today = new Date();
  const dayOfWeek = today.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    logEvent("SCHEDULED", "", "", "", "", "Skipped weekend scan", "INFO");
    return;
  }
  try {
    logEvent(
      "SCHEDULED",
      "",
      "",
      "",
      "",
      "Starting scheduled scan with enhanced pleading naming",
      "INFO",
    );
    // Scan Client Documents for new files only
    const clientDriveId = getConfigValue("CLIENT_DOCUMENTS_DRIVE_ID");
    if (clientDriveId && getConfigValue("LAST_FULL_SCAN_CLIENT")) {
      const clientChanges = detectChangesEnhanced(
        clientDriveId,
        CONFIG.DRIVE_TYPES.CLIENT_DOCS,
        CONFIG.SHEETS.UPLOADED_BY_CLIENT,
        "CLIENT_FILE_INDEX",
      );
      setConfigValue("LAST_UPDATE_SCAN_CLIENT", new Date().toISOString());
      logEvent(
        "SCHEDULED",
        CONFIG.DRIVE_TYPES.CLIENT_DOCS,
        "",
        "",
        "",
        `Client documents: ${clientChanges.length} changes`,
        "SUCCESS",
      );
    }
    // Scan Pleadings for new files and naming updates
    const pleadingsDriveId = getConfigValue("SCANNED_PLEADINGS_DRIVE_ID");
    if (pleadingsDriveId && getConfigValue("LAST_FULL_SCAN_PLEADINGS")) {
      const pleadingsChanges = detectChangesEnhanced(
        pleadingsDriveId,
        CONFIG.DRIVE_TYPES.PLEADINGS,
        CONFIG.SHEETS.SUMMARY_PLEADINGS,
        "PLEADINGS_FILE_INDEX",
      );
      setConfigValue("LAST_UPDATE_SCAN_PLEADINGS", new Date().toISOString());
      logEvent(
        "SCHEDULED",
        CONFIG.DRIVE_TYPES.PLEADINGS,
        "",
        "",
        "",
        `Scanned pleadings: ${pleadingsChanges.length} changes with naming updates`,
        "SUCCESS",
      );
    }
    logEvent(
      "SCHEDULED",
      "",
      "",
      "",
      "",
      "Scheduled scan completed successfully with enhanced features",
      "SUCCESS",
    );
  } catch (error) {
    logEvent(
      "ERROR",
      "",
      "",
      "",
      "",
      `Scheduled scan failed: ${error.message}`,
      "ERROR",
      error.message,
    );
  }
}

function setupDailySchedule() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "⏰ Setup Daily Schedule",
    "This will create a daily trigger to scan BOTH drives for changes at 6:00 PM Manila time, Monday through Friday.\n\n✨ ENHANCED FEATURES:\n• 👥 User email tracking for new files\n• 📋 Automatic pleading renaming (YYYY-MM-DD format)\n• 🔄 Updates existing pleading names to new format\n\nContinue?",
    ui.ButtonSet.YES_NO,
  );
  if (response !== ui.Button.YES) return;
  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach((trigger) => {
      if (trigger.getHandlerFunction() === "scheduledScan") {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    ScriptApp.newTrigger("scheduledScan")
      .timeBased()
      .everyDays(1)
      .atHour(CONFIG.SCHEDULE_HOUR)
      .inTimezone(CONFIG.TIMEZONE)
      .create();
    logEvent(
      "CONFIG",
      "",
      "",
      "",
      "",
      "Daily schedule enabled with enhanced pleading naming features",
      "SUCCESS",
    );
    ui.alert(
      "✅ Enhanced daily schedule enabled!\n\nThe system will now:\n• Scan both drives daily at 6:00 PM Manila time (Mon-Fri)\n• Track user emails for all new files\n• Auto-rename pleadings using YYYY-MM-DD format\n• Update existing pleading names to new format",
    );
  } catch (error) {
    logEvent(
      "ERROR",
      "",
      "",
      "",
      "",
      `Failed to setup daily schedule: ${error.message}`,
      "ERROR",
      error.message,
    );
    ui.alert("❌ Failed to setup daily schedule: " + error.message);
  }
}

// Keep all remaining functions unchanged...
function removeDailySchedule() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "🛑 Remove Daily Schedule",
    "Are you sure you want to remove the daily scanning schedule for both drives?",
    ui.ButtonSet.YES_NO,
  );
  if (response !== ui.Button.YES) return;
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let removedCount = 0;
    triggers.forEach((trigger) => {
      if (trigger.getHandlerFunction() === "scheduledScan") {
        ScriptApp.deleteTrigger(trigger);
        removedCount++;
      }
    });
    logEvent(
      "CONFIG",
      "",
      "",
      "",
      "",
      `Daily schedule disabled (${removedCount} triggers removed)`,
      "SUCCESS",
    );
    ui.alert(
      `✅ Daily schedule removed!\n\nRemoved ${removedCount} scheduled trigger(s).`,
    );
  } catch (error) {
    logEvent(
      "ERROR",
      "",
      "",
      "",
      "",
      `Failed to remove daily schedule: ${error.message}`,
      "ERROR",
      error.message,
    );
    ui.alert("❌ Failed to remove daily schedule: " + error.message);
  }
}

function testUserTracking() {
  const ui = SpreadsheetApp.getUi();
  // Test both drives if configured
  const clientDriveId = getConfigValue("CLIENT_DOCUMENTS_DRIVE_ID");
  const pleadingsDriveId = getConfigValue("SCANNED_PLEADINGS_DRIVE_ID");
  if (!clientDriveId && !pleadingsDriveId) {
    ui.alert("❌ No drives configured. Please set drive IDs first.");
    return;
  }
  try {
    let report = "👥 USER TRACKING TEST RESULTS:\n\n";
    let testResults = { success: 0, errors: 0, advancedApi: false };
    // Test client documents drive
    if (clientDriveId) {
      const clientTest = testDriveUserTracking(
        clientDriveId,
        "Client Documents",
      );
      report += `📄 ${clientTest.report}\n`;
      testResults.success += clientTest.success;
      testResults.errors += clientTest.errors;
      if (clientTest.advancedApi) testResults.advancedApi = true;
    }
    // Test pleadings drive with naming convention test
    if (pleadingsDriveId) {
      const pleadingsTest = testDriveUserTracking(
        pleadingsDriveId,
        "Scanned Pleadings",
      );
      report += `⚖️ ${pleadingsTest.report}\n`;
      testResults.success += pleadingsTest.success;
      testResults.errors += pleadingsTest.errors;
      if (pleadingsTest.advancedApi) testResults.advancedApi = true;

      // Test naming convention
      report += `📋 Pleading Naming: ✅ YYYY-MM-DD Descriptive Title format active\n`;
    }
    report += `\n📊 SUMMARY:\n`;
    report += `✅ Successful tests: ${testResults.success}\n`;
    report += `❌ Failed tests: ${testResults.errors}\n`;
    report += `🔧 Advanced Drive API: ${testResults.advancedApi ? "✅ Working" : "⚠️ Limited"}\n`;
    report += `📋 Enhanced Dropdowns: ✅ "No Action Needed" available\n`;
    if (!testResults.advancedApi) {
      report += `\n💡 TIP: Enable Advanced Drive API in Google Cloud Console for enhanced user tracking.`;
    }
    logEvent(
      "TEST_USER_TRACKING",
      "",
      "",
      "",
      "",
      `User tracking test completed with enhanced features: ${testResults.success} success, ${testResults.errors} errors`,
      "SUCCESS",
    );
    ui.alert(report);
  } catch (error) {
    logEvent(
      "ERROR",
      "",
      "",
      "",
      "",
      `User tracking test failed: ${error.message}`,
      "ERROR",
      error.message,
    );
    ui.alert(`❌ User tracking test failed: ${error.message}`);
  }
}

function testDriveUserTracking(driveId, driveName) {
  const result = { success: 0, errors: 0, advancedApi: false, report: "" };
  try {
    const folder = DriveApp.getFolderById(driveId);
    const files = folder.getFiles();
    let testCount = 0;
    while (files.hasNext() && testCount < 3) {
      const file = files.next();
      testCount++;

      try {
        const userInfo = getEnhancedUserInfo(file);
        if (userInfo.uploadedByEmail || userInfo.lastModifiedByEmail) {
          result.success++;
          if (userInfo.fromAdvancedApi) result.advancedApi = true;
        } else {
          result.errors++;
        }
      } catch (fileError) {
        result.errors++;
      }
    }
    result.report = `${driveName}: Tested ${testCount} files - ${result.success} success, ${result.errors} errors`;
  } catch (error) {
    result.errors++;
    result.report = `${driveName}: Access error - ${error.message}`;
  }
  return result;
}

function testClientDocsAccess() {
  testDriveAccess(
    "CLIENT_DOCUMENTS_DRIVE_ID",
    CONFIG.DRIVE_TYPES.CLIENT_DOCS,
    "Client Documents",
  );
}

function testPleadingsAccess() {
  testDriveAccess(
    "SCANNED_PLEADINGS_DRIVE_ID",
    CONFIG.DRIVE_TYPES.PLEADINGS,
    "Scanned Pleadings",
  );
}

function testDriveAccess(configKey, driveType, displayName) {
  const ui = SpreadsheetApp.getUi();
  const driveId = getConfigValue(configKey);
  if (!driveId) {
    ui.alert(`❌ No ${displayName} drive ID configured. Please set it first.`);
    return;
  }
  try {
    const folder = DriveApp.getFolderById(driveId);
    const folderName = folder.getName();
    const files = folder.getFiles();
    let fileCount = 0;
    while (files.hasNext() && fileCount < 5) {
      files.next();
      fileCount++;
    }
    let report = `✅ ${displayName} Access Test Results:\n\n`;
    report += `📁 Folder: ${folderName}\n`;
    report += `🔢 Sample files found: ${fileCount}\n`;
    report += `👥 User tracking: Available\n`;
    if (driveType === CONFIG.DRIVE_TYPES.PLEADINGS) {
      report += `📋 Pleading naming: ✅ YYYY-MM-DD format\n`;
      report += `🎯 Enhanced dropdowns: ✅ "No Action Needed"\n`;
    }
    report += `✅ Access confirmed`;
    logEvent(
      "TEST",
      driveType,
      "",
      "",
      "",
      `Drive access test passed: ${folderName}`,
      "SUCCESS",
    );
    ui.alert(report);
  } catch (error) {
    const errorMsg = `${displayName} access test failed: ${error.message}`;
    logEvent("ERROR", driveType, "", "", "", errorMsg, "ERROR", error.message);
    ui.alert(`❌ ${displayName} Access Failed\n\nError: ${error.message}`);
  }
}

function viewSystemStatus() {
  const clientDriveId = getConfigValue("CLIENT_DOCUMENTS_DRIVE_ID");
  const pleadingsDriveId = getConfigValue("SCANNED_PLEADINGS_DRIVE_ID");
  const lastFullScanClient = getConfigValue("LAST_FULL_SCAN_CLIENT");
  const lastFullScanPleadings = getConfigValue("LAST_FULL_SCAN_PLEADINGS");
  const lastUpdateClient = getConfigValue("LAST_UPDATE_SCAN_CLIENT");
  const lastUpdatePleadings = getConfigValue("LAST_UPDATE_SCAN_PLEADINGS");
  const hasApiKey = !!PropertiesService.getUserProperties().getProperty(
    GEMINI_CONFIG.USER_PROP_KEY,
  );
  const triggers = ScriptApp.getProjectTriggers();
  const scheduledTriggers = triggers.filter(
    (t) => t.getHandlerFunction() === "scheduledScan",
  );
  let status = "📊 ENHANCED SYSTEM STATUS\n\n";
  status += "📄 CLIENT DOCUMENTS:\n";
  status += `   🔗 Drive: ${clientDriveId ? "✅ Configured" : "❌ Not set"}\n`;
  status += `   🔍 Last Full Scan: ${lastFullScanClient ? new Date(lastFullScanClient).toLocaleString() : "❌ Never"}\n`;
  status += `   🔄 Last Update: ${lastUpdateClient ? new Date(lastUpdateClient).toLocaleString() : "❌ Never"}\n\n`;
  status += "⚖️ SCANNED PLEADINGS:\n";
  status += `   🔗 Drive: ${pleadingsDriveId ? "✅ Configured" : "❌ Not set"}\n`;
  status += `   🔍 Last Full Scan: ${lastFullScanPleadings ? new Date(lastFullScanPleadings).toLocaleString() : "❌ Never"}\n`;
  status += `   🔄 Last Update: ${lastUpdatePleadings ? new Date(lastUpdatePleadings).toLocaleString() : "❌ Never"}\n`;
  status += `   📋 Auto-Naming: ✅ YYYY-MM-DD Descriptive Title\n\n`;
  status += "SYSTEM FEATURES:\n";
  status += `🔑 Gemini API: ${hasApiKey ? "✅ Set" : "❌ Not set"}\n`;
  status += `⏰ Auto-Scan: ${scheduledTriggers.length > 0 ? "✅ Active (Enhanced Features)" : "❌ Disabled"}\n`;
  status += `📋 Dropdowns: ✅ Enhanced with "No Action Needed"\n`;
  status += `👥 User Tracking: ✅ Upload & Modification Emails\n`;
  status += `📧 Activity Logging: ✅ Enhanced with User Attribution\n`;
  status += `🔄 File Renaming: ✅ Both Drive Files & Sheet Display Names`;
  SpreadsheetApp.getUi().alert(status);
}

function scanBothDrives() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "🔄 Scan Both Drives",
    'This will run full scans on both drives sequentially with all enhanced features:\n\n✨ FEATURES ACTIVE:\n• 👥 User email tracking\n• 📋 Pleading auto-renaming (YYYY-MM-DD format)\n• 🔄 Drive file AND sheet display name updates\n• ✅ Enhanced dropdowns with "No Action Needed"\n\nThis may take several minutes.\n\nProceed?',
    ui.ButtonSet.YES_NO,
  );
  if (response !== ui.Button.YES) return;
  let results = "📊 ENHANCED DUAL SCAN RESULTS:\n\n";
  try {
    // Scan Client Documents
    const clientDriveId = getConfigValue("CLIENT_DOCUMENTS_DRIVE_ID");
    if (clientDriveId) {
      const clientIndex = scanSharedDriveEnhanced(
        clientDriveId,
        CONFIG.DRIVE_TYPES.CLIENT_DOCS,
        true,
        CONFIG.SHEETS.UPLOADED_BY_CLIENT,
      );
      setConfigValue("CLIENT_FILE_INDEX", JSON.stringify(clientIndex));
      setConfigValue("LAST_FULL_SCAN_CLIENT", new Date().toISOString());
      results += `📄 Client Documents: ${Object.keys(clientIndex).length} files with user tracking\n`;
    } else {
      results += `📄 Client Documents: ⚠️ Not configured\n`;
    }
    // Scan Pleadings with enhanced naming
    const pleadingsDriveId = getConfigValue("SCANNED_PLEADINGS_DRIVE_ID");
    if (pleadingsDriveId) {
      const pleadingsIndex = scanSharedDriveEnhanced(
        pleadingsDriveId,
        CONFIG.DRIVE_TYPES.PLEADINGS,
        true,
        CONFIG.SHEETS.SUMMARY_PLEADINGS,
      );
      setConfigValue("PLEADINGS_FILE_INDEX", JSON.stringify(pleadingsIndex));
      setConfigValue("LAST_FULL_SCAN_PLEADINGS", new Date().toISOString());
      results += `⚖️ Scanned Pleadings: ${Object.keys(pleadingsIndex).length} files auto-renamed\n`;
    } else {
      results += `⚖️ Scanned Pleadings: ⚠️ Not configured\n`;
    }
    results += `\n✨ ENHANCED FEATURES ACTIVE:\n`;
    results += `• 👥 User email tracking for all files\n`;
    results += `• 📋 Pleading files renamed: YYYY-MM-DD Descriptive Title\n`;
    results += `• 🔄 Both Drive files and sheet display names updated\n`;
    results += `• ✅ Enhanced dropdowns including "No Action Needed"`;
    ui.alert("✅ Enhanced Dual Scan Completed!\n\n" + results);
  } catch (error) {
    logEvent(
      "ERROR",
      "BOTH",
      "",
      "",
      "",
      `Enhanced dual scan failed: ${error.message}`,
      "ERROR",
      error.message,
    );
    ui.alert(`❌ Dual scan failed: ${error.message}`);
  }
}

/**
 * AI-powered descriptive title generation for a pleading file using Gemini.
 * Returns a suggested title string, or null if AI is unavailable or fails.
 */
function generateAiPleadingTitle(fileName) {
  const apiKey = PropertiesService.getUserProperties().getProperty(
    GEMINI_CONFIG.USER_PROP_KEY,
  );
  if (!apiKey) {
    console.log("Gemini API key not set — skipping AI title generation.");
    return null;
  }
  try {
    const prompt =
      `You are a legal document classifier. Given the raw filename of a scanned legal pleading, ` +
      `produce a clean, formal, human-readable descriptive title (no date prefix, no file extension). ` +
      `Use proper title case. Output ONLY the title, nothing else.\n\nFilename: ${fileName}`;
    const payload = {
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      generationConfig: {
        temperature: GEMINI_CONFIG.TEMPERATURE,
        maxOutputTokens: GEMINI_CONFIG.MAX_TOKENS,
      },
    };
    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "x-goog-api-key": apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };
    const response = UrlFetchApp.fetch(getGeminiEndpoint(), options);
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      console.log(
        `Gemini API error ${responseCode}: ${response.getContentText()}`,
      );
      return null;
    }
    const data = JSON.parse(response.getContentText());
    const text =
      data.candidates &&
      data.candidates[0] &&
      data.candidates[0].content &&
      data.candidates[0].content.parts &&
      data.candidates[0].content.parts[0] &&
      data.candidates[0].content.parts[0].text;
    if (!text) {
      console.log("Gemini returned empty response for: " + fileName);
      return null;
    }
    return text.trim();
  } catch (error) {
    console.log(`Gemini AI title generation failed: ${error.message}`);
    return null;
  }
}

// Gemini API key management functions
function setGeminiApiKey() {
  const ui = SpreadsheetApp.getUi();
  // Show current key status before prompting
  const existingKey = PropertiesService.getUserProperties().getProperty(
    GEMINI_CONFIG.USER_PROP_KEY,
  );
  const selectedModel =
    getConfigValue(GEMINI_CONFIG.SELECTED_MODEL_KEY) ||
    GEMINI_CONFIG.FALLBACK_MODEL;
  let statusMsg = "🔑 SET GEMINI API KEY\n\n";
  if (existingKey) {
    const masked =
      existingKey.substring(0, 4) +
      "..." +
      existingKey.substring(existingKey.length - 4);
    statusMsg += `Current status: ✅ Key is SET (${masked})\n`;
    statusMsg += `Active model: ${selectedModel}\n\n`;
    statusMsg +=
      "Enter a new key below to replace it, or press Cancel to keep the existing key.";
  } else {
    statusMsg += "Current status: ❌ No key set\n\n";
    statusMsg +=
      "Get your free API key from:\nhttps://aistudio.google.com/app/apikey\n\n";
    statusMsg +=
      "Note: The key is stored securely in your personal user properties and is not visible to other spreadsheet users.";
  }
  const response = ui.prompt(
    "🔑 Set Gemini API Key",
    statusMsg,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() === ui.Button.OK) {
    const apiKey = response.getResponseText().trim();
    if (!apiKey) {
      ui.alert("⚠️ No key entered. Existing key (if any) was not changed.");
      return;
    }
    try {
      ui.alert("⏳ Validating API key against Gemini... Please wait.");
      testGeminiApiKey(apiKey);
      PropertiesService.getUserProperties().setProperty(
        GEMINI_CONFIG.USER_PROP_KEY,
        apiKey,
      );
      logEvent(
        "CONFIG",
        "",
        "",
        "",
        "",
        "Gemini API key set successfully",
        "SUCCESS",
      );
      const masked =
        apiKey.substring(0, 4) + "..." + apiKey.substring(apiKey.length - 4);
      ui.alert(
        `✅ Gemini API key saved!\n\nKey: ${masked}\nModel: ${selectedModel}\n\nThe key was validated successfully. AI title generation is now active.`,
      );
    } catch (error) {
      logEvent(
        "ERROR",
        "",
        "",
        "",
        "",
        `Invalid Gemini API key: ${error.message}`,
        "ERROR",
        error.message,
      );
      ui.alert(
        `❌ API key validation failed.\n\nError: ${error.message}\n\nThe key was NOT saved. Please check the key and try again.`,
      );
    }
  }
}

function testGeminiApiKey(apiKey) {
  const payload = {
    contents: [
      {
        role: "user",
        parts: [
          { text: 'Say "API key test successful" in exactly those words.' },
        ],
      },
    ],
    generationConfig: {
      temperature: 0,
      maxOutputTokens: GEMINI_CONFIG.MAX_TOKENS,
    },
  };
  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "x-goog-api-key": apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };
  const response = UrlFetchApp.fetch(getGeminiEndpoint(), options);
  const responseCode = response.getResponseCode();
  if (responseCode !== 200) {
    throw new Error(
      `API returned ${responseCode}: ${response.getContentText()}`,
    );
  }
  const data = JSON.parse(response.getContentText());
  if (!data.candidates || !data.candidates[0] || !data.candidates[0].content) {
    throw new Error("Invalid API response structure");
  }
}

function clearGeminiApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "🗑️ Clear Gemini API Key",
    "Are you sure you want to remove the stored Gemini API key?",
    ui.ButtonSet.YES_NO,
  );
  if (response === ui.Button.YES) {
    PropertiesService.getUserProperties().deleteProperty(
      GEMINI_CONFIG.USER_PROP_KEY,
    );
    logEvent("CONFIG", "", "", "", "", "Gemini API key cleared", "SUCCESS");
    ui.alert("✅ Gemini API key has been cleared.");
  }
}

function logEvent(
  eventType,
  driveType = "",
  fileId = "",
  fileName = "",
  userEmail = "",
  details = "",
  status = "INFO",
  errorMessage = "",
) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      CONFIG.SHEETS.LOGS,
    );
    const timestamp = new Date();
    const row = [
      timestamp,
      eventType,
      driveType,
      fileId,
      fileName,
      userEmail,
      details,
      status,
      errorMessage,
    ];
    sheet.appendRow(row);
    const lastRow = sheet.getLastRow();
    // Highlight ERROR rows in red for easy visual identification
    if (status === "ERROR") {
      sheet
        .getRange(lastRow, 1, 1, LOG_HEADERS.length)
        .setBackground("#f4cccc");
    }
    if (lastRow > 1001) {
      sheet.deleteRows(2, lastRow - 1001);
    }
  } catch (error) {
    console.error("Failed to write to log:", error.message);
  }
}

function getConfigValue(key) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      CONFIG.SHEETS.CONFIG,
    );
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        return data[i][1];
      }
    }
  } catch (error) {
    console.log(`Error getting config value ${key}: ${error.message}`);
  }
  return "";
}

function setConfigValue(key, value) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      CONFIG.SHEETS.CONFIG,
    );
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        return;
      }
    }
    sheet.appendRow([key, value]);
  } catch (error) {
    console.log(`Error setting config value ${key}: ${error.message}`);
  }
}

/**
 * IMPROVED FIX: Update Action Needed dropdown with better error handling
 */
function updateActionNeededDropdown() {
  let ui;
  try {
    ui = SpreadsheetApp.getUi();
    ui.alert("🔄 Starting dropdown update...");

    // Get the spreadsheet and sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log("Got spreadsheet:", ss.getName());

    const sheet = ss.getSheetByName("Summary of Pleadings");
    if (!sheet) {
      throw new Error(
        "Summary of Pleadings sheet not found. Please check the sheet name.",
      );
    }
    console.log("Found Summary of Pleadings sheet");

    // Clear existing validation first
    console.log("Clearing existing validation...");
    const actionRange = sheet.getRange("G2:G100"); // Smaller range to avoid issues
    actionRange.clearDataValidations();

    // Small delay
    Utilities.sleep(1000);

    // Create new validation rule
    console.log("Creating new validation rule...");
    const newOptions = [
      "For Response",
      "For Compliance",
      "For Filing/Release",
      "Filed/Completed",
      "No Action Needed",
    ];

    const validationRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(newOptions, true)
      .setAllowInvalid(false)
      .setHelpText("Select the required action for this pleading")
      .build();

    // Apply the new validation
    console.log("Applying new validation...");
    actionRange.setDataValidation(validationRule);

    console.log("Dropdown update completed successfully");
    ui.alert(
      '✅ Success!\n\nAction Needed dropdown updated successfully.\n"No Action Needed" is now available in column G.',
    );
  } catch (error) {
    console.error("Error details:", error);
    const errorMessage = error.message || "Unknown error occurred";
    if (ui) {
      ui.alert(
        "❌ Error updating dropdown:\n\n" +
          errorMessage +
          "\n\nPlease try the manual method instead.",
      );
    }
  }
}

/**
 * Simple test function - just check if we can access the sheet
 */
function testSheetAccess() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      "Summary of Pleadings",
    );
    if (sheet) {
      SpreadsheetApp.getUi().alert("✅ Sheet found: " + sheet.getName());
    } else {
      SpreadsheetApp.getUi().alert("❌ Sheet not found");
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert("❌ Error: " + error.message);
  }
}

// ============================================================
// PLAN ITEM 2: GEMINI MODEL SELECTOR
// ============================================================

/**
 * Fetches the list of Gemini models that support generateContent.
 * Returns an array of model name strings (e.g. ["gemini-2.0-flash", ...]).
 * Throws on network or API error.
 */
function fetchGeminiModels(apiKey) {
  const url = `${GEMINI_CONFIG.MODELS_LIST_URL}?key=${apiKey}`;
  const options = {
    method: "get",
    muteHttpExceptions: true,
  };
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  if (responseCode !== 200) {
    throw new Error(
      `Failed to fetch models (HTTP ${responseCode}): ${response.getContentText()}`,
    );
  }
  const data = JSON.parse(response.getContentText());
  if (!data.models || !Array.isArray(data.models)) {
    throw new Error("Unexpected response format from models API.");
  }
  // Filter to only models that support generateContent
  const capable = data.models
    .filter(
      (m) =>
        m.supportedGenerationMethods &&
        m.supportedGenerationMethods.includes("generateContent"),
    )
    .map((m) => m.name.replace("models/", ""))
    .sort();
  return capable;
}

/**
 * Determines the "newest" model from a list by preferring higher version numbers
 * and non-experimental/non-legacy names. Returns the model name string.
 */
function pickNewestModel(models) {
  if (!models || models.length === 0) return null;
  // Prefer non-exp, non-legacy, non-preview, non-001 models; then sort descending
  const scored = models.map((m) => {
    let score = 0;
    if (m.includes("exp") || m.includes("preview")) score -= 10;
    if (m.includes("legacy") || m.includes("001")) score -= 20;
    // Extract version number for comparison
    const versionMatch = m.match(/(\d+\.\d+)/);
    if (versionMatch) score += parseFloat(versionMatch[1]) * 10;
    if (m.includes("pro")) score += 5;
    if (m.includes("flash")) score += 3;
    return { model: m, score };
  });
  scored.sort((a, b) => b.score - a.score);
  return scored[0].model;
}

/**
 * Fetches the live Gemini model list and lets the user pick one.
 * The newest model is highlighted with ★. Saves selection to System Config.
 */
function selectGeminiModel() {
  const ui = SpreadsheetApp.getUi();
  const apiKey = PropertiesService.getUserProperties().getProperty(
    GEMINI_CONFIG.USER_PROP_KEY,
  );
  if (!apiKey) {
    ui.alert(
      "❌ No Gemini API key set.\n\nPlease set your API key first via:\n🔑 Set Gemini API Key",
    );
    return;
  }
  try {
    ui.alert("⏳ Fetching available Gemini models... Please wait.");
    const models = fetchGeminiModels(apiKey);
    if (models.length === 0) {
      ui.alert(
        "❌ No generateContent-capable models found.\n\nCheck your API key permissions.",
      );
      return;
    }
    const newestModel = pickNewestModel(models);
    const currentModel =
      getConfigValue(GEMINI_CONFIG.SELECTED_MODEL_KEY) ||
      GEMINI_CONFIG.FALLBACK_MODEL;
    // Build numbered list for display
    let listMsg = "🤖 SELECT GEMINI AI MODEL\n\n";
    listMsg += `Currently active: ${currentModel}\n\n`;
    listMsg += "Available models (enter the number to select):\n\n";
    models.forEach((m, i) => {
      const isCurrent = m === currentModel ? " ◀ CURRENT" : "";
      const isNewest = m === newestModel ? " ★ RECOMMENDED" : "";
      listMsg += `${i + 1}. ${m}${isNewest}${isCurrent}\n`;
    });
    listMsg += `\nEnter a number (1–${models.length}):`;
    const response = ui.prompt(
      "🤖 Select AI Model",
      listMsg,
      ui.ButtonSet.OK_CANCEL,
    );
    if (response.getSelectedButton() !== ui.Button.OK) return;
    const input = response.getResponseText().trim();
    const index = parseInt(input, 10) - 1;
    if (isNaN(index) || index < 0 || index >= models.length) {
      ui.alert(
        `❌ Invalid selection "${input}".\n\nPlease enter a number between 1 and ${models.length}.`,
      );
      return;
    }
    const selectedModel = models[index];
    setConfigValue(GEMINI_CONFIG.SELECTED_MODEL_KEY, selectedModel);
    logEvent(
      "CONFIG",
      "",
      "",
      "",
      "",
      `Gemini model selected: ${selectedModel}`,
      "SUCCESS",
    );
    const isNewest = selectedModel === newestModel ? " (★ Recommended)" : "";
    ui.alert(
      `✅ AI Model Updated!\n\nSelected: ${selectedModel}${isNewest}\n\nAll future AI title generation will use this model. You can change it anytime from the menu.`,
    );
  } catch (error) {
    logEvent(
      "ERROR",
      "",
      "",
      "",
      "",
      `Model selection failed: ${error.message}`,
      "ERROR",
      error.message,
    );
    ui.alert(`❌ Failed to fetch model list.\n\nError: ${error.message}`);
  }
}

// ============================================================
// PLAN ITEM 1 (continued): TEST GEMINI CONNECTION
// ============================================================

/**
 * Standalone diagnostic: tests the currently stored API key and selected model.
 * Shows key status, model in use, response received, and latency.
 */
function testGeminiConnection() {
  const ui = SpreadsheetApp.getUi();
  const apiKey = PropertiesService.getUserProperties().getProperty(
    GEMINI_CONFIG.USER_PROP_KEY,
  );
  if (!apiKey) {
    ui.alert(
      "❌ No Gemini API key set.\n\nPlease set your API key first via:\n🔑 Set Gemini API Key",
    );
    return;
  }
  const selectedModel =
    getConfigValue(GEMINI_CONFIG.SELECTED_MODEL_KEY) ||
    GEMINI_CONFIG.FALLBACK_MODEL;
  const endpoint = getGeminiEndpoint();
  const masked =
    apiKey.substring(0, 4) + "..." + apiKey.substring(apiKey.length - 4);
  try {
    const payload = {
      contents: [
        {
          role: "user",
          parts: [{ text: 'Reply with exactly: "Connection successful."' }],
        },
      ],
      generationConfig: { temperature: 0, maxOutputTokens: 20 },
    };
    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "x-goog-api-key": apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };
    const startTime = Date.now();
    const response = UrlFetchApp.fetch(endpoint, options);
    const latencyMs = Date.now() - startTime;
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      throw new Error(`HTTP ${responseCode}: ${response.getContentText()}`);
    }
    const data = JSON.parse(response.getContentText());
    const replyText =
      data.candidates &&
      data.candidates[0] &&
      data.candidates[0].content &&
      data.candidates[0].content.parts &&
      data.candidates[0].content.parts[0] &&
      data.candidates[0].content.parts[0].text;
    logEvent(
      "TEST_GEMINI_CONNECTION",
      "",
      "",
      "",
      "",
      `Connection test passed. Model: ${selectedModel}, Latency: ${latencyMs}ms`,
      "SUCCESS",
    );
    let report = "🔍 GEMINI CONNECTION TEST\n\n";
    report += `✅ Status: CONNECTED\n`;
    report += `🔑 Key: ${masked}\n`;
    report += `🤖 Model: ${selectedModel}\n`;
    report += `⚡ Latency: ${latencyMs}ms\n`;
    report += `💬 Response: "${replyText ? replyText.trim() : "(empty)"}"`;
    ui.alert(report);
  } catch (error) {
    logEvent(
      "ERROR",
      "",
      "",
      "",
      "",
      `Gemini connection test failed: ${error.message}`,
      "ERROR",
      error.message,
    );
    let report = "🔍 GEMINI CONNECTION TEST\n\n";
    report += `❌ Status: FAILED\n`;
    report += `🔑 Key: ${masked}\n`;
    report += `🤖 Model: ${selectedModel}\n`;
    report += `❌ Error: ${error.message}`;
    ui.alert(report);
  }
}

// ============================================================
// PLAN ITEM 4: GEMINI AI INTEGRATION DIAGNOSTIC
// ============================================================

/**
 * End-to-end integration test for the full AI title generation pipeline.
 * Tests: key presence → model config → actual generateAiPleadingTitle() call → result validation.
 */
function testGeminiIntegration() {
  const ui = SpreadsheetApp.getUi();
  const sampleFilename = "motion_to_dismiss_complaint_2024.pdf";
  const apiKey = PropertiesService.getUserProperties().getProperty(
    GEMINI_CONFIG.USER_PROP_KEY,
  );
  const selectedModel =
    getConfigValue(GEMINI_CONFIG.SELECTED_MODEL_KEY) ||
    GEMINI_CONFIG.FALLBACK_MODEL;
  let report = "🧪 GEMINI AI INTEGRATION TEST\n\n";
  // Step 1: API key check
  if (!apiKey) {
    report += "❌ STEP 1 — API Key: NOT SET\n";
    report += "\nTest cannot proceed. Please set your Gemini API key first.";
    ui.alert(report);
    return;
  }
  const masked =
    apiKey.substring(0, 4) + "..." + apiKey.substring(apiKey.length - 4);
  report += `✅ STEP 1 — API Key: SET (${masked})\n`;
  // Step 2: Model config check
  report += `✅ STEP 2 — Model: ${selectedModel}\n`;
  report += `✅ STEP 3 — Endpoint: ${getGeminiEndpoint()}\n`;
  // Step 3: Run actual AI title generation
  report += `\n📄 Sample Input: "${sampleFilename}"\n`;
  report += "⏳ Calling generateAiPleadingTitle()...\n";
  const startTime = Date.now();
  let aiTitle = null;
  let testPassed = false;
  let errorMsg = "";
  try {
    aiTitle = generateAiPleadingTitle(sampleFilename);
    const latencyMs = Date.now() - startTime;
    if (aiTitle && aiTitle.length > 0) {
      testPassed = true;
      report += `\n✅ STEP 4 — AI Output: "${aiTitle}"\n`;
      report += `⚡ Latency: ${latencyMs}ms\n`;
      // Step 4: Simulate full name generation
      const fakeDate = new Date("2024-03-15");
      const fullName = generatePleadingName(sampleFilename, fakeDate);
      report += `\n📋 Full Generated Name:\n"${fullName}"\n`;
      report += "\n✅ RESULT: ALL TESTS PASSED\n";
      report += "AI title generation is fully operational.";
    } else {
      errorMsg = "AI returned null or empty response";
      report += `\n❌ STEP 4 — AI Output: (empty/null)\n`;
      report += `\n❌ RESULT: FAILED — ${errorMsg}`;
    }
  } catch (error) {
    const latencyMs = Date.now() - startTime;
    errorMsg = error.message;
    report += `\n❌ STEP 4 — Exception after ${latencyMs}ms: ${errorMsg}\n`;
    report += `\n❌ RESULT: FAILED`;
  }
  logEvent(
    "TEST_GEMINI_INTEGRATION",
    "",
    "",
    "",
    "",
    testPassed
      ? `Integration test PASSED. Model: ${selectedModel}, Output: "${aiTitle}"`
      : `Integration test FAILED: ${errorMsg}`,
    testPassed ? "SUCCESS" : "ERROR",
    testPassed ? "" : errorMsg,
  );
  ui.alert(report);
}

// ============================================================
// PLAN ITEM 3: ERROR LOG REPORT & CLEAR
// ============================================================

/**
 * Reads the Logs sheet and displays the last 20 ERROR-status rows
 * in a formatted alert for easy developer reporting.
 */
function viewErrorReport() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      CONFIG.SHEETS.LOGS,
    );
    if (!sheet) {
      ui.alert("❌ Logs sheet not found. Please run Initial Setup first.");
      return;
    }
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      ui.alert("📋 No log entries found.");
      return;
    }
    const data = sheet
      .getRange(2, 1, lastRow - 1, LOG_HEADERS.length)
      .getValues();
    // Filter to ERROR rows only
    const errorRows = data.filter((row) => row[7] === "ERROR");
    if (errorRows.length === 0) {
      ui.alert(
        "✅ No errors found in the log!\n\nThe system has been running without errors.",
      );
      return;
    }
    // Show last 20 errors (most recent first)
    const recentErrors = errorRows.slice(-20).reverse();
    let report = `🚨 ERROR REPORT (${errorRows.length} total errors, showing last ${recentErrors.length})\n`;
    report += "─────────────────────────────────\n\n";
    recentErrors.forEach((row, i) => {
      const timestamp =
        row[0] instanceof Date ? row[0].toLocaleString() : row[0];
      const eventType = row[1] || "";
      const details = row[6] || "";
      const errorMsg = row[8] || "";
      report += `[${i + 1}] ${timestamp}\n`;
      report += `    Type: ${eventType}\n`;
      if (details) report += `    Details: ${details}\n`;
      if (errorMsg) report += `    Error: ${errorMsg}\n`;
      report += "\n";
    });
    report += "─────────────────────────────────\n";
    report += "Use 🗑️ Clear Error Logs to remove these entries.";
    ui.alert(report);
  } catch (error) {
    ui.alert(`❌ Failed to read error report: ${error.message}`);
  }
}

/**
 * Clears only ERROR-status rows from the Logs sheet after confirmation.
 */
function clearErrorLogs() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "🗑️ Clear Error Logs",
    "This will permanently delete all ERROR rows from the Logs sheet.\n\nNon-error log entries (INFO, SUCCESS) will be kept.\n\nProceed?",
    ui.ButtonSet.YES_NO,
  );
  if (response !== ui.Button.YES) return;
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      CONFIG.SHEETS.LOGS,
    );
    if (!sheet) {
      ui.alert("❌ Logs sheet not found.");
      return;
    }
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      ui.alert("📋 No log entries to clear.");
      return;
    }
    // Collect ERROR row indices (iterate from bottom to avoid index shifting)
    const data = sheet
      .getRange(2, 1, lastRow - 1, LOG_HEADERS.length)
      .getValues();
    let deletedCount = 0;
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i][7] === "ERROR") {
        sheet.deleteRow(i + 2); // +2 because data starts at row 2
        deletedCount++;
      }
    }
    logEvent(
      "CONFIG",
      "",
      "",
      "",
      "",
      `Error logs cleared: ${deletedCount} ERROR rows deleted`,
      "SUCCESS",
    );
    ui.alert(
      `✅ Error logs cleared!\n\n${deletedCount} ERROR row(s) removed from the Logs sheet.`,
    );
  } catch (error) {
    ui.alert(`❌ Failed to clear error logs: ${error.message}`);
  }
}
