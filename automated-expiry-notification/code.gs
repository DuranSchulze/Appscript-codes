// =============================================================================
// AUTOMATED EXPIRY NOTIFICATION — Google Apps Script
// Sheet: "VISA automation" (headers row 2, data row 3+) | Logs: "LOGS"
// Entry point: runDailyCheck()  (run via time-based trigger)
// One-time setup: run installTrigger() manually once from the editor
// =============================================================================

// ---------------------------------------------------------------------------
// CONFIGURATION
// ---------------------------------------------------------------------------
var CONFIG = {
  SHEET_NAME: "VISA automation",
  LOGS_SHEET_NAME: "LOGS",
  AUTOMATION_SHEET_PROPERTY_KEY: "AUTOMATION_SHEET_NAME",
  HEADER_ROW: 2, // Row containing column headers
  DATA_START_ROW: 3, // First row of actual data
  TRIGGER_HOUR: 8, // 8 AM daily trigger
  REPLY_SCAN_TRIGGER_HOUR: 9, // 9 AM daily reply scan trigger
  SENDER_NAME: "DDS Office",
};

// Expected header names — must match row 2 of the sheet (case-insensitive trim)
var HEADERS = {
  NO: "No.",
  CLIENT_NAME: "Client Name",
  CLIENT_EMAIL: "Client Email",
  DOC_TYPE: "Type of ID/Document",
  EXPIRY_DATE: "Expiry Date",
  NOTICE_DATE: "Notice Date",
  REMARKS: "Remarks",
  ATTACHMENTS: "Attached Files",
  STATUS: "Status",
  STAFF_EMAIL: "Staff Email",
  SEND_MODE: "Send Mode",
  SENT_AT: "Sent At",
  SENT_THREAD_ID: "Sent Thread Id",
  SENT_MESSAGE_ID: "Sent Message Id",
  REPLY_STATUS: "Reply Status",
  REPLIED_AT: "Replied At",
  REPLY_KEYWORD: "Reply Keyword",
  OPEN_TOKEN: "Open Tracking Token",
  FIRST_OPENED_AT: "First Opened At",
  LAST_OPENED_AT: "Last Opened At",
  OPEN_COUNT: "Open Count",
};

// Optional accepted header aliases (mapped to the same logical field key)
var HEADER_ALIASES = {
  STAFF_EMAIL: ["Assigned Staff Email"],
  SENT_THREAD_ID: ["Sent Thread ID"],
  SENT_MESSAGE_ID: ["Sent Message ID"],
  SEND_MODE: ["Send Option", "Mode"],
  OPEN_TOKEN: ["Open Token"],
  FIRST_OPENED_AT: ["First Open At"],
  LAST_OPENED_AT: ["Last Open At"],
};

// Status values
var STATUS = {
  ACTIVE: "Active",
  SENT: "Sent",
  ERROR: "Error",
  SKIPPED: "Skipped",
};

// Row-level send mode values
var SEND_MODE = {
  AUTO: "Auto",
  HOLD: "Hold",
  MANUAL_ONLY: "Manual Only",
};

// Reply tracking status values
var REPLY_STATUS = {
  PENDING: "Pending",
  REPLIED: "Replied",
};

// Configurable properties
var PROP_KEYS = {
  SPREADSHEET_ID: "SPREADSHEET_ID",
  REPLY_KEYWORDS: "REPLY_KEYWORDS",
  AI_ENABLED: "AI_ENABLED",
  AI_PROVIDER: "AI_PROVIDER",
  AI_API_KEY: "AI_API_KEY",
  AI_MODEL: "AI_MODEL",
  FALLBACK_TEMPLATE_MODE: "FALLBACK_TEMPLATE_MODE",
  FALLBACK_TEMPLATE: "FALLBACK_TEMPLATE",
  OPEN_TRACKING_BASE_URL: "OPEN_TRACKING_BASE_URL",
};

var AI_PROVIDER = {
  GEMINI: "gemini",
};

var FALLBACK_TEMPLATE_MODE = {
  HARDCODED: "HARDCODED",
  PROPERTY: "PROPERTY",
};

var DEFAULT_REPLY_KEYWORDS = ["ACK", "RECEIVED", "OK"];
var DEFAULT_AI_MODEL = "models/gemini-1.5-flash";

// LOGS sheet column indices (1-based)
var LOG_COL = {
  TIMESTAMP: 1, // A
  CLIENT_NAME: 2, // B
  ACTION: 3, // C
  DETAIL: 4, // D
};

// =============================================================================
// CUSTOM MENU (appears in Google Sheets menu bar when the sheet is opened)
// =============================================================================

/**
 * Runs automatically when the spreadsheet is opened.
 * Adds the "Expiry Notifications" menu to the menu bar.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  rememberSpreadsheetId(SpreadsheetApp.getActiveSpreadsheet());
  ui.createMenu("Expiry Notifications")
    .addItem("Show Automation Status", "showAutomationStatus")
    .addItem("Initialize / Select Working Sheet", "initializeAutomationSheet")
    .addItem("Run Manual Check Now", "manualRunNow")
    .addSeparator()
    .addItem("Check Schedule Status", "checkScheduleStatus")
    .addItem("Activate Daily Schedule (8 AM)", "installTrigger")
    .addItem("Deactivate Daily Schedule", "removeTrigger")
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("Diagnostics")
        .addItem("Preview Target Dates (no emails sent)", "previewTargetDates")
        .addItem("Inspect Row...", "diagnosticInspectRow")
        .addItem("Send Test Email by No....", "diagnosticSendTestRow")
        .addItem(
          "Preview Effective Fallback Body",
          "previewFallbackTemplateBody",
        ),
    )
    .addSubMenu(
      ui
        .createMenu("Reply Tracking")
        .addItem("Run Reply Scan Now", "runReplyScanNow")
        .addItem("Set Reply Keywords", "setReplyKeywords")
        .addItem("View Reply Tracking Status", "showReplyTrackingStatus")
        .addSeparator()
        .addItem("Activate Reply Scan Schedule", "installReplyScanTrigger")
        .addItem("Deactivate Reply Scan Schedule", "removeReplyScanTrigger"),
    )
    .addSubMenu(
      ui
        .createMenu("AI Integration")
        .addItem("Toggle AI Generation", "toggleAiGeneration")
        .addItem("Set Gemini API Key", "setGeminiApiKey")
        .addItem("Select Gemini Model", "selectGeminiModel")
        .addItem("Test AI Connection", "testAiConnection")
        .addItem("View AI Status", "showAiStatus")
        .addSeparator()
        .addItem("Set Fallback Template", "setFallbackTemplate")
        .addItem("Toggle Fallback Source", "toggleFallbackTemplateSource")
        .addItem("Set Open Tracking URL", "setOpenTrackingBaseUrl"),
    )
    .addToUi();
}

/**
 * Menu alias for a more obvious status entry point.
 */
function showAutomationStatus() {
  checkScheduleStatus();
}

/**
 * One-time/anytime setup utility to select which sheet tab the automation should use.
 * Stores selection in document properties so tab renames can be reconfigured quickly.
 */
function initializeAutomationSheet() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets().filter(function (sheet) {
    return sheet.getName() !== CONFIG.LOGS_SHEET_NAME;
  });

  if (sheets.length === 0) {
    ui.alert(
      "Initialization",
      "No selectable sheet tabs were found.",
      ui.ButtonSet.OK,
    );
    return;
  }

  var currentSheetName = getConfiguredSheetName();
  var options = [];
  for (var i = 0; i < sheets.length; i++) {
    options.push(i + 1 + ". " + sheets[i].getName());
  }

  var response = ui.prompt(
    "Initialize Automation Sheet",
    'Select sheet tab by number (or type exact tab name):\n\nCurrent: "' +
      currentSheetName +
      '"\n\n' +
      options.join("\n"),
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var input = response.getResponseText().trim();
  if (!input) {
    ui.alert("Invalid input. Please enter a number or tab name.");
    return;
  }

  var index = parseInt(input, 10);
  var selectedSheetName = "";
  if (
    !isNaN(index) &&
    String(index) === input &&
    index >= 1 &&
    index <= sheets.length
  ) {
    selectedSheetName = sheets[index - 1].getName();
  } else {
    selectedSheetName = input;
  }

  var selectedSheet = ss.getSheetByName(selectedSheetName);
  if (!selectedSheet) {
    ui.alert(
      'Sheet "' + selectedSheetName + '" was not found. Please try again.',
    );
    return;
  }

  setConfiguredSheetName(selectedSheetName);
  ui.alert(
    "Initialization Complete",
    'Automation will now use sheet tab: "' + selectedSheetName + '".',
    ui.ButtonSet.OK,
  );
}

/**
 * Gets the selected automation sheet name from properties, falling back to CONFIG default.
 */
function getConfiguredSheetName() {
  var props = PropertiesService.getDocumentProperties();
  var saved = props.getProperty(CONFIG.AUTOMATION_SHEET_PROPERTY_KEY);
  return saved ? saved.trim() : CONFIG.SHEET_NAME;
}

/**
 * Persists automation sheet selection to document properties.
 */
function setConfiguredSheetName(sheetName) {
  var value = String(sheetName || "").trim();
  if (!value) return;
  PropertiesService.getDocumentProperties().setProperty(
    CONFIG.AUTOMATION_SHEET_PROPERTY_KEY,
    value,
  );
}

/**
 * Resolves configured automation sheet and returns both name and Sheet object.
 */
function resolveAutomationSheet(ss) {
  var sheetName = getConfiguredSheetName();
  return {
    sheetName: sheetName,
    sheet: ss.getSheetByName(sheetName),
  };
}

// =============================================================================
// SETTINGS / FEATURE CONFIG HELPERS
// =============================================================================

/**
 * Returns document properties object used by this automation.
 */
function getAutomationProperties() {
  return PropertiesService.getDocumentProperties();
}

/**
 * Saves current spreadsheet ID for trigger/webapp contexts.
 */
function rememberSpreadsheetId(ss) {
  try {
    if (!ss) return;
    var id = ss.getId();
    if (id) setPropString(PROP_KEYS.SPREADSHEET_ID, id);
  } catch (e) {}
}

/**
 * Returns spreadsheet in active context, or opens by stored ID.
 */
function getAutomationSpreadsheet() {
  var active = null;
  try {
    active = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}

  if (active) {
    rememberSpreadsheetId(active);
    return active;
  }

  var savedId = getPropString(PROP_KEYS.SPREADSHEET_ID, "");
  if (!savedId) {
    throw new Error(
      "Spreadsheet context unavailable. Open the spreadsheet once to initialize stored spreadsheet ID.",
    );
  }

  return SpreadsheetApp.openById(savedId);
}

/**
 * Returns a string property with optional fallback.
 */
function getPropString(key, fallbackValue) {
  var value = getAutomationProperties().getProperty(key);
  if (value === null || value === undefined || String(value).trim() === "") {
    return fallbackValue || "";
  }
  return String(value).trim();
}

/**
 * Saves a string property. Passing empty value clears the property.
 */
function setPropString(key, value) {
  var props = getAutomationProperties();
  var text = String(value === null || value === undefined ? "" : value).trim();
  if (!text) {
    props.deleteProperty(key);
    return;
  }
  props.setProperty(key, text);
}

/**
 * Returns boolean property value with fallback.
 */
function getPropBoolean(key, fallbackValue) {
  var raw = getAutomationProperties().getProperty(key);
  if (raw === null || raw === undefined || String(raw).trim() === "") {
    return !!fallbackValue;
  }
  return String(raw).toLowerCase().trim() === "true";
}

/**
 * Saves boolean property value.
 */
function setPropBoolean(key, value) {
  getAutomationProperties().setProperty(key, value ? "true" : "false");
}

/**
 * Returns configured send mode normalized to known values.
 */
function normalizeSendMode(modeValue) {
  var text = String(modeValue || "")
    .trim()
    .toLowerCase();
  if (!text) return SEND_MODE.AUTO;
  if (text === "auto") return SEND_MODE.AUTO;
  if (text === "hold") return SEND_MODE.HOLD;
  if (text === "manual only" || text === "manual-only" || text === "manual") {
    return SEND_MODE.MANUAL_ONLY;
  }
  return SEND_MODE.AUTO;
}

/**
 * Returns row-level send mode based on optional Send Mode column.
 */
function getRowSendMode(row, colMap) {
  if (!colMap.SEND_MODE) return SEND_MODE.AUTO;
  return normalizeSendMode(getCellStr(row, colMap.SEND_MODE));
}

/**
 * Returns skip reason for non-sendable modes, or empty string when sendable.
 */
function getSendModeSkipReason(sendMode) {
  if (sendMode === SEND_MODE.HOLD) {
    return "Skipped by Send Mode: Hold";
  }
  if (sendMode === SEND_MODE.MANUAL_ONLY) {
    return "Skipped by Send Mode: Manual Only";
  }
  return "";
}

/**
 * Gets configured reply keywords.
 */
function getReplyKeywords() {
  var raw = getPropString(PROP_KEYS.REPLY_KEYWORDS, "");
  if (!raw) return DEFAULT_REPLY_KEYWORDS.slice();
  var list = raw
    .split(",")
    .map(function (item) {
      return String(item || "")
        .trim()
        .toUpperCase();
    })
    .filter(function (item) {
      return !!item;
    });
  return list.length > 0 ? list : DEFAULT_REPLY_KEYWORDS.slice();
}

/**
 * Stores reply keywords from comma-separated text input.
 */
function setReplyKeywordsFromText(rawText) {
  var list = String(rawText || "")
    .split(",")
    .map(function (item) {
      return String(item || "")
        .trim()
        .toUpperCase();
    })
    .filter(function (item) {
      return !!item;
    });
  if (list.length === 0) list = DEFAULT_REPLY_KEYWORDS.slice();
  setPropString(PROP_KEYS.REPLY_KEYWORDS, list.join(", "));
  return list;
}

/**
 * Returns whether AI generation is enabled.
 */
function isAiGenerationEnabled() {
  return getPropBoolean(PROP_KEYS.AI_ENABLED, false);
}

/**
 * Returns configured AI model.
 */
function getAiModel() {
  return getPropString(PROP_KEYS.AI_MODEL, DEFAULT_AI_MODEL);
}

/**
 * Returns configured Gemini API key.
 */
function getAiApiKey() {
  return getPropString(PROP_KEYS.AI_API_KEY, "");
}

/**
 * Returns fallback template mode.
 */
function getFallbackTemplateMode() {
  var mode = getPropString(
    PROP_KEYS.FALLBACK_TEMPLATE_MODE,
    FALLBACK_TEMPLATE_MODE.HARDCODED,
  ).toUpperCase();
  return mode === FALLBACK_TEMPLATE_MODE.PROPERTY
    ? FALLBACK_TEMPLATE_MODE.PROPERTY
    : FALLBACK_TEMPLATE_MODE.HARDCODED;
}

/**
 * Returns fallback template text configured in properties.
 */
function getConfiguredFallbackTemplate() {
  return getPropString(PROP_KEYS.FALLBACK_TEMPLATE, "");
}

/**
 * Returns open tracking base URL configured for doGet tracking.
 */
function getOpenTrackingBaseUrl() {
  return getPropString(PROP_KEYS.OPEN_TRACKING_BASE_URL, "");
}

/**
 * Masks secrets for UI display.
 */
function maskSecret(value) {
  var text = String(value || "");
  if (!text) return "(not set)";
  if (text.length <= 8) return "********";
  return text.substring(0, 4) + "..." + text.substring(text.length - 4);
}

/**
 * Returns a new unique token for open tracking.
 */
function generateOpenTrackingToken() {
  return Utilities.getUuid();
}

/**
 * Shows whether the daily schedule is currently active and when it will next run.
 */
function checkScheduleStatus() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetConfig = resolveAutomationSheet(ss);
  var logsSheet = ensureLogsSheet(ss);
  var active = getTriggersByHandler("runDailyCheck");
  var replyTriggers = getTriggersByHandler("runReplyScan");

  var msg;
  if (active.length === 0) {
    msg =
      "Status: INACTIVE\n\nNo daily schedule is set up.\nUse 'Activate Daily Schedule (8 AM)' to enable it.";
  } else {
    var lines = ["Status: ACTIVE", "", active.length + " trigger(s) found:"];
    for (var j = 0; j < active.length; j++) {
      var t = active[j];
      lines.push(
        "  - Runs daily at " +
          CONFIG.TRIGGER_HOUR +
          ":00 Philippine Time (Asia/Manila)",
      );
    }
    if (active.length > 1) {
      lines.push("");
      lines.push(
        "Warning: " +
          active.length +
          " duplicate triggers detected. Run 'Deactivate' then 'Activate' to clean up.",
      );
    }
    msg = lines.join("\n");
  }

  msg +=
    '\n\nConfigured sheet: "' +
    sheetConfig.sheetName +
    '"' +
    (sheetConfig.sheet ? " (found)" : " (NOT FOUND)");
  if (!sheetConfig.sheet) {
    msg +=
      "\nUse 'Initialize / Select Working Sheet' to choose the correct tab.";
  }

  msg +=
    "\n\nReply tracking schedule: " +
    (replyTriggers.length > 0
      ? "ACTIVE (" +
        replyTriggers.length +
        " trigger(s), " +
        CONFIG.REPLY_SCAN_TRIGGER_HOUR +
        ":00 Asia/Manila)"
      : "INACTIVE");

  msg +=
    "\nAI generation: " +
    (isAiGenerationEnabled() ? "ENABLED" : "DISABLED") +
    " | Model: " +
    getAiModel();

  msg += "\n\n" + getLatestRunSummary(logsSheet);
  ui.alert("Daily Schedule Status", msg, ui.ButtonSet.OK);
}

/**
 * Returns all project triggers for a given handler function.
 */
function getTriggersByHandler(handlerName) {
  var triggers = ScriptApp.getProjectTriggers();
  var list = [];
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === handlerName) {
      list.push(triggers[i]);
    }
  }
  return list;
}

/**
 * Finds the newest SUMMARY entry in LOGS and returns it as a short status line.
 */
function getLatestRunSummary(logsSheet) {
  var lastRow = logsSheet.getLastRow();
  if (lastRow < 2) return "Last run: No run history yet.";

  var rows = logsSheet
    .getRange(2, LOG_COL.TIMESTAMP, lastRow - 1, LOG_COL.DETAIL)
    .getValues();

  for (var i = rows.length - 1; i >= 0; i--) {
    var action = String(rows[i][LOG_COL.ACTION - 1] || "")
      .trim()
      .toUpperCase();
    if (action !== "SUMMARY") continue;

    var timestamp = rows[i][LOG_COL.TIMESTAMP - 1];
    var detail = String(rows[i][LOG_COL.DETAIL - 1] || "").trim();
    var timestampText =
      timestamp instanceof Date
        ? Utilities.formatDate(
            timestamp,
            Session.getScriptTimeZone() || "Asia/Manila",
            "dd MMM yyyy hh:mm a",
          )
        : String(timestamp || "Unknown");

    return (
      "Last run: " + timestampText + "\n" + (detail || "(No summary detail)")
    );
  }

  return "Last run: No summary log yet.";
}

/**
 * Manual trigger: runs the daily check immediately and shows a confirmation dialog.
 * Called from the "Expiry Notifications > Run Manual Check Now" menu item.
 */
function manualRunNow() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    "Run Manual Check",
    "This will scan all Active/blank rows and send emails for any row whose target date is due (today or earlier), subject to Send Mode rules.\n\nProceed?",
    ui.ButtonSet.YES_NO,
  );
  if (confirm !== ui.Button.YES) return;

  try {
    runDailyCheck();
    ui.alert(
      "Done",
      "Manual check complete. Check the LOGS sheet for details.",
      ui.ButtonSet.OK,
    );
  } catch (e) {
    ui.alert("Error", "Manual check failed: " + e.message, ui.ButtonSet.OK);
  }
}

// =============================================================================
// ENTRY POINT
// =============================================================================

/**
 * Main function — called daily by the time-based trigger.
 * Reads headers from row 2, scans data from row 3+, computes target dates,
 * sends emails using remarks/AI/fallback body sources,
 * and updates status + optional metadata columns.
 */
function runDailyCheck() {
  var ss = getAutomationSpreadsheet();
  var sheetConfig = resolveAutomationSheet(ss);
  var visaSheet = sheetConfig.sheet;
  if (!visaSheet) {
    throw new Error(
      'Configured sheet "' +
        sheetConfig.sheetName +
        '" not found. Use "Initialize / Select Working Sheet" from the Expiry Notifications menu.',
    );
  }

  var logsSheet = ensureLogsSheet(ss);

  var colMap = buildColumnMap(visaSheet);
  var mapError = validateColumnMap(colMap);
  if (mapError) {
    appendLog(logsSheet, "", "ERROR", mapError);
    throw new Error(mapError);
  }

  var lastRow = visaSheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) {
    appendLog(logsSheet, "", "INFO", "No data rows found. Nothing to process.");
    return;
  }

  var numDataRows = lastRow - CONFIG.DATA_START_ROW + 1;
  var numCols = visaSheet.getLastColumn();
  var data = visaSheet
    .getRange(CONFIG.DATA_START_ROW, 1, numDataRows, numCols)
    .getValues();
  var totalRows = data.length;
  var today = getMidnight(new Date());
  var processed = 0,
    sent = 0,
    errors = 0;
  var autoActivated = 0;
  var skippedMode = 0;
  var skippedFuture = 0;
  var senderEmail = getSenderAccountEmail();
  var trackingEnabled = !!getOpenTrackingBaseUrl();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowIndex = CONFIG.DATA_START_ROW + i;

    var clientName = getCellStr(row, colMap.CLIENT_NAME);
    var clientEmail = getCellStr(row, colMap.CLIENT_EMAIL);
    var staffEmail = getCellStr(row, colMap.STAFF_EMAIL);
    var docType = getCellStr(row, colMap.DOC_TYPE);
    var expiryRaw = colMap.EXPIRY_DATE ? row[colMap.EXPIRY_DATE - 1] : "";
    var noticeStr = getCellStr(row, colMap.NOTICE_DATE);
    var remarks = getCellStr(row, colMap.REMARKS);
    var attachRaw = getCellStr(row, colMap.ATTACHMENTS);
    var status = getCellStr(row, colMap.STATUS);
    var sendMode = getRowSendMode(row, colMap);

    if (!isProcessableStatus(status)) continue;

    if (isStatusBlank(status)) {
      setStatus(visaSheet, rowIndex, colMap.STATUS, STATUS.ACTIVE);
      status = STATUS.ACTIVE;
      autoActivated++;
      appendLog(
        logsSheet,
        clientName,
        "INFO",
        "Blank Status auto-set to Active for processing.",
      );
    }

    var modeSkipReason = getSendModeSkipReason(sendMode);
    if (modeSkipReason) {
      appendLog(logsSheet, clientName, "SKIPPED", modeSkipReason);
      skippedMode++;
      continue;
    }

    processed++;

    var missing = [];
    if (!clientName) missing.push("Client Name");
    if (!clientEmail) missing.push("Client Email");
    if (!expiryRaw) missing.push("Expiry Date");
    if (!noticeStr) missing.push("Notice Date");
    if (missing.length > 0) {
      var errMsg = "Missing required field(s): " + missing.join(", ");
      setStatus(visaSheet, rowIndex, colMap.STATUS, STATUS.ERROR);
      appendLog(logsSheet, clientName, "ERROR", errMsg);
      errors++;
      continue;
    }

    var expiryDate =
      expiryRaw instanceof Date ? expiryRaw : new Date(expiryRaw);
    if (isNaN(expiryDate.getTime())) {
      setStatus(visaSheet, rowIndex, colMap.STATUS, STATUS.ERROR);
      appendLog(
        logsSheet,
        clientName,
        "ERROR",
        "Invalid Expiry Date: " + expiryRaw,
      );
      errors++;
      continue;
    }

    var offset = parseNoticeOffset(noticeStr);
    if (offset === null) {
      setStatus(visaSheet, rowIndex, colMap.STATUS, STATUS.ERROR);
      appendLog(
        logsSheet,
        clientName,
        "ERROR",
        'Cannot parse Notice Date: "' +
          noticeStr +
          '". Use: "N days/weeks/months before" or "On expiry date".',
      );
      errors++;
      continue;
    }

    var targetDate = computeTargetDate(expiryDate, offset);
    if (!isTargetDateDue(targetDate, today)) {
      skippedFuture++;
      continue;
    }

    var attachResult = resolveAttachments(attachRaw);
    if (attachResult.error) {
      setStatus(visaSheet, rowIndex, colMap.STATUS, STATUS.ERROR);
      appendLog(logsSheet, clientName, "ERROR", attachResult.error);
      errors++;
      continue;
    }

    var openToken = trackingEnabled ? generateOpenTrackingToken() : "";
    var emailContent = buildEmailContent(
      remarks,
      clientName,
      expiryDate,
      docType,
      openToken,
    );
    var subject = buildEmailSubject(docType, clientName, expiryDate);

    try {
      var sentMeta = sendReminderEmail(
        clientEmail,
        staffEmail,
        subject,
        emailContent.htmlBody,
        attachResult.blobs,
      );
      setStatus(visaSheet, rowIndex, colMap.STATUS, STATUS.SENT);
      setStaffEmail(visaSheet, rowIndex, colMap.STAFF_EMAIL, senderEmail);
      writePostSendMetadata(visaSheet, rowIndex, colMap, {
        sentAt: new Date(),
        senderEmail: senderEmail,
        openToken: openToken,
        threadId: sentMeta.threadId,
        messageId: sentMeta.messageId,
      });
      appendLog(
        logsSheet,
        clientName,
        "SENT",
        "Email sent to " +
          clientEmail +
          (staffEmail ? " (CC: " + staffEmail + ")" : "") +
          " | Mode: " +
          sendMode +
          " | Body: " +
          emailContent.source +
          (openToken ? " | Tracking: enabled" : "") +
          (senderEmail ? " | Sender: " + senderEmail : ""),
      );
      sent++;
    } catch (e) {
      setStatus(visaSheet, rowIndex, colMap.STATUS, STATUS.ERROR);
      appendLog(logsSheet, clientName, "ERROR", "Send failed: " + e.message);
      errors++;
    }
  }

  appendLog(
    logsSheet,
    "",
    "SUMMARY",
    "Run complete. Total Rows: " +
      totalRows +
      " | Eligible (Active/Blank): " +
      processed +
      " | Auto-Activated: " +
      autoActivated +
      " | Skipped (Mode): " +
      skippedMode +
      " | Skipped (Future): " +
      skippedFuture +
      " | Sent: " +
      sent +
      " | Errors: " +
      errors,
  );
}

// =============================================================================
// COLUMN MAP HELPERS
// =============================================================================

/**
 * Reads HEADER_ROW and returns a map of { HEADER_KEY: 1-based-col-index }.
 * Matching is case-insensitive and trims whitespace.
 */
function buildColumnMap(sheet) {
  var headerRow = sheet
    .getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  var reverseHeaders = {};
  for (var key in HEADERS) {
    reverseHeaders[HEADERS[key].toLowerCase().trim()] = key;
    var aliases = HEADER_ALIASES[key] || [];
    for (var a = 0; a < aliases.length; a++) {
      reverseHeaders[String(aliases[a]).toLowerCase().trim()] = key;
    }
  }
  var map = {};
  for (var c = 0; c < headerRow.length; c++) {
    var h = String(headerRow[c]).toLowerCase().trim();
    if (reverseHeaders[h]) {
      map[reverseHeaders[h]] = c + 1;
    }
  }
  return map;
}

/**
 * Validates that all required columns were found in the header row.
 * Returns an error string if any are missing, or null if OK.
 */
function validateColumnMap(colMap) {
  var required = [
    "CLIENT_NAME",
    "CLIENT_EMAIL",
    "EXPIRY_DATE",
    "NOTICE_DATE",
    "STATUS",
  ];
  var missing = [];
  for (var i = 0; i < required.length; i++) {
    if (!colMap[required[i]]) missing.push(HEADERS[required[i]]);
  }
  return missing.length > 0
    ? "Required column(s) not found in row " +
        CONFIG.HEADER_ROW +
        ": " +
        missing.join(", ")
    : null;
}

/**
 * Returns trimmed string value from a data row using a 1-based column index.
 */
function getCellStr(row, colIndex) {
  if (!colIndex) return "";
  return String(row[colIndex - 1] || "").trim();
}

/**
 * Returns true when a status value should be treated as Active.
 * Handles case differences from dropdown/manual entry (e.g., ACTIVE/active/Active).
 */
function isStatusActive(statusValue) {
  return (
    String(statusValue || "")
      .trim()
      .toLowerCase() === STATUS.ACTIVE.toLowerCase()
  );
}

/**
 * Returns true when status is blank/empty.
 */
function isStatusBlank(statusValue) {
  return String(statusValue || "").trim() === "";
}

/**
 * Returns true when a row should be considered for processing.
 * Current rule: process rows with Active or blank Status.
 */
function isProcessableStatus(statusValue) {
  return isStatusActive(statusValue) || isStatusBlank(statusValue);
}

/**
 * Compares a No. cell value with a user-entered No. value.
 * Supports text and numeric equivalence (e.g., 1 matches 1.0).
 */
function isSameNoValue(cellValue, inputValue) {
  var cellStr = String(
    cellValue === null || cellValue === undefined ? "" : cellValue,
  ).trim();
  var inputStr = String(
    inputValue === null || inputValue === undefined ? "" : inputValue,
  ).trim();

  if (!cellStr || !inputStr) return false;
  if (cellStr === inputStr) return true;

  var cellNum = Number(cellStr);
  var inputNum = Number(inputStr);
  if (!isNaN(cellNum) && !isNaN(inputNum)) return cellNum === inputNum;

  return cellStr.toLowerCase() === inputStr.toLowerCase();
}

/**
 * Finds a data row by No. column value.
 * Returns { rowNum, warning, error }.
 */
function findRowNumberByNo(sheet, colMap, noValue) {
  if (!colMap.NO) {
    return {
      rowNum: null,
      warning: "",
      error: 'Column "No." not found in row ' + CONFIG.HEADER_ROW + ".",
    };
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) {
    return { rowNum: null, warning: "", error: "No data rows found." };
  }

  var numDataRows = lastRow - CONFIG.DATA_START_ROW + 1;
  var noValues = sheet
    .getRange(CONFIG.DATA_START_ROW, colMap.NO, numDataRows, 1)
    .getValues();

  var matches = [];
  for (var i = 0; i < noValues.length; i++) {
    if (isSameNoValue(noValues[i][0], noValue)) {
      matches.push(CONFIG.DATA_START_ROW + i);
    }
  }

  if (matches.length === 0) {
    return {
      rowNum: null,
      warning: "",
      error: 'No row found for No. "' + noValue + '".',
    };
  }

  if (matches.length > 1) {
    return {
      rowNum: matches[0],
      warning:
        'Multiple rows found for No. "' +
        noValue +
        '". Using first match at row ' +
        matches[0] +
        ".",
      error: "",
    };
  }

  return { rowNum: matches[0], warning: "", error: "" };
}

/**
 * Writes a Status value to the Status column of a given row.
 */
function setStatus(sheet, rowIndex, statusColIndex, statusValue) {
  if (!statusColIndex) return;
  sheet.getRange(rowIndex, statusColIndex).setValue(statusValue);
}

/**
 * Writes the sender account email to the Staff Email column after successful send.
 */
function setStaffEmail(sheet, rowIndex, staffEmailColIndex, senderEmail) {
  if (!staffEmailColIndex || !senderEmail) return;
  sheet.getRange(rowIndex, staffEmailColIndex).setValue(senderEmail);
}

/**
 * Writes any value to a row/column if the column index exists.
 */
function setCellValueIfColumn(sheet, rowIndex, colIndex, value) {
  if (!colIndex) return;
  sheet.getRange(rowIndex, colIndex).setValue(value);
}

/**
 * Clears a row/column value if the column index exists.
 */
function clearCellValueIfColumn(sheet, rowIndex, colIndex) {
  if (!colIndex) return;
  sheet.getRange(rowIndex, colIndex).clearContent();
}

/**
 * Persists post-send metadata into optional columns.
 */
function writePostSendMetadata(sheet, rowIndex, colMap, meta) {
  var sentAt = meta && meta.sentAt ? meta.sentAt : new Date();
  var threadId = meta && meta.threadId ? String(meta.threadId) : "";
  var messageId = meta && meta.messageId ? String(meta.messageId) : "";
  var openToken = meta && meta.openToken ? String(meta.openToken) : "";

  setCellValueIfColumn(sheet, rowIndex, colMap.SENT_AT, sentAt);
  setCellValueIfColumn(sheet, rowIndex, colMap.SENT_THREAD_ID, threadId);
  setCellValueIfColumn(sheet, rowIndex, colMap.SENT_MESSAGE_ID, messageId);
  setCellValueIfColumn(sheet, rowIndex, colMap.OPEN_TOKEN, openToken);

  if (colMap.REPLY_STATUS) {
    setCellValueIfColumn(
      sheet,
      rowIndex,
      colMap.REPLY_STATUS,
      REPLY_STATUS.PENDING,
    );
  }
  clearCellValueIfColumn(sheet, rowIndex, colMap.REPLIED_AT);
  clearCellValueIfColumn(sheet, rowIndex, colMap.REPLY_KEYWORD);
}

/**
 * Returns the email account most likely used to send messages from this script.
 */
function getSenderAccountEmail() {
  try {
    var effectiveEmail = Session.getEffectiveUser().getEmail();
    if (effectiveEmail) return effectiveEmail;
  } catch (e) {}

  try {
    var activeEmail = Session.getActiveUser().getEmail();
    if (activeEmail) return activeEmail;
  } catch (e) {}

  return "";
}

// =============================================================================
// DATE HELPERS
// =============================================================================

/**
 * Returns a new Date set to midnight (00:00:00.000) for the given date.
 */
function getMidnight(date) {
  var d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

/**
 * Returns true if two dates fall on the same calendar day.
 */
function isSameDay(dateA, dateB) {
  var a = getMidnight(dateA);
  var b = getMidnight(dateB);
  return (
    a.getFullYear() === b.getFullYear() &&
    a.getMonth() === b.getMonth() &&
    a.getDate() === b.getDate()
  );
}

/**
 * Returns true when the target send date is due on or before reference date.
 */
function isTargetDateDue(targetDate, referenceDate) {
  return (
    getMidnight(targetDate).getTime() <= getMidnight(referenceDate).getTime()
  );
}

/**
 * Parses a Notice Date dropdown string into an offset object dynamically.
 * Supports any "N days/weeks/months before" pattern — no hardcoded list.
 * @returns {{ unit: string, value: number }} or null if unrecognized.
 */
function parseNoticeOffset(noticeStr) {
  var s = noticeStr.toLowerCase().trim();
  if (/^on(\s+the)?\s+expiry\s+date$/.test(s)) {
    return { unit: "days", value: 0 };
  }
  var m = s.match(/(\d+)\s*(day|days|week|weeks|month|months)\s+before/);
  if (m) {
    var value = parseInt(m[1], 10);
    var unitRaw = m[2];
    if (unitRaw === "week" || unitRaw === "weeks")
      return { unit: "days", value: value * 7 };
    if (unitRaw === "month" || unitRaw === "months")
      return { unit: "months", value: value };
    return { unit: "days", value: value };
  }
  return null;
}

/**
 * Computes the target send date from an expiry date and a parsed offset object.
 */
function computeTargetDate(expiryDate, offset) {
  var target = new Date(expiryDate);
  if (offset.unit === "days") {
    target.setDate(target.getDate() - offset.value);
  } else if (offset.unit === "months") {
    target.setMonth(target.getMonth() - offset.value);
  }
  return getMidnight(target);
}

/**
 * Formats a Date as "DD MMM YYYY" (e.g. "24 Mar 2026").
 */
function formatDate(date) {
  var months = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];
  return (
    date.getDate() + " " + months[date.getMonth()] + " " + date.getFullYear()
  );
}

// =============================================================================
// ATTACHMENT HELPERS
// =============================================================================

/**
 * Extracts a Google Drive file ID from a URL or returns the raw string as-is.
 * Handles:
 *   https://drive.google.com/file/d/FILE_ID/view
 *   https://drive.google.com/open?id=FILE_ID
 *   https://drive.google.com/uc?id=FILE_ID
 *   Raw file ID string
 */
function extractDriveFileId(urlOrId) {
  if (!urlOrId) return null;
  var s = urlOrId.trim();

  // /file/d/FILE_ID/
  var m = s.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];

  // ?id=FILE_ID or &id=FILE_ID
  m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];

  // Assume raw ID if it looks like one (alphanumeric + _ -)
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s)) return s;

  return null;
}

/**
 * Resolves a comma-separated list of Drive URLs/IDs into blobs.
 * @param {string} rawField - Cell value from "Attached Files" column.
 * @returns {{ blobs: Array, error: string|null }}
 */
function resolveAttachments(rawField) {
  if (!rawField || rawField.trim() === "") {
    return { blobs: [], error: null };
  }

  var entries = rawField.split(",");
  var blobs = [];

  for (var i = 0; i < entries.length; i++) {
    var entry = entries[i].trim();
    if (!entry) continue;

    var fileId = extractDriveFileId(entry);
    if (!fileId) {
      return {
        blobs: [],
        error: 'Cannot parse Drive file ID from: "' + entry + '"',
      };
    }

    try {
      var file = DriveApp.getFileById(fileId);
      blobs.push({ blob: file.getBlob(), name: file.getName() });
    } catch (e) {
      return {
        blobs: [],
        error:
          "Drive file not found or not accessible (ID: " +
          fileId +
          "): " +
          e.message,
      };
    }
  }

  return { blobs: blobs, error: null };
}

// =============================================================================
// EMAIL HELPERS
// =============================================================================

/**
 * Returns built-in fallback body template.
 */
function getDefaultFallbackTemplate() {
  return (
    "Good day, [Client Name],\n\n" +
    "This is a reminder that your [Document Type] is expiring on [Expiry Date].\n\n" +
    "Please take the necessary steps before the expiry date.\n\n" +
    "Thank you."
  );
}

/**
 * Returns configured fallback template text and source label.
 */
function resolveFallbackTemplateText() {
  var mode = getFallbackTemplateMode();
  if (mode === FALLBACK_TEMPLATE_MODE.PROPERTY) {
    var configured = getConfiguredFallbackTemplate();
    if (configured) {
      return { text: configured, source: "Fallback(Property)" };
    }
  }
  return { text: getDefaultFallbackTemplate(), source: "Fallback(Hardcoded)" };
}

/**
 * Builds full email content with source metadata.
 * Source priority:
 * 1) Remarks template
 * 2) AI generated body (when enabled and configured)
 * 3) Fallback template
 */
function buildEmailContent(
  remarks,
  clientName,
  expiryDate,
  docType,
  openToken,
) {
  var expiryStr = formatDate(expiryDate);
  var docTypeText = docType ? docType : "Visa/Permit";
  var bodyText = "";
  var source = "";

  if (remarks) {
    bodyText = applyTemplatePlaceholders(
      remarks,
      clientName,
      expiryStr,
      docTypeText,
    );
    source = "Remarks";
  } else {
    var aiResult = null;
    if (isAiGenerationEnabled()) {
      aiResult = tryGenerateAiEmailBody(clientName, expiryDate, docTypeText);
    }

    if (aiResult && aiResult.text) {
      bodyText = aiResult.text;
      source = "AI(" + aiResult.model + ")";
    } else {
      var fallback = resolveFallbackTemplateText();
      bodyText = applyTemplatePlaceholders(
        fallback.text,
        clientName,
        expiryStr,
        docTypeText,
      );
      source = fallback.source;
    }
  }

  var htmlBody = String(bodyText || "").replace(/\n/g, "<br>");
  htmlBody = injectOpenTrackingPixel(htmlBody, openToken);

  htmlBody = [
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#333;max-width:600px;line-height:1.6;">',
    htmlBody,
    '<hr style="border:none;border-top:1px solid #eee;margin:24px 0;">',
    '<p style="font-size:11px;color:#999;">This is an automated reminder. Please do not reply directly to this email.</p>',
    "</div>",
  ].join("\n");

  return {
    htmlBody: htmlBody,
    textBody: bodyText,
    source: source || "Unknown",
  };
}

/**
 * Backward-compatible helper that returns only HTML body.
 */
function buildEmailBody(remarks, clientName, expiryDate, docType) {
  return buildEmailContent(remarks, clientName, expiryDate, docType, "")
    .htmlBody;
}

/**
 * Injects open tracking pixel when token and tracking URL are configured.
 */
function injectOpenTrackingPixel(htmlBody, openToken) {
  var baseUrl = getOpenTrackingBaseUrl();
  if (!baseUrl || !openToken) return htmlBody;

  var separator = baseUrl.indexOf("?") >= 0 ? "&" : "?";
  var openUrl =
    baseUrl + separator + "mode=open&t=" + encodeURIComponent(openToken);

  return (
    htmlBody +
    '<br><img src="' +
    openUrl +
    '" width="1" height="1" style="display:none;" alt="" />'
  );
}

/**
 * Attempts to generate email body text via Gemini.
 * Returns null when AI is unavailable/unconfigured or generation fails.
 */
function tryGenerateAiEmailBody(clientName, expiryDate, docTypeText) {
  var apiKey = getAiApiKey();
  var model = getAiModel();
  if (!apiKey || !model) return null;

  var modelPath = model;
  if (modelPath.indexOf("models/") !== 0) {
    modelPath = "models/" + modelPath;
  }

  var prompt = [
    "Write a concise but professional visa/document expiry reminder email body.",
    "Tone: courteous, formal, and actionable.",
    "Do not include a subject line.",
    "Use plain text with short paragraphs.",
    "Client Name: " + clientName,
    "Document Type: " + docTypeText,
    "Expiry Date: " + formatDate(expiryDate),
  ].join("\n");

  var endpoint =
    "https://generativelanguage.googleapis.com/v1beta/" +
    modelPath +
    ":generateContent?key=" +
    encodeURIComponent(apiKey);

  var payload = {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.5,
      maxOutputTokens: 400,
    },
  };

  for (var attempt = 1; attempt <= 2; attempt++) {
    try {
      var response = UrlFetchApp.fetch(endpoint, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      });

      var statusCode = response.getResponseCode();
      var body = response.getContentText() || "";
      if (statusCode < 200 || statusCode >= 300) {
        if (attempt < 2) {
          Utilities.sleep(350);
          continue;
        }
        return null;
      }

      var parsed = JSON.parse(body);
      var text = extractGeminiText(parsed);
      if (!text) {
        if (attempt < 2) {
          Utilities.sleep(350);
          continue;
        }
        return null;
      }

      return {
        text: text.trim(),
        model: modelPath,
      };
    } catch (e) {
      if (attempt < 2) {
        Utilities.sleep(350);
        continue;
      }
      return null;
    }
  }

  return null;
}

/**
 * Extracts generated text from Gemini response payload.
 */
function extractGeminiText(payload) {
  if (!payload || !payload.candidates || payload.candidates.length === 0) {
    return "";
  }

  var candidate = payload.candidates[0];
  var parts =
    candidate && candidate.content && candidate.content.parts
      ? candidate.content.parts
      : [];

  var textChunks = [];
  for (var i = 0; i < parts.length; i++) {
    var partText = String(parts[i].text || "").trim();
    if (partText) textChunks.push(partText);
  }
  return textChunks.join("\n\n");
}

/**
 * Applies supported template placeholders in a case-insensitive way.
 */
function applyTemplatePlaceholders(
  templateText,
  clientName,
  expiryStr,
  docType,
) {
  return String(templateText || "")
    .replace(/\[\s*client\s*name\s*\]/gi, clientName || "")
    .replace(/\[\s*date\s*of\s*(expiration|expiry)\s*\]/gi, expiryStr || "")
    .replace(/\[\s*expiry\s*date\s*\]/gi, expiryStr || "")
    .replace(/\[\s*document\s*type\s*\]/gi, docType || "Visa/Permit");
}

/**
 * Builds the email subject line.
 * Uses Type of ID/Document if available, otherwise falls back to "Visa/Permit".
 */
function buildEmailSubject(docType, clientName, expiryDate) {
  var docLabel = docType ? docType : "Visa/Permit";
  return (
    "Reminder: " +
    docLabel +
    " Expiry on " +
    formatDate(expiryDate) +
    " \u2013 " +
    clientName
  );
}

/**
 * Sends the reminder email via GmailApp.
 * @param {string} clientEmail - To address.
 * @param {string} staffEmail  - CC address (may be empty).
 * @param {string} subject     - Email subject.
 * @param {string} htmlBody    - HTML email body.
 * @param {Array}  blobItems   - Array of {blob, name} objects.
 */
function sendReminderEmail(
  clientEmail,
  staffEmail,
  subject,
  htmlBody,
  blobItems,
) {
  var options = {
    htmlBody: htmlBody,
    name: CONFIG.SENDER_NAME,
  };

  if (staffEmail) {
    options.cc = staffEmail;
  }

  if (blobItems && blobItems.length > 0) {
    var attachments = blobItems.map(function (item) {
      return item.blob.setName(item.name);
    });
    options.attachments = attachments;
  }

  GmailApp.sendEmail(clientEmail, subject, "", options);

  // Best-effort metadata lookup from Sent mailbox.
  return lookupRecentSentMessageMeta(clientEmail, subject);
}

/**
 * Looks up recently-sent message metadata by recipient + subject.
 */
function lookupRecentSentMessageMeta(clientEmail, subject) {
  var meta = { threadId: "", messageId: "" };
  if (!clientEmail || !subject) return meta;

  try {
    var query =
      "in:sent to:(" +
      escapeGmailQueryValue(clientEmail) +
      ') subject:("' +
      escapeGmailQueryValue(subject) +
      '") newer_than:7d';
    var threads = GmailApp.search(query, 0, 5);
    if (!threads || threads.length === 0) return meta;

    var thread = threads[0];
    var messages = thread.getMessages();
    var latest = messages[messages.length - 1];
    meta.threadId = thread.getId() || "";
    meta.messageId = latest && latest.getId ? latest.getId() : "";
  } catch (e) {}

  return meta;
}

/**
 * Escapes double quotes in Gmail query fragments.
 */
function escapeGmailQueryValue(value) {
  return String(value || "").replace(/"/g, '\\"');
}

// =============================================================================
// LOGS SHEET HELPERS
// =============================================================================

/**
 * Ensures the LOGS sheet exists with the correct headers.
 * Creates it if missing. Returns the sheet.
 */
function ensureLogsSheet(ss) {
  var sheet = ss.getSheetByName(CONFIG.LOGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.LOGS_SHEET_NAME);
    var headers = ["Timestamp", "Client Name", "Action", "Detail"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet
      .getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#4a86e8")
      .setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(LOG_COL.TIMESTAMP, 160);
    sheet.setColumnWidth(LOG_COL.CLIENT_NAME, 200);
    sheet.setColumnWidth(LOG_COL.ACTION, 90);
    sheet.setColumnWidth(LOG_COL.DETAIL, 450);
  }
  return sheet;
}

/**
 * Appends one log row to the LOGS sheet with color-coded Action cell.
 * Signature: appendLog(logsSheet, clientName, action, detail)
 */
function appendLog(logsSheet, clientName, action, detail) {
  logsSheet.appendRow([new Date(), clientName, action, detail]);
  var lastRow = logsSheet.getLastRow();
  var actionCell = logsSheet.getRange(lastRow, LOG_COL.ACTION);
  if (action === "SENT") {
    actionCell.setBackground("#d9ead3").setFontColor("#274e13");
  } else if (action === "ERROR") {
    actionCell.setBackground("#fce8e6").setFontColor("#a61c00");
  } else if (action === "SKIPPED") {
    actionCell.setBackground("#fff2cc").setFontColor("#7f6000");
  } else if (action === "SUMMARY") {
    actionCell.setBackground("#e8eaf6").setFontColor("#1a237e");
  } else {
    actionCell.setBackground(null).setFontColor(null);
  }
}

// =============================================================================
// TRIGGER MANAGEMENT
// =============================================================================

/**
 * One-time setup: installs a daily 8 AM time-based trigger for runDailyCheck().
 * Run this function ONCE manually from the Apps Script editor (Run > installTrigger).
 */
function installTrigger() {
  rememberSpreadsheetId(SpreadsheetApp.getActiveSpreadsheet());
  removeTrigger(); // clean up any existing triggers first

  ScriptApp.newTrigger("runDailyCheck")
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.TRIGGER_HOUR)
    .inTimezone("Asia/Manila")
    .create();

  var msg =
    "Daily schedule activated. runDailyCheck() will run automatically every day at " +
    CONFIG.TRIGGER_HOUR +
    ":00 Philippine Time (Asia/Manila).";
  Logger.log(msg);
  try {
    SpreadsheetApp.getUi().alert(
      "Schedule Activated",
      msg,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  } catch (e) {}
}

/**
 * Installs daily reply scan trigger.
 */
function installReplyScanTrigger() {
  rememberSpreadsheetId(SpreadsheetApp.getActiveSpreadsheet());
  removeReplyScanTrigger();
  ScriptApp.newTrigger("runReplyScan")
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.REPLY_SCAN_TRIGGER_HOUR)
    .inTimezone("Asia/Manila")
    .create();

  var msg =
    "Reply scan schedule activated at " +
    CONFIG.REPLY_SCAN_TRIGGER_HOUR +
    ":00 Asia/Manila.";
  try {
    SpreadsheetApp.getUi().alert(
      "Reply Tracking",
      msg,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  } catch (e) {}
  Logger.log(msg);
}

/**
 * Removes reply scan triggers.
 */
function removeReplyScanTrigger() {
  var triggers = getTriggersByHandler("runReplyScan");
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  var msg =
    triggers.length > 0
      ? "Reply scan schedule deactivated. " +
        triggers.length +
        " trigger(s) removed."
      : "No active reply scan schedule found.";
  try {
    SpreadsheetApp.getUi().alert(
      "Reply Tracking",
      msg,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  } catch (e) {}
  Logger.log(msg);
}

/**
 * UI wrapper: run reply scan now.
 */
function runReplyScanNow() {
  var ui = SpreadsheetApp.getUi();
  try {
    var summary = runReplyScan();
    ui.alert("Reply Scan Complete", summary, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("Reply Scan Error", e.message, ui.ButtonSet.OK);
  }
}

/**
 * Trigger-safe reply scan routine.
 */
function runReplyScan() {
  var ss = getAutomationSpreadsheet();
  var sheetConfig = resolveAutomationSheet(ss);
  var sheet = sheetConfig.sheet;
  if (!sheet) {
    throw new Error(
      'Configured sheet "' +
        sheetConfig.sheetName +
        '" not found. Use "Initialize / Select Working Sheet".',
    );
  }

  var logsSheet = ensureLogsSheet(ss);
  var colMap = buildColumnMap(sheet);
  var mapError = validateColumnMap(colMap);
  if (mapError) {
    appendLog(logsSheet, "", "ERROR", mapError);
    throw new Error(mapError);
  }

  if (!colMap.SENT_THREAD_ID || !colMap.REPLY_STATUS) {
    var requiredMsg =
      'Reply scan skipped: add columns "Sent Thread Id" and "Reply Status" to use reply tracking.';
    appendLog(logsSheet, "", "INFO", requiredMsg);
    return requiredMsg;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) {
    var noRowsMsg = "Reply scan complete. No data rows found.";
    appendLog(logsSheet, "", "INFO", noRowsMsg);
    return noRowsMsg;
  }

  var numDataRows = lastRow - CONFIG.DATA_START_ROW + 1;
  var numCols = sheet.getLastColumn();
  var data = sheet
    .getRange(CONFIG.DATA_START_ROW, 1, numDataRows, numCols)
    .getValues();

  var keywords = getReplyKeywords();
  var scanned = 0;
  var updated = 0;
  var skipped = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowIndex = CONFIG.DATA_START_ROW + i;
    var status = getCellStr(row, colMap.STATUS);
    var replyStatus = getCellStr(row, colMap.REPLY_STATUS);
    var clientEmail = getCellStr(row, colMap.CLIENT_EMAIL);
    var clientName = getCellStr(row, colMap.CLIENT_NAME);
    var threadId = getCellStr(row, colMap.SENT_THREAD_ID);
    var sentAtRaw = colMap.SENT_AT ? row[colMap.SENT_AT - 1] : "";
    var sentAt = sentAtRaw instanceof Date ? sentAtRaw : new Date(sentAtRaw);

    if (!status || String(status).toLowerCase() !== STATUS.SENT.toLowerCase()) {
      skipped++;
      continue;
    }
    if (
      replyStatus &&
      String(replyStatus).toLowerCase() === REPLY_STATUS.REPLIED.toLowerCase()
    ) {
      skipped++;
      continue;
    }
    if (!threadId || !clientEmail) {
      skipped++;
      continue;
    }

    scanned++;
    var match = findReplyMatchForRow(threadId, clientEmail, keywords, sentAt);
    if (!match) continue;

    setCellValueIfColumn(
      sheet,
      rowIndex,
      colMap.REPLY_STATUS,
      REPLY_STATUS.REPLIED,
    );
    setCellValueIfColumn(
      sheet,
      rowIndex,
      colMap.REPLIED_AT,
      match.date || new Date(),
    );
    setCellValueIfColumn(
      sheet,
      rowIndex,
      colMap.REPLY_KEYWORD,
      match.keyword || "",
    );

    appendLog(
      logsSheet,
      clientName,
      "INFO",
      'Reply detected. Keyword: "' +
        (match.keyword || "") +
        '" | From: ' +
        (match.from || clientEmail),
    );
    updated++;
  }

  var summary =
    "Reply scan complete. Sent rows scanned: " +
    scanned +
    " | Updated: " +
    updated +
    " | Skipped: " +
    skipped;
  appendLog(logsSheet, "", "INFO", summary);
  return summary;
}

/**
 * Finds a matching reply in a sent thread based on sender + configured keywords.
 */
function findReplyMatchForRow(threadId, clientEmail, keywords, sentAt) {
  try {
    var thread = GmailApp.getThreadById(threadId);
    if (!thread) return null;
    var messages = thread.getMessages();
    var senderEmail = getSenderAccountEmail().toLowerCase();
    var clientLower = String(clientEmail || "").toLowerCase();

    for (var i = messages.length - 1; i >= 0; i--) {
      var msg = messages[i];
      if (sentAt instanceof Date && !isNaN(sentAt.getTime())) {
        if (msg.getDate().getTime() <= sentAt.getTime()) continue;
      }

      var from = String(msg.getFrom() || "").toLowerCase();
      if (senderEmail && from.indexOf(senderEmail) >= 0) continue;
      if (clientLower && from.indexOf(clientLower) < 0) continue;

      var haystack = (
        String(msg.getSubject() || "") +
        "\n" +
        String(msg.getPlainBody() || "")
      ).toUpperCase();
      var keyword = findMatchingKeyword(haystack, keywords);
      if (!keyword) continue;

      return {
        keyword: keyword,
        date: msg.getDate(),
        from: msg.getFrom(),
      };
    }
  } catch (e) {}

  return null;
}

/**
 * Returns the first matched keyword found in text.
 */
function findMatchingKeyword(text, keywords) {
  var upper = String(text || "").toUpperCase();
  for (var i = 0; i < keywords.length; i++) {
    var keyword = String(keywords[i] || "")
      .trim()
      .toUpperCase();
    if (!keyword) continue;
    var escaped = keyword.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    var regex = new RegExp("(^|\\b)" + escaped + "(\\b|$)", "i");
    if (regex.test(upper)) return keyword;
  }
  return "";
}

/**
 * Menu action: set reply keywords.
 */
function setReplyKeywords() {
  var ui = SpreadsheetApp.getUi();
  var current = getReplyKeywords().join(", ");
  var response = ui.prompt(
    "Set Reply Keywords",
    "Enter comma-separated keywords that mark a reply as acknowledged.\n\nCurrent: " +
      current,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var list = setReplyKeywordsFromText(response.getResponseText());
  ui.alert("Reply keywords saved: " + list.join(", "));
}

/**
 * Shows reply tracking status/config summary.
 */
function showReplyTrackingStatus() {
  var ui = SpreadsheetApp.getUi();
  var active = getTriggersByHandler("runReplyScan").length;
  var msg = [
    "Reply tracking schedule: " + (active > 0 ? "ACTIVE" : "INACTIVE"),
    "Trigger count: " + active,
    "Keywords: " + getReplyKeywords().join(", "),
  ].join("\n");
  ui.alert("Reply Tracking Status", msg, ui.ButtonSet.OK);
}

/**
 * Toggle AI generation on/off.
 */
function toggleAiGeneration() {
  var enabled = !isAiGenerationEnabled();
  setPropBoolean(PROP_KEYS.AI_ENABLED, enabled);
  SpreadsheetApp.getUi().alert(
    "AI Integration",
    "AI generation is now " + (enabled ? "ENABLED" : "DISABLED") + ".",
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

/**
 * Stores Gemini API key in properties.
 */
function setGeminiApiKey() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    "Set Gemini API Key",
    "Enter Gemini API key. Leave blank to clear.",
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var key = response.getResponseText().trim();
  setPropString(PROP_KEYS.AI_API_KEY, key);
  setPropString(PROP_KEYS.AI_PROVIDER, AI_PROVIDER.GEMINI);
  ui.alert(
    "AI Integration",
    key ? "Gemini key saved." : "Gemini key cleared.",
    ui.ButtonSet.OK,
  );
}

/**
 * Selects Gemini model.
 */
function selectGeminiModel() {
  var ui = SpreadsheetApp.getUi();
  var current = getAiModel();
  var apiKey = getAiApiKey();
  var models = apiKey ? listGeminiModels(apiKey) : [];

  var promptText =
    "Enter Gemini model name (e.g., models/gemini-1.5-flash).\n\nCurrent: " +
    current;
  if (models.length > 0) {
    promptText += "\n\nAvailable:\n- " + models.slice(0, 10).join("\n- ");
  }

  var response = ui.prompt(
    "Select Gemini Model",
    promptText,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var model = response.getResponseText().trim();
  if (!model) {
    ui.alert("Model cannot be empty.");
    return;
  }
  setPropString(PROP_KEYS.AI_MODEL, model);
  ui.alert("AI model saved: " + model);
}

/**
 * Calls Gemini models endpoint and returns model names.
 */
function listGeminiModels(apiKey) {
  try {
    var response = UrlFetchApp.fetch(
      "https://generativelanguage.googleapis.com/v1beta/models?key=" +
        encodeURIComponent(apiKey),
      { muteHttpExceptions: true },
    );
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
      return [];
    }
    var payload = JSON.parse(response.getContentText() || "{}");
    var models = payload.models || [];
    return models
      .map(function (item) {
        return String(item.name || "").trim();
      })
      .filter(function (name) {
        return name.indexOf("models/gemini") === 0;
      });
  } catch (e) {
    return [];
  }
}

/**
 * Validates AI setup by attempting a generation.
 */
function testAiConnection() {
  var ui = SpreadsheetApp.getUi();
  var result = tryGenerateAiEmailBody("Test Client", new Date(), "Visa/Permit");
  if (result && result.text) {
    ui.alert(
      "AI Integration",
      "AI connection successful. Model: " + result.model,
      ui.ButtonSet.OK,
    );
  } else {
    ui.alert(
      "AI Integration",
      "AI test failed. Check API key/model and try again.",
      ui.ButtonSet.OK,
    );
  }
}

/**
 * Displays AI configuration status.
 */
function showAiStatus() {
  var ui = SpreadsheetApp.getUi();
  var msg = [
    "AI Enabled: " + (isAiGenerationEnabled() ? "YES" : "NO"),
    "Provider: " + getPropString(PROP_KEYS.AI_PROVIDER, AI_PROVIDER.GEMINI),
    "Model: " + getAiModel(),
    "API Key: " + maskSecret(getAiApiKey()),
    "Fallback Source: " + getFallbackTemplateMode(),
    "Open Tracking URL: " + (getOpenTrackingBaseUrl() || "(not set)"),
  ].join("\n");
  ui.alert("AI Status", msg, ui.ButtonSet.OK);
}

/**
 * Stores property fallback template text.
 */
function setFallbackTemplate() {
  var ui = SpreadsheetApp.getUi();
  var current = getConfiguredFallbackTemplate();
  var response = ui.prompt(
    "Set Fallback Template",
    "Enter fallback body template text.\n\nCurrent value (first 300 chars):\n" +
      (current ? current.substring(0, 300) : "(empty)"),
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  setPropString(PROP_KEYS.FALLBACK_TEMPLATE, response.getResponseText());
  ui.alert("Fallback template saved.");
}

/**
 * Toggles fallback source between hardcoded and property mode.
 */
function toggleFallbackTemplateSource() {
  var current = getFallbackTemplateMode();
  var next =
    current === FALLBACK_TEMPLATE_MODE.HARDCODED
      ? FALLBACK_TEMPLATE_MODE.PROPERTY
      : FALLBACK_TEMPLATE_MODE.HARDCODED;
  setPropString(PROP_KEYS.FALLBACK_TEMPLATE_MODE, next);
  SpreadsheetApp.getUi().alert(
    "Fallback Source",
    "Fallback template mode is now: " + next,
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

/**
 * Sets open tracking base URL.
 */
function setOpenTrackingBaseUrl() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    "Set Open Tracking URL",
    "Enter deployed web app URL for tracking endpoint (doGet). Leave blank to disable.",
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  setPropString(PROP_KEYS.OPEN_TRACKING_BASE_URL, response.getResponseText());
  ui.alert("Open tracking URL updated.");
}

/**
 * Preview resolved fallback body using sample values.
 */
function previewFallbackTemplateBody() {
  var ui = SpreadsheetApp.getUi();
  var sampleDate = new Date();
  sampleDate.setMonth(sampleDate.getMonth() + 1);

  var fallback = resolveFallbackTemplateText();
  var rendered = applyTemplatePlaceholders(
    fallback.text,
    "Sample Client",
    formatDate(sampleDate),
    "Visa/Permit",
  );

  ui.alert(
    "Fallback Body Preview",
    "Source: " +
      fallback.source +
      "\n\n" +
      rendered.substring(0, 1500) +
      (rendered.length > 1500 ? "..." : ""),
    ui.ButtonSet.OK,
  );
}

/**
 * Web app endpoint for open/click tracking.
 */
function doGet(e) {
  var params = (e && e.parameter) || {};
  var mode = String(params.mode || "open").toLowerCase();
  var token = String(params.t || "").trim();

  if (token) {
    recordOpenTrackingEvent(token, mode);
  }

  if (mode === "click") {
    var url = String(params.u || "").trim();
    return HtmlService.createHtmlOutput(
      url
        ? '<meta http-equiv="refresh" content="0;url=' +
            sanitizeHtmlAttribute(url) +
            '">'
        : "Tracking click recorded.",
    );
  }

  return ContentService.createTextOutput(
    '<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"></svg>',
  ).setMimeType(ContentService.MimeType.XML);
}

/**
 * Sanitizes HTML attribute values in lightweight redirect responses.
 */
function sanitizeHtmlAttribute(value) {
  return String(value || "").replace(/["<>]/g, "");
}

/**
 * Records open/click events into sheet open tracking columns.
 */
function recordOpenTrackingEvent(token, mode) {
  try {
    var ss = getAutomationSpreadsheet();
    var sheetConfig = resolveAutomationSheet(ss);
    var sheet = sheetConfig.sheet;
    if (!sheet) return;

    var colMap = buildColumnMap(sheet);
    if (!colMap.OPEN_TOKEN) return;

    var rowNum = findRowNumberByToken(sheet, colMap, token);
    if (!rowNum) return;

    var now = new Date();
    if (colMap.FIRST_OPENED_AT) {
      var firstCell = sheet.getRange(rowNum, colMap.FIRST_OPENED_AT);
      if (!firstCell.getValue()) firstCell.setValue(now);
    }
    setCellValueIfColumn(sheet, rowNum, colMap.LAST_OPENED_AT, now);

    if (colMap.OPEN_COUNT) {
      var countCell = sheet.getRange(rowNum, colMap.OPEN_COUNT);
      var current = Number(countCell.getValue() || 0);
      countCell.setValue(isNaN(current) ? 1 : current + 1);
    }

    var logsSheet = ensureLogsSheet(ss);
    var clientName = colMap.CLIENT_NAME
      ? getCellStr(
          sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0],
          colMap.CLIENT_NAME,
        )
      : "";
    appendLog(
      logsSheet,
      clientName,
      "INFO",
      "Tracking event recorded: " + (mode || "open") + " | token=" + token,
    );
  } catch (e) {}
}

/**
 * Finds row number by open-tracking token.
 */
function findRowNumberByToken(sheet, colMap, token) {
  if (!colMap.OPEN_TOKEN || !token) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return 0;

  var numDataRows = lastRow - CONFIG.DATA_START_ROW + 1;
  var values = sheet
    .getRange(CONFIG.DATA_START_ROW, colMap.OPEN_TOKEN, numDataRows, 1)
    .getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || "").trim() === token) {
      return CONFIG.DATA_START_ROW + i;
    }
  }
  return 0;
}

/**
 * Removes all existing time-based triggers for runDailyCheck().
 * Safe to run multiple times.
 */
function removeTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var count = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "runDailyCheck") {
      ScriptApp.deleteTrigger(triggers[i]);
      count++;
    }
  }
  var msg =
    count > 0
      ? "Daily schedule deactivated. " + count + " trigger(s) removed."
      : "No active daily schedule found. Nothing to remove.";
  Logger.log(msg);
  try {
    SpreadsheetApp.getUi().alert(
      "Schedule Deactivated",
      msg,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  } catch (e) {}
}

// =============================================================================
// MANUAL TESTING HELPERS
// =============================================================================

/**
 * TEST HELPER: Runs the daily check immediately.
 * WARNING: This will actually send emails for any row whose target date is due (today or earlier).
 */
function testRunNow() {
  Logger.log("=== testRunNow: calling runDailyCheck() ===");
  runDailyCheck();
  Logger.log("=== testRunNow: complete. Check LOGS sheet and your inbox. ===");
}

/**
 * TEST HELPER: Logs the computed target date for every eligible row (Active or blank Status) without sending anything.
 * Also shows Send Mode gating results.
 * Use this to verify date calculations before going live.
 */
function previewTargetDates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetConfig = resolveAutomationSheet(ss);
  var sheet = sheetConfig.sheet;
  if (!sheet) {
    Logger.log(
      'Configured sheet "' +
        sheetConfig.sheetName +
        '" not found. Use "Initialize / Select Working Sheet".',
    );
    return;
  }

  var colMap = buildColumnMap(sheet);
  var mapError = validateColumnMap(colMap);
  if (mapError) {
    Logger.log("Column map error: " + mapError);
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) {
    Logger.log("No data rows.");
    return;
  }

  var numDataRows = lastRow - CONFIG.DATA_START_ROW + 1;
  var numCols = sheet.getLastColumn();
  var data = sheet
    .getRange(CONFIG.DATA_START_ROW, 1, numDataRows, numCols)
    .getValues();
  var today = getMidnight(new Date());

  Logger.log("Today: " + formatDate(today));
  Logger.log("---");

  data.forEach(function (row, i) {
    var rowIndex = CONFIG.DATA_START_ROW + i;
    var clientName = getCellStr(row, colMap.CLIENT_NAME);
    var expiryRaw = colMap.EXPIRY_DATE ? row[colMap.EXPIRY_DATE - 1] : "";
    var noticeStr = getCellStr(row, colMap.NOTICE_DATE);
    var status = getCellStr(row, colMap.STATUS);
    var sendMode = getRowSendMode(row, colMap);

    if (!isProcessableStatus(status)) return;

    var statusLabel = isStatusBlank(status)
      ? "(blank -> treated as Active)"
      : status;

    var modeSkipReason = getSendModeSkipReason(sendMode);
    if (modeSkipReason) {
      Logger.log(
        "Row " +
          rowIndex +
          " | " +
          clientName +
          " | Status: " +
          statusLabel +
          " | Mode: " +
          sendMode +
          " | " +
          modeSkipReason,
      );
      return;
    }

    var expiryDate =
      expiryRaw instanceof Date ? expiryRaw : new Date(expiryRaw);
    if (isNaN(expiryDate.getTime())) {
      Logger.log(
        "Row " + rowIndex + " [" + clientName + "]: INVALID expiry date",
      );
      return;
    }

    var offset = parseNoticeOffset(noticeStr);
    if (offset === null) {
      Logger.log(
        "Row " +
          rowIndex +
          " [" +
          clientName +
          "]: UNKNOWN notice option: " +
          noticeStr,
      );
      return;
    }

    var targetDate = computeTargetDate(expiryDate, offset);
    var dueNow = isTargetDateDue(targetDate, today)
      ? " <<< SENDS NOW (DUE/OVERDUE)"
      : "";
    Logger.log(
      "Row " +
        rowIndex +
        " | " +
        clientName +
        " | Status: " +
        statusLabel +
        " | Mode: " +
        sendMode +
        " | Expiry: " +
        formatDate(expiryDate) +
        " | Notice: " +
        noticeStr +
        " | Target: " +
        formatDate(targetDate) +
        dueNow,
    );
  });
}

/**
 * DIAGNOSTIC: Prompts for a row number and shows all parsed field values
 * for that row in a dialog — no email is sent.
 */
function diagnosticInspectRow() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    "Inspect Row",
    "Enter the row number to inspect (data starts at row " +
      CONFIG.DATA_START_ROW +
      "):",
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var rowNum = parseInt(response.getResponseText().trim(), 10);
  if (isNaN(rowNum) || rowNum < CONFIG.DATA_START_ROW) {
    ui.alert(
      "Invalid row number. Data starts at row " + CONFIG.DATA_START_ROW + ".",
    );
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetConfig = resolveAutomationSheet(ss);
  var sheet = sheetConfig.sheet;
  if (!sheet) {
    ui.alert(
      'Configured sheet "' +
        sheetConfig.sheetName +
        '" not found. Use "Initialize / Select Working Sheet".',
    );
    return;
  }

  var colMap = buildColumnMap(sheet);
  var mapError = validateColumnMap(colMap);
  if (mapError) {
    ui.alert("Column map error: " + mapError);
    return;
  }

  if (rowNum > sheet.getLastRow()) {
    ui.alert(
      "Row " +
        rowNum +
        " does not exist. Last row is " +
        sheet.getLastRow() +
        ".",
    );
    return;
  }

  var numCols = sheet.getLastColumn();
  var row = sheet.getRange(rowNum, 1, 1, numCols).getValues()[0];

  var clientName = getCellStr(row, colMap.CLIENT_NAME);
  var clientEmail = getCellStr(row, colMap.CLIENT_EMAIL);
  var staffEmail = getCellStr(row, colMap.STAFF_EMAIL);
  var docType = getCellStr(row, colMap.DOC_TYPE);
  var expiryRaw = colMap.EXPIRY_DATE ? row[colMap.EXPIRY_DATE - 1] : "";
  var noticeStr = getCellStr(row, colMap.NOTICE_DATE);
  var remarks = getCellStr(row, colMap.REMARKS);
  var attachRaw = getCellStr(row, colMap.ATTACHMENTS);
  var status = getCellStr(row, colMap.STATUS);
  var sendMode = getRowSendMode(row, colMap);
  var modeSkipReason = getSendModeSkipReason(sendMode);
  var replyStatus = getCellStr(row, colMap.REPLY_STATUS);
  var repliedAt = colMap.REPLIED_AT ? row[colMap.REPLIED_AT - 1] : "";
  var sentThreadId = getCellStr(row, colMap.SENT_THREAD_ID);
  var sentMessageId = getCellStr(row, colMap.SENT_MESSAGE_ID);
  var openToken = getCellStr(row, colMap.OPEN_TOKEN);
  var openCount = colMap.OPEN_COUNT ? row[colMap.OPEN_COUNT - 1] : "";

  var expiryDate = expiryRaw instanceof Date ? expiryRaw : new Date(expiryRaw);
  var expiryStr = isNaN(expiryDate.getTime())
    ? "INVALID (" + expiryRaw + ")"
    : formatDate(expiryDate);

  var offset = parseNoticeOffset(noticeStr);
  var targetStr =
    offset === null
      ? 'Cannot parse notice: "' + noticeStr + '"'
      : formatDate(computeTargetDate(expiryDate, offset));

  var today = getMidnight(new Date());
  var sendEligibleNow =
    offset !== null && !isNaN(expiryDate.getTime())
      ? isTargetDateDue(computeTargetDate(expiryDate, offset), today)
        ? "YES"
        : "No"
      : "N/A";

  var msg = [
    "Row: " + rowNum,
    "Status: " + (status || "(empty)"),
    "Send Mode: " +
      sendMode +
      (modeSkipReason ? " (" + modeSkipReason + ")" : ""),
    "Reply Status: " + (replyStatus || "(empty)"),
    "Replied At: " + (repliedAt || "(empty)"),
    "",
    "Client Name:  " + (clientName || "(empty)"),
    "Client Email: " + (clientEmail || "(empty)"),
    "Staff Email:  " + (staffEmail || "(empty)"),
    "Doc Type:     " + (docType || "(empty)"),
    "",
    "Expiry Date:  " + expiryStr,
    "Notice Date:  " + (noticeStr || "(empty)"),
    "Target Date:  " + targetStr,
    "Send Eligible Now (Due/Overdue):  " + sendEligibleNow,
    "",
    "Sent Thread Id: " + (sentThreadId || "(empty)"),
    "Sent Message Id: " + (sentMessageId || "(empty)"),
    "Open Token: " + (openToken || "(empty)"),
    "Open Count: " + (openCount || "(empty)"),
    "",
    "Attached Files: " + (attachRaw || "(none)"),
    "",
    "Remarks (first 200 chars):",
    remarks
      ? remarks.substring(0, 200) + (remarks.length > 200 ? "..." : "")
      : "(empty)",
  ].join("\n");

  ui.alert("Row " + rowNum + " Inspection", msg, ui.ButtonSet.OK);
}

/**
 * DIAGNOSTIC: Prompts for a No. value, finds the matching row,
 * shows a summary of what will be sent,
 * and asks for confirmation before actually sending the test email.
 * Ignores Status and target date — sends regardless.
 */
function diagnosticSendTestRow() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    "Send Test Email by No.",
    'Enter the value from column "No." to test (e.g., 15):\n\n' +
      "Note: Email will be sent regardless of Status or target date.",
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var noValue = response.getResponseText().trim();
  if (!noValue) {
    ui.alert("Invalid No. value. Please enter a value from column No.");
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetConfig = resolveAutomationSheet(ss);
  var sheet = sheetConfig.sheet;
  if (!sheet) {
    ui.alert(
      'Configured sheet "' +
        sheetConfig.sheetName +
        '" not found. Use "Initialize / Select Working Sheet".',
    );
    return;
  }

  var colMap = buildColumnMap(sheet);
  var mapError = validateColumnMap(colMap);
  if (mapError) {
    ui.alert("Column map error: " + mapError);
    return;
  }

  var lookup = findRowNumberByNo(sheet, colMap, noValue);
  if (lookup.error) {
    ui.alert(lookup.error);
    return;
  }

  var rowNum = lookup.rowNum;
  if (lookup.warning) {
    ui.alert("No. Lookup Notice", lookup.warning, ui.ButtonSet.OK);
  }

  var numCols = sheet.getLastColumn();
  var row = sheet.getRange(rowNum, 1, 1, numCols).getValues()[0];
  var rowNo = getCellStr(row, colMap.NO) || noValue;

  var clientName = getCellStr(row, colMap.CLIENT_NAME);
  var clientEmail = getCellStr(row, colMap.CLIENT_EMAIL);
  var staffEmail = getCellStr(row, colMap.STAFF_EMAIL);
  var docType = getCellStr(row, colMap.DOC_TYPE);
  var expiryRaw = colMap.EXPIRY_DATE ? row[colMap.EXPIRY_DATE - 1] : "";
  var remarks = getCellStr(row, colMap.REMARKS);
  var attachRaw = getCellStr(row, colMap.ATTACHMENTS);

  var missing = [];
  if (!clientName) missing.push("Client Name");
  if (!clientEmail) missing.push("Client Email");
  if (!expiryRaw) missing.push("Expiry Date");
  if (missing.length > 0) {
    ui.alert("Cannot send — missing required field(s): " + missing.join(", "));
    return;
  }

  var expiryDate = expiryRaw instanceof Date ? expiryRaw : new Date(expiryRaw);
  if (isNaN(expiryDate.getTime())) {
    ui.alert("Cannot send — invalid Expiry Date: " + expiryRaw);
    return;
  }

  var subject = buildEmailSubject(docType, clientName, expiryDate);

  var confirm = ui.alert(
    "Confirm Test Email",
    "This will send a REAL email for No. " +
      rowNo +
      " (row " +
      rowNum +
      "):\n\n" +
      "To:      " +
      clientEmail +
      "\n" +
      (staffEmail ? "CC:      " + staffEmail + "\n" : "") +
      "Subject: " +
      subject +
      "\n\n" +
      "Proceed?",
    ui.ButtonSet.YES_NO,
  );
  if (confirm !== ui.Button.YES) return;

  var attachResult = resolveAttachments(attachRaw);
  if (attachResult.error) {
    ui.alert("Attachment error: " + attachResult.error);
    return;
  }

  var openToken = getOpenTrackingBaseUrl() ? generateOpenTrackingToken() : "";
  var emailContent = buildEmailContent(
    remarks,
    clientName,
    expiryDate,
    docType,
    openToken,
  );

  try {
    var sentMeta = sendReminderEmail(
      clientEmail,
      staffEmail,
      subject,
      emailContent.htmlBody,
      attachResult.blobs,
    );
    var senderEmail = getSenderAccountEmail();
    setStaffEmail(sheet, rowNum, colMap.STAFF_EMAIL, senderEmail);
    writePostSendMetadata(sheet, rowNum, colMap, {
      sentAt: new Date(),
      senderEmail: senderEmail,
      openToken: openToken,
      threadId: sentMeta.threadId,
      messageId: sentMeta.messageId,
    });
    appendLog(
      ensureLogsSheet(ss),
      clientName,
      "INFO",
      "Test email sent by No. " +
        rowNo +
        " | To: " +
        clientEmail +
        (staffEmail ? " | CC: " + staffEmail : "") +
        " | Body: " +
        emailContent.source,
    );
    ui.alert(
      "Test email sent successfully to " +
        clientEmail +
        "." +
        "\n\nBody source: " +
        emailContent.source +
        (senderEmail
          ? "\n\nStaff Email updated with sender account: " + senderEmail
          : ""),
    );
  } catch (e) {
    ui.alert("Failed to send: " + e.message);
  }
}
