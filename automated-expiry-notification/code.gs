// System Version: March 9, 2026 8:08 am

var CONFIG = {
  SHEET_NAME: "VISA automation",
  LOGS_SHEET_NAME: "LOGS",
  AUTOMATION_SHEET_PROPERTY_KEY: "AUTOMATION_SHEET_NAMES",
  HEADER_ROW: 2,
  DATA_START_ROW: 3,
  TRIGGER_HOUR: 8,
  REPLY_SCAN_TRIGGER_HOUR: 9,
  SENDER_NAME: "Office",
  STATIC_REDIRECT_URL: "https://pastebin.com/8n85J6k6",
};

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
  FINAL_NOTICE_SENT_AT: "Final Notice Sent At",
  FINAL_NOTICE_THREAD_ID: "Final Notice Thread Id",
  FINAL_NOTICE_MESSAGE_ID: "Final Notice Message Id",
};

var HEADER_ALIASES = {
  CLIENT_NAME: ["Seller", "Seller Name", "Buyer", "Buyer Name"],
  CLIENT_EMAIL: [
    "Seller Email",
    "Buyer Email",
    "Seller E-mail",
    "Buyer E-mail",
    "Seller Mail",
    "Buyer Mail",
  ],
  DOC_TYPE: ["Services", "Service"],
  EXPIRY_DATE: ["Due Date"],
  NOTICE_DATE: ["Remaining Days", "Reminder Days"],
  REMARKS: [
    "Reminder (Email Content)",
    "Reminder Email Content",
    "Reminder Content",
  ],
  ATTACHMENTS: [
    "Attached File",
    "Gsheet",
    "GSheet",
    "Google Sheet",
    "Google Sheets",
  ],
  STATUS: ["Project Status"],
  STAFF_EMAIL: ["Assigned Staff Email"],
  SENT_THREAD_ID: ["Sent Thread ID"],
  SENT_MESSAGE_ID: ["Sent Message ID"],
  SEND_MODE: ["Send Option", "Mode"],
  OPEN_TOKEN: ["Open Token"],
  FIRST_OPENED_AT: ["First Open At"],
  LAST_OPENED_AT: ["Last Open At"],
  FINAL_NOTICE_SENT_AT: ["Final Notice Date", "Final Sent At"],
  FINAL_NOTICE_THREAD_ID: ["Final Notice Thread ID", "Final Thread ID"],
  FINAL_NOTICE_MESSAGE_ID: ["Final Notice Message ID", "Final Message ID"],
};

var STATUS = {
  ACTIVE: "Active",
  SENT: "Sent",
  ERROR: "Error",
  SKIPPED: "Skipped",
};

var SEND_MODE = {
  AUTO: "Auto",
  HOLD: "Hold",
  MANUAL_ONLY: "Manual Only",
};

var REPLY_STATUS = {
  PENDING: "Pending",
  REPLIED: "Replied",
};

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
  DEFAULT_CC_EMAILS: "DEFAULT_CC_EMAILS",
  DAILY_TRIGGER_HOUR: "DAILY_TRIGGER_HOUR",
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

var LOG_COL = {
  TIMESTAMP: 1,
  TAB: 2,
  CLIENT_NAME: 3,
  ACTION: 4,
  DETAIL: 5,
};

var TAB_CONFIG = {
  PREFIX: "TAB_CONFIG_",
  DEFAULT_HEADER_ROW: 2,
};

var TAB_CONFIG_KEYS = {
  COLUMN_MAP: "COLUMN_MAP",
  HEADER_ROW: "HEADER_ROW",
  DATA_START_ROW: "DATA_START_ROW",
  NOTICE_OPTIONS: "NOTICE_OPTIONS",
  STATUS_OPTIONS: "STATUS_OPTIONS",
  SEND_MODE_OPTIONS: "SEND_MODE_OPTIONS",
  LAST_SELECTED: "LAST_SELECTED",
};

var FLEXIBLE_HEADER_ALIASES = {
  NO: [
    "No.",
    "No",
    "Number",
    "#",
    "ID",
    "Ref",
    "Reference",
    "Ref No",
    "Ref. No.",
    "Reference No",
  ],
  CLIENT_NAME: [
    "Client Name",
    "Name",
    "Full Name",
    "Client",
    "Applicant Name",
    "Applicant",
    "Person",
    "Contact Name",
    "Seller",
    "Seller Name",
    "Buyer",
    "Buyer Name",
  ],
  CLIENT_EMAIL: [
    "Client Email",
    "Email",
    "Email Address",
    "E-mail",
    "E-mail Address",
    "Contact Email",
    "Mail",
    "Seller Email",
    "Buyer Email",
    "Seller E-mail",
    "Buyer E-mail",
    "Seller Mail",
    "Buyer Mail",
  ],
  DOC_TYPE: [
    "Type of ID/Document",
    "Document Type",
    "Doc Type",
    "Type",
    "ID Type",
    "Visa Type",
    "Permit Type",
    "Document",
    "Services",
    "Service",
  ],
  EXPIRY_DATE: [
    "Expiry Date",
    "Expiration Date",
    "Expires On",
    "Valid Until",
    "End Date",
    "Date of Expiration",
    "Expiry",
    "Due Date",
  ],
  NOTICE_DATE: [
    "Notice Date",
    "Notice",
    "Reminder Date",
    "Send On",
    "Notify On",
    "Advance Notice",
    "Remaining Days",
    "Reminder Days",
  ],
  REMARKS: [
    "Remarks",
    "Notes",
    "Comments",
    "Message",
    "Body",
    "Email Body",
    "Note",
    "Reminder (Email Content)",
    "Reminder Email Content",
    "Reminder Content",
  ],
  ATTACHMENTS: [
    "Attached Files",
    "Attachments",
    "Files",
    "Docs",
    "Documents",
    "Attached",
    "Drive Links",
    "Attached File",
    "Gsheet",
    "GSheet",
    "Google Sheet",
    "Google Sheets",
  ],
  STATUS: [
    "Status",
    "State",
    "Send Status",
    "Processing Status",
    "Project Status",
  ],
  STAFF_EMAIL: [
    "Staff Email",
    "Assigned Staff Email",
    "Staff",
    "Handler Email",
    "Assigned To",
    "Owner Email",
  ],
  SEND_MODE: [
    "Send Mode",
    "Send Option",
    "Mode",
    "Send",
    "Auto Send",
    "Processing Mode",
  ],
  SENT_AT: ["Sent At", "Date Sent", "Sent On", "Processed At", "Last Sent"],
  SENT_THREAD_ID: [
    "Sent Thread Id",
    "Sent Thread ID",
    "Thread ID",
    "Gmail Thread",
    "Thread",
  ],
  SENT_MESSAGE_ID: [
    "Sent Message Id",
    "Sent Message ID",
    "Message ID",
    "Gmail Message",
  ],
  REPLY_STATUS: ["Reply Status", "Response Status", "Acknowledged", "Replied"],
  REPLIED_AT: ["Replied At", "Reply Date", "Response Date", "Replied On"],
  REPLY_KEYWORD: ["Reply Keyword", "Keyword", "Ack Keyword"],
  OPEN_TOKEN: ["Open Tracking Token", "Open Token", "Tracking Token", "Token"],
  FIRST_OPENED_AT: ["First Opened At", "First Open At", "First Viewed"],
  LAST_OPENED_AT: ["Last Opened At", "Last Open At", "Last Viewed"],
  OPEN_COUNT: ["Open Count", "View Count", "Times Opened"],
  FINAL_NOTICE_SENT_AT: [
    "Final Notice Sent At",
    "Final Notice Date",
    "Final Sent At",
  ],
  FINAL_NOTICE_THREAD_ID: [
    "Final Notice Thread Id",
    "Final Notice Thread ID",
    "Final Thread ID",
  ],
  FINAL_NOTICE_MESSAGE_ID: [
    "Final Notice Message Id",
    "Final Notice Message ID",
    "Final Message ID",
  ],
};

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  rememberSpreadsheetId(SpreadsheetApp.getActiveSpreadsheet());
  ui.createMenu("🔔 Expiry Notifications")
    .addItem("🚀 Setup This Sheet for Automation", "runSetupWizard")
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("Status & Overview")
        .addItem("Show System Status", "showSystemStatus")
        .addItem("View Last Run Summary", "checkScheduleStatus")
        .addItem("Show All Tabs Info", "showAllTabsInfo"),
    )
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("Tab Management")
        .addItem("Configure Automation Tabs", "configureAutomationSheets")
        .addItem("Select Working Tab", "selectWorkingTab")
        .addItem("Setup Tab Dropdowns", "setupSheetDropdowns")
        .addItem("Map Tab Columns", "mapTabColumns")
        .addItem("Set Tab Header Row", "setTabHeaderRowDialog"),
    )
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("Automation Settings")
        .addItem("Activate Daily Schedule", "installTrigger")
        .addItem("Set Send Time", "setDailyTriggerHourDialog")
        .addItem("Deactivate Daily Schedule", "removeTrigger")
        .addItem("Set Default CC Emails", "setDefaultCcEmailsDialog")
        .addSeparator()
        .addItem("Set Reply Keywords", "setReplyKeywords")
        .addItem("Activate Reply Scan (2x Daily)", "installReplyScanTrigger")
        .addItem("Deactivate Reply Scan", "removeReplyScanTrigger")
        .addSeparator()
        .addItem("Set Open Tracking URL", "setOpenTrackingBaseUrl"),
    )
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("Run & Diagnostics")
        .addItem("Run Manual Check Now", "manualRunNow")
        .addItem("Preview Target Dates", "previewTargetDates")
        .addSeparator()
        .addItem("Inspect Row...", "diagnosticInspectRow")
        .addItem("Send Test Email by No....", "diagnosticSendTestRow")
        .addSeparator()
        .addItem("Test Gmail Send", "testGmailSend")
        .addItem("Test Drive Access", "testDriveAccess")
        .addItem("Test All Connections", "testAllConnections")
        .addSeparator()
        .addItem("Check Column Mappings", "checkColumnMappings")
        .addItem("Validate Tab Structure", "validateActiveTabStructure")
        .addItem("View Tab Configuration", "viewTabConfiguration")
        .addItem("Check Reply Tracking Setup", "checkReplyTrackingSetup")
        .addItem("Preview Fallback Template", "previewFallbackTemplateBody")
        .addSeparator()
        .addItem("System Diagnostics", "runSystemDiagnostics"),
    )
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("Help")
        .addItem(
          "Want to integrate this to another google sheet?",
          "showIntegrationLinkDialog",
        )
        .addSeparator()
        .addItem("View Documentation", "viewDocumentation")
        .addItem("About", "showAbout"),
    )
    .addToUi();
}

// ─── SETUP WIZARD ────────────────────────────────────────────────────────────

/**
 * Entry point for the guided setup wizard.
 * Walks through: tab selection/creation → column check → dropdowns → schedule.
 */
function runSetupWizard() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  ui.alert(
    "🚀 Setup This Sheet for Automation",
    "Welcome! This wizard will guide you through setting up a sheet tab for the Expiry Notification automation.\n\nYou will:\n  Step 1 — Select or create a sheet tab\n  Step 2 — Verify column headers\n  Step 3 — Apply dropdowns\n  Step 4 — Activate the daily schedule\n\nClick OK to begin.",
    ui.ButtonSet.OK,
  );

  var context = {
    tabName: "",
    sheet: null,
    columnsOk: false,
    dropsApplied: false,
    scheduleActive: false,
  };

  context = wizardStep1Tab(ss, ui, context);
  if (!context) return;

  context = wizardStep2Columns(ss, ui, context);
  if (!context) return;

  context = wizardStep3Dropdowns(ss, ui, context);
  if (!context) return;

  context = wizardStep4Schedule(ui, context);
  if (!context) return;

  wizardStep5Summary(ui, context);
}

/**
 * Wizard Step 1 — select an existing tab or create a new one.
 */
function wizardStep1Tab(ss, ui, context) {
  var choice = ui.prompt(
    "Step 1 of 4 — Sheet Tab",
    "Which tab should be used for automation?\n\n  1. Use an existing tab\n  2. Create a new tab\n\nEnter 1 or 2:",
    ui.ButtonSet.OK_CANCEL,
  );
  if (choice.getSelectedButton() !== ui.Button.OK) return null;

  var input = choice.getResponseText().trim();

  if (input === "2") {
    // ── Create new tab ──
    var nameResp = ui.prompt(
      "Step 1 of 4 — Create New Tab",
      "Enter a name for the new sheet tab:",
      ui.ButtonSet.OK_CANCEL,
    );
    if (nameResp.getSelectedButton() !== ui.Button.OK) return null;

    var newName = nameResp.getResponseText().trim();
    if (!newName) {
      ui.alert("No name entered. Setup cancelled.");
      return null;
    }

    if (ss.getSheetByName(newName)) {
      ui.alert(
        'A tab named "' +
          newName +
          "\" already exists. Setup cancelled.\n\nRe-run the wizard and choose 'Use an existing tab' to configure it.",
      );
      return null;
    }

    var newSheet = ss.insertSheet(newName);

    // Write required headers into row 1
    var requiredHeaders = [
      HEADERS.NO,
      HEADERS.CLIENT_NAME,
      HEADERS.CLIENT_EMAIL,
      HEADERS.DOC_TYPE,
      HEADERS.EXPIRY_DATE,
      HEADERS.NOTICE_DATE,
      HEADERS.REMARKS,
      HEADERS.ATTACHMENTS,
      HEADERS.STATUS,
      HEADERS.STAFF_EMAIL,
      HEADERS.SEND_MODE,
    ];
    newSheet
      .getRange(1, 1, 1, requiredHeaders.length)
      .setValues([requiredHeaders]);

    // Bold + freeze header row
    newSheet.getRange(1, 1, 1, requiredHeaders.length).setFontWeight("bold");
    newSheet.setFrozenRows(1);

    // Register tab by ID+name
    var existingEntries = getConfiguredTabEntries();
    var existingIds = existingEntries.map(function (e) {
      return e.id;
    });
    if (existingIds.indexOf(newSheet.getSheetId()) < 0) {
      existingEntries.push({ id: newSheet.getSheetId(), name: newName });
      setConfiguredTabEntries(existingEntries);
    }

    // Set header row = 1, data start = 2
    setTabHeaderRow(newName, 1);
    setTabDataStartRow(newName, 2);

    context.tabName = newName;
    context.sheet = newSheet;

    ui.alert(
      "Step 1 Complete ✓",
      'Tab "' + newName + '" created with headers.\n\nProceeding to Step 2.',
      ui.ButtonSet.OK,
    );
    return context;
  } else {
    // ── Use existing tab ── show ALL tabs
    var sheets = ss.getSheets();

    var currentEntries = getConfiguredTabEntries();
    var currentIds = currentEntries.map(function (e) {
      return e.id;
    });
    var currentNames = currentEntries.map(function (e) {
      return e.name;
    });

    var options = [];
    for (var i = 0; i < sheets.length; i++) {
      var isActive =
        currentIds.indexOf(sheets[i].getSheetId()) >= 0 ||
        currentNames.indexOf(sheets[i].getName()) >= 0;
      options.push(
        i +
          1 +
          ". " +
          sheets[i].getName() +
          (isActive ? " [already registered]" : ""),
      );
    }

    var pickResp = ui.prompt(
      "Step 1 of 4 — Select Existing Tab",
      "Available tabs:\n\n" +
        options.join("\n") +
        "\n\nEnter the number of the tab to use:",
      ui.ButtonSet.OK_CANCEL,
    );
    if (pickResp.getSelectedButton() !== ui.Button.OK) return null;

    var idx = parseInt(pickResp.getResponseText().trim(), 10);
    if (isNaN(idx) || idx < 1 || idx > sheets.length) {
      ui.alert("Invalid selection. Setup cancelled.");
      return null;
    }

    var selectedSheet = sheets[idx - 1];
    var selectedName = selectedSheet.getName();
    var selectedId = selectedSheet.getSheetId();

    // Register by ID+name if not already registered
    var alreadyRegistered = currentIds.indexOf(selectedId) >= 0;
    if (!alreadyRegistered) {
      currentEntries.push({ id: selectedId, name: selectedName });
      setConfiguredTabEntries(currentEntries);
    }

    context.tabName = selectedName;
    context.sheet = selectedSheet;

    ui.alert(
      "Step 1 Complete ✓",
      'Tab "' + selectedName + '" registered.\n\nProceeding to Step 2.',
      ui.ButtonSet.OK,
    );
    return context;
  }
}

/**
 * Wizard Step 2 — verify columns, offer manual mapping if needed.
 */
function wizardStep2Columns(ss, ui, context) {
  var flexMap = buildFlexibleColumnMap(context.sheet, context.tabName);
  var required = [
    "CLIENT_NAME",
    "CLIENT_EMAIL",
    "EXPIRY_DATE",
    "NOTICE_DATE",
    "STATUS",
  ];
  var missing = [];
  for (var i = 0; i < required.length; i++) {
    if (!flexMap.map[required[i]]) missing.push(HEADERS[required[i]]);
  }

  if (missing.length === 0) {
    context.columnsOk = true;
    ui.alert(
      "Step 2 Complete ✓",
      "All required columns detected automatically.\n\nProceeding to Step 3.",
      ui.ButtonSet.OK,
    );
    return context;
  }

  // Some columns missing
  var mapNow = ui.alert(
    "Step 2 of 4 — Column Check",
    "The following required columns were not detected:\n\n  • " +
      missing.join("\n  • ") +
      "\n\nWould you like to map columns manually now?",
    ui.ButtonSet.YES_NO,
  );

  if (mapNow === ui.Button.YES) {
    mapTabColumns();
    // Re-check after mapping
    var recheck = buildFlexibleColumnMap(context.sheet, context.tabName);
    var stillMissing = [];
    for (var j = 0; j < required.length; j++) {
      if (!recheck.map[required[j]]) stillMissing.push(HEADERS[required[j]]);
    }
    context.columnsOk = stillMissing.length === 0;
  } else {
    context.columnsOk = false;
    ui.alert(
      "Step 2 — Warning",
      "Continuing without all required columns. The automation may not work correctly until columns are mapped.\n\nYou can fix this later via Tab Management → Map Tab Columns.",
      ui.ButtonSet.OK,
    );
  }

  return context;
}

/**
 * Wizard Step 3 — apply dropdowns to Status, Send Mode, Notice Date.
 */
function wizardStep3Dropdowns(ss, ui, context) {
  var apply = ui.alert(
    "Step 3 of 4 — Dropdowns",
    'Apply dropdown options to the Status, Send Mode, and Notice Date columns in "' +
      context.tabName +
      '"?\n\nThis makes data entry easier. Default values will be used.',
    ui.ButtonSet.YES_NO,
  );

  if (apply !== ui.Button.YES) {
    ui.alert(
      "Step 3 Skipped",
      "Dropdowns skipped. You can apply them later via Tab Management → Setup Tab Dropdowns.",
      ui.ButtonSet.OK,
    );
    context.dropsApplied = false;
    return context;
  }

  var sheet = context.sheet;
  var tabName = context.tabName;
  var colMap = buildColumnMap(sheet, tabName);
  var dataStartRow = getTabDataStartRow(tabName);
  var lastRow = sheet.getLastRow();
  var dataLastRow = Math.max(lastRow, dataStartRow + 100);
  var applied = [];

  if (colMap.STATUS) {
    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(
        [STATUS.ACTIVE, STATUS.SENT, STATUS.ERROR, STATUS.SKIPPED],
        true,
      )
      .setAllowInvalid(true)
      .build();
    sheet
      .getRange(dataStartRow, colMap.STATUS, dataLastRow - dataStartRow + 1, 1)
      .setDataValidation(statusRule);
    applied.push("Status");
  }

  if (colMap.SEND_MODE) {
    var sendModeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(
        [SEND_MODE.AUTO, SEND_MODE.HOLD, SEND_MODE.MANUAL_ONLY],
        true,
      )
      .setAllowInvalid(true)
      .build();
    sheet
      .getRange(
        dataStartRow,
        colMap.SEND_MODE,
        dataLastRow - dataStartRow + 1,
        1,
      )
      .setDataValidation(sendModeRule);
    applied.push("Send Mode");
  }

  if (colMap.NOTICE_DATE) {
    var noticeOptions = getNoticeOptionsForTab(tabName);
    var noticeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(noticeOptions, true)
      .setAllowInvalid(true)
      .build();
    sheet
      .getRange(
        dataStartRow,
        colMap.NOTICE_DATE,
        dataLastRow - dataStartRow + 1,
        1,
      )
      .setDataValidation(noticeRule);
    applied.push("Notice Date");
  }

  context.dropsApplied = applied.length > 0;

  ui.alert(
    "Step 3 Complete ✓",
    applied.length > 0
      ? "Dropdowns applied to: " +
          applied.join(", ") +
          ".\n\nProceeding to Step 4."
      : "No matching columns found for dropdowns. You can retry later.\n\nProceeding to Step 4.",
    ui.ButtonSet.OK,
  );

  return context;
}

/**
 * Wizard Step 4 — offer to activate the daily 8 AM trigger.
 */
function wizardStep4Schedule(ui, context) {
  var triggerCount = getTriggersByHandler("runDailyCheck").length;
  var currentStatus =
    triggerCount > 0
      ? "ACTIVE (" + triggerCount + " trigger already set)"
      : "INACTIVE";

  var activate = ui.alert(
    "Step 4 of 4 — Daily Schedule",
    "Current daily schedule status: " +
      currentStatus +
      "\n\nActivate the daily 8 AM email schedule now?",
    ui.ButtonSet.YES_NO,
  );

  if (activate === ui.Button.YES) {
    if (triggerCount === 0) {
      installTrigger();
    }
    context.scheduleActive = true;
  } else {
    context.scheduleActive = triggerCount > 0;
    ui.alert(
      "Step 4 Skipped",
      "Schedule not changed. You can activate it later via Automation Settings → Activate Daily Schedule.",
      ui.ButtonSet.OK,
    );
  }

  return context;
}

/**
 * Wizard Step 5 — summary alert.
 */
function wizardStep5Summary(ui, context) {
  var lines = [
    "✅ Setup Complete!",
    "",
    "Tab:             " + context.tabName,
    "Columns:         " +
      (context.columnsOk
        ? "✓ All required columns OK"
        : "⚠ Some columns missing — map them via Tab Management"),
    "Dropdowns:       " + (context.dropsApplied ? "✓ Applied" : "— Skipped"),
    "Daily Schedule:  " +
      (context.scheduleActive ? "✓ Active" : "— Not activated"),
    "",
    "You're ready to go!",
    'Add your client data to the "' +
      context.tabName +
      '" tab and the automation will handle the rest.',
  ];

  ui.alert("🚀 Setup Complete", lines.join("\n"), ui.ButtonSet.OK);
}

// ─────────────────────────────────────────────────────────────────────────────

/**
 * Shows comprehensive system status across all configured tabs.
 */
function showSystemStatus() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logsSheet = ensureLogsSheet(ss);

  var sheetConfigs = resolveAutomationSheets(ss);
  var triggerCount = getTriggersByHandler("runDailyCheck").length;
  var replyTriggerCount = getTriggersByHandler("runReplyScan").length;

  var lines = [
    "=== System Status ===",
    "",
    "Daily Schedule: " +
      (triggerCount > 0 ? "ACTIVE (" + triggerCount + " trigger)" : "INACTIVE"),
    "Reply Scan: " +
      (replyTriggerCount > 0
        ? "ACTIVE (" + replyTriggerCount + " trigger)"
        : "INACTIVE"),
    "",
    "=== Configured Tabs (" + sheetConfigs.length + ") ===",
    "",
  ];

  for (var i = 0; i < sheetConfigs.length; i++) {
    var config = sheetConfigs[i];
    var sheet = config.sheet;
    var statusIcon = sheet ? "✓" : "✗";
    var colStatus = "";

    if (sheet) {
      var flexMap = buildFlexibleColumnMap(sheet, config.sheetName);
      var missing = [];
      var required = [
        "CLIENT_NAME",
        "CLIENT_EMAIL",
        "EXPIRY_DATE",
        "NOTICE_DATE",
        "STATUS",
      ];
      for (var r = 0; r < required.length; r++) {
        if (!flexMap.map[required[r]]) missing.push(required[r]);
      }

      if (missing.length === 0) {
        colStatus = " [columns OK]";
      } else {
        colStatus = " [⚠ missing: " + missing.join(", ") + "]";
      }
    } else {
      colStatus = " [NOT FOUND]";
    }

    lines.push(
      "  " + (i + 1) + ". " + statusIcon + " " + config.sheetName + colStatus,
    );
  }

  lines.push("");
  lines.push(getLatestRunSummary(logsSheet));

  ui.alert("System Status", lines.join("\n"), ui.ButtonSet.OK);
}

/**
 * Shows detailed information about all tabs including column mappings.
 */
function showAllTabsInfo() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetConfigs = resolveAutomationSheets(ss);

  var lines = ["=== All Tabs Information ===", ""];

  for (var i = 0; i < sheetConfigs.length; i++) {
    var config = sheetConfigs[i];
    var sheet = config.sheet;

    lines.push("Tab: " + config.sheetName);

    if (!sheet) {
      lines.push("  Status: ✗ NOT FOUND");
      lines.push("");
      continue;
    }

    var headerRow = getTabHeaderRow(config.sheetName);
    var dataStartRow = getTabDataStartRow(config.sheetName);
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var dataRows = lastRow >= dataStartRow ? lastRow - dataStartRow + 1 : 0;

    lines.push("  Status: ✓ Found");
    lines.push("  Header Row: " + headerRow);
    lines.push("  Data Start Row: " + dataStartRow);
    lines.push("  Total Rows: " + lastRow);
    lines.push("  Data Rows: " + dataRows);
    lines.push("  Columns: " + lastCol);

    var flexMap = buildFlexibleColumnMap(sheet, config.sheetName);
    lines.push("  Column Source: " + flexMap.source);

    if (flexMap.warnings.length > 0) {
      lines.push("  Warnings: " + flexMap.warnings.join("; "));
    }

    // Show mapped columns
    var mappedCols = [];
    for (var key in flexMap.map) {
      if (flexMap.map[key]) {
        mappedCols.push(key + "→" + flexMap.map[key]);
      }
    }
    if (mappedCols.length > 0) {
      lines.push(
        "  Mapped Columns: " +
          mappedCols.slice(0, 5).join(", ") +
          (mappedCols.length > 5 ? "..." : ""),
      );
    }

    lines.push("");
  }

  ui.alert(
    "All Tabs Info",
    lines.join("\n").substring(0, 1800),
    ui.ButtonSet.OK,
  );
}

/**
 * Allows user to select a working tab with visual indicators.
 */
function selectWorkingTab() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // resolveAutomationSheets auto-purges deleted tabs and returns only live ones
  var sheetConfigs = resolveAutomationSheets(ss);

  if (sheetConfigs.length === 0) {
    ui.alert(
      "No Active Tabs Found",
      "No configured tabs exist (or all were deleted/renamed).\n\nUse 'Configure Automation Tabs' to register tabs.",
      ui.ButtonSet.OK,
    );
    return;
  }

  var lastSelected = getPropString(
    getTabConfigKey("_GLOBAL", TAB_CONFIG_KEYS.LAST_SELECTED),
    "",
  );

  var options = [];
  for (var i = 0; i < sheetConfigs.length; i++) {
    var config = sheetConfigs[i];
    var sheet = config.sheet; // always non-null here (purged above)
    var icon = "✓";
    var status = "";

    var flexMap = buildFlexibleColumnMap(sheet, config.sheetName);
    var hasMissing = flexMap.warnings.some(function (w) {
      return w.indexOf("Missing required") >= 0;
    });
    if (hasMissing) {
      icon = "⚠";
      status = " (missing columns)";
    }

    var marker = config.sheetName === lastSelected ? " ★ [current]" : "";
    options.push(
      i + 1 + ". " + icon + " " + config.sheetName + status + marker,
    );
  }

  var response = ui.prompt(
    "Select Working Tab",
    "✓ = ready, ⚠ = needs column setup   ★ = currently selected\n\n" +
      options.join("\n") +
      "\n\nEnter number:",
    ui.ButtonSet.OK_CANCEL,
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  var idx = parseInt(response.getResponseText().trim(), 10);
  if (isNaN(idx) || idx < 1 || idx > sheetConfigs.length) {
    ui.alert("Invalid selection.");
    return;
  }

  var selected = sheetConfigs[idx - 1];
  setPropString(
    getTabConfigKey("_GLOBAL", TAB_CONFIG_KEYS.LAST_SELECTED),
    selected.sheetName,
  );

  ui.alert(
    "Working Tab Set",
    '"' +
      selected.sheetName +
      '" is now your working tab.\n\nThis tab will be used for subsequent operations.',
    ui.ButtonSet.OK,
  );
}

/**
 * Gets the last selected working tab config, or prompts if none.
 */
function getWorkingTabConfig(ss) {
  var lastSelected = getPropString(
    getTabConfigKey("_GLOBAL", TAB_CONFIG_KEYS.LAST_SELECTED),
    "",
  );
  var sheetConfigs = resolveAutomationSheets(ss);

  if (lastSelected && sheetConfigs.length > 0) {
    for (var i = 0; i < sheetConfigs.length; i++) {
      if (sheetConfigs[i].sheetName === lastSelected) {
        return sheetConfigs[i];
      }
    }
  }

  // If only one tab, use it
  if (sheetConfigs.length === 1) {
    return sheetConfigs[0];
  }

  // Otherwise prompt
  return promptSelectConfiguredSheet(ss, "Select Working Tab");
}

/**
 * Sends a real test email to a user-specified recipient to confirm Gmail sending works.
 */
function testGmailSend() {
  var ui = SpreadsheetApp.getUi();
  var senderEmail = getSenderAccountEmail();
  var senderName = getSenderDisplayName(senderEmail);

  var response = ui.prompt(
    "Test Gmail Send",
    "Sending from: " +
      (senderEmail || "(unknown)") +
      "\n\nEnter recipient email address:",
    ui.ButtonSet.OK_CANCEL,
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  var recipient = response.getResponseText().trim();
  if (!recipient || !isValidEmail(recipient)) {
    ui.alert(
      "Test Gmail Send",
      "Invalid or empty email address. Aborted.",
      ui.ButtonSet.OK,
    );
    return;
  }

  var now = new Date();
  var subject = "[TEST] Expiry Notification – Connection Test";
  var htmlBody =
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#333;max-width:600px;line-height:1.6;">' +
    "<p>This is a <strong>test email</strong> sent from the Expiry Notification automation.</p>" +
    "<p>If you received this, Gmail sending is working correctly.</p>" +
    '<hr style="border:none;border-top:1px solid #eee;margin:16px 0;">' +
    '<p style="font-size:12px;color:#888;">' +
    "Sent by: " +
    sanitizeHtmlContent(senderEmail || "(unknown)") +
    "<br>" +
    "Timestamp: " +
    now.toLocaleString() +
    "</p></div>";

  try {
    GmailApp.sendEmail(recipient, subject, "", {
      htmlBody: htmlBody,
      name: senderName || CONFIG.SENDER_NAME,
    });
    ui.alert(
      "Test Gmail Send",
      "✓ Test email sent successfully!\n\nFrom: " +
        (senderEmail || "(unknown)") +
        "\nTo: " +
        recipient,
      ui.ButtonSet.OK,
    );
  } catch (e) {
    ui.alert(
      "Test Gmail Send",
      "✗ Failed to send test email:\n" + e.message,
      ui.ButtonSet.OK,
    );
  }
}

/**
 * Tests Drive access by attempting to get root folder.
 */
function testDriveAccess() {
  var ui = SpreadsheetApp.getUi();
  try {
    var root = DriveApp.getRootFolder();
    var files = DriveApp.getFilesByType("application/pdf");
    var hasFiles = files.hasNext();
    ui.alert(
      "Drive Access Test",
      "✓ Drive access successful!\n\nRoot folder: " +
        root.getName() +
        "\nCan read files: " +
        (hasFiles ? "YES" : "No files found"),
      ui.ButtonSet.OK,
    );
  } catch (e) {
    ui.alert(
      "Drive Access Test",
      "✗ Drive access failed:\n" + e.message,
      ui.ButtonSet.OK,
    );
  }
}

/**
 * Runs all connection tests (Gmail service + Drive access).
 */
function testAllConnections() {
  var ui = SpreadsheetApp.getUi();
  var results = [];

  // Test Gmail service availability
  try {
    var senderEmail = getSenderAccountEmail();
    results.push(
      "✓ Gmail: accessible (" + (senderEmail || "unknown account") + ")",
    );
  } catch (e) {
    results.push("✗ Gmail: " + e.message);
  }

  // Test Drive
  try {
    var root = DriveApp.getRootFolder();
    results.push("✓ Drive: accessible (" + root.getName() + ")");
  } catch (e) {
    results.push("✗ Drive: " + e.message);
  }

  // Check configured tabs
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configs = resolveAutomationSheets(ss);
    var found = configs.filter(function (c) {
      return !!c.sheet;
    }).length;
    results.push(
      "✓ Automation tabs: " + found + "/" + configs.length + " found",
    );
  } catch (e) {
    results.push("✗ Tabs: " + e.message);
  }

  results.push("");
  results.push("Use 'Test Gmail Send' to confirm actual email delivery.");

  ui.alert("All Connection Tests", results.join("\n"), ui.ButtonSet.OK);
}

/**
 * Validates the active/selected tab structure.
 */
function validateActiveTabStructure() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = getWorkingTabConfig(ss);

  if (!config) return;

  if (!config.sheet) {
    ui.alert("Tab '" + config.sheetName + "' not found!");
    return;
  }

  var flexMap = buildFlexibleColumnMap(config.sheet, config.sheetName);
  var validation = validateFlexibleColumnMap(flexMap);

  var lines = [
    "Tab: " + config.sheetName,
    "Header Row: " + flexMap.headerRow,
    "Column Source: " + flexMap.source,
    "",
    validation
      ? "✗ Validation Failed:\n" + validation
      : "✓ All required columns found!",
    "",
    "=== Column Mappings ===",
  ];

  var required = [
    "CLIENT_NAME",
    "CLIENT_EMAIL",
    "EXPIRY_DATE",
    "NOTICE_DATE",
    "STATUS",
  ];
  var optional = [
    "NO",
    "DOC_TYPE",
    "REMARKS",
    "ATTACHMENTS",
    "STAFF_EMAIL",
    "SEND_MODE",
    "SENT_AT",
    "REPLY_STATUS",
  ];

  for (var i = 0; i < required.length; i++) {
    var key = required[i];
    var col = flexMap.map[key];
    lines.push(
      "  " +
        (col ? "✓" : "✗") +
        " " +
        key +
        ": " +
        (col ? "Column " + col : "NOT FOUND"),
    );
  }

  lines.push("");
  lines.push("Optional Columns:");
  for (var i = 0; i < optional.length; i++) {
    var key = optional[i];
    var col = flexMap.map[key];
    if (col) lines.push("  ✓ " + key + ": Column " + col);
  }

  ui.alert(
    "Tab Structure Validation",
    lines.join("\n").substring(0, 1800),
    ui.ButtonSet.OK,
  );
}

/**
 * Shows column mappings for the selected tab.
 */
function checkColumnMappings() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = getWorkingTabConfig(ss);

  if (!config) return;

  var stored = getTabColumnMapping(config.sheetName);
  var flexMap = buildFlexibleColumnMap(config.sheet, config.sheetName);

  var lines = [
    "Tab: " + config.sheetName,
    "Configured Header Row: " + getTabHeaderRow(config.sheetName),
    "Effective Header Row: " + flexMap.headerRow,
    "Mapping Source: " + flexMap.source,
    "Stored Mapping: " +
      (Object.keys(stored).length > 0 ? "Yes" : "No (using auto-detect)"),
    "",
    "=== Current Mappings ===",
    "",
  ];

  for (var key in flexMap.map) {
    lines.push(key + " → Column " + flexMap.map[key]);
  }

  if (flexMap.warnings.length > 0) {
    lines.push("");
    lines.push("=== Warnings ===");
    for (var i = 0; i < flexMap.warnings.length; i++) {
      lines.push("⚠ " + flexMap.warnings[i]);
    }
  }

  ui.alert(
    "Column Mappings",
    lines.join("\n").substring(0, 1800),
    ui.ButtonSet.OK,
  );
}

/**
 * Returns the list of logical columns shown in the interactive tab-mapping UI.
 */
function getMapTabColumnKeys() {
  return [
    "CLIENT_NAME",
    "CLIENT_EMAIL",
    "DOC_TYPE",
    "EXPIRY_DATE",
    "NOTICE_DATE",
    "STATUS",
    "REMARKS",
    "ATTACHMENTS",
    "NO",
    "STAFF_EMAIL",
    "SEND_MODE",
  ];
}

/**
 * Builds available header choices from a row of header values.
 */
function buildAvailableHeaderChoices(headerValues) {
  var choices = [];
  for (var i = 0; i < headerValues.length; i++) {
    var text = String(headerValues[i] || "").trim();
    if (!text) continue;
    choices.push({ index: i + 1, header: text });
  }
  return choices;
}

/**
 * Builds a map of { logicalKey: suggestedHeaderText } from detection suggestions.
 */
function buildSuggestedHeaderMap(suggestions) {
  var bestByKey = {};
  for (var i = 0; i < suggestions.length; i++) {
    var item = suggestions[i];
    var current = bestByKey[item.suggestedLogicalKey];
    if (!current || item.confidence > current.confidence) {
      bestByKey[item.suggestedLogicalKey] = item;
    }
  }

  var map = {};
  for (var key in bestByKey) {
    map[key] = bestByKey[key].actualHeader;
  }
  return map;
}

/**
 * Resolves user input to a valid header choice.
 * Supports column number or exact header text.
 */
function resolveHeaderSelectionInput(inputText, availableHeaders) {
  var input = String(inputText || "").trim();
  if (!input) return "";

  var asNumber = parseInt(input, 10);
  if (!isNaN(asNumber) && String(asNumber) === input) {
    for (var i = 0; i < availableHeaders.length; i++) {
      if (availableHeaders[i].index === asNumber) {
        return availableHeaders[i].header;
      }
    }
  }

  var normalizedInput = normalizeHeaderName(input);
  for (var i = 0; i < availableHeaders.length; i++) {
    if (normalizeHeaderName(availableHeaders[i].header) === normalizedInput) {
      return availableHeaders[i].header;
    }
  }

  return null;
}

/**
 * Formats a compact list of available header options for prompt dialogs.
 */
function formatHeaderOptionsForPrompt(availableHeaders, maxItems) {
  var limit = maxItems || 20;
  var lines = [];
  var count = Math.min(availableHeaders.length, limit);
  for (var i = 0; i < count; i++) {
    lines.push(availableHeaders[i].index + ". " + availableHeaders[i].header);
  }
  if (availableHeaders.length > limit) {
    lines.push("... +" + (availableHeaders.length - limit) + " more header(s)");
  }
  return lines.join("\n");
}

/**
 * Interactive column mapping UI for a tab.
 */
function mapTabColumns() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = promptSelectConfiguredSheet(ss, "Map Tab Columns");

  if (!config || !config.sheet) {
    ui.alert("Tab not found.");
    return;
  }

  var headerInfo = resolveEffectiveHeaderRow(config.sheet, config.sheetName);
  var headerValues = headerInfo.headers || [];
  var availableHeaders = buildAvailableHeaderChoices(headerValues);
  if (availableHeaders.length === 0) {
    ui.alert(
      "No usable headers found in tab '" +
        config.sheetName +
        "'. Set a valid header row first.",
    );
    return;
  }

  var suggestions = detectColumnMappings(config.sheet, config.sheetName);
  var suggestedMap = buildSuggestedHeaderMap(suggestions);
  var currentMapping = getTabColumnMapping(config.sheetName);
  var mapKeys = getMapTabColumnKeys();

  var lines = [
    "Tab: " + config.sheetName,
    "Configured Header Row: " + getTabHeaderRow(config.sheetName),
    "Effective Header Row: " + headerInfo.headerRow,
  ];
  if (headerInfo.usedFallback && headerInfo.reason) {
    lines.push("Note: " + headerInfo.reason);
  }

  lines.push("");
  lines.push("Available headers:");
  lines.push(formatHeaderOptionsForPrompt(availableHeaders, 20));
  lines.push("");
  lines.push("You will map " + mapKeys.length + " logical automation fields.");
  lines.push("Use number or exact header text.");
  lines.push(
    "Leave blank = keep current/suggested value. Enter 0 = clear field.",
  );

  ui.alert(
    "Map Tab Columns",
    lines.join("\n").substring(0, 1800),
    ui.ButtonSet.OK,
  );

  var newMapping = {};
  var optionsText = formatHeaderOptionsForPrompt(availableHeaders, 12);

  for (var k = 0; k < mapKeys.length; k++) {
    var logicalKey = mapKeys[k];
    var label = HEADERS[logicalKey] || logicalKey;
    var defaultHeader =
      currentMapping[logicalKey] || suggestedMap[logicalKey] || "";

    while (true) {
      var promptLines = [
        "Field " + (k + 1) + " of " + mapKeys.length,
        "Logical Key: " + logicalKey,
        "Expected Label: " + label,
        "Current/Suggested: " + (defaultHeader || "(none)"),
        "",
        "Available headers:",
        optionsText,
        "",
        "Input: number or exact header text",
        "Blank: keep current/suggested",
        "0: clear this mapping",
      ];

      var response = ui.prompt(
        "Map Column - " + logicalKey,
        promptLines.join("\n").substring(0, 1800),
        ui.ButtonSet.OK_CANCEL,
      );

      if (response.getSelectedButton() !== ui.Button.OK) {
        ui.alert("Column mapping cancelled. No changes were saved.");
        return;
      }

      var input = response.getResponseText().trim();
      if (!input) {
        if (defaultHeader) newMapping[logicalKey] = defaultHeader;
        break;
      }

      if (input === "0") {
        break;
      }

      var resolved = resolveHeaderSelectionInput(input, availableHeaders);
      if (resolved) {
        newMapping[logicalKey] = resolved;
        break;
      }

      ui.alert(
        "Invalid selection",
        "Could not resolve '" +
          input +
          "'. Enter a valid column number or exact header text.",
        ui.ButtonSet.OK,
      );
    }
  }

  saveTabColumnMapping(config.sheetName, newMapping);

  var savedFlexMap = buildFlexibleColumnMap(config.sheet, config.sheetName);
  var validationMessage = validateFlexibleColumnMap(savedFlexMap);

  var saveLines = [
    "Tab: " + config.sheetName,
    "Saved mappings: " + Object.keys(newMapping).length,
    "Effective header row: " + savedFlexMap.headerRow,
  ];

  if (validationMessage) {
    saveLines.push("");
    saveLines.push("Validation warning:");
    saveLines.push(validationMessage);
  } else {
    saveLines.push("");
    saveLines.push("All required columns are mapped.");
  }

  ui.alert(
    "Column Mapping Saved",
    saveLines.join("\n").substring(0, 1800),
    ui.ButtonSet.OK,
  );
}

/**
 * Dialog to set header row for a tab.
 */
function setTabHeaderRowDialog() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = promptSelectConfiguredSheet(ss, "Set Header Row");

  if (!config || !config.sheet) {
    ui.alert("Tab not found.");
    return;
  }

  var currentRow = getTabHeaderRow(config.sheetName);

  var response = ui.prompt(
    "Set Header Row for '" + config.sheetName + "'",
    "Current header row: " +
      currentRow +
      "\n\nEnter new header row number (default is 2):",
    ui.ButtonSet.OK_CANCEL,
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  var newRow = parseInt(response.getResponseText().trim(), 10);
  if (isNaN(newRow) || newRow < 1) {
    ui.alert("Invalid row number.");
    return;
  }

  var finalRow = newRow;
  var headerInfo = resolveEffectiveHeaderRow(
    config.sheet,
    config.sheetName,
    newRow,
  );
  if (headerInfo.usedFallback && headerInfo.headerRow !== newRow) {
    var useSuggested = ui.alert(
      "Header Row Suggestion",
      "Row " +
        newRow +
        " appears to be a divider/non-header row.\n\n" +
        "Suggested header row: " +
        headerInfo.headerRow +
        "\n\nUse the suggested row?",
      ui.ButtonSet.YES_NO,
    );
    if (useSuggested === ui.Button.YES) {
      finalRow = headerInfo.headerRow;
    }
  }

  setTabHeaderRow(config.sheetName, finalRow);
  setTabDataStartRow(config.sheetName, finalRow + 1);

  ui.alert(
    "Header row set to " +
      finalRow +
      " for '" +
      config.sheetName +
      "'.\nData start row set to " +
      (finalRow + 1) +
      ".",
  );
}

/**
 * View tab configuration details.
 */
function viewTabConfiguration() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = getWorkingTabConfig(ss);

  if (!config) return;

  var headerRow = getTabHeaderRow(config.sheetName);
  var dataStartRow = getTabDataStartRow(config.sheetName);
  var mapping = getTabColumnMapping(config.sheetName);
  var noticeOpts = getNoticeOptionsForTab(config.sheetName);

  var lines = [
    "Tab: " + config.sheetName,
    "",
    "Header Row: " + headerRow,
    "Data Start Row: " + dataStartRow,
    "",
    "Stored Column Mapping: " +
      (Object.keys(mapping).length > 0
        ? "Yes (" + Object.keys(mapping).length + " mappings)"
        : "No (auto-detect)"),
    "",
    "Notice Date Options:",
    noticeOpts.join(", "),
  ];

  ui.alert("Tab Configuration", lines.join("\n"), ui.ButtonSet.OK);
}

/**
 * Check reply tracking setup.
 */
function checkReplyTrackingSetup() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = getWorkingTabConfig(ss);

  if (!config || !config.sheet) {
    ui.alert("No tab selected.");
    return;
  }

  var flexMap = buildFlexibleColumnMap(config.sheet, config.sheetName);
  var hasThreadId = !!flexMap.map.SENT_THREAD_ID;
  var hasReplyStatus = !!flexMap.map.REPLY_STATUS;
  var keywords = getReplyKeywords();
  var triggerActive = getTriggersByHandler("runReplyScan").length > 0;

  var lines = [
    "Tab: " + config.sheetName,
    "",
    "Required Columns:",
    "  " + (hasThreadId ? "✓" : "✗") + " Sent Thread Id column",
    "  " + (hasReplyStatus ? "✓" : "✗") + " Reply Status column",
    "",
    "Reply Keywords: " + keywords.join(", "),
    "Reply Scan Trigger: " + (triggerActive ? "ACTIVE" : "INACTIVE"),
    "",
    hasThreadId && hasReplyStatus
      ? "✓ Reply tracking is ready!"
      : "⚠ Add missing columns to enable reply tracking.",
  ];

  ui.alert("Reply Tracking Setup", lines.join("\n"), ui.ButtonSet.OK);
}

/**
 * Validate date parsing across tabs.
 */
function formatParsedNoticeOffset(offset) {
  if (!offset) return "INVALID";
  if (offset.unit === "absolute_date") {
    return "absolute_date(" + formatDate(offset.value) + ")";
  }
  return offset.value + " " + offset.unit;
}

function validateDateParsing() {
  var ui = SpreadsheetApp.getUi();

  var testCases = [
    { input: "7 days before", valid: true, unit: "days", value: 7 },
    { input: "2 weeks before", valid: true, unit: "days", value: 14 },
    { input: "1 month before", valid: true, unit: "months", value: 1 },
    { input: "1 year before", valid: true, unit: "years", value: 1 },
    { input: "2 years before", valid: true, unit: "years", value: 2 },
    { input: "1 yr before", valid: true, unit: "years", value: 1 },
    { input: "7", valid: true, unit: "days", value: 7 },
    { input: "On expiry date", valid: true, unit: "days", value: 0 },
    { input: "2026-12-31", valid: true, unit: "absolute_date" },
    { input: "invalid", valid: false },
  ];

  var lines = ["Date Parsing Validation:", ""];
  var passedCount = 0;

  for (var i = 0; i < testCases.length; i++) {
    var tc = testCases[i];
    var result = parseNoticeOffset(tc.input);
    var passed = tc.valid ? !!result : result === null;
    if (passed && tc.valid && tc.unit && result.unit !== tc.unit) {
      passed = false;
    }
    if (
      passed &&
      tc.valid &&
      typeof tc.value === "number" &&
      result.value !== tc.value
    ) {
      passed = false;
    }

    if (passed) passedCount++;

    lines.push(
      (passed ? "✓" : "✗") +
        " '" +
        tc.input +
        "' → " +
        formatParsedNoticeOffset(result),
    );
  }

  lines.push("");
  lines.push("Passed: " + passedCount + "/" + testCases.length);

  ui.alert("Date Parsing", lines.join("\n"), ui.ButtonSet.OK);
}

/**
 * Test notice options parsing.
 */
function testNoticeOptions() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = getWorkingTabConfig(ss);

  if (!config) return;

  var opts = getNoticeOptionsForTab(config.sheetName);

  var lines = ["Tab: " + config.sheetName, "", "Available Notice Options:"];

  var invalid = [];
  for (var i = 0; i < opts.length; i++) {
    var option = String(opts[i] || "").trim();
    var parsed = parseNoticeOffset(option);
    if (parsed === null) {
      invalid.push(option);
      lines.push("✗ " + option + " → INVALID");
    } else {
      lines.push("✓ " + option + " → " + formatParsedNoticeOffset(parsed));
    }
  }

  lines.push("");
  lines.push(
    "All options are parseable: " + (invalid.length === 0 ? "YES" : "NO"),
  );
  if (invalid.length > 0) {
    lines.push("Invalid option(s): " + invalid.join(", "));
    lines.push(getSupportedNoticeDateHint());
  }

  ui.alert(
    "Notice Options",
    lines.join("\n").substring(0, 1800),
    ui.ButtonSet.OK,
  );
}

/**
 * Run comprehensive system diagnostics.
 */
function runSystemDiagnostics() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var results = [];

  // Check spreadsheet access
  try {
    var id = ss.getId();
    results.push("✓ Spreadsheet access: OK");
  } catch (e) {
    results.push("✗ Spreadsheet access: " + e.message);
  }

  // Check configured tabs
  var configs = resolveAutomationSheets(ss);
  results.push("✓ Configured tabs: " + configs.length);

  var readyTabs = 0;
  for (var i = 0; i < configs.length; i++) {
    if (configs[i].sheet) {
      var flexMap = buildFlexibleColumnMap(
        configs[i].sheet,
        configs[i].sheetName,
      );
      var hasAllRequired = flexMap.warnings.every(function (w) {
        return w.indexOf("Missing required") < 0;
      });
      if (hasAllRequired) readyTabs++;
    }
  }
  results.push("✓ Ready tabs: " + readyTabs + "/" + configs.length);

  // Check triggers
  var dailyTriggers = getTriggersByHandler("runDailyCheck").length;
  var replyTriggers = getTriggersByHandler("runReplyScan").length;
  results.push("✓ Daily triggers: " + dailyTriggers);
  results.push("✓ Reply triggers: " + replyTriggers);

  // Check properties
  var props = getAutomationProperties();
  var propKeys = props.getKeys();
  results.push("✓ Stored properties: " + propKeys.length);

  ui.alert("System Diagnostics", results.join("\n"), ui.ButtonSet.OK);
}

/**
 * View documentation placeholder.
 */
function showIntegrationLinkDialog() {
  var ui = SpreadsheetApp.getUi();
  var link = String(CONFIG.STATIC_REDIRECT_URL || "").trim();

  if (!link) {
    ui.alert(
      "Integration Link",
      "No integration link is configured. Set CONFIG.STATIC_REDIRECT_URL in the code first.",
      ui.ButtonSet.OK,
    );
    return;
  }

  var safeLink = sanitizeHtmlAttribute(link);
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;padding:16px;line-height:1.5;">' +
      '<h3 style="margin-top:0;">Integrate with Another Google Sheet</h3>' +
      "<p>Open this link in a new tab to continue the integration setup:</p>" +
      '<div style="margin:12px 0;padding:10px;border:1px solid #dadce0;border-radius:4px;background:#f8f9fa;font-size:12px;color:#444;word-break:break-all;">' +
      safeLink +
      "</div>" +
      '<p style="margin:0 0 12px 0;">' +
      '<a href="' +
      safeLink +
      '" target="_blank" rel="noopener noreferrer">' +
      safeLink +
      "</a>" +
      "</p>" +
      "<button onclick=\"window.open('" +
      safeLink +
      "', '_blank');\" " +
      'style="cursor:pointer;padding:10px 14px;background:#1a73e8;color:#fff;border:none;border-radius:4px;">Open Link in New Tab</button>' +
      "</div>",
  )
    .setWidth(420)
    .setHeight(260);

  ui.showModalDialog(html, "Integration Link");
}

function viewDocumentation() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    "Documentation",
    "Automated Expiry Notification System\n\n" +
      "This system sends automated email reminders for document expiries.\n\n" +
      "Key Features:\n" +
      "- Multi-tab support for different document types\n" +
      "- Flexible column name mapping\n" +
      "- Two-stage reminders: Notice Date and final Expiry Date reminder\n" +
      "- Global default CC emails from the Automation Settings menu\n" +
      "- Row-level Staff Email is also included in CC\n" +
      "- AI-powered email generation\n" +
      "- Reply tracking with auto-created Reply Status column\n" +
      "- Open tracking\n\n" +
      "Static Redirect Link:\n" +
      "- Set CONFIG.STATIC_REDIRECT_URL in the code to your target link\n" +
      "- If set, emails include a clickable link that opens that URL\n" +
      "- If Open Tracking URL is configured, the click routes through the script redirect first\n\n" +
      "Remarks Template Fields:\n" +
      "- [Client Name]\n" +
      "- [Document Type]\n" +
      "- [Expiry Date]\n" +
      "- [Date of Expiration]\n" +
      "- [Date of Expiry]\n" +
      "- Put these directly in the Remarks column to auto-fill row data\n" +
      "- Example: Good day [Client Name], your [Document Type] expires on [Expiry Date].\n\n" +
      "Reminder Rules:\n" +
      "- First reminder sends on the computed Notice Date target\n" +
      "- Final reminder sends on the exact Expiry Date\n" +
      "- A row is fully complete only after the final reminder is sent\n\n" +
      "Reply Tracking:\n" +
      "- The sheet auto-creates a Reply Status column next to Status when needed\n" +
      "- Pending = email was sent and the system is waiting for a reply\n" +
      "- Replied = a matching client reply was detected\n\n" +
      "For setup:\n" +
      "1. Configure automation tabs\n" +
      "2. Map columns if needed\n" +
      "3. Set up dropdowns\n" +
      "4. Set Default CC Emails (optional)\n" +
      "5. Activate schedule\n\n" +
      "See the project README for full documentation.",
    ui.ButtonSet.OK,
  );
}

/**
 * Show about dialog.
 */
function showAbout() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    "About",
    "Automated Expiry Notification System\n" +
      "Version 2.0 - Flexible Multi-Tab Edition\n\n" +
      "This sheet runs a daily scan across configured tabs.\n" +
      "It sends one reminder based on the Notice Date\n" +
      "(for example, 90 days before expiry), and one final\n" +
      "reminder on the exact Expiry Date.\n\n" +
      "Default CC emails can be configured from\n" +
      "Automation Settings > Set Default CC Emails.\n" +
      "Any Staff Email in the row is also CC'd.\n\n" +
      "You can also set CONFIG.STATIC_REDIRECT_URL in the code.\n" +
      "If it has a value, the email includes a clickable link.\n\n" +
      "When emails are sent, the sheet can auto-create a\n" +
      "Reply Status column next to Status.\n" +
      "Pending means waiting for reply.\n" +
      "Replied means the client replied and was detected.\n\n" +
      "In the Remarks column, users can write the email body\n" +
      "and use these fields:\n" +
      "[Client Name], [Document Type], [Expiry Date],\n" +
      "[Date of Expiration], [Date of Expiry]\n\n" +
      "Example:\n" +
      "Good day [Client Name], your [Document Type]\n" +
      "expires on [Expiry Date].\n\n" +
      "Rows stay Active until the final expiry-day reminder\n" +
      "is sent, then they are marked Sent.",
    ui.ButtonSet.OK,
  );
}

/**
 * Backward-compatible alias for showAutomationStatus.
 */
function showAutomationStatus() {
  showSystemStatus();
}

/**
 * Backward-compatible alias — kept so any saved references still work.
 */
function initializeAutomationSheet() {
  configureAutomationSheets();
}

/**
 * Multi-tab setup: lets the user pick one or more sheet tabs for the automation to process.
 * Stores selection as a JSON array in document properties.
 * Accepts comma-separated numbers and/or tab names in a single prompt.
 */
function configureAutomationSheets() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Show ALL tabs — user decides what to include
  var sheets = ss.getSheets();

  if (sheets.length === 0) {
    ui.alert("Configure Sheets", "No sheet tabs were found.", ui.ButtonSet.OK);
    return;
  }

  // Resolve configured tabs — this auto-purges any deleted tabs from storage
  var liveConfigs = resolveAutomationSheets(ss);
  var liveIds = liveConfigs.map(function (c) {
    return c.sheet ? c.sheet.getSheetId() : null;
  });
  var liveNames = liveConfigs.map(function (c) {
    return c.sheetName;
  });

  var options = [];
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var isActive =
      liveIds.indexOf(sh.getSheetId()) >= 0 ||
      liveNames.indexOf(sh.getName()) >= 0;
    options.push(i + 1 + ". " + sh.getName() + (isActive ? " ★ [active]" : ""));
  }

  var currentLabel = liveNames.length > 0 ? liveNames.join(", ") : "(none)";

  var response = ui.prompt(
    "Configure Automation Sheet(s)",
    "Enter tab numbers separated by commas. You can select multiple.\n" +
      "Example: 1, 3\n\n" +
      "Currently active: " +
      currentLabel +
      "\n\n" +
      options.join("\n"),
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var input = response.getResponseText().trim();
  if (!input) {
    ui.alert("No input provided. Configuration unchanged.");
    return;
  }

  var parts = input.split(",");
  var selectedEntries = [];
  var errors = [];

  for (var p = 0; p < parts.length; p++) {
    var part = parts[p].trim();
    if (!part) continue;

    var idx = parseInt(part, 10);
    var resolvedSheet = null;

    if (!isNaN(idx) && idx >= 1 && idx <= sheets.length) {
      // Number input — index into the shown list
      resolvedSheet = sheets[idx - 1];
    } else {
      // Name input — case-insensitive match
      for (var s = 0; s < sheets.length; s++) {
        if (sheets[s].getName().toLowerCase() === part.toLowerCase()) {
          resolvedSheet = sheets[s];
          break;
        }
      }
      if (!resolvedSheet) errors.push('"' + part + '"');
    }
    // Validate the resolved sheet actually exists
    if (resolvedSheet && !ss.getSheetByName(resolvedSheet.getName())) {
      errors.push('"' + resolvedSheet.getName() + '" (deleted)');
      resolvedSheet = null;
    }

    if (resolvedSheet) {
      var alreadyAdded = false;
      for (var x = 0; x < selectedEntries.length; x++) {
        if (selectedEntries[x].id === resolvedSheet.getSheetId()) {
          alreadyAdded = true;
          break;
        }
      }
      if (!alreadyAdded) {
        selectedEntries.push({
          id: resolvedSheet.getSheetId(),
          name: resolvedSheet.getName(),
        });
      }
    }
  }

  if (selectedEntries.length === 0) {
    var errMsg =
      errors.length > 0
        ? "Could not find tab(s): " +
          errors.join(", ") +
          "\n\nPlease check the numbers and try again."
        : "No valid tabs selected. Configuration unchanged.";
    ui.alert("Nothing Saved", errMsg, ui.ButtonSet.OK);
    return;
  }

  setConfiguredTabEntries(selectedEntries);

  var summary = selectedEntries
    .map(function (e, i) {
      return i + 1 + ". " + e.name + " (ID: " + e.id + ")";
    })
    .join("\n");

  var warningLine =
    errors.length > 0 ? "\n\n⚠ Not found (skipped): " + errors.join(", ") : "";

  ui.alert(
    "Configuration Saved",
    "Automation will now process " +
      selectedEntries.length +
      " tab(s):\n\n" +
      summary +
      warningLine,
    ui.ButtonSet.OK,
  );
}

/**
 * Returns the configured automation tabs as an array of {id, name} objects.
 * Handles three legacy formats:
 *   - new format: JSON array of {id, name} objects
 *   - old format: JSON array of plain strings
 *   - legacy: plain comma-separated string
 */
function getConfiguredTabEntries() {
  var props = PropertiesService.getDocumentProperties();
  var saved = props.getProperty(CONFIG.AUTOMATION_SHEET_PROPERTY_KEY);
  if (!saved || saved.trim() === "") {
    return [{ id: null, name: CONFIG.SHEET_NAME }];
  }

  var trimmed = saved.trim();
  if (trimmed.charAt(0) === "[") {
    try {
      var parsed = JSON.parse(trimmed);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed
          .filter(function (entry) {
            return (
              entry && (entry.name || entry.id || typeof entry === "string")
            );
          })
          .map(function (entry) {
            // Legacy: plain string inside the array
            if (typeof entry === "string")
              return { id: null, name: entry.trim() };
            return {
              id: entry.id || null,
              name: String(entry.name || "").trim(),
            };
          })
          .filter(function (e) {
            return e.name || e.id;
          });
      }
    } catch (e) {}
  }
  // Legacy plain string
  return [{ id: null, name: trimmed }];
}

/**
 * Returns configured automation sheet names as plain array (backward compat).
 */
function getConfiguredSheetNames() {
  return getConfiguredTabEntries().map(function (e) {
    return e.name;
  });
}

/**
 * Backward-compat getter for single-tab callers (returns first configured name).
 */
function getConfiguredSheetName() {
  return getConfiguredSheetNames()[0];
}

/**
 * Persists a list of {id, name} tab entries to document properties.
 */
function setConfiguredTabEntries(entries) {
  var clean = (entries || []).filter(function (e) {
    return e && (e.name || e.id);
  });
  if (clean.length === 0) return;
  PropertiesService.getDocumentProperties().setProperty(
    CONFIG.AUTOMATION_SHEET_PROPERTY_KEY,
    JSON.stringify(clean),
  );
}

/**
 * Persists the list of configured sheet names (backward compat — wraps into {id,name} entries).
 * Tries to resolve sheet IDs from the active spreadsheet when possible.
 */
function setConfiguredSheetNames(namesArray) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var entries = (namesArray || [])
    .filter(function (n) {
      return !!String(n || "").trim();
    })
    .map(function (name) {
      var sheet = ss.getSheetByName(name);
      return {
        id: sheet ? sheet.getSheetId() : null,
        name: String(name).trim(),
      };
    });
  setConfiguredTabEntries(entries);
}

/**
 * Backward-compat setter for single-tab callers.
 */
function setConfiguredSheetName(sheetName) {
  var value = String(sheetName || "").trim();
  if (!value) return;
  setConfiguredSheetNames([value]);
}

/**
 * Resolves all configured automation sheets.
 * Tries to find each tab by sheet ID first (rename-safe).
 * If found by ID but name changed, updates the stored name automatically.
 * Returns array of { sheetName, sheet } objects (sheet may be null if not found).
 */
function resolveAutomationSheets(ss) {
  var entries = getConfiguredTabEntries();
  var needsUpdate = false;
  var results = [];

  for (var i = 0; i < entries.length; i++) {
    var entry = entries[i];
    var foundSheet = null;

    // Try by numeric sheet ID first
    if (entry.id !== null && entry.id !== undefined) {
      var allSheets = ss.getSheets();
      for (var j = 0; j < allSheets.length; j++) {
        if (allSheets[j].getSheetId() === entry.id) {
          foundSheet = allSheets[j];
          break;
        }
      }
      // If found by ID but name has changed, update stored name
      if (foundSheet && foundSheet.getName() !== entry.name) {
        entry.name = foundSheet.getName();
        entries[i] = entry;
        needsUpdate = true;
      }
    }

    // Fallback: find by name when ID lookup failed
    if (!foundSheet && entry.name) {
      foundSheet = ss.getSheetByName(entry.name);
      if (foundSheet) {
        // Tab exists by name — update stored ID (handles deleted+recreated same-name tabs)
        entry.id = foundSheet.getSheetId();
        entries[i] = entry;
        needsUpdate = true;
      }
    }

    results.push({
      sheetName: entry.name,
      sheet: foundSheet,
      resolvedId: foundSheet ? foundSheet.getSheetId() : null,
    });
  }

  // Auto-purge entries for tabs that no longer exist
  // Also deduplicate: if two entries resolved to the same sheet ID, keep only the first one
  var surviving = [];
  var seenIds = [];
  for (var k = 0; k < results.length; k++) {
    if (!results[k].sheet) {
      // Tab not found — drop it
      needsUpdate = true;
      continue;
    }
    var resolvedId = results[k].resolvedId;
    if (seenIds.indexOf(resolvedId) >= 0) {
      // Duplicate — same physical sheet registered twice, drop the extra entry
      needsUpdate = true;
      continue;
    }
    seenIds.push(resolvedId);
    surviving.push(entries[k]);
  }

  // Persist updated names/IDs silently (including purge + dedup)
  if (needsUpdate) {
    setConfiguredTabEntries(surviving);
  }

  // Return only found, deduplicated sheets
  return results.filter(function (r) {
    return !!r.sheet && seenIds.indexOf(r.resolvedId) >= 0;
  });
}

/**
 * Backward-compat: resolves first configured sheet (single-tab callers).
 */
function resolveAutomationSheet(ss) {
  var results = resolveAutomationSheets(ss);
  return results[0] || { sheetName: CONFIG.SHEET_NAME, sheet: null };
}

/**
 * Prompts user to pick one of the configured tabs (for diagnostic functions).
 * Returns { sheetName, sheet } or null if cancelled.
 */
function promptSelectConfiguredSheet(ss, title) {
  var ui = SpreadsheetApp.getUi();

  // resolveAutomationSheets auto-purges deleted tabs; only live tabs returned
  var configs = resolveAutomationSheets(ss);

  if (configs.length === 0) {
    ui.alert(
      "No Tabs Found",
      "No configured tabs exist. Use 'Configure Automation Tabs' to register tabs.",
      ui.ButtonSet.OK,
    );
    return null;
  }

  if (configs.length === 1) {
    return configs[0];
  }

  var lastSelected = getPropString(
    getTabConfigKey("_GLOBAL", TAB_CONFIG_KEYS.LAST_SELECTED),
    "",
  );

  var options = configs.map(function (c, i) {
    var marker = c.sheetName === lastSelected ? " ★ [current]" : "";
    return i + 1 + ". " + c.sheetName + marker;
  });

  var response = ui.prompt(
    title || "Select Sheet Tab",
    "Select a tab by number:  ★ = currently selected working tab\n\n" +
      options.join("\n"),
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return null;

  var idx = parseInt(response.getResponseText().trim(), 10);
  if (isNaN(idx) || idx < 1 || idx > configs.length) {
    ui.alert("Invalid selection.");
    return null;
  }
  return configs[idx - 1];
}

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

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || "").trim());
}

function normalizeEmailList(value) {
  var rawValues = [];

  if (Array.isArray(value)) {
    rawValues = value;
  } else {
    var text = String(value || "");
    if (!text.trim()) return [];
    rawValues = text.split(/[,;\n]/);
  }

  var seen = {};
  var result = [];

  for (var i = 0; i < rawValues.length; i++) {
    var email = String(rawValues[i] || "").trim();
    if (!email) continue;

    var key = email.toLowerCase();
    if (seen[key]) continue;
    seen[key] = true;
    result.push(email);
  }

  return result;
}

function parseClientEmails(rawValue) {
  var list = normalizeEmailList(rawValue);
  return list.length > 0 ? list : [];
}

function validateEmailList(emails) {
  var invalid = [];
  for (var i = 0; i < emails.length; i++) {
    if (!isValidEmail(emails[i])) invalid.push(emails[i]);
  }
  return invalid;
}

function mergeUniqueEmails() {
  var merged = [];
  var seen = {};

  for (var i = 0; i < arguments.length; i++) {
    var source = normalizeEmailList(arguments[i]);
    for (var j = 0; j < source.length; j++) {
      var email = source[j];
      var key = email.toLowerCase();
      if (seen[key]) continue;
      seen[key] = true;
      merged.push(email);
    }
  }

  return merged;
}

function getDefaultCcEmails() {
  var raw = getPropString(PROP_KEYS.DEFAULT_CC_EMAILS, "");
  if (!raw) return [];

  try {
    var parsed = JSON.parse(raw);
    return normalizeEmailList(parsed);
  } catch (e) {
    return normalizeEmailList(raw);
  }
}

function setDefaultCcEmails(emails) {
  var normalized = normalizeEmailList(emails);
  if (normalized.length === 0) {
    setPropString(PROP_KEYS.DEFAULT_CC_EMAILS, "");
    return;
  }

  setPropString(PROP_KEYS.DEFAULT_CC_EMAILS, JSON.stringify(normalized));
}

function resolveCcEmails(clientEmails, staffEmail) {
  var combined = mergeUniqueEmails(getDefaultCcEmails(), staffEmail);
  var clientList = Array.isArray(clientEmails)
    ? clientEmails
    : normalizeEmailList(clientEmails);
  var clientKeys = {};
  for (var k = 0; k < clientList.length; k++) {
    clientKeys[clientList[k].toLowerCase()] = true;
  }
  if (Object.keys(clientKeys).length === 0) return combined;

  var filtered = [];
  for (var i = 0; i < combined.length; i++) {
    if (!clientKeys[combined[i].toLowerCase()]) filtered.push(combined[i]);
  }
  return filtered;
}

function formatCcDisplay(ccEmails) {
  var ccList = normalizeEmailList(ccEmails);
  return ccList.length > 0 ? ccList.join(", ") : "(none)";
}

function setDefaultCcEmailsDialog() {
  var ui = SpreadsheetApp.getUi();
  var current = getDefaultCcEmails();
  var response = ui.prompt(
    "Set Default CC Emails",
    "Enter one or more default CC email addresses.\n" +
      "Use commas, semicolons, or new lines.\n" +
      "Leave blank to clear.\n\n" +
      "Current: " +
      formatCcDisplay(current),
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var emails = normalizeEmailList(response.getResponseText());
  var invalid = validateEmailList(emails);
  if (invalid.length > 0) {
    ui.alert(
      "Invalid Email(s)",
      "These email addresses are invalid:\n" + invalid.join("\n"),
      ui.ButtonSet.OK,
    );
    return;
  }

  setDefaultCcEmails(emails);
  ui.alert(
    "Default CC Emails Saved",
    emails.length > 0
      ? "Default CC list:\n" + emails.join("\n")
      : "Default CC emails cleared.",
    ui.ButtonSet.OK,
  );
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
  var logsSheet = ensureLogsSheet(ss);
  var active = getTriggersByHandler("runDailyCheck");
  var replyTriggers = getTriggersByHandler("runReplyScan");

  var configuredHour = getDailyTriggerHour();
  var msg;
  if (active.length === 0) {
    msg =
      "Status: INACTIVE\n\nNo daily schedule is set up.\nConfigured send time: " +
      configuredHour +
      ":00\nUse 'Activate Daily Schedule' to enable it.";
  } else {
    var lines = [
      "Status: ACTIVE",
      "",
      "Send time: " + configuredHour + ":00 Philippine Time (Asia/Manila)",
      active.length + " trigger(s) active.",
    ];
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

  var sheetConfigs = resolveAutomationSheets(ss);
  var sheetLines = sheetConfigs.map(function (c, i) {
    return (
      "  " +
      (i + 1) +
      ". " +
      c.sheetName +
      (c.sheet ? " (found)" : " (NOT FOUND)")
    );
  });
  var anyMissing = sheetConfigs.some(function (c) {
    return !c.sheet;
  });
  msg +=
    "\n\nConfigured tab(s) (" +
    sheetConfigs.length +
    "):\n" +
    sheetLines.join("\n");
  if (anyMissing) {
    msg += "\nUse 'Configure Automation Sheet(s)' to fix missing tabs.";
  }

  msg +=
    "\n\nReply tracking schedule: " +
    (replyTriggers.length > 0
      ? "ACTIVE (" +
        replyTriggers.length +
        " trigger(s) — 9:00 AM & 3:00 PM Asia/Manila)"
      : "INACTIVE (use Automation Settings → Activate Reply Scan)");

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

  var numCols = LOG_COL.DETAIL; // read up to Detail column
  var rows = logsSheet
    .getRange(2, LOG_COL.TIMESTAMP, lastRow - 1, numCols)
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

/**
 * Main function — called daily by the time-based trigger.
 * Loops over all configured sheet tabs, reads headers from row 2, scans data from row 3+,
 * computes target dates, sends emails using remarks/AI/fallback body sources,
 * and updates status + optional metadata columns.
 */
function runDailyCheck() {
  var ss = getAutomationSpreadsheet();
  var logsSheet = ensureLogsSheet(ss);
  var sheetConfigs = resolveAutomationSheets(ss);
  var senderEmail = getSenderAccountEmail();
  var trackingEnabled = !!getOpenTrackingBaseUrl();
  var today = getMidnight(new Date());

  var totalAllTabs = 0,
    sentAllTabs = 0,
    errorsAllTabs = 0;

  for (var t = 0; t < sheetConfigs.length; t++) {
    var sheetConfig = sheetConfigs[t];
    var tabName = sheetConfig.sheetName;
    var visaSheet = sheetConfig.sheet;

    if (!visaSheet) {
      appendLog(
        logsSheet,
        tabName,
        "",
        "ERROR",
        'Sheet "' +
          tabName +
          '" not found. Use "Configure Automation Sheet(s)" to fix.',
      );
      continue;
    }

    // Get per-tab configuration
    var headerRow = getTabHeaderRow(tabName);
    var dataStartRow = getTabDataStartRow(tabName);

    var colMap = buildColumnMap(visaSheet, tabName);
    var mapError = validateColumnMap(colMap, headerRow);
    if (mapError) {
      appendLog(
        logsSheet,
        tabName,
        "",
        "ERROR",
        "[" + tabName + "] " + mapError,
      );
      continue;
    }

    var lastRow = visaSheet.getLastRow();
    if (lastRow < dataStartRow) {
      appendLog(
        logsSheet,
        tabName,
        "",
        "INFO",
        "[" + tabName + "] No data rows found. Skipping.",
      );
      continue;
    }

    var numDataRows = lastRow - dataStartRow + 1;
    var numCols = visaSheet.getLastColumn();
    if (numCols === 0) {
      appendLog(
        logsSheet,
        tabName,
        "",
        "INFO",
        "[" + tabName + "] Tab has no columns. Skipping.",
      );
      continue;
    }
    var data = visaSheet
      .getRange(dataStartRow, 1, numDataRows, numCols)
      .getValues();
    var totalRows = data.length;
    var processed = 0,
      sent = 0,
      errors = 0,
      autoActivated = 0,
      skippedMode = 0,
      skippedFuture = 0;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rowIndex = dataStartRow + i;

      var clientName = getCellStr(row, colMap.CLIENT_NAME);
      var clientEmailRaw = getCellStr(row, colMap.CLIENT_EMAIL);
      var clientEmailList = parseClientEmails(clientEmailRaw);
      var staffEmail = getCellStr(row, colMap.STAFF_EMAIL);
      var docType = getCellStr(row, colMap.DOC_TYPE);
      var expiryRaw = colMap.EXPIRY_DATE ? row[colMap.EXPIRY_DATE - 1] : "";
      var noticeStr = getCellStr(row, colMap.NOTICE_DATE);
      var remarks = getCellStr(row, colMap.REMARKS);
      var attachRaw = getCellStr(row, colMap.ATTACHMENTS);
      var status = getCellStr(row, colMap.STATUS);
      var replyStatus = getCellStr(row, colMap.REPLY_STATUS);
      var sendMode = getRowSendMode(row, colMap);

      if (!isProcessableStatus(status)) continue;

      if (isStatusBlank(status)) {
        if (colMap.STATUS) {
          setResolvedStatus(
            visaSheet,
            rowIndex,
            colMap,
            tabName,
            STATUS.ACTIVE,
          );
          status = resolveStatusValueForTab(
            visaSheet,
            tabName,
            colMap,
            STATUS.ACTIVE,
          );
          autoActivated++;
          appendLog(
            logsSheet,
            tabName,
            clientName,
            "INFO",
            "Blank Status auto-set to Active for processing.",
          );
        } else {
          status = STATUS.ACTIVE;
        }
      }

      var modeSkipReason = getSendModeSkipReason(sendMode);
      if (modeSkipReason) {
        appendLog(logsSheet, tabName, clientName, "SKIPPED", modeSkipReason);
        skippedMode++;
        continue;
      }

      processed++;

      var missing = [];
      if (!clientName) missing.push("Client Name");
      if (clientEmailList.length === 0) missing.push("Client Email");
      if (!expiryRaw) missing.push("Expiry Date");
      if (!noticeStr) missing.push("Notice Date");
      if (missing.length > 0) {
        var errMsg = "Missing required field(s): " + missing.join(", ");
        setResolvedStatus(visaSheet, rowIndex, colMap, tabName, STATUS.ERROR);
        appendLog(logsSheet, tabName, clientName, "ERROR", errMsg);
        errors++;
        continue;
      }

      var expiryDate =
        expiryRaw instanceof Date ? expiryRaw : new Date(expiryRaw);
      if (isNaN(expiryDate.getTime())) {
        setResolvedStatus(visaSheet, rowIndex, colMap, tabName, STATUS.ERROR);
        appendLog(
          logsSheet,
          tabName,
          clientName,
          "ERROR",
          "Invalid Expiry Date: " + expiryRaw,
        );
        errors++;
        continue;
      }

      var offset = parseNoticeOffset(noticeStr);
      if (offset === null) {
        setResolvedStatus(visaSheet, rowIndex, colMap, tabName, STATUS.ERROR);
        appendLog(
          logsSheet,
          tabName,
          clientName,
          "ERROR",
          'Cannot parse Notice Date: "' +
            noticeStr +
            '". ' +
            getSupportedNoticeDateHint(),
        );
        errors++;
        continue;
      }

      var targetDate = computeTargetDate(expiryDate, offset);
      var isNoticeDue = isTargetDateDue(targetDate, today);
      var isExpiryDay = isSameDay(expiryDate, today);
      var sameDayFinal = isSameDay(targetDate, expiryDate);
      var firstReminderSent = !!(colMap.SENT_AT && row[colMap.SENT_AT - 1]);
      var finalReminderSent = !!(
        colMap.FINAL_NOTICE_SENT_AT && row[colMap.FINAL_NOTICE_SENT_AT - 1]
      );

      var shouldSendNotice = isNoticeDue && !firstReminderSent;
      var shouldSendFinal = isExpiryDay && !finalReminderSent;

      if (isReplyStatusReplied(replyStatus)) {
        shouldSendNotice = false;
        if (!shouldSendFinal) {
          appendLog(
            logsSheet,
            tabName,
            clientName,
            "SKIPPED",
            "Reply keyword already received. Future reminder emails are suppressed for this row until the final expiry-day email.",
          );
        }
      }

      if (!shouldSendNotice && !shouldSendFinal) {
        skippedFuture++;
        continue;
      }

      var attachResult = resolveAttachments(attachRaw);
      if (attachResult.fatalError) {
        setResolvedStatus(visaSheet, rowIndex, colMap, tabName, STATUS.ERROR);
        appendLog(
          logsSheet,
          tabName,
          clientName,
          "ERROR",
          attachResult.fatalError,
        );
        errors++;
        continue;
      }
      if (attachResult.warnings && attachResult.warnings.length > 0) {
        appendLog(
          logsSheet,
          tabName,
          clientName,
          "INFO",
          "Attachment warning(s): " + attachResult.warnings.join("; "),
        );
      }
      var ccEmails = resolveCcEmails(clientEmailList, staffEmail);
      var baseSubject = buildEmailSubject(docType, clientName, expiryDate);
      var clientEmailDisplay = clientEmailList.join(", ");

      try {
        var sentThisRow = false;

        if (shouldSendNotice) {
          var noticeToken = trackingEnabled ? generateOpenTrackingToken() : "";
          var noticeContent = buildStageEmailContent(
            remarks,
            clientName,
            expiryDate,
            docType,
            noticeToken,
            "notice",
          );
          var noticeSubject = buildStageSubject(baseSubject, "notice");
          var displayName = getSenderDisplayName(senderEmail);
          var noticeFallbackHtml = buildFallbackLinksHtml(
            attachResult.failedLinks,
          );
          var noticeHtmlBody = noticeFallbackHtml
            ? noticeContent.htmlBody + noticeFallbackHtml
            : noticeContent.htmlBody;
          var noticeMeta = { threadId: "", messageId: "" };
          for (var ei = 0; ei < clientEmailList.length; ei++) {
            var eMeta = sendReminderEmail(
              clientEmailList[ei],
              ccEmails,
              noticeSubject,
              noticeHtmlBody,
              attachResult.blobs,
              displayName,
            );
            if (ei === 0) noticeMeta = eMeta;
          }

          colMap = ensureReplyStatusColumn(visaSheet, tabName, colMap);
          writePostSendMetadata(visaSheet, rowIndex, colMap, {
            sentAt: new Date(),
            senderEmail: senderEmail,
            openToken: noticeToken,
            threadId: noticeMeta.threadId,
            messageId: noticeMeta.messageId,
          });

          if (sameDayFinal && shouldSendFinal) {
            colMap = ensureFinalNoticeColumns(visaSheet, tabName, colMap);
            writeFinalNoticeMetadata(visaSheet, rowIndex, colMap, {
              sentAt: new Date(),
              threadId: noticeMeta.threadId,
              messageId: noticeMeta.messageId,
            });
            shouldSendFinal = false;
            setResolvedStatus(
              visaSheet,
              rowIndex,
              colMap,
              tabName,
              STATUS.SENT,
            );
          } else {
            setResolvedStatus(
              visaSheet,
              rowIndex,
              colMap,
              tabName,
              STATUS.ACTIVE,
            );
          }

          setStaffEmail(visaSheet, rowIndex, colMap.STAFF_EMAIL, senderEmail);
          appendLog(
            logsSheet,
            tabName,
            clientName,
            sameDayFinal ? "SENT_NOTICE_FINAL" : "SENT_NOTICE",
            "Email sent to " +
              clientEmailDisplay +
              (ccEmails.length > 0
                ? " (CC: " + ccEmails.join(", ") + ")"
                : "") +
              " | Stage: " +
              (sameDayFinal ? "notice+final" : "notice") +
              " | Mode: " +
              sendMode +
              " | Body: " +
              noticeContent.source +
              (attachResult.blobs.length > 0
                ? " | Attachments: " + attachResult.blobs.length
                : "") +
              (noticeToken ? " | Tracking: enabled" : "") +
              (senderEmail ? " | Sender: " + senderEmail : ""),
          );
          sent++;
          sentThisRow = true;
        }

        if (shouldSendFinal) {
          colMap = ensureFinalNoticeColumns(visaSheet, tabName, colMap);
          var finalToken = trackingEnabled ? generateOpenTrackingToken() : "";
          var finalContent = buildStageEmailContent(
            remarks,
            clientName,
            expiryDate,
            docType,
            finalToken,
            "final",
          );
          var finalSubject = buildStageSubject(baseSubject, "final");
          var finalFallbackHtml = buildFallbackLinksHtml(
            attachResult.failedLinks,
          );
          var finalHtmlBody = finalFallbackHtml
            ? finalContent.htmlBody + finalFallbackHtml
            : finalContent.htmlBody;
          var finalMeta = { threadId: "", messageId: "" };
          for (var fi = 0; fi < clientEmailList.length; fi++) {
            var fMeta = sendReminderEmail(
              clientEmailList[fi],
              ccEmails,
              finalSubject,
              finalHtmlBody,
              attachResult.blobs,
              displayName,
            );
            if (fi === 0) finalMeta = fMeta;
          }

          writeFinalNoticeMetadata(visaSheet, rowIndex, colMap, {
            sentAt: new Date(),
            threadId: finalMeta.threadId,
            messageId: finalMeta.messageId,
          });

          if (!firstReminderSent && !shouldSendNotice) {
            colMap = ensureReplyStatusColumn(visaSheet, tabName, colMap);
            writePostSendMetadata(visaSheet, rowIndex, colMap, {
              sentAt: new Date(),
              senderEmail: senderEmail,
              openToken: finalToken,
              threadId: finalMeta.threadId,
              messageId: finalMeta.messageId,
            });
          } else if (finalToken) {
            setCellValueIfColumn(
              visaSheet,
              rowIndex,
              colMap.OPEN_TOKEN,
              finalToken,
            );
          }

          setResolvedStatus(visaSheet, rowIndex, colMap, tabName, STATUS.SENT);
          setStaffEmail(visaSheet, rowIndex, colMap.STAFF_EMAIL, senderEmail);
          appendLog(
            logsSheet,
            tabName,
            clientName,
            "SENT_FINAL",
            "Email sent to " +
              clientEmailDisplay +
              (ccEmails.length > 0
                ? " (CC: " + ccEmails.join(", ") + ")"
                : "") +
              " | Stage: final" +
              " | Mode: " +
              sendMode +
              " | Body: " +
              finalContent.source +
              (attachResult.blobs.length > 0
                ? " | Attachments: " + attachResult.blobs.length
                : "") +
              (finalToken ? " | Tracking: enabled" : "") +
              (senderEmail ? " | Sender: " + senderEmail : ""),
          );
          sent++;
          sentThisRow = true;
        }

        if (!sentThisRow) {
          skippedFuture++;
        }
      } catch (e) {
        setResolvedStatus(visaSheet, rowIndex, colMap, tabName, STATUS.ERROR);
        appendLog(
          logsSheet,
          tabName,
          clientName,
          "ERROR",
          "Send failed: " + e.message,
        );
        errors++;
      }
    }

    appendLog(
      logsSheet,
      tabName,
      "",
      "SUMMARY",
      "[" +
        tabName +
        "] Total: " +
        totalRows +
        " | Eligible: " +
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

    totalAllTabs += totalRows;
    sentAllTabs += sent;
    errorsAllTabs += errors;
  }

  if (sheetConfigs.length > 1) {
    appendLog(
      logsSheet,
      "",
      "",
      "SUMMARY",
      "All tabs complete. Tabs processed: " +
        sheetConfigs.length +
        " | Total Rows: " +
        totalAllTabs +
        " | Total Sent: " +
        sentAllTabs +
        " | Total Errors: " +
        errorsAllTabs,
    );
  }
}

/**
 * Backward-compatible wrapper: returns a { HEADER_KEY: 1-based-col-index } map.
 * When tabName is provided, uses the flexible 3-tier mapping system
 * (stored mapping → global HEADERS/aliases → fuzzy detection).
 * When tabName is omitted, falls back to legacy CONFIG.HEADER_ROW behavior.
 * @param {Sheet} sheet - The sheet to map.
 * @param {string} [tabName] - Tab name for per-tab flexible mapping.
 */
function buildColumnMap(sheet, tabName) {
  if (tabName) {
    var flexible = buildFlexibleColumnMap(sheet, tabName);
    return flexible.map;
  }
  return buildColumnMapLegacy(sheet);
}

/**
 * Validates that all required columns were found in the header row.
 * Returns an error string if any are missing, or null if OK.
 * @param {Object} colMap - Column map to validate.
 * @param {number} [headerRow] - Actual header row number for the error message (optional).
 */
function validateColumnMap(colMap, headerRow) {
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
  var rowLabel = headerRow || CONFIG.HEADER_ROW;
  return missing.length > 0
    ? "Required column(s) not found in row " +
        rowLabel +
        ": " +
        missing.join(", ")
    : null;
}

/**
 * Returns the property key for a specific tab configuration.
 */
function getTabConfigKey(tabName, configType) {
  return TAB_CONFIG.PREFIX + tabName + "_" + configType;
}

/**
 * Saves column mapping for a specific tab.
 * mapping: { logicalKey: "Actual Header Name", ... }
 */
function saveTabColumnMapping(tabName, mapping) {
  var key = getTabConfigKey(tabName, TAB_CONFIG_KEYS.COLUMN_MAP);
  setPropString(key, JSON.stringify(mapping || {}));
}

/**
 * Gets stored column mapping for a specific tab.
 * Returns {} if none stored.
 */
function getTabColumnMapping(tabName) {
  var key = getTabConfigKey(tabName, TAB_CONFIG_KEYS.COLUMN_MAP);
  var raw = getPropString(key, "");
  if (!raw) return {};
  try {
    var parsed = JSON.parse(raw);
    return typeof parsed === "object" && parsed !== null ? parsed : {};
  } catch (e) {
    return {};
  }
}

/**
 * Clears column mapping for a specific tab.
 */
function clearTabColumnMapping(tabName) {
  var key = getTabConfigKey(tabName, TAB_CONFIG_KEYS.COLUMN_MAP);
  getAutomationProperties().deleteProperty(key);
}

/**
 * Gets the configured header row for a tab (defaults to TAB_CONFIG.DEFAULT_HEADER_ROW).
 */
function getTabHeaderRow(tabName) {
  var key = getTabConfigKey(tabName, TAB_CONFIG_KEYS.HEADER_ROW);
  var raw = getPropString(key, "");
  var row = parseInt(raw, 10);
  return isNaN(row) || row < 1 ? TAB_CONFIG.DEFAULT_HEADER_ROW : row;
}

/**
 * Sets the header row for a specific tab.
 */
function setTabHeaderRow(tabName, rowNum) {
  var row = parseInt(rowNum, 10);
  if (isNaN(row) || row < 1) row = TAB_CONFIG.DEFAULT_HEADER_ROW;
  var key = getTabConfigKey(tabName, TAB_CONFIG_KEYS.HEADER_ROW);
  setPropString(key, String(row));
}

/**
 * Gets the data start row for a tab (defaults to headerRow + 1).
 */
function getTabDataStartRow(tabName) {
  var key = getTabConfigKey(tabName, TAB_CONFIG_KEYS.DATA_START_ROW);
  var raw = getPropString(key, "");
  var row = parseInt(raw, 10);
  if (!isNaN(row) && row > 0) return row;
  return getTabHeaderRow(tabName) + 1;
}

/**
 * Sets the data start row for a specific tab.
 */
function setTabDataStartRow(tabName, rowNum) {
  var row = parseInt(rowNum, 10);
  if (isNaN(row) || row < 1) return;
  var key = getTabConfigKey(tabName, TAB_CONFIG_KEYS.DATA_START_ROW);
  setPropString(key, String(row));
}

/**
 * Normalizes header text for matching.
 */
function normalizeHeaderName(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[._-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Reads one full row (up to lastCol) and returns raw cell values.
 */
function getRowValues(sheet, rowNum, lastCol) {
  if (rowNum < 1 || lastCol < 1) return [];
  if (rowNum > sheet.getLastRow()) return [];
  return sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
}

/**
 * Builds a lookup of known header names from required labels and aliases.
 */
function buildKnownHeaderLookup() {
  var lookup = {};

  function addHeaderName(value) {
    var normalized = normalizeHeaderName(value);
    if (normalized) lookup[normalized] = true;
  }

  for (var key in HEADERS) {
    addHeaderName(HEADERS[key]);

    var aliases = HEADER_ALIASES[key] || [];
    for (var i = 0; i < aliases.length; i++) {
      addHeaderName(aliases[i]);
    }

    var flexibleAliases = FLEXIBLE_HEADER_ALIASES[key] || [];
    for (var j = 0; j < flexibleAliases.length; j++) {
      addHeaderName(flexibleAliases[j]);
    }
  }

  return lookup;
}

/**
 * Scores how likely a row is to be a header row.
 */
function scoreHeaderRowValues(headerValues, knownLookup) {
  var nonEmpty = 0;
  var exactRecognized = 0;
  var fuzzyRecognized = 0;

  for (var i = 0; i < headerValues.length; i++) {
    var text = String(headerValues[i] || "").trim();
    if (!text) continue;

    nonEmpty++;
    var normalized = normalizeHeaderName(text);
    if (knownLookup[normalized]) {
      exactRecognized++;
      continue;
    }

    var fuzzy = fuzzyMatchHeader(text, 0.82);
    if (fuzzy) fuzzyRecognized++;
  }

  return {
    nonEmpty: nonEmpty,
    exactRecognized: exactRecognized,
    fuzzyRecognized: fuzzyRecognized,
    score:
      exactRecognized * 3 + fuzzyRecognized * 1.5 + Math.min(nonEmpty, 8) * 0.1,
  };
}

/**
 * Resolves the effective header row for a tab.
 * If configured row looks like a divider/non-header row, scans forward and picks a better candidate.
 */
function resolveEffectiveHeaderRow(sheet, tabName, rowOverride) {
  var configuredRow = parseInt(rowOverride, 10);
  if (isNaN(configuredRow) || configuredRow < 1) {
    configuredRow = getTabHeaderRow(tabName);
  }

  var result = {
    configuredRow: configuredRow,
    headerRow: configuredRow,
    usedFallback: false,
    reason: "",
    headers: [],
  };

  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  if (lastCol === 0 || lastRow < configuredRow) {
    result.headers = [];
    return result;
  }

  var knownLookup = buildKnownHeaderLookup();
  var preferredHeaders = getRowValues(sheet, configuredRow, lastCol);
  var preferredScore = scoreHeaderRowValues(preferredHeaders, knownLookup);

  var bestRow = configuredRow;
  var bestHeaders = preferredHeaders;
  var bestScore = preferredScore;

  var scanUntil = Math.min(lastRow, configuredRow + 6);
  for (var row = configuredRow + 1; row <= scanUntil; row++) {
    var candidateHeaders = getRowValues(sheet, row, lastCol);
    var candidateScore = scoreHeaderRowValues(candidateHeaders, knownLookup);
    if (candidateScore.score > bestScore.score) {
      bestRow = row;
      bestHeaders = candidateHeaders;
      bestScore = candidateScore;
    }
  }

  var configuredLooksWeak =
    preferredScore.nonEmpty <= 1 || preferredScore.exactRecognized === 0;
  var bestLooksHeader =
    bestScore.nonEmpty >= 2 &&
    (bestScore.exactRecognized >= 1 || bestScore.fuzzyRecognized >= 2);
  var significantlyBetter =
    bestRow !== configuredRow && bestScore.score >= preferredScore.score + 2;

  if (bestRow !== configuredRow && bestLooksHeader) {
    if (configuredLooksWeak || significantlyBetter) {
      result.headerRow = bestRow;
      result.usedFallback = true;
      result.reason =
        "Configured header row " +
        configuredRow +
        " looked weak (non-empty cells: " +
        preferredScore.nonEmpty +
        "). Using row " +
        bestRow +
        " for column detection.";
      result.headers = bestHeaders;
      return result;
    }
  }

  result.headers = preferredHeaders;
  return result;
}

/**
 * Calculates fuzzy similarity between two strings (0-1 scale).
 * Uses Levenshtein distance normalized by max length.
 */
function fuzzyStringSimilarity(str1, str2) {
  var s1 = String(str1 || "")
    .toLowerCase()
    .trim();
  var s2 = String(str2 || "")
    .toLowerCase()
    .trim();

  if (s1 === s2) return 1.0;
  if (!s1 || !s2) return 0.0;

  // Check for substring match (strong indicator)
  if (s1.indexOf(s2) >= 0 || s2.indexOf(s1) >= 0) {
    return 0.8;
  }

  // Calculate Levenshtein distance
  var len1 = s1.length;
  var len2 = s2.length;
  var maxLen = Math.max(len1, len2);

  if (maxLen === 0) return 1.0;

  // Simple Levenshtein
  var matrix = [];
  for (var i = 0; i <= len1; i++) {
    matrix[i] = [i];
  }
  for (var j = 0; j <= len2; j++) {
    matrix[0][j] = j;
  }

  for (var i = 1; i <= len1; i++) {
    for (var j = 1; j <= len2; j++) {
      var cost = s1.charAt(i - 1) === s2.charAt(j - 1) ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,
        matrix[i][j - 1] + 1,
        matrix[i - 1][j - 1] + cost,
      );
    }
  }

  var distance = matrix[len1][len2];
  return 1.0 - distance / maxLen;
}

/**
 * Attempts to match an actual header to known logical columns using fuzzy matching.
 * Returns { logicalKey, similarity, matchedHeader } or null if no good match.
 */
function fuzzyMatchHeader(actualHeader, minSimilarity) {
  var minSim = minSimilarity || 0.6;
  var bestMatch = null;
  var bestScore = 0;

  for (var logicalKey in FLEXIBLE_HEADER_ALIASES) {
    var aliases = FLEXIBLE_HEADER_ALIASES[logicalKey];
    for (var i = 0; i < aliases.length; i++) {
      var score = fuzzyStringSimilarity(actualHeader, aliases[i]);
      if (score > bestScore && score >= minSim) {
        bestScore = score;
        bestMatch = {
          logicalKey: logicalKey,
          similarity: score,
          matchedHeader: aliases[i],
        };
      }
    }
  }

  return bestMatch;
}

/**
 * Auto-detects column mappings from actual sheet headers.
 * Returns array of { actualHeader, suggestedLogicalKey, confidence }.
 */
function detectColumnMappings(sheet, tabName) {
  var headerInfo = resolveEffectiveHeaderRow(sheet, tabName);
  var headerRow = headerInfo.headerRow;
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];

  var headers = headerInfo.headers || [];
  if (headers.length === 0) {
    headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  }
  var suggestions = [];

  for (var i = 0; i < headers.length; i++) {
    var actualHeader = String(headers[i] || "").trim();
    if (!actualHeader) continue;

    var match = fuzzyMatchHeader(actualHeader, 0.5);
    if (match) {
      suggestions.push({
        colIndex: i + 1,
        actualHeader: actualHeader,
        suggestedLogicalKey: match.logicalKey,
        confidence: match.similarity,
        matchedTo: match.matchedHeader,
      });
    }
  }

  // Sort by confidence descending
  suggestions.sort(function (a, b) {
    return b.confidence - a.confidence;
  });

  return suggestions;
}

/**
 * Builds a flexible column map for a sheet/tab combination.
 * Priority: 1) Stored per-tab mapping, 2) Global HEADERS + HEADER_ALIASES, 3) Fuzzy auto-detection.
 * Returns { map: {}, source: "stored|global|fuzzy|mixed", warnings: [] }
 */
function buildFlexibleColumnMap(sheet, tabName) {
  var headerInfo = resolveEffectiveHeaderRow(sheet, tabName);
  var result = {
    map: {},
    source: "",
    warnings: [],
    headerRow: headerInfo.headerRow,
    configuredHeaderRow: headerInfo.configuredRow,
  };

  if (headerInfo.usedFallback && headerInfo.reason) {
    result.warnings.push(headerInfo.reason);
  }

  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    result.warnings.push("Sheet has no columns");
    return result;
  }

  // Read actual headers from sheet (using effective row)
  var actualHeaders = headerInfo.headers || [];
  if (actualHeaders.length === 0) {
    actualHeaders = getRowValues(sheet, result.headerRow, lastCol);
  }

  // Build reverse lookup from actual header to column index
  var headerToIndex = {};
  for (var i = 0; i < actualHeaders.length; i++) {
    var h = normalizeHeaderName(actualHeaders[i]);
    if (h) headerToIndex[h] = i + 1;
  }

  // 1) Try stored per-tab mapping
  var storedMapping = getTabColumnMapping(tabName);
  var usedStored = false;
  var usedGlobal = false;
  var usedFuzzy = false;

  for (var logicalKey in storedMapping) {
    var expectedHeader = normalizeHeaderName(storedMapping[logicalKey]);
    if (expectedHeader && headerToIndex[expectedHeader]) {
      result.map[logicalKey] = headerToIndex[expectedHeader];
      usedStored = true;
    }
  }

  // 2) Fill gaps with global HEADERS + HEADER_ALIASES
  var globalMap = {};
  for (var key in HEADERS) {
    globalMap[normalizeHeaderName(HEADERS[key])] = key;
    var aliases = HEADER_ALIASES[key] || [];
    for (var a = 0; a < aliases.length; a++) {
      globalMap[normalizeHeaderName(aliases[a])] = key;
    }
  }

  for (var h in headerToIndex) {
    var globalKey = globalMap[h];
    if (globalKey && !result.map[globalKey]) {
      result.map[globalKey] = headerToIndex[h];
      usedGlobal = true;
    }
  }

  // 3) Fill remaining gaps with fuzzy matching for critical columns
  var criticalColumns = [
    "CLIENT_NAME",
    "CLIENT_EMAIL",
    "EXPIRY_DATE",
    "NOTICE_DATE",
    "STATUS",
  ];
  for (var i = 0; i < criticalColumns.length; i++) {
    var key = criticalColumns[i];
    if (!result.map[key]) {
      for (var normalizedHeader in headerToIndex) {
        var match = fuzzyMatchHeader(
          actualHeaders[headerToIndex[normalizedHeader] - 1],
          0.75,
        );
        if (match && match.logicalKey === key) {
          result.map[key] = headerToIndex[normalizedHeader];
          usedFuzzy = true;
          break;
        }
      }
    }
  }

  // Determine source label
  var sources = [];
  if (usedStored) sources.push("stored");
  if (usedGlobal) sources.push("global");
  if (usedFuzzy) sources.push("fuzzy");
  result.source = sources.join("+") || "none";

  // Validate required columns
  var required = [
    "CLIENT_NAME",
    "CLIENT_EMAIL",
    "EXPIRY_DATE",
    "NOTICE_DATE",
    "STATUS",
  ];
  for (var i = 0; i < required.length; i++) {
    if (!result.map[required[i]]) {
      result.warnings.push("Missing required column: " + required[i]);
    }
  }

  return result;
}

/**
 * Original buildColumnMap implementation for backward compatibility.
 */
function buildColumnMapLegacy(sheet) {
  var headerRow = sheet
    .getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  var reverseHeaders = {};
  for (var key in HEADERS) {
    reverseHeaders[normalizeHeaderName(HEADERS[key])] = key;
    var aliases = HEADER_ALIASES[key] || [];
    for (var a = 0; a < aliases.length; a++) {
      reverseHeaders[normalizeHeaderName(aliases[a])] = key;
    }
  }
  var map = {};
  for (var c = 0; c < headerRow.length; c++) {
    var h = normalizeHeaderName(headerRow[c]);
    if (reverseHeaders[h]) {
      map[reverseHeaders[h]] = c + 1;
    }
  }
  return map;
}

/**
 * Validates flexible column map result.
 * Returns null if valid, error string if invalid.
 */
function validateFlexibleColumnMap(flexibleResult) {
  if (!flexibleResult || !flexibleResult.map) {
    return "No column map available";
  }

  var required = [
    "CLIENT_NAME",
    "CLIENT_EMAIL",
    "EXPIRY_DATE",
    "NOTICE_DATE",
    "STATUS",
  ];
  var missing = [];

  for (var i = 0; i < required.length; i++) {
    if (!flexibleResult.map[required[i]]) {
      missing.push(required[i]);
    }
  }

  if (missing.length > 0) {
    return (
      "Missing required columns: " +
      missing.join(", ") +
      "\n\nDetected source: " +
      flexibleResult.source +
      (flexibleResult.warnings.length > 0
        ? "\nWarnings: " + flexibleResult.warnings.join("; ")
        : "")
    );
  }

  return null;
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

function isReplyStatusReplied(replyStatusValue) {
  return (
    String(replyStatusValue || "").trim().toLowerCase() ===
    REPLY_STATUS.REPLIED.toLowerCase()
  );
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
 * @param {Sheet} sheet - The sheet to search.
 * @param {Object} colMap - Column map for the sheet.
 * @param {*} noValue - The No. value to search for.
 * @param {number} [dataStartRow] - First data row (per-tab). Defaults to CONFIG.DATA_START_ROW.
 */
function findRowNumberByNo(sheet, colMap, noValue, dataStartRow) {
  var startRow = dataStartRow || CONFIG.DATA_START_ROW;
  if (!colMap.NO) {
    return {
      rowNum: null,
      warning: "",
      error: 'Column "No." not found in the header row.',
    };
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
    return { rowNum: null, warning: "", error: "No data rows found." };
  }

  var numDataRows = lastRow - startRow + 1;
  var noValues = sheet
    .getRange(startRow, colMap.NO, numDataRows, 1)
    .getValues();

  var matches = [];
  for (var i = 0; i < noValues.length; i++) {
    if (isSameNoValue(noValues[i][0], noValue)) {
      matches.push(startRow + i);
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

function getDefaultStatusOptions() {
  return [STATUS.ACTIVE, STATUS.SENT, STATUS.ERROR, STATUS.SKIPPED];
}

function getStatusOptionsForTab(sheet, tabName, colMap) {
  if (!sheet || !colMap || !colMap.STATUS) return getDefaultStatusOptions();

  var dataStartRow = getTabDataStartRow(tabName);
  var sampleRow = Math.max(dataStartRow, 1);

  try {
    var rule = sheet.getRange(sampleRow, colMap.STATUS).getDataValidation();
    if (!rule) return getDefaultStatusOptions();

    var criteriaType = rule.getCriteriaType();
    var criteriaValues = rule.getCriteriaValues();
    if (
      criteriaType !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST ||
      !criteriaValues ||
      !criteriaValues[0]
    ) {
      return getDefaultStatusOptions();
    }

    var options = criteriaValues[0]
      .map(function (value) {
        return String(value || "").trim();
      })
      .filter(function (value) {
        return !!value;
      });

    return options.length > 0 ? options : getDefaultStatusOptions();
  } catch (e) {
    return getDefaultStatusOptions();
  }
}

function resolveStatusValueForTab(sheet, tabName, colMap, desiredStatus) {
  var target = String(desiredStatus || "").trim();
  if (!target) return target;

  var options = getStatusOptionsForTab(sheet, tabName, colMap);
  var targetLower = target.toLowerCase();

  for (var i = 0; i < options.length; i++) {
    if (
      String(options[i] || "")
        .trim()
        .toLowerCase() === targetLower
    ) {
      return options[i];
    }
  }

  return target;
}

function setResolvedStatus(sheet, rowIndex, colMap, tabName, desiredStatus) {
  if (!colMap || !colMap.STATUS) return;
  setStatus(
    sheet,
    rowIndex,
    colMap.STATUS,
    resolveStatusValueForTab(sheet, tabName, colMap, desiredStatus),
  );
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

function applyReplyStatusValidation(sheet, tabName, colMap) {
  if (!sheet || !colMap || !colMap.REPLY_STATUS) return;

  var dataStartRow = getTabDataStartRow(tabName);
  var lastRow = sheet.getLastRow();
  var dataLastRow = Math.max(lastRow, dataStartRow + 100);
  var replyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([REPLY_STATUS.PENDING, REPLY_STATUS.REPLIED], true)
    .setAllowInvalid(true)
    .build();

  sheet
    .getRange(
      dataStartRow,
      colMap.REPLY_STATUS,
      dataLastRow - dataStartRow + 1,
      1,
    )
    .setDataValidation(replyRule);
}

function ensureReplyStatusColumn(sheet, tabName, colMap) {
  if (!sheet || !colMap || !colMap.STATUS) return colMap || {};
  if (colMap.REPLY_STATUS) {
    applyReplyStatusValidation(sheet, tabName, colMap);
    return colMap;
  }

  var headerRow = getTabHeaderRow(tabName);
  sheet.insertColumnAfter(colMap.STATUS);
  sheet.getRange(headerRow, colMap.STATUS + 1).setValue(HEADERS.REPLY_STATUS);

  var updatedMap = buildColumnMap(sheet, tabName);
  applyReplyStatusValidation(sheet, tabName, updatedMap);
  return updatedMap;
}

function ensureColumnExists(sheet, tabName, logicalKey) {
  if (logicalKey && buildColumnMap(sheet, tabName)[logicalKey]) {
    return buildColumnMap(sheet, tabName)[logicalKey];
  }

  var headerRow = getTabHeaderRow(tabName);
  var colIndex = sheet.getLastColumn() + 1;
  sheet.getRange(headerRow, colIndex).setValue(HEADERS[logicalKey]);
  return buildColumnMap(sheet, tabName)[logicalKey] || colIndex;
}

function ensureReplyMetadataColumns(sheet, tabName, colMap) {
  var updatedMap = colMap || {};

  updatedMap = ensureReplyStatusColumn(sheet, tabName, updatedMap);

  if (!updatedMap.REPLIED_AT) {
    updatedMap.REPLIED_AT = ensureColumnExists(sheet, tabName, "REPLIED_AT");
  }

  return buildColumnMap(sheet, tabName);
}

function ensureFinalNoticeColumns(sheet, tabName, colMap) {
  var keys = [
    "FINAL_NOTICE_SENT_AT",
    "FINAL_NOTICE_THREAD_ID",
    "FINAL_NOTICE_MESSAGE_ID",
  ];

  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    if (!colMap[key]) {
      colMap[key] = ensureColumnExists(sheet, tabName, key);
    }
  }

  return colMap;
}

function writeFinalNoticeMetadata(sheet, rowIndex, colMap, meta) {
  var finalAt = meta && meta.sentAt ? meta.sentAt : new Date();
  var threadId = meta && meta.threadId ? String(meta.threadId) : "";
  var messageId = meta && meta.messageId ? String(meta.messageId) : "";

  setCellValueIfColumn(sheet, rowIndex, colMap.FINAL_NOTICE_SENT_AT, finalAt);
  setCellValueIfColumn(
    sheet,
    rowIndex,
    colMap.FINAL_NOTICE_THREAD_ID,
    threadId,
  );
  setCellValueIfColumn(
    sheet,
    rowIndex,
    colMap.FINAL_NOTICE_MESSAGE_ID,
    messageId,
  );
}

function buildStageSubject(baseSubject, stage) {
  if (stage === "final") {
    return "Final Reminder: " + baseSubject + " (Expires Today)";
  }
  return baseSubject;
}

function buildStageEmailContent(
  remarks,
  clientName,
  expiryDate,
  docType,
  openToken,
  stage,
) {
  var content = buildEmailContent(
    remarks,
    clientName,
    expiryDate,
    docType,
    openToken,
  );

  if (stage === "final") {
    content.htmlBody =
      "<p><strong>This is your final reminder. The document expires today.</strong></p>" +
      content.htmlBody;
  }

  return content;
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

function getSenderDisplayName(email) {
  var raw = String(email || "").trim();
  var atIndex = raw.indexOf("@");
  if (atIndex < 0) return CONFIG.SENDER_NAME;

  var domain = raw.slice(atIndex + 1);
  var domainBase = domain.split(".")[0];
  if (!domainBase) return CONFIG.SENDER_NAME;

  return domainBase.charAt(0).toUpperCase() + domainBase.slice(1).toLowerCase();
}

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

function getSupportedNoticeDateHint() {
  return (
    'Supported formats: "N days/weeks/months/years before", ' +
    '"N days/weeks/months/years", numeric day count (e.g. "7"), ' +
    '"On expiry date", or an explicit date (e.g. "2026-12-31").'
  );
}

function parseIsoDateString(value) {
  var text = String(value || "").trim();
  var match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!match) return null;

  var year = parseInt(match[1], 10);
  var month = parseInt(match[2], 10);
  var day = parseInt(match[3], 10);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;

  var parsed = new Date(year, month - 1, day);
  if (
    parsed.getFullYear() !== year ||
    parsed.getMonth() !== month - 1 ||
    parsed.getDate() !== day
  ) {
    return null;
  }

  return parsed;
}

/**
 * Parses a Notice Date dropdown string into an offset object dynamically.
 * Supports any "N days/weeks/months before" pattern — no hardcoded list.
 * @returns {{ unit: string, value: number }} or null if unrecognized.
 */
function parseNoticeOffset(noticeStr) {
  var s = String(noticeStr || "")
    .toLowerCase()
    .trim();
  if (!s) return null;

  if (/^on(\s+the)?\s+expiry\s+date$/.test(s)) {
    return { unit: "days", value: 0 };
  }

  if (/^\d+$/.test(s)) {
    return { unit: "days", value: parseInt(s, 10) };
  }

  var relative = s.match(
    /^(\d+)\s*(day|days|week|weeks|month|months|year|years|yr|yrs)(?:\s+before)?$/,
  );
  if (relative) {
    var value = parseInt(relative[1], 10);
    var unitRaw = relative[2];

    if (unitRaw === "week" || unitRaw === "weeks") {
      return { unit: "days", value: value * 7 };
    }
    if (unitRaw === "month" || unitRaw === "months") {
      return { unit: "months", value: value };
    }
    if (
      unitRaw === "year" ||
      unitRaw === "years" ||
      unitRaw === "yr" ||
      unitRaw === "yrs"
    ) {
      return { unit: "years", value: value };
    }

    return { unit: "days", value: value };
  }

  var isoDate = parseIsoDateString(s);
  if (isoDate) {
    return { unit: "absolute_date", value: getMidnight(isoDate) };
  }

  var directDate = new Date(s);
  if (!isNaN(directDate.getTime())) {
    return { unit: "absolute_date", value: getMidnight(directDate) };
  }

  return null;
}

/**
 * Computes the target send date from an expiry date and a parsed offset object.
 */
function computeTargetDate(expiryDate, offset) {
  if (offset.unit === "absolute_date") {
    return getMidnight(offset.value);
  }

  var target = new Date(expiryDate);
  if (offset.unit === "days") {
    target.setDate(target.getDate() - offset.value);
  } else if (offset.unit === "months") {
    target.setMonth(target.getMonth() - offset.value);
  } else if (offset.unit === "years") {
    target.setFullYear(target.getFullYear() - offset.value);
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
 * Splits a raw "Attached Files" cell value into individual entries.
 * Supports comma-separated and newline-separated Drive URLs/IDs.
 */
function splitAttachmentEntries(rawField) {
  if (!rawField || rawField.trim() === "") return [];
  // Normalize: treat newlines as delimiters alongside commas
  var normalized = String(rawField).replace(/[\r\n]+/g, ",");
  return normalized
    .split(",")
    .map(function (e) {
      return e.trim();
    })
    .filter(function (e) {
      return !!e;
    });
}

/**
 * Resolves a comma/newline-separated list of Drive URLs/IDs into blobs.
 * Partial failure mode: collects warnings for bad entries but continues fetching valid files.
 * Only sets fatalError when no entries were provided at all (not a Drive link issue).
 * @param {string} rawField - Cell value from "Attached Files" column (comma or newline separated).
 * @returns {{ blobs: Array, warnings: Array, fatalError: string|null }}
 *   blobs: array of { blob, name, fileId }
 *   warnings: per-entry error messages for entries that failed
 *   fatalError: set only for unparseable non-URL entries (no fileId extractable)
 */
function resolveAttachments(rawField) {
  if (!rawField || rawField.trim() === "") {
    return { blobs: [], warnings: [], failedLinks: [], fatalError: null };
  }

  var entries = splitAttachmentEntries(rawField);
  if (entries.length === 0) {
    return { blobs: [], warnings: [], failedLinks: [], fatalError: null };
  }

  var blobs = [];
  var warnings = [];
  var failedLinks = [];

  for (var i = 0; i < entries.length; i++) {
    var entry = entries[i];
    var fileId = extractDriveFileId(entry);

    if (!fileId) {
      warnings.push('Cannot parse Drive file ID from: "' + entry + '"');
      failedLinks.push({ label: entry, url: null });
      continue;
    }

    try {
      var file = DriveApp.getFileById(fileId);
      blobs.push({
        blob: file.getBlob(),
        name: file.getName(),
        fileId: fileId,
      });
    } catch (e) {
      warnings.push(
        "File not found or inaccessible (ID: " + fileId + "): " + e.message,
      );
      var originalUrl =
        entry.indexOf("drive.google.com") >= 0
          ? entry
          : "https://drive.google.com/file/d/" + fileId + "/view";
      failedLinks.push({ label: fileId, url: originalUrl });
    }
  }

  return {
    blobs: blobs,
    warnings: warnings,
    failedLinks: failedLinks,
    fatalError: null,
  };
}

function buildFallbackLinksHtml(failedLinks) {
  if (!failedLinks || failedLinks.length === 0) return "";

  var items = [];
  for (var i = 0; i < failedLinks.length; i++) {
    var fl = failedLinks[i];
    if (fl.url) {
      items.push(
        '<li><a href="' +
          sanitizeHtmlAttribute(fl.url) +
          '" target="_blank" rel="noopener noreferrer">' +
          sanitizeHtmlContent(fl.label) +
          "</a></li>",
      );
    } else {
      items.push("<li>" + sanitizeHtmlContent(fl.label) + "</li>");
    }
  }

  return (
    '<div style="margin-top:16px;padding:12px;background:#fff8e1;border-left:3px solid #f9a825;font-size:13px;">' +
    '<p style="margin:0 0 8px 0;font-weight:bold;color:#7a5c00;">Some files could not be attached. You can access them via the links below:</p>' +
    '<ul style="margin:0;padding-left:20px;">' +
    items.join("") +
    "</ul></div>"
  );
}

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
    var fallback = resolveFallbackTemplateText();
    bodyText = applyTemplatePlaceholders(
      fallback.text,
      clientName,
      expiryStr,
      docTypeText,
    );
    source = fallback.source;
  }

  var htmlBody = String(bodyText || "").replace(/\n/g, "<br>");
  htmlBody = injectOpenTrackingPixel(htmlBody, openToken);

  htmlBody = [
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#333;max-width:600px;line-height:1.6;">',
    htmlBody,
    '<hr style="border:none;border-top:1px solid #eee;margin:24px 0;">',
    '<p style="font-size:11px;color:#999;">This is an automated reminder. Email opens may be tracked for delivery visibility.</p>',
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

function getStaticRedirectLinkUrl() {
  var targetUrl = String(CONFIG.STATIC_REDIRECT_URL || "").trim();
  if (!targetUrl) return "";

  var baseUrl = getOpenTrackingBaseUrl();
  if (!baseUrl) return targetUrl;

  var separator = baseUrl.indexOf("?") >= 0 ? "&" : "?";
  return baseUrl + separator + "mode=click&u=" + encodeURIComponent(targetUrl);
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
 * @param {Array|string} ccEmails - CC addresses (may be empty).
 * @param {string} subject     - Email subject.
 * @param {string} htmlBody    - HTML email body.
 * @param {Array}  blobItems   - Array of {blob, name} objects.
 */
function sendReminderEmail(
  clientEmail,
  ccEmails,
  subject,
  htmlBody,
  blobItems,
  senderName,
) {
  var options = {
    htmlBody: htmlBody,
    name: senderName || CONFIG.SENDER_NAME,
  };

  var normalizedCc = normalizeEmailList(ccEmails);
  if (normalizedCc.length > 0) {
    options.cc = normalizedCc.join(",");
  }

  if (blobItems && blobItems.length > 0) {
    var attachments = blobItems.map(function (item) {
      return item.blob.setName(item.name || "attachment");
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

function escapeForQuotedPrintable(value) {
  return String(value || "").replace(/"/g, '\\"');
}

function initializeLogsSheetLayout(sheet) {
  var headers = ["Timestamp", "Tab", "Client Name", "Action", "Detail"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet
    .getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#4a86e8")
    .setFontColor("#ffffff");
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(LOG_COL.TIMESTAMP, 160);
  sheet.setColumnWidth(LOG_COL.TAB, 130);
  sheet.setColumnWidth(LOG_COL.CLIENT_NAME, 200);
  sheet.setColumnWidth(LOG_COL.ACTION, 90);
  sheet.setColumnWidth(LOG_COL.DETAIL, 450);
}

function ensureLogsSheet(ss) {
  var sheet = ss.getSheetByName(CONFIG.LOGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.LOGS_SHEET_NAME);
    initializeLogsSheetLayout(sheet);
    return sheet;
  }

  if (sheet.getLastColumn() < 1) {
    initializeLogsSheetLayout(sheet);
    return sheet;
  }

  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var hasTabCol = false;
  for (var h = 0; h < headerRow.length; h++) {
    if (
      String(headerRow[h] || "")
        .trim()
        .toLowerCase() === "tab"
    ) {
      hasTabCol = true;
      break;
    }
  }
  if (!hasTabCol) {
    sheet.insertColumnBefore(LOG_COL.TAB);
    sheet
      .getRange(1, LOG_COL.TAB)
      .setValue("Tab")
      .setFontWeight("bold")
      .setBackground("#4a86e8")
      .setFontColor("#ffffff");
    sheet.setColumnWidth(LOG_COL.TAB, 130);
  }

  return sheet;
}

function appendLog(logsSheet, tabName, clientName, action, detail) {
  logsSheet.appendRow([
    new Date(),
    tabName || "",
    clientName || "",
    action,
    detail,
  ]);
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

/**
 * Returns the configured daily trigger hour (defaults to CONFIG.TRIGGER_HOUR).
 */
function getDailyTriggerHour() {
  var stored = getPropString(PROP_KEYS.DAILY_TRIGGER_HOUR, "");
  var parsed = parseInt(stored, 10);
  return !isNaN(parsed) && parsed >= 0 && parsed <= 23
    ? parsed
    : CONFIG.TRIGGER_HOUR;
}

/**
 * Dialog to let the user set the daily send time.
 */
function setDailyTriggerHourDialog() {
  var ui = SpreadsheetApp.getUi();
  var current = getDailyTriggerHour();
  var response = ui.prompt(
    "Set Daily Send Time",
    "Enter the hour (0–23) for the daily email schedule.\n\nExamples: 8 = 8:00 AM, 14 = 2:00 PM\n\nCurrent: " +
      current +
      ":00",
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var input = parseInt(response.getResponseText().trim(), 10);
  if (isNaN(input) || input < 0 || input > 23) {
    ui.alert("Invalid hour. Please enter a number between 0 and 23.");
    return;
  }
  setPropString(PROP_KEYS.DAILY_TRIGGER_HOUR, String(input));
  ui.alert(
    "Send Time Updated",
    "Daily send time set to " +
      input +
      ":00 Philippine Time.\n\nClick 'Activate Daily Schedule' to apply the new time.",
    ui.ButtonSet.OK,
  );
}

function installTrigger() {
  rememberSpreadsheetId(SpreadsheetApp.getActiveSpreadsheet());
  removeTrigger();

  var hour = getDailyTriggerHour();

  ScriptApp.newTrigger("runDailyCheck")
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .inTimezone("Asia/Manila")
    .create();

  var msg =
    "Daily schedule activated. runDailyCheck() will run automatically every day at " +
    hour +
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
 * Installs two daily reply scan triggers: 9 AM and 3 PM.
 */
function installReplyScanTrigger() {
  rememberSpreadsheetId(SpreadsheetApp.getActiveSpreadsheet());
  removeReplyScanTrigger();

  ScriptApp.newTrigger("runReplyScan")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .inTimezone("Asia/Manila")
    .create();

  ScriptApp.newTrigger("runReplyScan")
    .timeBased()
    .everyDays(1)
    .atHour(15)
    .inTimezone("Asia/Manila")
    .create();

  var msg =
    "Reply scan activated: runs twice daily at 9:00 AM and 3:00 PM Philippine Time.";
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
 * Trigger-safe reply scan routine. Loops over all configured tabs.
 */
function runReplyScan() {
  var ss = getAutomationSpreadsheet();
  var logsSheet = ensureLogsSheet(ss);
  var sheetConfigs = resolveAutomationSheets(ss);
  var keywords = getReplyKeywords();

  var totalScanned = 0,
    totalUpdated = 0,
    totalSkipped = 0;
  var tabSummaries = [];

  for (var t = 0; t < sheetConfigs.length; t++) {
    var sheetConfig = sheetConfigs[t];
    var tabName = sheetConfig.sheetName;
    var sheet = sheetConfig.sheet;

    if (!sheet) {
      appendLog(
        logsSheet,
        tabName,
        "",
        "ERROR",
        "[" + tabName + "] Sheet not found. Skipping reply scan.",
      );
      continue;
    }

    // Get per-tab configuration
    var headerRow = getTabHeaderRow(tabName);
    var dataStartRow = getTabDataStartRow(tabName);

    var colMap = buildColumnMap(sheet, tabName);
    var mapError = validateColumnMap(colMap, headerRow);
    if (mapError) {
      appendLog(
        logsSheet,
        tabName,
        "",
        "ERROR",
        "[" + tabName + "] " + mapError,
      );
      continue;
    }

    colMap = ensureReplyStatusColumn(sheet, tabName, colMap);

    if (!colMap.SENT_THREAD_ID) {
      var skipMsg =
        "[" + tabName + '] Reply scan skipped: add "Sent Thread Id" column.';
      appendLog(logsSheet, tabName, "", "INFO", skipMsg);
      tabSummaries.push(skipMsg);
      continue;
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < dataStartRow) {
      appendLog(
        logsSheet,
        tabName,
        "",
        "INFO",
        "[" + tabName + "] No data rows.",
      );
      continue;
    }

    var numDataRows = lastRow - dataStartRow + 1;
    var numCols = sheet.getLastColumn();
    if (numCols === 0) {
      appendLog(
        logsSheet,
        tabName,
        "",
        "INFO",
        "[" + tabName + "] Tab has no columns. Skipping reply scan.",
      );
      continue;
    }
    var data = sheet
      .getRange(dataStartRow, 1, numDataRows, numCols)
      .getValues();

    var scanned = 0,
      updated = 0,
      skipped = 0;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rowIndex = dataStartRow + i;
      var status = getCellStr(row, colMap.STATUS);
      var replyStatus = getCellStr(row, colMap.REPLY_STATUS);
      var clientEmail = getCellStr(row, colMap.CLIENT_EMAIL);
      var clientName = getCellStr(row, colMap.CLIENT_NAME);
      var threadId = getCellStr(row, colMap.SENT_THREAD_ID);
      var sentAtRaw = colMap.SENT_AT ? row[colMap.SENT_AT - 1] : "";
      var sentAt = sentAtRaw instanceof Date ? sentAtRaw : new Date(sentAtRaw);

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
      if (!(sentAt instanceof Date) || isNaN(sentAt.getTime())) {
        skipped++;
        continue;
      }

      scanned++;
      var match = findReplyMatchForRow(threadId, clientEmail, keywords, sentAt);
      if (!match) continue;

      colMap = ensureReplyMetadataColumns(sheet, tabName, colMap);

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
        tabName,
        clientName,
        "INFO",
        'Reply detected. Keyword: "' +
          (match.keyword || "") +
          '" | From: ' +
          (match.from || clientEmail),
      );
      updated++;
    }

    var tabSummary =
      "[" +
      tabName +
      "] Scanned: " +
      scanned +
      " | Updated: " +
      updated +
      " | Skipped: " +
      skipped;
    appendLog(logsSheet, tabName, "", "INFO", tabSummary);
    tabSummaries.push(tabSummary);
    totalScanned += scanned;
    totalUpdated += updated;
    totalSkipped += skipped;
  }

  var overallSummary =
    sheetConfigs.length > 1
      ? "Reply scan complete (" +
        sheetConfigs.length +
        " tabs). " +
        "Total scanned: " +
        totalScanned +
        " | Updated: " +
        totalUpdated +
        " | Skipped: " +
        totalSkipped
      : tabSummaries[0] || "Reply scan complete.";

  if (sheetConfigs.length > 1) {
    appendLog(logsSheet, "", "", "INFO", overallSummary);
  }

  return overallSummary;
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
  var haystack = String(text || "");
  for (var i = 0; i < keywords.length; i++) {
    var keyword = String(keywords[i] || "").trim();
    if (!keyword) continue;
    var escaped = keyword.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    var regex = new RegExp("(^|\\b)" + escaped + "(\\b|$)", "i");
    if (regex.test(haystack)) return keyword.toUpperCase();
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
    "Reply Status column: auto-created next to Status when needed",
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

var DEFAULT_NOTICE_OPTIONS = [
  "7 days before",
  "14 days before",
  "30 days before",
  "60 days before",
  "90 days before",
  "1 year before",
  "2 years before",
  "On expiry date",
];

var PROP_KEY_NOTICE_OPTIONS_PREFIX = "NOTICE_OPTIONS_";

/**
 * Returns stored Notice Date options for a given sheet tab, or defaults.
 */
function getNoticeOptionsForTab(tabName) {
  var key = PROP_KEY_NOTICE_OPTIONS_PREFIX + tabName;
  var raw = getPropString(key, "");
  if (!raw) return DEFAULT_NOTICE_OPTIONS.slice();
  var list = raw
    .split(",")
    .map(function (s) {
      return s.trim();
    })
    .filter(function (s) {
      return !!s;
    });
  return list.length > 0 ? list : DEFAULT_NOTICE_OPTIONS.slice();
}

/**
 * Saves Notice Date options for a given sheet tab.
 */
function setNoticeOptionsForTab(tabName, optionsArray) {
  var key = PROP_KEY_NOTICE_OPTIONS_PREFIX + tabName;
  var clean = (optionsArray || []).filter(function (s) {
    return !!String(s || "").trim();
  });
  setPropString(key, clean.join(", "));
}

function findInvalidNoticeOptions(optionsArray) {
  var invalid = [];
  var options = optionsArray || [];

  for (var i = 0; i < options.length; i++) {
    var option = String(options[i] || "").trim();
    if (!option) continue;
    if (parseNoticeOffset(option) === null) {
      invalid.push(option);
    }
  }

  return invalid;
}

/**
 * DIAGNOSTIC: Applies dropdown DataValidation rules to Status, Send Mode, and Notice Date columns
 * on a selected sheet tab. Scoped to that tab only — each tab can have different Notice Date options.
 * Purely cosmetic — the automation parsing works regardless of dropdown setup.
 */
function setupSheetDropdowns() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetConfig = promptSelectConfiguredSheet(
    ss,
    "Setup Dropdowns — Select Sheet",
  );
  if (!sheetConfig) return;

  var sheet = sheetConfig.sheet;
  var tabName = sheetConfig.sheetName;
  if (!sheet) {
    ui.alert(
      'Sheet "' + tabName + '" not found. Use "Configure Automation Sheet(s)".',
    );
    return;
  }

  // Get per-tab configuration
  var dataStartRow = getTabDataStartRow(tabName);

  var colMap = buildColumnMap(sheet, tabName);
  var mapError = validateColumnMap(colMap, getTabHeaderRow(tabName));
  if (mapError) {
    ui.alert("Column map error: " + mapError);
    return;
  }

  // Prompt for custom Notice Date options for this tab
  var currentOptions = getNoticeOptionsForTab(tabName);
  var optionResponse = ui.prompt(
    "Notice Date Options — " + tabName,
    "Enter comma-separated Notice Date options for this tab.\n" +
      "Leave blank to use the standard defaults.\n\n" +
      "Current:\n" +
      currentOptions.join(", "),
    ui.ButtonSet.OK_CANCEL,
  );
  if (optionResponse.getSelectedButton() !== ui.Button.OK) return;

  var customInput = optionResponse.getResponseText().trim();
  var noticeOptions;
  var shouldSaveCustomOptions = false;
  if (customInput) {
    noticeOptions = customInput
      .split(",")
      .map(function (s) {
        return s.trim();
      })
      .filter(function (s) {
        return !!s;
      });
    if (noticeOptions.length > 0) {
      shouldSaveCustomOptions = true;
    } else {
      noticeOptions = DEFAULT_NOTICE_OPTIONS.slice();
    }
  } else {
    noticeOptions = currentOptions;
  }

  var invalidNoticeOptions = findInvalidNoticeOptions(noticeOptions);
  if (invalidNoticeOptions.length > 0) {
    ui.alert(
      "Invalid Notice Date Options",
      "These options are not parseable:\n- " +
        invalidNoticeOptions.join("\n- ") +
        "\n\n" +
        getSupportedNoticeDateHint(),
      ui.ButtonSet.OK,
    );
    return;
  }

  if (shouldSaveCustomOptions) {
    setNoticeOptionsForTab(tabName, noticeOptions);
  }

  var lastRow = sheet.getLastRow();
  var dataLastRow = Math.max(lastRow, dataStartRow + 100); // apply to at least 100 rows ahead

  var applied = [];

  // Status dropdown
  if (colMap.STATUS) {
    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(
        [STATUS.ACTIVE, STATUS.SENT, STATUS.ERROR, STATUS.SKIPPED],
        true,
      )
      .setAllowInvalid(true)
      .build();
    sheet
      .getRange(dataStartRow, colMap.STATUS, dataLastRow - dataStartRow + 1, 1)
      .setDataValidation(statusRule);
    applied.push("Status");
  }

  // Send Mode dropdown
  if (colMap.SEND_MODE) {
    var sendModeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(
        [SEND_MODE.AUTO, SEND_MODE.HOLD, SEND_MODE.MANUAL_ONLY],
        true,
      )
      .setAllowInvalid(true)
      .build();
    sheet
      .getRange(
        dataStartRow,
        colMap.SEND_MODE,
        dataLastRow - dataStartRow + 1,
        1,
      )
      .setDataValidation(sendModeRule);
    applied.push("Send Mode");
  }

  // Notice Date dropdown (per-tab options)
  if (colMap.NOTICE_DATE) {
    var noticeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(noticeOptions, true)
      .setAllowInvalid(true)
      .build();
    sheet
      .getRange(
        dataStartRow,
        colMap.NOTICE_DATE,
        dataLastRow - dataStartRow + 1,
        1,
      )
      .setDataValidation(noticeRule);
    applied.push("Notice Date (" + noticeOptions.length + " options)");
  }

  if (applied.length === 0) {
    ui.alert(
      "No matching columns found for dropdown setup in sheet: " +
        tabName +
        "\n\nRequired columns: Status, Send Mode, Notice Date (any found will get dropdowns).",
    );
    return;
  }

  ui.alert(
    "Dropdowns Applied — " + tabName,
    "Applied to rows " +
      dataStartRow +
      "-" +
      dataLastRow +
      ":\n\n" +
      applied
        .map(function (a, i) {
          return i + 1 + ". " + a;
        })
        .join("\n") +
      "\n\nNote: setAllowInvalid(true) — users can still type custom values.",
    ui.ButtonSet.OK,
  );
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
function sanitizeHtmlContent(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function sanitizeHtmlAttribute(value) {
  return String(value || "").replace(/["<>]/g, "");
}

/**
 * Records open/click events into sheet open tracking columns.
 */
function recordOpenTrackingEvent(token, mode) {
  try {
    var ss = getAutomationSpreadsheet();
    var sheetConfigs = resolveAutomationSheets(ss);

    // Search all configured tabs for the token
    for (var i = 0; i < sheetConfigs.length; i++) {
      var sheetConfig = sheetConfigs[i];
      var sheet = sheetConfig.sheet;
      var tabName = sheetConfig.sheetName;
      if (!sheet) continue;

      var colMap = buildColumnMap(sheet, tabName);
      if (!colMap.OPEN_TOKEN) continue;

      var dataStartRow = getTabDataStartRow(tabName);
      var rowNum = findRowNumberByToken(sheet, colMap, token, dataStartRow);
      if (!rowNum) continue;

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

      colMap = ensureReplyMetadataColumns(sheet, tabName, colMap);
      setCellValueIfColumn(sheet, rowNum, colMap.REPLY_STATUS, REPLY_STATUS.REPLIED);
      if (colMap.REPLIED_AT) {
        var repliedAtCell = sheet.getRange(rowNum, colMap.REPLIED_AT);
        if (!repliedAtCell.getValue()) repliedAtCell.setValue(now);
      }
      setCellValueIfColumn(
        sheet,
        rowNum,
        colMap.REPLY_KEYWORD,
        mode === "click" ? "CLICK_TRACKED" : "OPEN_TRACKED",
      );

      var logsSheet = ensureLogsSheet(ss);
      var clientName = colMap.CLIENT_NAME
        ? getCellStr(
            sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0],
            colMap.CLIENT_NAME,
          )
        : "";
      appendLog(
        logsSheet,
        tabName,
        clientName,
        "INFO",
        "Tracking event recorded: " +
          (mode || "open") +
          " | token=" +
          token +
          " | Reply Status set to Replied",
      );
      break; // Found and recorded, no need to check other tabs
    }
  } catch (e) {}
}

/**
 * Finds row number by open-tracking token.
 */
function findRowNumberByToken(sheet, colMap, token, dataStartRow) {
  if (!colMap.OPEN_TOKEN || !token) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return 0;

  var numDataRows = lastRow - dataStartRow + 1;
  var values = sheet
    .getRange(dataStartRow, colMap.OPEN_TOKEN, numDataRows, 1)
    .getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || "").trim() === token) {
      return dataStartRow + i;
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
  var sheetConfig = promptSelectConfiguredSheet(
    ss,
    "Preview Target Dates — Select Sheet",
  );
  if (!sheetConfig) return;
  var sheet = sheetConfig.sheet;
  if (!sheet) {
    Logger.log(
      'Configured sheet "' +
        sheetConfig.sheetName +
        '" not found. Use "Configure Automation Sheet(s)".',
    );
    return;
  }

  var tabName = sheetConfig.sheetName;
  var dataStartRow = getTabDataStartRow(tabName);

  var colMap = buildColumnMap(sheet, tabName);
  var mapError = validateColumnMap(colMap, getTabHeaderRow(tabName));
  if (mapError) {
    Logger.log("Column map error: " + mapError);
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) {
    Logger.log("No data rows.");
    return;
  }

  var numDataRows = lastRow - dataStartRow + 1;
  var numCols = sheet.getLastColumn();
  var data = sheet.getRange(dataStartRow, 1, numDataRows, numCols).getValues();
  var today = getMidnight(new Date());

  Logger.log("Today: " + formatDate(today));
  Logger.log("Tab: " + tabName + " (data starts at row " + dataStartRow + ")");
  Logger.log("---");

  data.forEach(function (row, i) {
    var rowIndex = dataStartRow + i;
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
          noticeStr +
          " | " +
          getSupportedNoticeDateHint(),
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
 * DIAGNOSTIC: Prompts for a sheet tab (if multiple), then a row number,
 * and shows all parsed field values for that row in a dialog — no email is sent.
 */
function diagnosticInspectRow() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetConfig = promptSelectConfiguredSheet(
    ss,
    "Inspect Row — Select Sheet",
  );
  if (!sheetConfig) return;

  var sheet = sheetConfig.sheet;
  var tabName = sheetConfig.sheetName;
  if (!sheet) {
    ui.alert(
      'Sheet "' +
        sheetConfig.sheetName +
        '" not found. Use "Configure Automation Sheet(s)".',
    );
    return;
  }

  var dataStartRow = getTabDataStartRow(tabName);

  var response = ui.prompt(
    "Inspect Row — " + sheetConfig.sheetName,
    "Enter the row number to inspect (data starts at row " +
      dataStartRow +
      "):",
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var rowNum = parseInt(response.getResponseText().trim(), 10);
  if (isNaN(rowNum) || rowNum < dataStartRow) {
    ui.alert("Invalid row number. Data starts at row " + dataStartRow + ".");
    return;
  }

  var colMap = buildColumnMap(sheet, tabName);
  var mapError = validateColumnMap(colMap, getTabHeaderRow(tabName));
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
  var clientEmailRaw = getCellStr(row, colMap.CLIENT_EMAIL);
  var clientEmailList = parseClientEmails(clientEmailRaw);
  var clientEmailDisplay =
    clientEmailList.length > 0 ? clientEmailList.join(", ") : "(empty)";
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
  var finalNoticeSentAt = colMap.FINAL_NOTICE_SENT_AT
    ? row[colMap.FINAL_NOTICE_SENT_AT - 1]
    : "";
  var finalNoticeThreadId = getCellStr(row, colMap.FINAL_NOTICE_THREAD_ID);
  var finalNoticeMessageId = getCellStr(row, colMap.FINAL_NOTICE_MESSAGE_ID);
  var openToken = getCellStr(row, colMap.OPEN_TOKEN);
  var openCount = colMap.OPEN_COUNT ? row[colMap.OPEN_COUNT - 1] : "";
  var effectiveCc = resolveCcEmails(clientEmailList, staffEmail);

  var expiryDate = expiryRaw instanceof Date ? expiryRaw : new Date(expiryRaw);
  var expiryStr = isNaN(expiryDate.getTime())
    ? "INVALID (" + expiryRaw + ")"
    : formatDate(expiryDate);

  var offset = parseNoticeOffset(noticeStr);
  var targetStr =
    offset === null
      ? 'Cannot parse notice: "' +
        noticeStr +
        '". ' +
        getSupportedNoticeDateHint()
      : formatDate(computeTargetDate(expiryDate, offset));

  var today = getMidnight(new Date());
  var sendEligibleNow =
    offset !== null && !isNaN(expiryDate.getTime())
      ? isTargetDateDue(computeTargetDate(expiryDate, offset), today)
        ? "YES"
        : "No"
      : "N/A";
  var finalReminderDueNow =
    !isNaN(expiryDate.getTime()) && isSameDay(expiryDate, today) ? "YES" : "No";
  var firstReminderSent =
    colMap.SENT_AT && row[colMap.SENT_AT - 1] ? "YES" : "No";
  var finalReminderSent = finalNoticeSentAt ? "YES" : "No";

  // Resolve attachments for per-file breakdown
  var attachDisplay = "(none)";
  if (attachRaw) {
    var attachResult = resolveAttachments(attachRaw);
    var lines = [];
    var entries = splitAttachmentEntries(attachRaw);
    for (var ai = 0; ai < entries.length; ai++) {
      var entry = entries[ai];
      var fileId = extractDriveFileId(entry);
      if (!fileId) {
        lines.push("  [" + (ai + 1) + "] UNPARSEABLE: " + entry);
        continue;
      }
      var resolved = false;
      for (var bi = 0; bi < attachResult.blobs.length; bi++) {
        if (attachResult.blobs[bi].fileId === fileId) {
          lines.push(
            "  [" +
              (ai + 1) +
              "] OK: " +
              attachResult.blobs[bi].name +
              " (ID: " +
              fileId +
              ")",
          );
          resolved = true;
          break;
        }
      }
      if (!resolved) {
        lines.push(
          "  [" +
            (ai + 1) +
            "] ERROR: " +
            fileId +
            (attachResult.warnings
              ? " — " + (attachResult.warnings[ai] || "fetch failed")
              : ""),
        );
      }
    }
    attachDisplay = entries.length + " file(s):\n" + lines.join("\n");
  }

  var msg = [
    "Sheet: " + sheetConfig.sheetName,
    "Row: " + rowNum,
    "Status: " + (status || "(empty)"),
    "Send Mode: " +
      sendMode +
      (modeSkipReason ? " (" + modeSkipReason + ")" : ""),
    "Reply Status: " + (replyStatus || "(empty)"),
    "Replied At: " + (repliedAt || "(empty)"),
    "",
    "Client Name:  " + (clientName || "(empty)"),
    "Client Email: " + clientEmailDisplay,
    "Staff Email:  " + (staffEmail || "(empty)"),
    "Effective CC: " + formatCcDisplay(effectiveCc),
    "Doc Type:     " + (docType || "(empty)"),
    "",
    "Expiry Date:  " + expiryStr,
    "Notice Date:  " + (noticeStr || "(empty)"),
    "Target Date:  " + targetStr,
    "Notice Reminder Due Now: " + sendEligibleNow,
    "Final Reminder Due Now:  " + finalReminderDueNow,
    "First Reminder Sent:     " + firstReminderSent,
    "Final Reminder Sent:     " + finalReminderSent,
    "",
    "Sent Thread Id: " + (sentThreadId || "(empty)"),
    "Sent Message Id: " + (sentMessageId || "(empty)"),
    "Final Notice Sent At: " + (finalNoticeSentAt || "(empty)"),
    "Final Notice Thread Id: " + (finalNoticeThreadId || "(empty)"),
    "Final Notice Message Id: " + (finalNoticeMessageId || "(empty)"),
    "Open Token: " + (openToken || "(empty)"),
    "Open Count: " + (openCount || "(empty)"),
    "",
    "Attached Files: " + attachDisplay,
    "",
    "Remarks (first 200 chars):",
    remarks
      ? remarks.substring(0, 200) + (remarks.length > 200 ? "..." : "")
      : "(empty)",
  ].join("\n");

  ui.alert("Row " + rowNum + " Inspection", msg, ui.ButtonSet.OK);
}

/**
 * DIAGNOSTIC: Prompts for a sheet tab (if multiple), then a No. value, finds the matching row,
 * shows a summary of what will be sent,
 * and asks for confirmation before actually sending the test email.
 * Ignores Status and target date — sends regardless.
 */
function diagnosticSendTestRow() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetConfig = promptSelectConfiguredSheet(
    ss,
    "Send Test Email — Select Sheet",
  );
  if (!sheetConfig) return;

  var sheet = sheetConfig.sheet;
  var tabName = sheetConfig.sheetName;
  if (!sheet) {
    ui.alert(
      'Sheet "' +
        sheetConfig.sheetName +
        '" not found. Use "Configure Automation Sheet(s)".',
    );
    return;
  }

  var response = ui.prompt(
    "Send Test Email by No. — " + sheetConfig.sheetName,
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

  var colMap = buildColumnMap(sheet, tabName);
  var tabHeaderRow = getTabHeaderRow(tabName);
  var tabDataStartRow = getTabDataStartRow(tabName);
  var mapError = validateColumnMap(colMap, tabHeaderRow);
  if (mapError) {
    ui.alert("Column map error: " + mapError);
    return;
  }

  var lookup = findRowNumberByNo(sheet, colMap, noValue, tabDataStartRow);
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
  var clientEmailRaw = getCellStr(row, colMap.CLIENT_EMAIL);
  var clientEmailList = parseClientEmails(clientEmailRaw);
  var staffEmail = getCellStr(row, colMap.STAFF_EMAIL);
  var ccEmails = resolveCcEmails(clientEmailList, staffEmail);
  var docType = getCellStr(row, colMap.DOC_TYPE);
  var expiryRaw = colMap.EXPIRY_DATE ? row[colMap.EXPIRY_DATE - 1] : "";
  var noticeStr = getCellStr(row, colMap.NOTICE_DATE);
  var remarks = getCellStr(row, colMap.REMARKS);
  var attachRaw = getCellStr(row, colMap.ATTACHMENTS);
  var firstReminderSent = !!(colMap.SENT_AT && row[colMap.SENT_AT - 1]);
  var finalReminderSent = !!(
    colMap.FINAL_NOTICE_SENT_AT && row[colMap.FINAL_NOTICE_SENT_AT - 1]
  );

  var missing = [];
  if (!clientName) missing.push("Client Name");
  if (clientEmailList.length === 0) missing.push("Client Email");
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
      "Notice:  " +
      (noticeStr || "(empty)") +
      "\n" +
      "To:      " +
      clientEmailList.join(", ") +
      "\n" +
      "CC:      " +
      formatCcDisplay(ccEmails) +
      "\n" +
      "First Sent: " +
      (firstReminderSent ? "YES" : "No") +
      "\n" +
      "Final Sent: " +
      (finalReminderSent ? "YES" : "No") +
      "\n" +
      "Subject: " +
      subject +
      "\n\n" +
      "Proceed?",
    ui.ButtonSet.YES_NO,
  );
  if (confirm !== ui.Button.YES) return;

  var attachResult = resolveAttachments(attachRaw);
  if (attachResult.fatalError) {
    ui.alert("Attachment error: " + attachResult.fatalError);
    return;
  }
  if (attachResult.warnings && attachResult.warnings.length > 0) {
    var warnConfirm = ui.alert(
      "Attachment Warning(s)",
      "Some files could not be loaded:\n\n" +
        attachResult.warnings.join("\n") +
        "\n\nContinue sending with " +
        attachResult.blobs.length +
        " valid file(s)?",
      ui.ButtonSet.YES_NO,
    );
    if (warnConfirm !== ui.Button.YES) return;
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
    var senderEmail = getSenderAccountEmail();
    var displayName = getSenderDisplayName(senderEmail);
    var testFallbackHtml = buildFallbackLinksHtml(attachResult.failedLinks);
    var testHtmlBody = testFallbackHtml
      ? emailContent.htmlBody + testFallbackHtml
      : emailContent.htmlBody;
    var sentMeta = { threadId: "", messageId: "" };
    for (var ti = 0; ti < clientEmailList.length; ti++) {
      var tMeta = sendReminderEmail(
        clientEmailList[ti],
        ccEmails,
        subject,
        testHtmlBody,
        attachResult.blobs,
        displayName,
      );
      if (ti === 0) sentMeta = tMeta;
    }
    setStaffEmail(sheet, rowNum, colMap.STAFF_EMAIL, senderEmail);
    colMap = ensureReplyStatusColumn(sheet, tabName, colMap);
    writePostSendMetadata(sheet, rowNum, colMap, {
      sentAt: new Date(),
      senderEmail: senderEmail,
      openToken: openToken,
      threadId: sentMeta.threadId,
      messageId: sentMeta.messageId,
    });
    if (isSameDay(expiryDate, getMidnight(new Date()))) {
      colMap = ensureFinalNoticeColumns(sheet, tabName, colMap);
      writeFinalNoticeMetadata(sheet, rowNum, colMap, {
        sentAt: new Date(),
        threadId: sentMeta.threadId,
        messageId: sentMeta.messageId,
      });
      setResolvedStatus(sheet, rowNum, colMap, tabName, STATUS.SENT);
    } else {
      setResolvedStatus(sheet, rowNum, colMap, tabName, STATUS.ACTIVE);
    }
    appendLog(
      ensureLogsSheet(ss),
      sheetConfig.sheetName,
      clientName,
      "INFO",
      "Test email sent by No. " +
        rowNo +
        " | To: " +
        clientEmailList.join(", ") +
        (ccEmails.length > 0 ? " | CC: " + ccEmails.join(", ") : "") +
        " | Body: " +
        emailContent.source,
    );
    ui.alert(
      "Test email sent successfully to " +
        clientEmailList.join(", ") +
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
