
// ═══════════════════════════════════════════════════════════════════════════
// MENU — entry point UI and read-only info dialogs
// ═══════════════════════════════════════════════════════════════════════════


// ═══════════════════════════════════════════════════════════════════════════
// Menu — onOpen + spreadsheet menu spec
// ═══════════════════════════════════════════════════════════════════════════

var ROOT_MENU_NAME = "🔔 Expiry Notifications";

var ROOT_MENU_SPEC = [
  { type: "item", label: "🚀 Setup This Sheet for Automation", fn: "runSetupWizard" },
  { type: "separator" },
  {
    type: "submenu",
    label: "Status & Overview",
    items: [
      { type: "item", label: "Show System Status", fn: "showSystemStatus" },
      { type: "item", label: "View Last Run Summary", fn: "checkScheduleStatus" },
      { type: "item", label: "Show All Tabs Info", fn: "showAllTabsInfo" },
    ],
  },
  { type: "separator" },
  {
    type: "submenu",
    label: "Tab Management",
    items: [
      { type: "item", label: "Configure Automation Tabs (multi)", fn: "configureAutomationSheets" },
      { type: "item", label: "Select Working Tab(s)", fn: "selectWorkingTab" },
      { type: "item", label: "Setup Tab Dropdowns (multi)", fn: "setupSheetDropdowns" },
      { type: "item", label: "Map Tab Columns (multi)", fn: "mapTabColumns" },
      { type: "item", label: "Set Tab Header Row (multi)", fn: "setTabHeaderRowDialog" },
    ],
  },
  { type: "separator" },
  {
    type: "submenu",
    label: "Automation Settings",
    items: [
      { type: "item", label: "Activate Daily Schedule", fn: "installTrigger" },
      { type: "item", label: "Set Send Time", fn: "setDailyTriggerHourDialog" },
      { type: "item", label: "Deactivate Daily Schedule", fn: "removeTrigger" },
      { type: "item", label: "Set Default CC Emails", fn: "setDefaultCcEmailsDialog" },
      { type: "separator" },
      { type: "item", label: "Set Reply Keywords", fn: "setReplyKeywords" },
      {
        type: "item",
        label: "Activate Reply Scan (2x Daily)",
        fn: "installReplyScanTrigger",
      },
      { type: "item", label: "Deactivate Reply Scan", fn: "removeReplyScanTrigger" },
      { type: "separator" },
      { type: "item", label: "Set Open Tracking URL", fn: "setOpenTrackingBaseUrl" },
    ],
  },
  { type: "separator" },
  {
    type: "submenu",
    label: "Run & Diagnostics",
    items: [
      { type: "item", label: "Run Manual Check Now", fn: "manualRunNow" },
      { type: "item", label: "Preview Target Dates", fn: "previewTargetDates" },
      { type: "separator" },
      { type: "item", label: "Inspect Row...", fn: "diagnosticInspectRow" },
      { type: "item", label: "Send Test Email by No....", fn: "diagnosticSendTestRow" },
      { type: "separator" },
      { type: "item", label: "Test Gmail Send", fn: "testGmailSend" },
      { type: "item", label: "Test Drive Access", fn: "testDriveAccess" },
      { type: "item", label: "Test All Connections", fn: "testAllConnections" },
      { type: "separator" },
      { type: "item", label: "Check Column Mappings", fn: "checkColumnMappings" },
      { type: "item", label: "Validate Tab Structure", fn: "validateActiveTabStructure" },
      { type: "item", label: "View Tab Configuration", fn: "viewTabConfiguration" },
      { type: "item", label: "Check Reply Tracking Setup", fn: "checkReplyTrackingSetup" },
      {
        type: "item",
        label: "Preview Fallback Template",
        fn: "previewFallbackTemplateBody",
      },
      { type: "separator" },
      { type: "item", label: "System Diagnostics", fn: "runSystemDiagnostics" },
    ],
  },
  { type: "separator" },
  {
    type: "submenu",
    label: "Help",
    items: [
      {
        type: "item",
        label: "Want to integrate this to another google sheet?",
        fn: "showIntegrationLinkDialog",
      },
      { type: "separator" },
      { type: "item", label: "View Documentation", fn: "viewDocumentation" },
      { type: "item", label: "About", fn: "showAbout" },
    ],
  },
];

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  rememberSpreadsheetId(SpreadsheetApp.getActiveSpreadsheet());
  buildRootMenu_(ui, ROOT_MENU_NAME, ROOT_MENU_SPEC).addToUi();
}

function buildRootMenu_(ui, menuName, spec) {
  var menu = ui.createMenu(menuName);
  applyMenuSpec_(ui, menu, spec);
  return menu;
}

function applyMenuSpec_(ui, menu, spec) {
  for (var i = 0; i < spec.length; i++) {
    var item = spec[i];
    if (!item || !item.type) continue;

    if (item.type === "separator") {
      menu.addSeparator();
    } else if (item.type === "item") {
      menu.addItem(item.label, item.fn);
    } else if (item.type === "submenu") {
      var subMenu = ui.createMenu(item.label);
      applyMenuSpec_(ui, subMenu, item.items || []);
      menu.addSubMenu(subMenu);
    }
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// StatusUi — status / docs / about / integration link dialogs
// ═══════════════════════════════════════════════════════════════════════════

// 11 Status Ui

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

function checkScheduleStatus() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logsSheet = ensureLogsSheet(ss);
  var active = getTriggersByHandler("runDailyCheck");
  var replyTriggers = getTriggersByHandler("runReplyScan");

  var configuredHour = getDailyTriggerHour();
  var configuredMinute = getDailyTriggerMinute();
  var configuredTime = formatDailyTriggerTime(configuredHour, configuredMinute);
  var msg;
  if (active.length === 0) {
    msg =
      "Status: INACTIVE\n\nNo daily schedule is set up.\nConfigured send time: " +
      configuredTime +
      "\nUse 'Activate Daily Schedule' to enable it.";
  } else {
    var lines = [
      "Status: ACTIVE",
      "",
      "Send time: " + configuredTime + " Philippine Time (Asia/Manila)",
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
      "- [Any Other Column Header]\n" +
      "- Put these directly in the Remarks column to auto-fill row data\n" +
      "- Example custom fields: [Company], [Issued Date], [ID No. /Document No.]\n" +
      "- Example: Good day [Client Name], your [Document Type] expires on [Expiry Date].\n\n" +
      "Reminder Rules:\n" +
      "- First reminder sends on the computed Notice Date target\n" +
      "- After the first reminder, Status becomes Notice Sent to block duplicate notice emails\n" +
      "- Final reminder sends on the exact Expiry Date and then marks Status as Sent\n" +
      "- No reminder emails are sent after the Expiry Date has passed\n" +
      "- A row is fully complete only after the final reminder is sent\n\n" +
      "Setup Behavior:\n" +
      "- Setup now creates or verifies automation-managed columns on existing tabs\n" +
      "- LOGS is the run-history page for sends, skips, and errors\n\n" +
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
      "Status drives the send flow:\n" +
      "Active = first notice can send\n" +
      "Notice Sent = final expiry-day reminder only\n" +
      "Sent = fully completed\n\n" +
      "Setup also prepares automation-managed columns on the selected tab.\n" +
      "LOGS remains the run-history sheet.\n\n" +
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
      "[Date of Expiration], [Date of Expiry],\n" +
      "or any other sheet header like [Company]\n\n" +
      "Example:\n" +
      "Good day [Client Name], your [Document Type]\n" +
      "expires on [Expiry Date].\n\n" +
      "Rows stay Active until the final expiry-day reminder\n" +
      "is sent, then they are marked Sent.",
    ui.ButtonSet.OK,
  );
}

function showAutomationStatus() {
  showSystemStatus();
}
