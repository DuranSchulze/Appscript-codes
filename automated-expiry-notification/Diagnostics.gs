// 44 Diagnostics

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

function testRunNow() {
  Logger.log("=== testRunNow: calling runDailyCheck() ===");
  runDailyCheck();
  Logger.log("=== testRunNow: complete. Check LOGS sheet and your inbox. ===");
}

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
  var pastExpiry =
    !isNaN(expiryDate.getTime()) &&
    getMidnight(today).getTime() > getMidnight(expiryDate).getTime()
      ? "YES"
      : "No";
  var firstReminderSent =
    colMap.SENT_AT && row[colMap.SENT_AT - 1] ? "YES" : "No";
  var finalReminderSent = finalNoticeSentAt ? "YES" : "No";
  var stageEligibility = [
    "Notice eligible now: " +
      (offset !== null &&
      !isNaN(expiryDate.getTime()) &&
      isTargetDateDue(computeTargetDate(expiryDate, offset), today) &&
      (isStatusBlank(status) || isStatusActive(status))
        ? "YES"
        : "No"),
    "Final eligible now:  " +
      (!isNaN(expiryDate.getTime()) &&
      isSameDay(expiryDate, today) &&
      (isStatusBlank(status) ||
        isStatusActive(status) ||
        isStatusNoticeSent(status))
        ? "YES"
        : "No"),
  ];

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
    "Past Expiry Date:        " + pastExpiry,
    stageEligibility[0],
    stageEligibility[1],
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

  var trackingEnabled = !!getOpenTrackingBaseUrl();
  if (trackingEnabled) {
    colMap = ensureOpenTrackingColumns(sheet, tabName, colMap);
  }
  var openToken = trackingEnabled ? generateOpenTrackingToken() : "";
  var templateContext = buildRowTemplateContext(sheet, tabName, row);
  var emailContent = buildEmailContent(
    remarks,
    clientName,
    expiryDate,
    docType,
    openToken,
    templateContext,
  );

  try {
    var senderEmail = getSenderAccountEmail();
    var displayName = getSenderDisplayName(senderEmail);
    var testFallbackHtml = buildFallbackLinksHtml(attachResult.failedLinks);
    var testHtmlBody = testFallbackHtml
      ? emailContent.htmlBody + testFallbackHtml
      : emailContent.htmlBody;
    var testSendResult = sendReminderEmails(
      clientEmailList,
      ccEmails,
      subject,
      testHtmlBody,
      attachResult.blobs,
      displayName,
    );
    var sentMeta = testSendResult.meta;
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
        testSendResult.success.join(", ") +
        (ccEmails.length > 0 ? " | CC: " + ccEmails.join(", ") : "") +
        (testSendResult.failed.length > 0
          ? " | Failed Recipients: " +
            testSendResult.failed
              .map(function (item) {
                return item.email + " (" + item.error + ")";
              })
              .join("; ")
          : "") +
        " | Body: " +
        emailContent.source,
    );
    ui.alert(
      "Test email sent successfully to " +
        testSendResult.success.join(", ") +
        "." +
        (testSendResult.failed.length > 0
          ? "\n\nFailed recipients:\n" +
            testSendResult.failed
              .map(function (item) {
                return item.email + " - " + item.error;
              })
              .join("\n")
          : "") +
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

// ═══════════════════════════════════════════════════════════════════════════
// HealthCheck — verify daily trigger + staff Send-As alias coverage
// ═══════════════════════════════════════════════════════════════════════════

function runHealthCheck() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logsSheet = ensureLogsSheet(ss);

  var lines = ["HEALTH CHECK"];
  var nowStr = Utilities.formatDate(
    new Date(),
    "Asia/Manila",
    "yyyy-MM-dd HH:mm 'PHT'",
  );
  lines.push(nowStr);
  lines.push("");

  var failCount = 0;
  var warnCount = 0;

  // 1. Daily send trigger
  var dailyTriggers = getTriggersByHandler("runDailyCheck");
  var configuredTime = formatDailyTriggerTime(
    getDailyTriggerHour(),
    getDailyTriggerMinute(),
  );
  if (dailyTriggers.length === 1) {
    lines.push(
      "[PASS] Daily send trigger installed (" + configuredTime + " PHT)",
    );
  } else if (dailyTriggers.length === 0) {
    lines.push(
      "[FAIL] No daily send trigger installed. " +
        "Run: Automation Settings → Activate Daily Schedule.",
    );
    failCount++;
  } else {
    lines.push(
      "[FAIL] " +
        dailyTriggers.length +
        " duplicate daily triggers detected. " +
        "Run: Deactivate Daily Schedule, then Activate Daily Schedule.",
    );
    failCount++;
  }

  // 2. Reply scan trigger (informational)
  var replyTriggers = getTriggersByHandler("runReplyScan");
  if (replyTriggers.length === 2) {
    lines.push("[INFO] Reply scan triggers installed (9:00, 15:00 PHT)");
  } else if (replyTriggers.length === 0) {
    lines.push(
      "[INFO] Reply scan not installed (optional). " +
        "Use: Activate Reply Scan (2x Daily).",
    );
  } else {
    lines.push(
      "[WARN] Reply scan trigger count is " +
        replyTriggers.length +
        " (expected 2). Re-activate to clean up.",
    );
    warnCount++;
  }

  lines.push("");

  // 3. Assigned Staff Email is verified Send-As alias on runner account
  resetVerifiedSenderAliasCache();
  var runnerEmail = getSenderAccountEmail() || "(unknown)";
  lines.push("Script runner account: " + runnerEmail);

  var sheetConfigs = resolveAutomationSheets(ss);
  var seenEmails = {}; // normalized email -> { tab, row }
  var unverified = [];
  var totalRowsChecked = 0;
  var tabsScanned = 0;

  for (var i = 0; i < sheetConfigs.length; i++) {
    var config = sheetConfigs[i];
    var sheet = config.sheet;
    var tabName = config.sheetName;
    if (!sheet) continue;

    var colMap = buildColumnMap(sheet, tabName);
    if (!colMap.STAFF_EMAIL) {
      lines.push(
        "[WARN] Tab '" + tabName + "' has no Assigned Staff Email column.",
      );
      warnCount++;
      continue;
    }

    var dataStartRow = getTabDataStartRow(tabName);
    var lastRow = sheet.getLastRow();
    if (lastRow < dataStartRow) {
      tabsScanned++;
      continue;
    }

    var values = sheet
      .getRange(dataStartRow, colMap.STAFF_EMAIL, lastRow - dataStartRow + 1, 1)
      .getValues();

    tabsScanned++;
    for (var r = 0; r < values.length; r++) {
      var raw = String(values[r][0] || "").trim();
      if (!raw) continue;
      totalRowsChecked++;
      var normalized = raw.toLowerCase();
      if (seenEmails[normalized]) continue;
      seenEmails[normalized] = { tab: tabName, row: dataStartRow + r };

      if (!canSendAs(raw)) {
        unverified.push({
          email: raw,
          tab: tabName,
          row: dataStartRow + r,
        });
      }
    }
  }

  var uniqueCount = Object.keys(seenEmails).length;
  lines.push(
    "Scanned " +
      tabsScanned +
      " tab(s), " +
      totalRowsChecked +
      " staff-email cell(s), " +
      uniqueCount +
      " unique address(es).",
  );

  if (uniqueCount === 0) {
    lines.push(
      "[INFO] No Assigned Staff Email values found in the configured tabs.",
    );
  } else if (unverified.length === 0) {
    lines.push(
      "[PASS] All " +
        uniqueCount +
        " staff email(s) are verified Send-As aliases.",
    );
  } else {
    lines.push(
      "[FAIL] " +
        unverified.length +
        " of " +
        uniqueCount +
        " staff email(s) are NOT verified Send-As aliases:",
    );
    for (var u = 0; u < unverified.length; u++) {
      var item = unverified[u];
      lines.push(
        "       - " +
          item.email +
          "  (tab: " +
          item.tab +
          ", first seen row " +
          item.row +
          ")",
      );
    }
    lines.push("");
    lines.push("Fix: on the script-runner account (" + runnerEmail + "),");
    lines.push(
      "     Gmail → Settings → Accounts → Send mail as → Add another email address.",
    );
    failCount++;
  }

  lines.push("");
  if (failCount === 0 && warnCount === 0) {
    lines.push("Result: ALL CHECKS PASSED");
  } else {
    lines.push(
      "Result: " + failCount + " failure(s), " + warnCount + " warning(s)",
    );
  }

  var report = lines.join("\n");
  try {
    appendLog(
      logsSheet,
      "",
      "",
      failCount > 0 ? "ERROR" : "INFO",
      "Health Check: " +
        (failCount === 0 && warnCount === 0
          ? "all checks passed"
          : failCount + " fail / " + warnCount + " warn"),
    );
  } catch (e) {}

  ui.alert("Health Check", report.substring(0, 1800), ui.ButtonSet.OK);
}
