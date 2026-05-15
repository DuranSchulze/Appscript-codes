// ═══════════════════════════════════════════════════════════════════════════
// ORCHESTRATION — runtime loop, scheduling, log sheet
// ═══════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════
// Logs — LOGS sheet bootstrap and append helpers
// ═══════════════════════════════════════════════════════════════════════════

// 43 Logs

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

// Batched log writer — accumulates entries in memory and writes them to the
// LOGS sheet in a single setValues call, then applies action-column colours
// in one batch. Use flush() at the end of each tab's processing block to
// commit. This replaces N individual appendRow calls with 2 API calls per tab.
function createLogBuffer(logsSheet) {
  var pending = [];

  function actionStyle(action) {
    if (
      action === "SENT" ||
      action === "SENT_NOTICE" ||
      action === "SENT_FINAL" ||
      action === "SENT_NOTICE_FINAL"
    ) {
      return { bg: "#d9ead3", fg: "#274e13" };
    }
    if (action === "ERROR") return { bg: "#fce8e6", fg: "#a61c00" };
    if (action === "SKIPPED") return { bg: "#fff2cc", fg: "#7f6000" };
    if (action === "SUMMARY") return { bg: "#e8eaf6", fg: "#1a237e" };
    return { bg: null, fg: null };
  }

  return {
    add: function (tabName, clientName, action, detail) {
      pending.push([
        new Date(),
        tabName || "",
        clientName || "",
        action,
        detail,
      ]);
    },
    flush: function () {
      if (pending.length === 0) return;
      var startRow = logsSheet.getLastRow() + 1;
      logsSheet.getRange(startRow, 1, pending.length, 5).setValues(pending);
      var bgs = [],
        fgs = [];
      for (var i = 0; i < pending.length; i++) {
        var s = actionStyle(pending[i][3]);
        bgs.push([s.bg]);
        fgs.push([s.fg]);
      }
      logsSheet
        .getRange(startRow, LOG_COL.ACTION, pending.length, 1)
        .setBackgrounds(bgs)
        .setFontColors(fgs);
      pending = [];
    },
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// Triggers — trigger install/remove and schedule parsing
// ═══════════════════════════════════════════════════════════════════════════

// 42 Triggers

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

function getDailyTriggerHour() {
  var stored = getPropString(PROP_KEYS.DAILY_TRIGGER_HOUR, "");
  var parsed = parseInt(stored, 10);
  return !isNaN(parsed) && parsed >= 0 && parsed <= 23
    ? parsed
    : CONFIG.TRIGGER_HOUR;
}

function getDailyTriggerMinute() {
  var stored = getPropString(PROP_KEYS.DAILY_TRIGGER_MINUTE, "");
  var parsed = parseInt(stored, 10);
  return !isNaN(parsed) && parsed >= 0 && parsed <= 59 ? parsed : 0;
}

function formatTwoDigits(value) {
  var num = parseInt(value, 10);
  if (isNaN(num) || num < 0) num = 0;
  return num < 10 ? "0" + num : String(num);
}

function formatDailyTriggerTime(hour, minute) {
  return hour + ":" + formatTwoDigits(minute);
}

function parseDailyTriggerTimeInput(rawText) {
  var text = String(rawText || "").trim();
  if (!text) return null;

  if (/^\d{1,2}$/.test(text)) {
    var hourOnly = parseInt(text, 10);
    if (hourOnly >= 0 && hourOnly <= 23) {
      return { hour: hourOnly, minute: 0 };
    }
    return null;
  }

  var match = text.match(/^(\d{1,2}):(\d{1,2})$/);
  if (!match) return null;

  var hour = parseInt(match[1], 10);
  var minute = parseInt(match[2], 10);
  if (hour < 0 || hour > 23 || minute < 0 || minute > 59) return null;

  return { hour: hour, minute: minute };
}

function setDailyTriggerHourDialog() {
  var ui = SpreadsheetApp.getUi();
  var current = getDailyTriggerHour();
  var currentMinute = getDailyTriggerMinute();
  var response = ui.prompt(
    "Set Daily Send Time",
    "Enter the daily email schedule time in 24-hour format.\n\nExamples: 8 = 8:00 AM, 8:30 = 8:30 AM, 14:15 = 2:15 PM\n\nCurrent: " +
      formatDailyTriggerTime(current, currentMinute),
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var parsed = parseDailyTriggerTimeInput(response.getResponseText());
  if (!parsed) {
    ui.alert(
      "Invalid time. Please enter a valid 24-hour time like 8, 8:30, or 14:15.",
    );
    return;
  }
  setPropString(PROP_KEYS.DAILY_TRIGGER_HOUR, String(parsed.hour));
  setPropString(PROP_KEYS.DAILY_TRIGGER_MINUTE, String(parsed.minute));
  ui.alert(
    "Send Time Updated",
    "Daily send time set to " +
      formatDailyTriggerTime(parsed.hour, parsed.minute) +
      " Philippine Time.\n\nClick 'Activate Daily Schedule' to apply the new time.",
    ui.ButtonSet.OK,
  );
}

function installTrigger() {
  rememberSpreadsheetId(SpreadsheetApp.getActiveSpreadsheet());
  removeTrigger();

  var hour = getDailyTriggerHour();
  var minute = getDailyTriggerMinute();

  ScriptApp.newTrigger("runDailyCheck")
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .nearMinute(minute)
    .inTimezone("Asia/Manila")
    .create();

  var msg =
    "Daily schedule activated. runDailyCheck() will run automatically every day at " +
    formatDailyTriggerTime(hour, minute) +
    " Philippine Time (Asia/Manila).\n\nNote: Apps Script time-based triggers run in a scheduled window, so the actual fire time may be near that minute rather than exact to the second.";
  Logger.log(msg);
  try {
    SpreadsheetApp.getUi().alert(
      "Schedule Activated",
      msg,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  } catch (e) {}
}

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

// ═══════════════════════════════════════════════════════════════════════════
// DailyRun — runDailyCheck + manualRunNow orchestration
// ═══════════════════════════════════════════════════════════════════════════

// 30 Run Daily Check

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

function runDailyCheck() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log(
      "runDailyCheck skipped because another automation run is still active.",
    );
    return;
  }
  try {
    return runDailyCheckLocked_();
  } finally {
    lock.releaseLock();
  }
}

function runDailyCheckLocked_() {
  var ss = getAutomationSpreadsheet();
  var logsSheet = ensureLogsSheet(ss);
  var sheetConfigs = resolveAutomationSheets(ss);
  var trackingEnabled = !!getOpenTrackingBaseUrl();
  var today = getMidnight(new Date());

  // Emails are always sent FROM the automation runner's account.
  // Staff emails are CC'd — no "Send mail as" alias needed.
  var senderEmail = getSenderAccountEmail();
  var displayName = getSenderDisplayName(senderEmail);

  // Buffered log writer — one setValues flush per tab instead of N appendRow calls.
  var logBuffer = createLogBuffer(logsSheet);

  logBuffer.add(
    "",
    "",
    "INFO",
    "Automation run started by: " + (senderEmail || "(unknown)") +
      " | Tabs queued: " + sheetConfigs.length,
  );
  logBuffer.flush();

  var totalAllTabs = 0,
    sentAllTabs = 0,
    errorsAllTabs = 0;

  for (var t = 0; t < sheetConfigs.length; t++) {
    var sheetConfig = sheetConfigs[t];
    var tabName = sheetConfig.sheetName;
    var visaSheet = sheetConfig.sheet;

    if (!visaSheet) {
      logBuffer.add(
        tabName,
        "",
        "ERROR",
        'Sheet "' +
          tabName +
          '" not found. Use "Configure Automation Sheet(s)" to fix.',
      );
      logBuffer.flush();
      continue;
    }

    var headerRow = getTabHeaderRow(tabName);
    var dataStartRow = getTabDataStartRow(tabName);

    var colMap = buildColumnMap(visaSheet, tabName);
    colMap = ensureSetupAutomationColumns(visaSheet, tabName, colMap);
    var mapError = validateColumnMap(colMap, headerRow);
    if (mapError) {
      logBuffer.add(tabName, "", "ERROR", "[" + tabName + "] " + mapError);
      logBuffer.flush();
      continue;
    }

    var lastRow = visaSheet.getLastRow();
    if (lastRow < dataStartRow) {
      logBuffer.add(
        tabName,
        "",
        "INFO",
        "[" + tabName + "] No data rows found. Skipping.",
      );
      logBuffer.flush();
      continue;
    }

    var numDataRows = lastRow - dataStartRow + 1;
    var numCols = visaSheet.getLastColumn();
    if (numCols === 0) {
      logBuffer.add(
        tabName,
        "",
        "INFO",
        "[" + tabName + "] Tab has no columns. Skipping.",
      );
      logBuffer.flush();
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
      var templateContext = buildRowTemplateContext(visaSheet, tabName, row);

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
          logBuffer.add(
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
        logBuffer.add(tabName, clientName, "SKIPPED", modeSkipReason);
        skippedMode++;
        continue;
      }

      processed++;

      // Only client name, email, expiry date, and notice date are required.
      // Staff name/email are optional — staffEmail is CC'd when present.
      var missing = [];
      if (!clientName) missing.push("Client Name");
      if (clientEmailList.length === 0) missing.push("Client Email");
      if (!expiryRaw) missing.push("Expiry Date");
      if (!noticeStr) missing.push("Notice Date");
      if (missing.length > 0) {
        var errMsg = "Missing required field(s): " + missing.join(", ");
        setResolvedStatus(visaSheet, rowIndex, colMap, tabName, STATUS.ERROR);
        logBuffer.add(tabName, clientName, "ERROR", errMsg);
        errors++;
        continue;
      }

      var expiryDate =
        expiryRaw instanceof Date ? expiryRaw : new Date(expiryRaw);
      if (isNaN(expiryDate.getTime())) {
        setResolvedStatus(visaSheet, rowIndex, colMap, tabName, STATUS.ERROR);
        logBuffer.add(
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
        logBuffer.add(
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
      var isPastExpiry =
        getMidnight(today).getTime() > getMidnight(expiryDate).getTime();
      var sameDayFinal = isSameDay(targetDate, expiryDate);
      var statusAllowsNotice = isStatusBlank(status) || isStatusActive(status);
      var statusAllowsFinal = statusAllowsNotice || isStatusNoticeSent(status);

      var shouldSendNotice = isNoticeDue && statusAllowsNotice && !isPastExpiry;
      var shouldSendFinal = isExpiryDay && statusAllowsFinal && !isPastExpiry;

      if (isPastExpiry) {
        logBuffer.add(
          tabName,
          clientName,
          "SKIPPED",
          "Expiry date has already passed. No further reminder emails will be sent for this row.",
        );
      }

      if (isReplyStatusReplied(replyStatus)) {
        shouldSendNotice = false;
        if (!shouldSendFinal) {
          logBuffer.add(
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

      // Always extract actual URLs first so links are shown regardless of attachment success
      var attachmentUrls = colMap.ATTACHMENTS
        ? extractAttachmentUrls(visaSheet, rowIndex, colMap.ATTACHMENTS)
        : [];
      var attachResult =
        attachmentUrls.length > 0
          ? resolveAttachments(attachmentUrls.join(","))
          : resolveAttachments(attachRaw);
      if (attachResult.fatalError) {
        setResolvedStatus(visaSheet, rowIndex, colMap, tabName, STATUS.ERROR);
        logBuffer.add(tabName, clientName, "ERROR", attachResult.fatalError);
        errors++;
        continue;
      }
      if (attachResult.warnings && attachResult.warnings.length > 0) {
        logBuffer.add(
          tabName,
          clientName,
          "INFO",
          "Attachment warning(s): " + attachResult.warnings.join("; "),
        );
      }

      // Sender is the automation runner. staffEmail (if any) is merged into CC
      // alongside any configured Default CC emails.
      var ccEmails = resolveCcEmails(clientEmailList, staffEmail);
      if (trackingEnabled) {
        colMap = ensureOpenTrackingColumns(visaSheet, tabName, colMap);
      }
      var baseSubject = buildEmailSubject(
        docType,
        clientName,
        expiryDate,
        tabName,
      );

      function buildLogDetail(sendResult, stage, token, contentSource) {
        return (
          "To: " +
          sendResult.success.join(", ") +
          (ccEmails.length > 0 ? " | CC: " + ccEmails.join(", ") : "") +
          " | From: " +
          senderEmail +
          " | Stage: " +
          stage +
          " | Mode: " +
          sendMode +
          " | Body: " +
          contentSource +
          (attachResult.blobs.length > 0
            ? " | Attachments: " + attachResult.blobs.length
            : "") +
          (token ? " | Tracking: enabled" : "") +
          (sendResult.failed.length > 0
            ? " | Failed: " +
              sendResult.failed
                .map(function (f) {
                  return f.email + " (" + f.error + ")";
                })
                .join("; ")
            : "")
        );
      }

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
            templateContext,
            "notice",
            tabName,
          );
          var noticeSubject = buildStageSubject(baseSubject, "notice");
          var noticeLinksHtml = buildAttachmentLinksHtml(
            attachResult.successLinks,
            attachResult.failedLinks,
          );
          var noticeHtmlBody = noticeLinksHtml
            ? noticeContent.htmlBody + noticeLinksHtml
            : noticeContent.htmlBody;

          var noticeSendResult = sendReminderEmails(
            clientEmailList,
            ccEmails,
            noticeSubject,
            noticeHtmlBody,
            attachResult.blobs,
            displayName,
            senderEmail,
          );
          Utilities.sleep(200);
          var noticeMeta = noticeSendResult.meta;

          colMap = ensureReplyStatusColumn(visaSheet, tabName, colMap);
          writePostSendMetadata(visaSheet, rowIndex, colMap, {
            sentAt: new Date(),
            senderEmail: senderEmail,
            openToken: noticeToken,
            threadId: noticeMeta.threadId,
            messageId: noticeMeta.messageId,
          });
          applyGmailLabelToThread(tabName, noticeMeta.threadId);

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
              STATUS.NOTICE_SENT,
            );
          }

          logBuffer.add(
            tabName,
            clientName,
            sameDayFinal ? "SENT_NOTICE_FINAL" : "SENT_NOTICE",
            buildLogDetail(
              noticeSendResult,
              sameDayFinal ? "notice+final" : "notice",
              noticeToken,
              noticeContent.source,
            ),
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
            templateContext,
            "final",
            tabName,
          );
          var finalSubject = buildStageSubject(baseSubject, "final");
          var finalLinksHtml = buildAttachmentLinksHtml(
            attachResult.successLinks,
            attachResult.failedLinks,
          );
          var finalHtmlBody = finalLinksHtml
            ? finalContent.htmlBody + finalLinksHtml
            : finalContent.htmlBody;

          var finalSendResult = sendReminderEmails(
            clientEmailList,
            ccEmails,
            finalSubject,
            finalHtmlBody,
            attachResult.blobs,
            displayName,
            senderEmail,
          );
          Utilities.sleep(200);
          var finalMeta = finalSendResult.meta;

          writeFinalNoticeMetadata(visaSheet, rowIndex, colMap, {
            sentAt: new Date(),
            threadId: finalMeta.threadId,
            messageId: finalMeta.messageId,
          });
          applyGmailLabelToThread(tabName, finalMeta.threadId);

          if (!isStatusNoticeSent(status) && !shouldSendNotice) {
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
          logBuffer.add(
            tabName,
            clientName,
            "SENT_FINAL",
            buildLogDetail(
              finalSendResult,
              "final",
              finalToken,
              finalContent.source,
            ),
          );
          sent++;
          sentThisRow = true;
        }

        if (!sentThisRow) skippedFuture++;
      } catch (e) {
        setResolvedStatus(visaSheet, rowIndex, colMap, tabName, STATUS.ERROR);
        logBuffer.add(
          tabName,
          clientName,
          "ERROR",
          "Send failed: " + e.message,
        );
        errors++;
      }
    }

    logBuffer.add(
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
    logBuffer.flush();

    totalAllTabs += totalRows;
    sentAllTabs += sent;
    errorsAllTabs += errors;
  }

  if (sheetConfigs.length > 1) {
    logBuffer.add(
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
    logBuffer.flush();
  }
}
