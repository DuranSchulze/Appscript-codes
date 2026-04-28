
// ═══════════════════════════════════════════════════════════════════════════
// TRACKING — post-send observation (open pixel + reply scan)
// ═══════════════════════════════════════════════════════════════════════════


// ═══════════════════════════════════════════════════════════════════════════
// OpenTracking — tracking URL config, doGet, token lookup, open writes
// ═══════════════════════════════════════════════════════════════════════════

// 41 Open Tracking

function getOpenTrackingBaseUrl() {
  var configured = getPropString(PROP_KEYS.OPEN_TRACKING_BASE_URL, "");
  if (configured) return configured;

  try {
    var serviceUrl = ScriptApp.getService().getUrl();
    if (serviceUrl) return serviceUrl;
  } catch (e) {}

  return "";
}

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

function generateOpenTrackingToken() {
  return Utilities.getUuid();
}

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

function sanitizeHtmlContent(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}


function sanitizeHtmlAttribute(value) {
  return String(value || "").replace(/["<>]/g, "");
}

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
      setCellValueIfColumn(
        sheet,
        rowNum,
        colMap.REPLY_STATUS,
        REPLY_STATUS.REPLIED,
      );
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

// ═══════════════════════════════════════════════════════════════════════════
// ReplyTracking — reply scan orchestration and reply matching
// ═══════════════════════════════════════════════════════════════════════════

// 40 Reply Tracking

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

function runReplyScanNow() {
  var ui = SpreadsheetApp.getUi();
  try {
    var summary = runReplyScan();
    ui.alert("Reply Scan Complete", summary, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("Reply Scan Error", e.message, ui.ButtonSet.OK);
  }
}

function runReplyScan() {
  var ss = getAutomationSpreadsheet();
  var logsSheet = ensureLogsSheet(ss);
  var sheetConfigs = resolveAutomationSheets(ss);
  var keywords = getReplyKeywords();

  resetVerifiedSenderAliasCache();

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

function findReplyMatchForRow(threadId, clientEmail, keywords, sentAt) {
  try {
    var thread = GmailApp.getThreadById(threadId);
    if (!thread) return null;
    var messages = thread.getMessages();
    // Outgoing messages may come from any verified alias (Send mail as), not
    // just the runner account. Filter all of them as "self".
    var ourAddresses = listVerifiedSenderAliases();
    var clientLower = String(clientEmail || "").toLowerCase();

    for (var i = messages.length - 1; i >= 0; i--) {
      var msg = messages[i];
      if (sentAt instanceof Date && !isNaN(sentAt.getTime())) {
        if (msg.getDate().getTime() <= sentAt.getTime()) continue;
      }

      var from = String(msg.getFrom() || "").toLowerCase();
      var isFromUs = false;
      for (var a = 0; a < ourAddresses.length; a++) {
        if (ourAddresses[a] && from.indexOf(ourAddresses[a]) >= 0) {
          isFromUs = true;
          break;
        }
      }
      if (isFromUs) continue;
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
