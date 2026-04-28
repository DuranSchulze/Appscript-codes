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
