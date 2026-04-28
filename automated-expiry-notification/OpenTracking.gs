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
