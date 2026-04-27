// 23 Sheet Row Ops

function getCellStr(row, colIndex) {
  if (!colIndex) return "";
  return String(row[colIndex - 1] || "").trim();
}

function isStatusActive(statusValue) {
  return (
    String(statusValue || "")
      .trim()
      .toLowerCase() === STATUS.ACTIVE.toLowerCase()
  );
}


function isStatusNoticeSent(statusValue) {
  return (
    String(statusValue || "")
      .trim()
      .toLowerCase() === STATUS.NOTICE_SENT.toLowerCase()
  );
}

function isStatusBlank(statusValue) {
  return String(statusValue || "").trim() === "";
}

function isProcessableStatus(statusValue) {
  return (
    isStatusActive(statusValue) ||
    isStatusNoticeSent(statusValue) ||
    isStatusBlank(statusValue)
  );
}


function isReplyStatusReplied(replyStatusValue) {
  return (
    String(replyStatusValue || "")
      .trim()
      .toLowerCase() === REPLY_STATUS.REPLIED.toLowerCase()
  );
}

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

function setStatus(sheet, rowIndex, statusColIndex, statusValue) {
  if (!statusColIndex) return;
  sheet.getRange(rowIndex, statusColIndex).setValue(statusValue);
}


function getDefaultStatusOptions() {
  return [
    STATUS.ACTIVE,
    STATUS.NOTICE_SENT,
    STATUS.SENT,
    STATUS.ERROR,
    STATUS.SKIPPED,
  ];
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

function setStaffEmail(sheet, rowIndex, staffEmailColIndex, senderEmail) {
  if (!staffEmailColIndex || !senderEmail) return;
  sheet.getRange(rowIndex, staffEmailColIndex).setValue(senderEmail);
}

function setCellValueIfColumn(sheet, rowIndex, colIndex, value) {
  if (!colIndex) return;
  sheet.getRange(rowIndex, colIndex).setValue(value);
}

function clearCellValueIfColumn(sheet, rowIndex, colIndex) {
  if (!colIndex) return;
  sheet.getRange(rowIndex, colIndex).clearContent();
}

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


function ensureOpenTrackingColumns(sheet, tabName, colMap) {
  var keys = ["OPEN_TOKEN", "FIRST_OPENED_AT", "LAST_OPENED_AT", "OPEN_COUNT"];
  var updatedMap = colMap || {};

  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    if (!updatedMap[key]) {
      updatedMap[key] = ensureColumnExists(sheet, tabName, key);
    }
  }

  return buildColumnMap(sheet, tabName);
}


function ensureSetupAutomationColumns(sheet, tabName, colMap) {
  var updatedMap = colMap || buildColumnMap(sheet, tabName);
  var directKeys = [
    "STATUS",
    "STAFF_EMAIL",
    "SEND_MODE",
    "SENT_AT",
    "SENT_THREAD_ID",
    "SENT_MESSAGE_ID",
  ];

  for (var i = 0; i < directKeys.length; i++) {
    var key = directKeys[i];
    if (!updatedMap[key]) {
      updatedMap[key] = ensureColumnExists(sheet, tabName, key);
    }
  }

  updatedMap = buildColumnMap(sheet, tabName);
  updatedMap = ensureReplyMetadataColumns(sheet, tabName, updatedMap);
  updatedMap.REPLY_KEYWORD =
    updatedMap.REPLY_KEYWORD || ensureColumnExists(sheet, tabName, "REPLY_KEYWORD");
  updatedMap = buildColumnMap(sheet, tabName);
  updatedMap = ensureFinalNoticeColumns(sheet, tabName, updatedMap);
  updatedMap = ensureOpenTrackingColumns(sheet, tabName, updatedMap);

  return buildColumnMap(sheet, tabName);
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
