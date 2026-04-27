// 14 Dropdowns

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
        [
          STATUS.ACTIVE,
          STATUS.NOTICE_SENT,
          STATUS.SENT,
          STATUS.ERROR,
          STATUS.SKIPPED,
        ],
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
