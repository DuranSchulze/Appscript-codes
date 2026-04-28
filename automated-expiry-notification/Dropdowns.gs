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

  var selectedNames = promptSelectTabs({
    ss: ss,
    title: "Setup Tab Dropdowns",
    source: "configured",
    prompt: "Pick one or more tabs. The same Notice Date options are applied to each.",
  });
  if (selectedNames.length === 0) return;

  // One Notice Date prompt covers the whole batch.
  var firstTab = selectedNames[0];
  var currentOptions = getNoticeOptionsForTab(firstTab);
  var optionResponse = ui.prompt(
    "Notice Date Options",
    "Comma-separated Notice Date options to apply to: " + selectedNames.join(", ") +
      "\n\nLeave blank to keep each tab's existing options (defaults used for unset tabs).\n\n" +
      "Reference (from " + firstTab + "):\n" + currentOptions.join(", "),
    ui.ButtonSet.OK_CANCEL,
  );
  if (optionResponse.getSelectedButton() !== ui.Button.OK) return;

  var customInput = optionResponse.getResponseText().trim();
  var sharedNoticeOptions = null;
  if (customInput) {
    sharedNoticeOptions = customInput
      .split(",")
      .map(function (s) { return s.trim(); })
      .filter(function (s) { return !!s; });
    var invalid = findInvalidNoticeOptions(sharedNoticeOptions);
    if (invalid.length > 0) {
      ui.alert(
        "Invalid Notice Date Options",
        "These options are not parseable:\n- " + invalid.join("\n- ") + "\n\n" +
          getSupportedNoticeDateHint(),
        ui.ButtonSet.OK,
      );
      return;
    }
  }

  var summaries = [];
  for (var s = 0; s < selectedNames.length; s++) {
    var tabName = selectedNames[s];
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      summaries.push("✗ " + tabName + " — not found");
      continue;
    }
    summaries.push(applyDropdownsForTab(sheet, tabName, sharedNoticeOptions));
  }

  ui.alert("Dropdowns — Summary", summaries.join("\n"), ui.ButtonSet.OK);
}

function applyDropdownsForTab(sheet, tabName, sharedNoticeOptions) {
  var dataStartRow = getTabDataStartRow(tabName);
  var colMap = buildColumnMap(sheet, tabName);
  var mapError = validateColumnMap(colMap, getTabHeaderRow(tabName));
  if (mapError) return "⚠ " + tabName + " — " + mapError;

  var noticeOptions = sharedNoticeOptions;
  if (noticeOptions && noticeOptions.length > 0) {
    setNoticeOptionsForTab(tabName, noticeOptions);
  } else {
    noticeOptions = getNoticeOptionsForTab(tabName);
  }

  var lastRow = sheet.getLastRow();
  var dataLastRow = Math.max(lastRow, dataStartRow + 100);
  var applied = [];

  if (colMap.STATUS) {
    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(
        [STATUS.ACTIVE, STATUS.NOTICE_SENT, STATUS.SENT, STATUS.ERROR, STATUS.SKIPPED],
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
      .getRange(dataStartRow, colMap.SEND_MODE, dataLastRow - dataStartRow + 1, 1)
      .setDataValidation(sendModeRule);
    applied.push("Send Mode");
  }

  if (colMap.NOTICE_DATE) {
    var noticeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(noticeOptions, true)
      .setAllowInvalid(true)
      .build();
    sheet
      .getRange(dataStartRow, colMap.NOTICE_DATE, dataLastRow - dataStartRow + 1, 1)
      .setDataValidation(noticeRule);
    applied.push("Notice Date (" + noticeOptions.length + ")");
  }

  return applied.length === 0
    ? "⚠ " + tabName + " — no matching dropdown columns"
    : "✓ " + tabName + " — " + applied.join(", ");
}
