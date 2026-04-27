// 13 Tab Mapping

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

function buildAvailableHeaderChoices(headerValues) {
  var choices = [];
  for (var i = 0; i < headerValues.length; i++) {
    var text = String(headerValues[i] || "").trim();
    if (!text) continue;
    choices.push({ index: i + 1, header: text });
  }
  return choices;
}

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
