
// ═══════════════════════════════════════════════════════════════════════════
// SETUP — everything the user clicks while configuring a sheet
// ═══════════════════════════════════════════════════════════════════════════


// ═══════════════════════════════════════════════════════════════════════════
// UiPrompts — multi-tab selection helper used by all setup menu items
// ═══════════════════════════════════════════════════════════════════════════

// 91 UI Prompts — multi-tab selection helper
//
// Most setup menu items operate per-tab. This helper lets the user pick
// many tabs at once with the same comma-list syntax already used by
// "Configure Automation Tabs" — e.g. "1,3,4" or "1-3" or names. Menu
// items then loop their action across the selection.

// Returns an array of resolved sheet names, or [] if the user cancelled.
// Options:
//   ss              SpreadsheetApp instance (required)
//   title           Dialog title (required)
//   prompt          Prompt body shown above the tab list
//   source          "configured" → only registered automation tabs
//                   "all"        → every tab in the spreadsheet
function promptSelectTabs(options) {
  var ui = SpreadsheetApp.getUi();
  var ss = options.ss;
  var sourceMode = options.source === "all" ? "all" : "configured";

  var entries = (sourceMode === "all")
    ? ss.getSheets().map(function (sh) {
        return { sheet: sh, sheetName: sh.getName() };
      })
    : resolveAutomationSheets(ss);

  if (!entries || entries.length === 0) {
    ui.alert(
      "No Tabs Available",
      sourceMode === "configured"
        ? "No tabs are configured for automation. Use 'Configure Automation Tabs' first."
        : "This spreadsheet has no tabs.",
      ui.ButtonSet.OK,
    );
    return [];
  }

  var optionLines = entries.map(function (cfg, i) {
    return (i + 1) + ". " + cfg.sheetName;
  });

  var promptBody =
    (options.prompt || "Enter one or more tab numbers, separated by commas. Ranges OK (e.g. 1-3).") +
    "\n\nAvailable tabs:\n" + optionLines.join("\n") +
    "\n\nLeave blank to cancel.";

  var response = ui.prompt(options.title || "Select Tab(s)", promptBody, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return [];

  var raw = String(response.getResponseText() || "").trim();
  if (!raw) return [];

  var parsed = parseTabSelectionInput(raw, entries);
  if (parsed.errors.length > 0 && parsed.selected.length === 0) {
    ui.alert(
      "Invalid Selection",
      "Could not resolve: " + parsed.errors.join(", ") +
        "\n\nValid range is 1-" + entries.length + ".",
      ui.ButtonSet.OK,
    );
    return [];
  }

  if (parsed.errors.length > 0) {
    ui.alert(
      "Some Selections Skipped",
      "Could not resolve: " + parsed.errors.join(", ") +
        "\n\nProceeding with: " + parsed.selected.map(function (e) { return e.sheetName; }).join(", "),
      ui.ButtonSet.OK,
    );
  }

  return parsed.selected.map(function (e) { return e.sheetName; });
}

// Parses "1,3,5-7,Some Tab Name" against a list of {sheetName} entries.
// Returns { selected: [entry,...], errors: [token,...] }, de-duplicated.
function parseTabSelectionInput(input, entries) {
  var tokens = String(input || "").split(",");
  var seen = {};
  var selected = [];
  var errors = [];

  function pushEntry(entry) {
    if (!entry) return;
    var key = entry.sheetName;
    if (seen[key]) return;
    seen[key] = true;
    selected.push(entry);
  }

  for (var i = 0; i < tokens.length; i++) {
    var token = String(tokens[i] || "").trim();
    if (!token) continue;

    // Range form: "1-3"
    var rangeMatch = token.match(/^(\d+)\s*-\s*(\d+)$/);
    if (rangeMatch) {
      var lo = parseInt(rangeMatch[1], 10);
      var hi = parseInt(rangeMatch[2], 10);
      if (lo > hi) { var tmp = lo; lo = hi; hi = tmp; }
      var anyValid = false;
      for (var n = lo; n <= hi; n++) {
        if (n >= 1 && n <= entries.length) {
          pushEntry(entries[n - 1]);
          anyValid = true;
        }
      }
      if (!anyValid) errors.push(token);
      continue;
    }

    // Plain number
    var asNum = parseInt(token, 10);
    if (!isNaN(asNum) && String(asNum) === token) {
      if (asNum >= 1 && asNum <= entries.length) {
        pushEntry(entries[asNum - 1]);
      } else {
        errors.push(token);
      }
      continue;
    }

    // Name match (case-insensitive)
    var matched = null;
    for (var e = 0; e < entries.length; e++) {
      if (entries[e].sheetName.toLowerCase() === token.toLowerCase()) {
        matched = entries[e];
        break;
      }
    }
    if (matched) pushEntry(matched);
    else errors.push('"' + token + '"');
  }

  return { selected: selected, errors: errors };
}

// ═══════════════════════════════════════════════════════════════════════════
// TabManagement — configure + select automation tabs
// ═══════════════════════════════════════════════════════════════════════════

// 12 Tab Management

function selectWorkingTab() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var selectedNames = promptSelectTabs({
    ss: ss,
    title: "Select Working Tab(s)",
    source: "configured",
    prompt:
      "Pick one or more configured tabs. The first becomes the default 'working tab' " +
      "for diagnostics; the rest stay registered and continue to be processed by the daily run.\n" +
      "Enter numbers (e.g. 1,3) or a range (1-3).",
  });
  if (selectedNames.length === 0) return;

  var primary = selectedNames[0];
  setPropString(
    getTabConfigKey("_GLOBAL", TAB_CONFIG_KEYS.LAST_SELECTED),
    primary,
  );

  ui.alert(
    "Working Tab Set",
    selectedNames.length === 1
      ? '"' + primary + '" is now your working tab.'
      : '"' + primary + '" is now the primary working tab.\n\nAlso selected: ' +
          selectedNames.slice(1).join(", "),
    ui.ButtonSet.OK,
  );
}

function getWorkingTabConfig(ss) {
  var lastSelected = getPropString(
    getTabConfigKey("_GLOBAL", TAB_CONFIG_KEYS.LAST_SELECTED),
    "",
  );
  var sheetConfigs = resolveAutomationSheets(ss);

  if (lastSelected && sheetConfigs.length > 0) {
    for (var i = 0; i < sheetConfigs.length; i++) {
      if (sheetConfigs[i].sheetName === lastSelected) {
        return sheetConfigs[i];
      }
    }
  }

  // If only one tab, use it
  if (sheetConfigs.length === 1) {
    return sheetConfigs[0];
  }

  // Otherwise prompt
  return promptSelectConfiguredSheet(ss, "Select Working Tab");
}

function initializeAutomationSheet() {
  configureAutomationSheets();
}

function configureAutomationSheets() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Show ALL tabs — user decides what to include
  var sheets = ss.getSheets();

  if (sheets.length === 0) {
    ui.alert("Configure Sheets", "No sheet tabs were found.", ui.ButtonSet.OK);
    return;
  }

  // Resolve configured tabs — this auto-purges any deleted tabs from storage
  var liveConfigs = resolveAutomationSheets(ss);
  var liveIds = liveConfigs.map(function (c) {
    return c.sheet ? c.sheet.getSheetId() : null;
  });
  var liveNames = liveConfigs.map(function (c) {
    return c.sheetName;
  });

  var options = [];
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var isActive =
      liveIds.indexOf(sh.getSheetId()) >= 0 ||
      liveNames.indexOf(sh.getName()) >= 0;
    options.push(i + 1 + ". " + sh.getName() + (isActive ? " ★ [active]" : ""));
  }

  var currentLabel = liveNames.length > 0 ? liveNames.join(", ") : "(none)";

  var response = ui.prompt(
    "Configure Automation Sheet(s)",
    "Enter tab numbers separated by commas. You can select multiple.\n" +
      "Example: 1, 3\n\n" +
      "Currently active: " +
      currentLabel +
      "\n\n" +
      options.join("\n"),
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var input = response.getResponseText().trim();
  if (!input) {
    ui.alert("No input provided. Configuration unchanged.");
    return;
  }

  var parts = input.split(",");
  var selectedEntries = [];
  var errors = [];

  for (var p = 0; p < parts.length; p++) {
    var part = parts[p].trim();
    if (!part) continue;

    var idx = parseInt(part, 10);
    var resolvedSheet = null;

    if (!isNaN(idx) && idx >= 1 && idx <= sheets.length) {
      // Number input — index into the shown list
      resolvedSheet = sheets[idx - 1];
    } else {
      // Name input — case-insensitive match
      for (var s = 0; s < sheets.length; s++) {
        if (sheets[s].getName().toLowerCase() === part.toLowerCase()) {
          resolvedSheet = sheets[s];
          break;
        }
      }
      if (!resolvedSheet) errors.push('"' + part + '"');
    }
    // Validate the resolved sheet actually exists
    if (resolvedSheet && !ss.getSheetByName(resolvedSheet.getName())) {
      errors.push('"' + resolvedSheet.getName() + '" (deleted)');
      resolvedSheet = null;
    }

    if (resolvedSheet) {
      var alreadyAdded = false;
      for (var x = 0; x < selectedEntries.length; x++) {
        if (selectedEntries[x].id === resolvedSheet.getSheetId()) {
          alreadyAdded = true;
          break;
        }
      }
      if (!alreadyAdded) {
        selectedEntries.push({
          id: resolvedSheet.getSheetId(),
          name: resolvedSheet.getName(),
        });
      }
    }
  }

  if (selectedEntries.length === 0) {
    var errMsg =
      errors.length > 0
        ? "Could not find tab(s): " +
          errors.join(", ") +
          "\n\nPlease check the numbers and try again."
        : "No valid tabs selected. Configuration unchanged.";
    ui.alert("Nothing Saved", errMsg, ui.ButtonSet.OK);
    return;
  }

  setConfiguredTabEntries(selectedEntries);

  var summary = selectedEntries
    .map(function (e, i) {
      return i + 1 + ". " + e.name + " (ID: " + e.id + ")";
    })
    .join("\n");

  var warningLine =
    errors.length > 0 ? "\n\n⚠ Not found (skipped): " + errors.join(", ") : "";

  ui.alert(
    "Configuration Saved",
    "Automation will now process " +
      selectedEntries.length +
      " tab(s):\n\n" +
      summary +
      warningLine,
    ui.ButtonSet.OK,
  );
}

function promptSelectConfiguredSheet(ss, title) {
  var ui = SpreadsheetApp.getUi();

  // resolveAutomationSheets auto-purges deleted tabs; only live tabs returned
  var configs = resolveAutomationSheets(ss);

  if (configs.length === 0) {
    ui.alert(
      "No Tabs Found",
      "No configured tabs exist. Use 'Configure Automation Tabs' to register tabs.",
      ui.ButtonSet.OK,
    );
    return null;
  }

  if (configs.length === 1) {
    return configs[0];
  }

  var lastSelected = getPropString(
    getTabConfigKey("_GLOBAL", TAB_CONFIG_KEYS.LAST_SELECTED),
    "",
  );

  var options = configs.map(function (c, i) {
    var marker = c.sheetName === lastSelected ? " ★ [current]" : "";
    return i + 1 + ". " + c.sheetName + marker;
  });

  var response = ui.prompt(
    title || "Select Sheet Tab",
    "Select a tab by number:  ★ = currently selected working tab\n\n" +
      options.join("\n"),
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return null;

  var idx = parseInt(response.getResponseText().trim(), 10);
  if (isNaN(idx) || idx < 1 || idx > configs.length) {
    ui.alert("Invalid selection.");
    return null;
  }
  return configs[idx - 1];
}

// ═══════════════════════════════════════════════════════════════════════════
// ColumnMappingCore — column-map persistence, alias matching, fuzzy detection
// ═══════════════════════════════════════════════════════════════════════════

// 22 Column Mapping Core

function buildColumnMap(sheet, tabName) {
  if (tabName) {
    var flexible = buildFlexibleColumnMap(sheet, tabName);
    return flexible.map;
  }
  return buildColumnMapLegacy(sheet);
}

function validateColumnMap(colMap, headerRow) {
  // Validates that all Team A (user-input) columns exist. Managed columns
  // (Status, etc.) are not validated here — they are auto-created by
  // ensureSetupAutomationColumns before this check runs.
  var missing = [];
  for (var i = 0; i < REQUIRED_USER_COLUMNS.length; i++) {
    var key = REQUIRED_USER_COLUMNS[i];
    if (!colMap[key]) missing.push(HEADERS[key]);
  }
  var rowLabel = headerRow || CONFIG.HEADER_ROW;
  return missing.length > 0
    ? "Required user-input column(s) not found in row " +
        rowLabel +
        ": " +
        missing.join(", ")
    : null;
}

function getTabConfigKey(tabName, configType) {
  return TAB_CONFIG.PREFIX + tabName + "_" + configType;
}

function saveTabColumnMapping(tabName, mapping) {
  var key = getTabConfigKey(tabName, TAB_CONFIG_KEYS.COLUMN_MAP);
  setPropString(key, JSON.stringify(mapping || {}));
}

function getTabColumnMapping(tabName) {
  var key = getTabConfigKey(tabName, TAB_CONFIG_KEYS.COLUMN_MAP);
  var raw = getPropString(key, "");
  if (!raw) return {};
  try {
    var parsed = JSON.parse(raw);
    return typeof parsed === "object" && parsed !== null ? parsed : {};
  } catch (e) {
    return {};
  }
}

function clearTabColumnMapping(tabName) {
  var key = getTabConfigKey(tabName, TAB_CONFIG_KEYS.COLUMN_MAP);
  getAutomationProperties().deleteProperty(key);
}

function getTabHeaderRow(tabName) {
  var key = getTabConfigKey(tabName, TAB_CONFIG_KEYS.HEADER_ROW);
  var raw = getPropString(key, "");
  var row = parseInt(raw, 10);
  return isNaN(row) || row < 1 ? TAB_CONFIG.DEFAULT_HEADER_ROW : row;
}

function setTabHeaderRow(tabName, rowNum) {
  var row = parseInt(rowNum, 10);
  if (isNaN(row) || row < 1) row = TAB_CONFIG.DEFAULT_HEADER_ROW;
  var key = getTabConfigKey(tabName, TAB_CONFIG_KEYS.HEADER_ROW);
  setPropString(key, String(row));
}

function getTabDataStartRow(tabName) {
  var key = getTabConfigKey(tabName, TAB_CONFIG_KEYS.DATA_START_ROW);
  var raw = getPropString(key, "");
  var row = parseInt(raw, 10);
  if (!isNaN(row) && row > 0) return row;
  return getTabHeaderRow(tabName) + 1;
}

function setTabDataStartRow(tabName, rowNum) {
  var row = parseInt(rowNum, 10);
  if (isNaN(row) || row < 1) return;
  var key = getTabConfigKey(tabName, TAB_CONFIG_KEYS.DATA_START_ROW);
  setPropString(key, String(row));
}

function normalizeHeaderName(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function getRowValues(sheet, rowNum, lastCol) {
  if (rowNum < 1 || lastCol < 1) return [];
  if (rowNum > sheet.getLastRow()) return [];
  return sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
}

function buildKnownHeaderLookup() {
  var lookup = {};

  function addHeaderName(value) {
    var normalized = normalizeHeaderName(value);
    if (normalized) lookup[normalized] = true;
  }

  for (var key in HEADERS) {
    addHeaderName(HEADERS[key]);

    var aliases = HEADER_ALIASES[key] || [];
    for (var i = 0; i < aliases.length; i++) {
      addHeaderName(aliases[i]);
    }

    var flexibleAliases = FLEXIBLE_HEADER_ALIASES[key] || [];
    for (var j = 0; j < flexibleAliases.length; j++) {
      addHeaderName(flexibleAliases[j]);
    }
  }

  return lookup;
}

function scoreHeaderRowValues(headerValues, knownLookup) {
  var nonEmpty = 0;
  var exactRecognized = 0;
  var fuzzyRecognized = 0;

  for (var i = 0; i < headerValues.length; i++) {
    var text = String(headerValues[i] || "").trim();
    if (!text) continue;

    nonEmpty++;
    var normalized = normalizeHeaderName(text);
    if (knownLookup[normalized]) {
      exactRecognized++;
      continue;
    }

    var fuzzy = fuzzyMatchHeader(text, 0.82);
    if (fuzzy) fuzzyRecognized++;
  }

  return {
    nonEmpty: nonEmpty,
    exactRecognized: exactRecognized,
    fuzzyRecognized: fuzzyRecognized,
    score:
      exactRecognized * 3 + fuzzyRecognized * 1.5 + Math.min(nonEmpty, 8) * 0.1,
  };
}

function resolveEffectiveHeaderRow(sheet, tabName, rowOverride) {
  var configuredRow = parseInt(rowOverride, 10);
  if (isNaN(configuredRow) || configuredRow < 1) {
    configuredRow = getTabHeaderRow(tabName);
  }

  var result = {
    configuredRow: configuredRow,
    headerRow: configuredRow,
    usedFallback: false,
    reason: "",
    headers: [],
  };

  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  if (lastCol === 0 || lastRow < configuredRow) {
    result.headers = [];
    return result;
  }

  var knownLookup = buildKnownHeaderLookup();
  var preferredHeaders = getRowValues(sheet, configuredRow, lastCol);
  var preferredScore = scoreHeaderRowValues(preferredHeaders, knownLookup);

  var bestRow = configuredRow;
  var bestHeaders = preferredHeaders;
  var bestScore = preferredScore;

  var scanUntil = Math.min(lastRow, configuredRow + 6);
  for (var row = configuredRow + 1; row <= scanUntil; row++) {
    var candidateHeaders = getRowValues(sheet, row, lastCol);
    var candidateScore = scoreHeaderRowValues(candidateHeaders, knownLookup);
    if (candidateScore.score > bestScore.score) {
      bestRow = row;
      bestHeaders = candidateHeaders;
      bestScore = candidateScore;
    }
  }

  var configuredLooksWeak =
    preferredScore.nonEmpty <= 1 || preferredScore.exactRecognized === 0;
  var bestLooksHeader =
    bestScore.nonEmpty >= 2 &&
    (bestScore.exactRecognized >= 1 || bestScore.fuzzyRecognized >= 2);
  var significantlyBetter =
    bestRow !== configuredRow && bestScore.score >= preferredScore.score + 2;

  if (bestRow !== configuredRow && bestLooksHeader) {
    if (configuredLooksWeak || significantlyBetter) {
      result.headerRow = bestRow;
      result.usedFallback = true;
      result.reason =
        "Configured header row " +
        configuredRow +
        " looked weak (non-empty cells: " +
        preferredScore.nonEmpty +
        "). Using row " +
        bestRow +
        " for column detection.";
      result.headers = bestHeaders;
      return result;
    }
  }

  result.headers = preferredHeaders;
  return result;
}

function fuzzyStringSimilarity(str1, str2) {
  var s1 = String(str1 || "")
    .toLowerCase()
    .trim();
  var s2 = String(str2 || "")
    .toLowerCase()
    .trim();

  if (s1 === s2) return 1.0;
  if (!s1 || !s2) return 0.0;

  // Check for substring match (strong indicator)
  if (s1.indexOf(s2) >= 0 || s2.indexOf(s1) >= 0) {
    return 0.8;
  }

  // Calculate Levenshtein distance
  var len1 = s1.length;
  var len2 = s2.length;
  var maxLen = Math.max(len1, len2);

  if (maxLen === 0) return 1.0;

  // Simple Levenshtein
  var matrix = [];
  for (var i = 0; i <= len1; i++) {
    matrix[i] = [i];
  }
  for (var j = 0; j <= len2; j++) {
    matrix[0][j] = j;
  }

  for (var i = 1; i <= len1; i++) {
    for (var j = 1; j <= len2; j++) {
      var cost = s1.charAt(i - 1) === s2.charAt(j - 1) ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,
        matrix[i][j - 1] + 1,
        matrix[i - 1][j - 1] + cost,
      );
    }
  }

  var distance = matrix[len1][len2];
  return 1.0 - distance / maxLen;
}

function fuzzyMatchHeader(actualHeader, minSimilarity) {
  var minSim = minSimilarity || 0.6;
  var bestMatch = null;
  var bestScore = 0;

  for (var logicalKey in FLEXIBLE_HEADER_ALIASES) {
    var aliases = FLEXIBLE_HEADER_ALIASES[logicalKey];
    for (var i = 0; i < aliases.length; i++) {
      var score = fuzzyStringSimilarity(actualHeader, aliases[i]);
      if (score > bestScore && score >= minSim) {
        bestScore = score;
        bestMatch = {
          logicalKey: logicalKey,
          similarity: score,
          matchedHeader: aliases[i],
        };
      }
    }
  }

  return bestMatch;
}

function detectColumnMappings(sheet, tabName) {
  var headerInfo = resolveEffectiveHeaderRow(sheet, tabName);
  var headerRow = headerInfo.headerRow;
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];

  var headers = headerInfo.headers || [];
  if (headers.length === 0) {
    headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  }
  var suggestions = [];

  for (var i = 0; i < headers.length; i++) {
    var actualHeader = String(headers[i] || "").trim();
    if (!actualHeader) continue;

    var match = fuzzyMatchHeader(actualHeader, 0.5);
    if (match) {
      suggestions.push({
        colIndex: i + 1,
        actualHeader: actualHeader,
        suggestedLogicalKey: match.logicalKey,
        confidence: match.similarity,
        matchedTo: match.matchedHeader,
      });
    }
  }

  // Sort by confidence descending
  suggestions.sort(function (a, b) {
    return b.confidence - a.confidence;
  });

  return suggestions;
}

function buildFlexibleColumnMap(sheet, tabName) {
  var headerInfo = resolveEffectiveHeaderRow(sheet, tabName);
  var result = {
    map: {},
    source: "",
    warnings: [],
    headerRow: headerInfo.headerRow,
    configuredHeaderRow: headerInfo.configuredRow,
  };

  if (headerInfo.usedFallback && headerInfo.reason) {
    result.warnings.push(headerInfo.reason);
  }

  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    result.warnings.push("Sheet has no columns");
    return result;
  }

  // Read actual headers from sheet (using effective row)
  var actualHeaders = headerInfo.headers || [];
  if (actualHeaders.length === 0) {
    actualHeaders = getRowValues(sheet, result.headerRow, lastCol);
  }

  // Build reverse lookup from actual header to column index
  var headerToIndex = {};
  for (var i = 0; i < actualHeaders.length; i++) {
    var h = normalizeHeaderName(actualHeaders[i]);
    if (h) headerToIndex[h] = i + 1;
  }

  // 1) Try stored per-tab mapping
  var storedMapping = getTabColumnMapping(tabName);
  var usedStored = false;
  var usedGlobal = false;
  var usedFuzzy = false;

  for (var logicalKey in storedMapping) {
    var expectedHeader = normalizeHeaderName(storedMapping[logicalKey]);
    if (expectedHeader && headerToIndex[expectedHeader]) {
      result.map[logicalKey] = headerToIndex[expectedHeader];
      usedStored = true;
    }
  }

  // 2) Fill gaps with global HEADERS + HEADER_ALIASES
  var globalMap = {};
  for (var key in HEADERS) {
    globalMap[normalizeHeaderName(HEADERS[key])] = key;
    var aliases = HEADER_ALIASES[key] || [];
    for (var a = 0; a < aliases.length; a++) {
      globalMap[normalizeHeaderName(aliases[a])] = key;
    }
  }

  for (var h in headerToIndex) {
    var globalKey = globalMap[h];
    if (globalKey && !result.map[globalKey]) {
      result.map[globalKey] = headerToIndex[h];
      usedGlobal = true;
    }
  }

  // 3) Fill remaining gaps with fuzzy matching for required user-input columns
  for (var i = 0; i < REQUIRED_USER_COLUMNS.length; i++) {
    var key = REQUIRED_USER_COLUMNS[i];
    if (!result.map[key]) {
      for (var normalizedHeader in headerToIndex) {
        var match = fuzzyMatchHeader(
          actualHeaders[headerToIndex[normalizedHeader] - 1],
          0.75,
        );
        if (match && match.logicalKey === key) {
          result.map[key] = headerToIndex[normalizedHeader];
          usedFuzzy = true;
          break;
        }
      }
    }
  }

  // Determine source label
  var sources = [];
  if (usedStored) sources.push("stored");
  if (usedGlobal) sources.push("global");
  if (usedFuzzy) sources.push("fuzzy");
  result.source = sources.join("+") || "none";

  // Validate required user-input columns
  for (var i = 0; i < REQUIRED_USER_COLUMNS.length; i++) {
    var reqKey = REQUIRED_USER_COLUMNS[i];
    if (!result.map[reqKey]) {
      result.warnings.push("Missing required column: " + reqKey);
    }
  }

  return result;
}

function buildColumnMapLegacy(sheet) {
  var headerRow = sheet
    .getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  var reverseHeaders = {};
  for (var key in HEADERS) {
    reverseHeaders[normalizeHeaderName(HEADERS[key])] = key;
    var aliases = HEADER_ALIASES[key] || [];
    for (var a = 0; a < aliases.length; a++) {
      reverseHeaders[normalizeHeaderName(aliases[a])] = key;
    }
  }
  var map = {};
  for (var c = 0; c < headerRow.length; c++) {
    var h = normalizeHeaderName(headerRow[c]);
    if (reverseHeaders[h]) {
      map[reverseHeaders[h]] = c + 1;
    }
  }
  return map;
}

function validateFlexibleColumnMap(flexibleResult) {
  if (!flexibleResult || !flexibleResult.map) {
    return "No column map available";
  }

  var missing = [];
  for (var i = 0; i < REQUIRED_USER_COLUMNS.length; i++) {
    var reqKey = REQUIRED_USER_COLUMNS[i];
    if (!flexibleResult.map[reqKey]) {
      missing.push(reqKey);
    }
  }

  if (missing.length > 0) {
    return (
      "Missing required columns: " +
      missing.join(", ") +
      "\n\nDetected source: " +
      flexibleResult.source +
      (flexibleResult.warnings.length > 0
        ? "\nWarnings: " + flexibleResult.warnings.join("; ")
        : "")
    );
  }

  return null;
}

// ═══════════════════════════════════════════════════════════════════════════
// ColumnMappingUi — Map Tab Columns + header-row dialogs
// ═══════════════════════════════════════════════════════════════════════════

// 13 Tab Mapping

function validateActiveTabStructure() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = getWorkingTabConfig(ss);

  if (!config) return;
  var sheet = config.sheet;
  var sheetName = config.sheetName;

  if (!sheet) {
    ui.alert("Tab '" + sheetName + "' not found!");
    return;
  }

  var flexMap = buildFlexibleColumnMap(sheet, sheetName);
  var validation = validateFlexibleColumnMap(flexMap);

  var lines = [
    "Tab: " + sheetName,
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
  var sheet = config.sheet;
  var sheetName = config.sheetName;

  var stored = getTabColumnMapping(sheetName);
  var flexMap = buildFlexibleColumnMap(sheet, sheetName);

  var lines = [
    "Tab: " + sheetName,
    "Configured Header Row: " + getTabHeaderRow(sheetName),
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

  var selectedNames = promptSelectTabs({
    ss: ss,
    title: "Map Tab Columns",
    source: "configured",
    prompt: "Pick one or more tabs to map. The mapping dialog runs once per tab.",
  });
  if (selectedNames.length === 0) return;

  var summaries = [];
  for (var s = 0; s < selectedNames.length; s++) {
    var sheetName = selectedNames[s];
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      summaries.push("✗ " + sheetName + " — tab not found");
      continue;
    }
    var result = mapColumnsForTab(ui, sheet, sheetName, s + 1, selectedNames.length);
    summaries.push(result);
    if (result.indexOf("CANCELLED") >= 0) break;
  }

  ui.alert(
    "Map Tab Columns — Summary",
    summaries.join("\n"),
    ui.ButtonSet.OK,
  );
}

function mapColumnsForTab(ui, sheet, sheetName, ordinal, total) {
  var prefix = "[" + ordinal + "/" + total + "] " + sheetName;
  var headerInfo = resolveEffectiveHeaderRow(sheet, sheetName);
  var headerValues = headerInfo.headers || [];
  var availableHeaders = buildAvailableHeaderChoices(headerValues);
  if (availableHeaders.length === 0) {
    ui.alert(
      prefix + " — No usable headers found. Set a valid header row first.",
    );
    return "✗ " + sheetName + " — no headers";
  }

  var suggestions = detectColumnMappings(sheet, sheetName);
  var suggestedMap = buildSuggestedHeaderMap(suggestions);
  var currentMapping = getTabColumnMapping(sheetName);
  var mapKeys = getMapTabColumnKeys();

  var lines = [
    "Tab: " + sheetName,
    "Configured Header Row: " + getTabHeaderRow(sheetName),
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
    "Map Tab Columns — " + prefix,
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
        ui.alert("Column mapping cancelled for " + sheetName + ". No changes saved.");
        return "CANCELLED " + sheetName;
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

  saveTabColumnMapping(sheetName, newMapping);

  var savedFlexMap = buildFlexibleColumnMap(sheet, sheetName);
  var validationMessage = validateFlexibleColumnMap(savedFlexMap);

  return validationMessage
    ? "⚠ " + sheetName + " — saved " + Object.keys(newMapping).length +
        " mapping(s); " + validationMessage.split("\n")[0]
    : "✓ " + sheetName + " — saved " + Object.keys(newMapping).length +
        " mapping(s); all required columns mapped";
}

function setTabHeaderRowDialog() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var selectedNames = promptSelectTabs({
    ss: ss,
    title: "Set Tab Header Row",
    source: "configured",
    prompt: "Pick one or more tabs. The same header row number is applied to each.",
  });
  if (selectedNames.length === 0) return;

  var rowResp = ui.prompt(
    "Set Header Row",
    "Enter the header row number to apply to: " + selectedNames.join(", ") +
      "\n\n(Default is 2.)",
    ui.ButtonSet.OK_CANCEL,
  );
  if (rowResp.getSelectedButton() !== ui.Button.OK) return;

  var newRow = parseInt(rowResp.getResponseText().trim(), 10);
  if (isNaN(newRow) || newRow < 1) {
    ui.alert("Invalid row number.");
    return;
  }

  var summaries = [];
  for (var s = 0; s < selectedNames.length; s++) {
    var sheetName = selectedNames[s];
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      summaries.push("✗ " + sheetName + " — not found");
      continue;
    }

    var finalRow = newRow;
    var headerInfo = resolveEffectiveHeaderRow(sheet, sheetName, newRow);
    if (headerInfo.usedFallback && headerInfo.headerRow !== newRow) {
      var useSuggested = ui.alert(
        "Header Row Suggestion — " + sheetName,
        "Row " + newRow +
          " looks like a divider/non-header row in this tab.\n\n" +
          "Suggested header row: " + headerInfo.headerRow +
          "\n\nUse the suggested row for this tab?",
        ui.ButtonSet.YES_NO,
      );
      if (useSuggested === ui.Button.YES) {
        finalRow = headerInfo.headerRow;
      }
    }

    setTabHeaderRow(sheetName, finalRow);
    setTabDataStartRow(sheetName, finalRow + 1);
    summaries.push("✓ " + sheetName + " — header row " + finalRow + ", data starts row " + (finalRow + 1));
  }

  ui.alert("Header Rows Updated", summaries.join("\n"), ui.ButtonSet.OK);
}

function viewTabConfiguration() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = getWorkingTabConfig(ss);

  if (!config) return;
  var sheetName = config.sheetName;

  var headerRow = getTabHeaderRow(sheetName);
  var dataStartRow = getTabDataStartRow(sheetName);
  var mapping = getTabColumnMapping(sheetName);
  var noticeOpts = getNoticeOptionsForTab(sheetName);

  var lines = [
    "Tab: " + sheetName,
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

  if (!config) return;
  var sheet = config.sheet;
  var sheetName = config.sheetName;
  if (!sheet) {
    ui.alert("No tab selected.");
    return;
  }

  var flexMap = buildFlexibleColumnMap(sheet, sheetName);
  var hasThreadId = !!flexMap.map.SENT_THREAD_ID;
  var hasReplyStatus = !!flexMap.map.REPLY_STATUS;
  var keywords = getReplyKeywords();
  var triggerActive = getTriggersByHandler("runReplyScan").length > 0;

  var lines = [
    "Tab: " + sheetName,
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

// ═══════════════════════════════════════════════════════════════════════════
// Dropdowns — Status / Send Mode / Notice Date dropdown setup
// ═══════════════════════════════════════════════════════════════════════════

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

// ═══════════════════════════════════════════════════════════════════════════
// ValidateUserColumns — Team A presence + staff-alias classification
// ═══════════════════════════════════════════════════════════════════════════

// 15 Validate User Columns
//
// Team A user-input columns must all be present in a configured tab. The
// setup wizard calls this to surface what's missing and offer to add the
// headers automatically. Aliases are honored — a sheet that already has
// "Staff Email" or "Remarks" passes without renaming.

function findMissingUserInputColumns(sheet, tabName) {
  var flexMap = buildFlexibleColumnMap(sheet, tabName);
  var missing = [];
  for (var i = 0; i < REQUIRED_USER_COLUMNS.length; i++) {
    var key = REQUIRED_USER_COLUMNS[i];
    if (!flexMap.map[key]) {
      missing.push({ key: key, label: HEADERS[key] });
    }
  }
  return missing;
}

// Returns true iff every Team A column is present (directly or via alias).
function userInputColumnsAreComplete(sheet, tabName) {
  return findMissingUserInputColumns(sheet, tabName).length === 0;
}

// Lists distinct, non-empty Assigned Staff Email values across the data
// rows of the given tabs. Used by the wizard to warn about staff emails
// that aren't verified Gmail aliases.
function listDistinctStaffEmailsForTabs(ss, tabNames) {
  var seen = {};
  var ordered = [];

  for (var t = 0; t < tabNames.length; t++) {
    var tabName = tabNames[t];
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) continue;

    var colMap = buildColumnMap(sheet, tabName);
    if (!colMap.STAFF_EMAIL) continue;

    var dataStartRow = getTabDataStartRow(tabName);
    var lastRow = sheet.getLastRow();
    if (lastRow < dataStartRow) continue;

    var values = sheet
      .getRange(dataStartRow, colMap.STAFF_EMAIL, lastRow - dataStartRow + 1, 1)
      .getValues();

    for (var i = 0; i < values.length; i++) {
      var raw = String(values[i][0] || "").trim().toLowerCase();
      if (!raw) continue;
      if (!seen[raw]) {
        seen[raw] = true;
        ordered.push(raw);
      }
    }
  }

  return ordered;
}

// Splits a list of staff email addresses into verified (usable as From)
// and unverified (would mis-attribute or fail). The returned object is
// passed into the wizard's pre-flight summary.
function classifyStaffEmailsByAliasVerification(staffEmails) {
  resetVerifiedSenderAliasCache();
  var verified = [];
  var unverified = [];
  for (var i = 0; i < staffEmails.length; i++) {
    if (canSendAs(staffEmails[i])) {
      verified.push(staffEmails[i]);
    } else {
      unverified.push(staffEmails[i]);
    }
  }
  return { verified: verified, unverified: unverified };
}

// ═══════════════════════════════════════════════════════════════════════════
// SetupWizard — guided setup flow (orchestrates everything above)
// ═══════════════════════════════════════════════════════════════════════════

// 10 Setup Wizard

function runSetupWizard() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  ui.alert(
    "🚀 Setup This Sheet for Automation",
    "Welcome! This wizard will guide you through setting up a sheet tab for the Expiry Notification automation.\n\nYou will:\n  Step 1 — Select or create a sheet tab\n  Step 2 — Verify column headers and create automation-managed columns\n  Step 3 — Apply dropdowns\n  Step 4 — Activate the daily schedule\n\nClick OK to begin.",
    ui.ButtonSet.OK,
  );

  var context = {
    tabName: "",
    sheet: null,
    columnsOk: false,
    setupColumnsApplied: false,
    dropsApplied: false,
    scheduleActive: false,
  };

  context = wizardStep1Tab(ss, ui, context);
  if (!context) return;

  context = wizardStep2Columns(ss, ui, context);
  if (!context) return;

  context = wizardStep3Dropdowns(ss, ui, context);
  if (!context) return;

  wizardStepStaffAliasCheck(ss, ui, context);

  context = wizardStep4Schedule(ui, context);
  if (!context) return;

  wizardStep5Summary(ui, context);
}

// Pre-flight check: every distinct Assigned Staff Email across configured
// tabs must be a verified Gmail "Send mail as" alias of the script runner.
// Otherwise those rows will error at send time. We only warn here — the
// user may add aliases in Gmail then re-run.
function wizardStepStaffAliasCheck(ss, ui, context) {
  var configured = getConfiguredTabEntries();
  var tabNames = configured.length > 0
    ? configured.map(function (e) { return e.name; })
    : [context.tabName];

  var staffEmails = listDistinctStaffEmailsForTabs(ss, tabNames);
  if (staffEmails.length === 0) {
    context.aliasCheckPassed = true;
    return context;
  }

  var classified = classifyStaffEmailsByAliasVerification(staffEmails);
  context.aliasCheckPassed = classified.unverified.length === 0;

  if (context.aliasCheckPassed) return context;

  ui.alert(
    "Sender Alias Check ⚠",
    "Outgoing emails are sent FROM the row's Assigned Staff Email. The following addresses are NOT registered as Gmail \"Send mail as\" aliases on this script's account, so rows using them will be marked Error at send time:\n\n  • " +
      classified.unverified.join("\n  • ") +
      "\n\nFix it in Gmail → Settings → Accounts and Import → Send mail as → Add another email address. Then re-run this wizard.",
    ui.ButtonSet.OK,
  );

  return context;
}

function wizardStep1Tab(ss, ui, context) {
  var choice = ui.prompt(
    "Step 1 of 4 — Sheet Tab",
    "Which tab should be used for automation?\n\n  1. Use an existing tab\n  2. Create a new tab\n\nEnter 1 or 2:",
    ui.ButtonSet.OK_CANCEL,
  );
  if (choice.getSelectedButton() !== ui.Button.OK) return null;

  var input = choice.getResponseText().trim();

  if (input === "2") {
    // ── Create new tab ──
    var nameResp = ui.prompt(
      "Step 1 of 4 — Create New Tab",
      "Enter a name for the new sheet tab:",
      ui.ButtonSet.OK_CANCEL,
    );
    if (nameResp.getSelectedButton() !== ui.Button.OK) return null;

    var newName = nameResp.getResponseText().trim();
    if (!newName) {
      ui.alert("No name entered. Setup cancelled.");
      return null;
    }

    if (ss.getSheetByName(newName)) {
      ui.alert(
        'A tab named "' +
          newName +
          "\" already exists. Setup cancelled.\n\nRe-run the wizard and choose 'Use an existing tab' to configure it.",
      );
      return null;
    }

    var newSheet = ss.insertSheet(newName);

    // Write required user-input headers into row 1, in the canonical order
    // a person would naturally fill in. Code-managed columns are added
    // later by ensureSetupAutomationColumns.
    var requiredHeaders = [
      HEADERS.NO,
      HEADERS.CLIENT_NAME,
      HEADERS.CLIENT_EMAIL,
      HEADERS.DOC_TYPE,
      HEADERS.EXPIRY_DATE,
      HEADERS.NOTICE_DATE,
      HEADERS.REMARKS,
      HEADERS.ATTACHMENTS,
      HEADERS.STAFF_NAME,
      HEADERS.STAFF_EMAIL,
    ];
    newSheet
      .getRange(1, 1, 1, requiredHeaders.length)
      .setValues([requiredHeaders]);

    // Bold + freeze header row
    newSheet.getRange(1, 1, 1, requiredHeaders.length).setFontWeight("bold");
    newSheet.setFrozenRows(1);

    // Register tab by ID+name
    var existingEntries = getConfiguredTabEntries();
    var existingIds = existingEntries.map(function (e) {
      return e.id;
    });
    if (existingIds.indexOf(newSheet.getSheetId()) < 0) {
      existingEntries.push({ id: newSheet.getSheetId(), name: newName });
      setConfiguredTabEntries(existingEntries);
    }

    // Set header row = 1, data start = 2
    setTabHeaderRow(newName, 1);
    setTabDataStartRow(newName, 2);

    context.tabName = newName;
    context.sheet = newSheet;

    ui.alert(
      "Step 1 Complete ✓",
      'Tab "' + newName + '" created with headers.\n\nProceeding to Step 2.',
      ui.ButtonSet.OK,
    );
    return context;
  } else {
    // ── Use existing tab ── show ALL tabs
    var sheets = ss.getSheets();

    var currentEntries = getConfiguredTabEntries();
    var currentIds = currentEntries.map(function (e) {
      return e.id;
    });
    var currentNames = currentEntries.map(function (e) {
      return e.name;
    });

    var options = [];
    for (var i = 0; i < sheets.length; i++) {
      var isActive =
        currentIds.indexOf(sheets[i].getSheetId()) >= 0 ||
        currentNames.indexOf(sheets[i].getName()) >= 0;
      options.push(
        i +
          1 +
          ". " +
          sheets[i].getName() +
          (isActive ? " [already registered]" : ""),
      );
    }

    var pickResp = ui.prompt(
      "Step 1 of 4 — Select Existing Tab(s)",
      "Available tabs:\n\n" +
        options.join("\n") +
        "\n\nEnter ONE OR MORE tab numbers, separated by commas (e.g. 1,2,3,4) " +
        "or as a range (e.g. 1-3). Names also work.",
      ui.ButtonSet.OK_CANCEL,
    );
    if (pickResp.getSelectedButton() !== ui.Button.OK) return null;

    var rawInput = String(pickResp.getResponseText() || "").trim();
    if (!rawInput) {
      ui.alert("No selection entered. Setup cancelled.");
      return null;
    }

    // Use the same multi-select parser the other menu items use, so
    // "1,2,3,4" registers ALL four tabs (not just the first).
    var sheetEntries = sheets.map(function (sh) {
      return { sheet: sh, sheetName: sh.getName() };
    });
    var parsed = parseTabSelectionInput(rawInput, sheetEntries);

    if (parsed.selected.length === 0) {
      ui.alert(
        "Invalid Selection",
        "Could not resolve: " + parsed.errors.join(", ") +
          "\n\nValid range is 1-" + sheets.length + ". Setup cancelled.",
        ui.ButtonSet.OK,
      );
      return null;
    }

    // Register every newly-selected tab.
    var addedNames = [];
    for (var n = 0; n < parsed.selected.length; n++) {
      var sel = parsed.selected[n].sheet;
      if (currentIds.indexOf(sel.getSheetId()) < 0) {
        currentEntries.push({ id: sel.getSheetId(), name: sel.getName() });
        currentIds.push(sel.getSheetId());
        addedNames.push(sel.getName());
      }
    }
    if (addedNames.length > 0) {
      setConfiguredTabEntries(currentEntries);
    }

    // The first selected tab drives the wizard's per-tab dialogs;
    // allTabNames lets later steps loop over every tab the user picked.
    var primarySheet = parsed.selected[0].sheet;
    context.tabName = primarySheet.getName();
    context.sheet = primarySheet;
    context.allTabNames = parsed.selected.map(function (e) { return e.sheetName; });

    var allLabel = context.allTabNames.join(", ");
    var skippedNote = parsed.errors.length > 0
      ? "\n\n⚠ Skipped (could not resolve): " + parsed.errors.join(", ")
      : "";

    ui.alert(
      "Step 1 Complete ✓",
      "Selected " + context.allTabNames.length + " tab(s): " + allLabel +
        "\n\nProceeding to Step 2 (will run per tab)." +
        skippedNote,
      ui.ButtonSet.OK,
    );
    return context;
  }
}

function wizardStep2Columns(ss, ui, context) {
  // If the user picked multiple tabs in Step 1, run the column check for
  // every tab. Otherwise just the primary.
  var tabNames = context.allTabNames && context.allTabNames.length > 0
    ? context.allTabNames
    : [context.tabName];

  var allOk = true;
  var allApplied = true;
  var perTabSummary = [];

  for (var t = 0; t < tabNames.length; t++) {
    var tabName = tabNames[t];
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      perTabSummary.push("✗ " + tabName + " — sheet not found");
      allOk = false;
      allApplied = false;
      continue;
    }

    var result = wizardStep2ColumnsForTab(ui, sheet, tabName, t + 1, tabNames.length);
    if (result === null) return null; // user cancelled
    perTabSummary.push(result.summary);
    if (!result.columnsOk) allOk = false;
    if (!result.setupColumnsApplied) allApplied = false;
  }

  context.columnsOk = allOk;
  context.setupColumnsApplied = allApplied;

  if (tabNames.length > 1) {
    ui.alert(
      "Step 2 Summary (" + tabNames.length + " tabs)",
      perTabSummary.join("\n") + "\n\nProceeding to Step 3.",
      ui.ButtonSet.OK,
    );
  }
  return context;
}

// Runs Step 2 for ONE tab. Returns { columnsOk, setupColumnsApplied, summary }
// or null if the user cancelled out.
function wizardStep2ColumnsForTab(ui, sheet, tabName, ordinal, total) {
  var prefix = total > 1 ? "[" + ordinal + "/" + total + "] " + tabName + " — " : "";
  var missing = findMissingUserInputColumns(sheet, tabName);

  if (missing.length === 0) {
    ensureSetupAutomationColumns(sheet, tabName, buildColumnMap(sheet, tabName));
    if (total === 1) {
      ui.alert(
        "Step 2 Complete ✓",
        "All " + REQUIRED_USER_COLUMNS.length +
          " required user-input columns are present.\n\nCode-managed columns were created/verified for this tab.\n\nProceeding to Step 3.",
        ui.ButtonSet.OK,
      );
    }
    return { columnsOk: true, setupColumnsApplied: true,
             summary: "✓ " + tabName + " — all columns OK" };
  }

  var missingLabels = missing.map(function (m) { return m.label; });
  var addNow = ui.alert(
    "Step 2 — Missing User-Input Columns",
    prefix + "These required user-input columns are missing:\n\n  • " +
      missingLabels.join("\n  • ") +
      "\n\nAdd them automatically (headers appended at the end of the row)?" +
      "\n\nChoose No to skip and continue.",
    ui.ButtonSet.YES_NO_CANCEL,
  );
  if (addNow === ui.Button.CANCEL) return null;

  if (addNow === ui.Button.YES) {
    ensureUserInputColumns(sheet, tabName);
  }

  var stillMissing = findMissingUserInputColumns(sheet, tabName);
  var ok = stillMissing.length === 0;

  if (ok) {
    ensureSetupAutomationColumns(sheet, tabName, buildColumnMap(sheet, tabName));
    return { columnsOk: true, setupColumnsApplied: true,
             summary: "✓ " + tabName + " — columns added" };
  }

  return { columnsOk: false, setupColumnsApplied: false,
           summary: "⚠ " + tabName + " — still missing: " +
             stillMissing.map(function (m) { return m.label; }).join(", ") };
}

function wizardStep3Dropdowns(ss, ui, context) {
  var tabNames = context.allTabNames && context.allTabNames.length > 0
    ? context.allTabNames
    : [context.tabName];
  var label = tabNames.length === 1 ? '"' + tabNames[0] + '"'
                                    : tabNames.length + " selected tabs";

  var apply = ui.alert(
    "Step 3 of 4 — Dropdowns",
    "Apply dropdown options to the Status, Send Mode, and Notice Date columns in " +
      label + "?\n\nThis makes data entry easier. Default values will be used.",
    ui.ButtonSet.YES_NO,
  );

  if (apply !== ui.Button.YES) {
    ui.alert(
      "Step 3 Skipped",
      "Dropdowns skipped. You can apply them later via Tab Management → Setup Tab Dropdowns.",
      ui.ButtonSet.OK,
    );
    context.dropsApplied = false;
    return context;
  }

  var perTabSummary = [];
  var anyApplied = false;
  for (var t = 0; t < tabNames.length; t++) {
    var tabName = tabNames[t];
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      perTabSummary.push("✗ " + tabName + " — sheet not found");
      continue;
    }
    var applied = wizardStep3DropdownsForTab(sheet, tabName);
    if (applied.length > 0) {
      anyApplied = true;
      perTabSummary.push("✓ " + tabName + " — " + applied.join(", "));
    } else {
      perTabSummary.push("⚠ " + tabName + " — no matching columns");
    }
  }
  context.dropsApplied = anyApplied;

  ui.alert("Step 3 Complete",
    perTabSummary.join("\n") + "\n\nProceeding to Step 4.", ui.ButtonSet.OK);
  return context;
}

// Applies Status / Send Mode / Notice Date dropdowns for ONE tab.
function wizardStep3DropdownsForTab(sheet, tabName) {
  var colMap = buildColumnMap(sheet, tabName);
  var dataStartRow = getTabDataStartRow(tabName);
  var dataLastRow = Math.max(sheet.getLastRow(), dataStartRow + 100);
  var applied = [];

  if (colMap.STATUS) {
    sheet.getRange(dataStartRow, colMap.STATUS, dataLastRow - dataStartRow + 1, 1)
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(
            [STATUS.ACTIVE, STATUS.NOTICE_SENT, STATUS.SENT, STATUS.ERROR, STATUS.SKIPPED],
            true)
          .setAllowInvalid(true).build());
    applied.push("Status");
  }
  if (colMap.SEND_MODE) {
    sheet.getRange(dataStartRow, colMap.SEND_MODE, dataLastRow - dataStartRow + 1, 1)
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(
            [SEND_MODE.AUTO, SEND_MODE.HOLD, SEND_MODE.MANUAL_ONLY], true)
          .setAllowInvalid(true).build());
    applied.push("Send Mode");
  }
  if (colMap.NOTICE_DATE) {
    sheet.getRange(dataStartRow, colMap.NOTICE_DATE, dataLastRow - dataStartRow + 1, 1)
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(getNoticeOptionsForTab(tabName), true)
          .setAllowInvalid(true).build());
    applied.push("Notice Date");
  }
  return applied;
}

function wizardStep4Schedule(ui, context) {
  var triggerCount = getTriggersByHandler("runDailyCheck").length;
  var scheduledTime = formatDailyTriggerTime(
    getDailyTriggerHour(),
    getDailyTriggerMinute(),
  );
  var currentStatus =
    triggerCount > 0
      ? "ACTIVE (" + triggerCount + " trigger already set)"
      : "INACTIVE";

  var activate = ui.alert(
    "Step 4 of 4 — Daily Schedule",
    "Current daily schedule status: " +
      currentStatus +
      "\n\nActivate the daily " +
      scheduledTime +
      " email schedule now?",
    ui.ButtonSet.YES_NO,
  );

  if (activate === ui.Button.YES) {
    if (triggerCount === 0) {
      installTrigger();
    }
    context.scheduleActive = true;
  } else {
    context.scheduleActive = triggerCount > 0;
    ui.alert(
      "Step 4 Skipped",
      "Schedule not changed. You can activate it later via Automation Settings → Activate Daily Schedule.",
      ui.ButtonSet.OK,
    );
  }

  return context;
}

function wizardStep5Summary(ui, context) {
  var lines = [
    "✅ Setup Complete!",
    "",
    "Tab(s):          " + ((context.allTabNames && context.allTabNames.length > 1) ? (context.allTabNames.length + " tabs: " + context.allTabNames.join(", ")) : context.tabName),
    "Columns:         " +
      (context.columnsOk
        ? "✓ All required columns OK"
        : "⚠ Some columns missing — map them via Tab Management"),
    "Automation Cols: " +
      (context.setupColumnsApplied
        ? "✓ Created/verified"
        : "— Not fully applied"),
    "Dropdowns:       " + (context.dropsApplied ? "✓ Applied" : "— Skipped"),
    "Sender Aliases:  " +
      (context.aliasCheckPassed === false
        ? "⚠ Some staff emails not verified — fix in Gmail Send-As"
        : "✓ All staff emails verified"),
    "Daily Schedule:  " +
      (context.scheduleActive ? "✓ Active" : "— Not activated"),
    "",
    "You're ready to go!",
    'Add your client data to the "' +
      context.tabName +
      '" tab and the automation will handle the rest.',
  ];

  ui.alert("🚀 Setup Complete", lines.join("\n"), ui.ButtonSet.OK);
}

// ─────────────────────────────────────────────────────────────────────────────
