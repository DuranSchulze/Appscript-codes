// 22 Column Mapping Core

function buildColumnMap(sheet, tabName) {
  if (tabName) {
    var flexible = buildFlexibleColumnMap(sheet, tabName);
    return flexible.map;
  }
  return buildColumnMapLegacy(sheet);
}

function validateColumnMap(colMap, headerRow) {
  var required = [
    "CLIENT_NAME",
    "CLIENT_EMAIL",
    "EXPIRY_DATE",
    "NOTICE_DATE",
    "STATUS",
  ];
  var missing = [];
  for (var i = 0; i < required.length; i++) {
    if (!colMap[required[i]]) missing.push(HEADERS[required[i]]);
  }
  var rowLabel = headerRow || CONFIG.HEADER_ROW;
  return missing.length > 0
    ? "Required column(s) not found in row " +
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

  // 3) Fill remaining gaps with fuzzy matching for critical columns
  var criticalColumns = [
    "CLIENT_NAME",
    "CLIENT_EMAIL",
    "EXPIRY_DATE",
    "NOTICE_DATE",
    "STATUS",
  ];
  for (var i = 0; i < criticalColumns.length; i++) {
    var key = criticalColumns[i];
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

  // Validate required columns
  var required = [
    "CLIENT_NAME",
    "CLIENT_EMAIL",
    "EXPIRY_DATE",
    "NOTICE_DATE",
    "STATUS",
  ];
  for (var i = 0; i < required.length; i++) {
    if (!result.map[required[i]]) {
      result.warnings.push("Missing required column: " + required[i]);
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

  var required = [
    "CLIENT_NAME",
    "CLIENT_EMAIL",
    "EXPIRY_DATE",
    "NOTICE_DATE",
    "STATUS",
  ];
  var missing = [];

  for (var i = 0; i < required.length; i++) {
    if (!flexibleResult.map[required[i]]) {
      missing.push(required[i]);
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
