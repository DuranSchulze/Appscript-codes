// 21 Sheet Resolution

function getConfiguredTabEntries() {
  var saved = getAutomationProperties().getProperty(
    CONFIG.AUTOMATION_SHEET_PROPERTY_KEY,
  );
  if (!saved || saved.trim() === "") {
    return [{ id: null, name: CONFIG.SHEET_NAME }];
  }

  var trimmed = saved.trim();
  if (trimmed.charAt(0) === "[") {
    try {
      var parsed = JSON.parse(trimmed);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed
          .filter(function (entry) {
            return (
              entry && (entry.name || entry.id || typeof entry === "string")
            );
          })
          .map(function (entry) {
            // Legacy: plain string inside the array
            if (typeof entry === "string")
              return { id: null, name: entry.trim() };
            return {
              id: entry.id || null,
              name: String(entry.name || "").trim(),
            };
          })
          .filter(function (e) {
            return e.name || e.id;
          });
      }
    } catch (e) {}
  }
  // Legacy plain string
  return [{ id: null, name: trimmed }];
}

function getConfiguredSheetNames() {
  return getConfiguredTabEntries().map(function (e) {
    return e.name;
  });
}

function getConfiguredSheetName() {
  return getConfiguredSheetNames()[0];
}

function setConfiguredTabEntries(entries) {
  var clean = (entries || []).filter(function (e) {
    return e && (e.name || e.id);
  });
  if (clean.length === 0) return;
  getAutomationProperties().setProperty(
    CONFIG.AUTOMATION_SHEET_PROPERTY_KEY,
    JSON.stringify(clean),
  );
}

function setConfiguredSheetNames(namesArray) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var entries = (namesArray || [])
    .filter(function (n) {
      return !!String(n || "").trim();
    })
    .map(function (name) {
      var sheet = ss.getSheetByName(name);
      return {
        id: sheet ? sheet.getSheetId() : null,
        name: String(name).trim(),
      };
    });
  setConfiguredTabEntries(entries);
}

function setConfiguredSheetName(sheetName) {
  var value = String(sheetName || "").trim();
  if (!value) return;
  setConfiguredSheetNames([value]);
}

function resolveAutomationSheets(ss) {
  var entries = getConfiguredTabEntries();
  var needsUpdate = false;
  var results = [];

  for (var i = 0; i < entries.length; i++) {
    var entry = entries[i];
    var foundSheet = null;

    // Try by numeric sheet ID first
    if (entry.id !== null && entry.id !== undefined) {
      var allSheets = ss.getSheets();
      for (var j = 0; j < allSheets.length; j++) {
        if (allSheets[j].getSheetId() === entry.id) {
          foundSheet = allSheets[j];
          break;
        }
      }
      // If found by ID but name has changed, update stored name
      if (foundSheet && foundSheet.getName() !== entry.name) {
        entry.name = foundSheet.getName();
        entries[i] = entry;
        needsUpdate = true;
      }
    }

    // Fallback: find by name when ID lookup failed
    if (!foundSheet && entry.name) {
      foundSheet = ss.getSheetByName(entry.name);
      if (foundSheet) {
        // Tab exists by name — update stored ID (handles deleted+recreated same-name tabs)
        entry.id = foundSheet.getSheetId();
        entries[i] = entry;
        needsUpdate = true;
      }
    }

    results.push({
      sheetName: entry.name,
      sheet: foundSheet,
      resolvedId: foundSheet ? foundSheet.getSheetId() : null,
    });
  }

  // Auto-purge entries for tabs that no longer exist
  // Also deduplicate: if two entries resolved to the same sheet ID, keep only the first one
  var surviving = [];
  var seenIds = [];
  for (var k = 0; k < results.length; k++) {
    if (!results[k].sheet) {
      // Tab not found — drop it
      needsUpdate = true;
      continue;
    }
    var resolvedId = results[k].resolvedId;
    if (seenIds.indexOf(resolvedId) >= 0) {
      // Duplicate — same physical sheet registered twice, drop the extra entry
      needsUpdate = true;
      continue;
    }
    seenIds.push(resolvedId);
    surviving.push(entries[k]);
  }

  // Persist updated names/IDs silently (including purge + dedup)
  if (needsUpdate) {
    setConfiguredTabEntries(surviving);
  }

  // Return only found, deduplicated sheets
  return results.filter(function (r) {
    return !!r.sheet && seenIds.indexOf(r.resolvedId) >= 0;
  });
}

function resolveAutomationSheet(ss) {
  var results = resolveAutomationSheets(ss);
  return results[0] || { sheetName: CONFIG.SHEET_NAME, sheet: null };
}

function getAutomationSpreadsheet() {
  var active = null;
  try {
    active = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}

  if (active) {
    rememberSpreadsheetId(active);
    return active;
  }

  var savedId = getPropString(PROP_KEYS.SPREADSHEET_ID, "");
  if (!savedId) {
    throw new Error(
      "Spreadsheet context unavailable. Open the spreadsheet once to initialize stored spreadsheet ID.",
    );
  }

  return SpreadsheetApp.openById(savedId);
}
