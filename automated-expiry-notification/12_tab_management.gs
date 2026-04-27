// 12 Tab Management

function selectWorkingTab() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // resolveAutomationSheets auto-purges deleted tabs and returns only live ones
  var sheetConfigs = resolveAutomationSheets(ss);

  if (sheetConfigs.length === 0) {
    ui.alert(
      "No Active Tabs Found",
      "No configured tabs exist (or all were deleted/renamed).\n\nUse 'Configure Automation Tabs' to register tabs.",
      ui.ButtonSet.OK,
    );
    return;
  }

  var lastSelected = getPropString(
    getTabConfigKey("_GLOBAL", TAB_CONFIG_KEYS.LAST_SELECTED),
    "",
  );

  var options = [];
  for (var i = 0; i < sheetConfigs.length; i++) {
    var config = sheetConfigs[i];
    var sheet = config.sheet; // always non-null here (purged above)
    var icon = "✓";
    var status = "";

    var flexMap = buildFlexibleColumnMap(sheet, config.sheetName);
    var hasMissing = flexMap.warnings.some(function (w) {
      return w.indexOf("Missing required") >= 0;
    });
    if (hasMissing) {
      icon = "⚠";
      status = " (missing columns)";
    }

    var marker = config.sheetName === lastSelected ? " ★ [current]" : "";
    options.push(
      i + 1 + ". " + icon + " " + config.sheetName + status + marker,
    );
  }

  var response = ui.prompt(
    "Select Working Tab",
    "✓ = ready, ⚠ = needs column setup   ★ = currently selected\n\n" +
      options.join("\n") +
      "\n\nEnter number:",
    ui.ButtonSet.OK_CANCEL,
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  var idx = parseInt(response.getResponseText().trim(), 10);
  if (isNaN(idx) || idx < 1 || idx > sheetConfigs.length) {
    ui.alert("Invalid selection.");
    return;
  }

  var selected = sheetConfigs[idx - 1];
  setPropString(
    getTabConfigKey("_GLOBAL", TAB_CONFIG_KEYS.LAST_SELECTED),
    selected.sheetName,
  );

  ui.alert(
    "Working Tab Set",
    '"' +
      selected.sheetName +
      '" is now your working tab.\n\nThis tab will be used for subsequent operations.',
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
