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
