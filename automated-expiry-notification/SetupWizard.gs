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
      "Step 1 of 4 — Select Existing Tab",
      "Available tabs:\n\n" +
        options.join("\n") +
        "\n\nEnter the number of the tab to use:",
      ui.ButtonSet.OK_CANCEL,
    );
    if (pickResp.getSelectedButton() !== ui.Button.OK) return null;

    var idx = parseInt(pickResp.getResponseText().trim(), 10);
    if (isNaN(idx) || idx < 1 || idx > sheets.length) {
      ui.alert("Invalid selection. Setup cancelled.");
      return null;
    }

    var selectedSheet = sheets[idx - 1];
    var selectedName = selectedSheet.getName();
    var selectedId = selectedSheet.getSheetId();

    // Register by ID+name if not already registered
    var alreadyRegistered = currentIds.indexOf(selectedId) >= 0;
    if (!alreadyRegistered) {
      currentEntries.push({ id: selectedId, name: selectedName });
      setConfiguredTabEntries(currentEntries);
    }

    context.tabName = selectedName;
    context.sheet = selectedSheet;

    ui.alert(
      "Step 1 Complete ✓",
      'Tab "' + selectedName + '" registered.\n\nProceeding to Step 2.',
      ui.ButtonSet.OK,
    );
    return context;
  }
}

function wizardStep2Columns(ss, ui, context) {
  var missing = findMissingUserInputColumns(context.sheet, context.tabName);

  if (missing.length === 0) {
    ensureSetupAutomationColumns(context.sheet, context.tabName,
      buildColumnMap(context.sheet, context.tabName));
    context.setupColumnsApplied = true;
    context.columnsOk = true;
    ui.alert(
      "Step 2 Complete ✓",
      "All " + REQUIRED_USER_COLUMNS.length +
        " required user-input columns are present.\n\nCode-managed columns were created/verified for this tab.\n\nProceeding to Step 3.",
      ui.ButtonSet.OK,
    );
    return context;
  }

  var missingLabels = missing.map(function (m) { return m.label; });

  var addNow = ui.alert(
    "Step 2 of 4 — Missing User-Input Columns",
    "These required user-input columns are missing from \"" +
      context.tabName + "\":\n\n  • " + missingLabels.join("\n  • ") +
      "\n\nAdd them automatically (headers appended at the end of the row)?\n\nChoose No to map them manually instead.",
    ui.ButtonSet.YES_NO_CANCEL,
  );

  if (addNow === ui.Button.CANCEL) return null;

  if (addNow === ui.Button.YES) {
    ensureUserInputColumns(context.sheet, context.tabName);
  } else {
    mapTabColumns();
  }

  // Re-check after either path.
  var stillMissing = findMissingUserInputColumns(context.sheet, context.tabName);
  context.columnsOk = stillMissing.length === 0;

  if (context.columnsOk) {
    ensureSetupAutomationColumns(context.sheet, context.tabName,
      buildColumnMap(context.sheet, context.tabName));
    context.setupColumnsApplied = true;
    ui.alert(
      "Step 2 Complete ✓",
      "All required user-input columns are in place.\n\nCode-managed columns were created/verified for this tab.",
      ui.ButtonSet.OK,
    );
  } else {
    var stillMissingLabels = stillMissing.map(function (m) { return m.label; });
    context.setupColumnsApplied = false;
    ui.alert(
      "Step 2 — Warning",
      "These columns are still missing:\n\n  • " +
        stillMissingLabels.join("\n  • ") +
        "\n\nThe automation will skip rows until these are addressed.",
      ui.ButtonSet.OK,
    );
  }

  return context;
}

function wizardStep3Dropdowns(ss, ui, context) {
  var apply = ui.alert(
    "Step 3 of 4 — Dropdowns",
    'Apply dropdown options to the Status, Send Mode, and Notice Date columns in "' +
      context.tabName +
      '"?\n\nThis makes data entry easier. Default values will be used.',
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

  var sheet = context.sheet;
  var tabName = context.tabName;
  var colMap = buildColumnMap(sheet, tabName);
  var dataStartRow = getTabDataStartRow(tabName);
  var lastRow = sheet.getLastRow();
  var dataLastRow = Math.max(lastRow, dataStartRow + 100);
  var applied = [];

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

  if (colMap.NOTICE_DATE) {
    var noticeOptions = getNoticeOptionsForTab(tabName);
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
    applied.push("Notice Date");
  }

  context.dropsApplied = applied.length > 0;

  ui.alert(
    "Step 3 Complete ✓",
    applied.length > 0
      ? "Dropdowns applied to: " +
          applied.join(", ") +
          ".\n\nProceeding to Step 4."
      : "No matching columns found for dropdowns. You can retry later.\n\nProceeding to Step 4.",
    ui.ButtonSet.OK,
  );

  return context;
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
    "Tab:             " + context.tabName,
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
