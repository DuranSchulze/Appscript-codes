/**
 * MCAD AUTOFORWARD — OPERATING GUIDELINES AND PLATFORM LIMITS
 *
 * LABEL MEANING
 * Gmail labels are applied to whole conversations (threads), not to individual
 * messages. They are operational indicators for people reviewing the mailbox:
 *
 *   AutoForward/Detected
 *     The conversation contains at least one message that matched a rule.
 *
 *   AutoForward/Forwarded
 *     Gmail accepted a forward request for at least one message in the
 *     conversation. This does not guarantee final delivery or that every
 *     message in the conversation was sent.
 *
 *   AutoForward/Failed
 *     At least one message in the conversation is waiting for another retry.
 *
 *   AutoForward/Retry Exhausted
 *     At least one message reached MAX_RETRIES and needs manual review.
 *
 * The per-user message-ID records in User Properties are the source of truth
 * for duplicate prevention, accepted forward requests, retry count, and
 * backoff.
 * Labels must never be used by themselves to decide that every message in a
 * conversation has been processed.
 *
 * DEPLOYMENT GUIDELINES
 * 1. Use this only as a spreadsheet-bound Apps Script.
 * 2. Each mailbox owner must create/authorize their own installable triggers.
 * 3. Pilot with one controlled sender and recipient before production use.
 * 4. Keep recipients at 50 or fewer per rule.
 * 5. Review Failed and Retry Exhausted labels plus Apps Script Executions.
 * 6. Do not run an older forwarding script for the same mailbox in parallel.
 * 7. When distributing this workbook, give each user their own Google Sheets
 *    copy. They must set up their rule tab and start automation once using
 *    their own Gmail account.
 * 8. Self-repair works while at least one managed trigger still runs. If all
 *    triggers are deleted or authorization is revoked, the user must choose
 *    Start or repair once; no script can restart itself with no execution path.
 *
 * OUTSIDE THIS SCRIPT'S CONTROL
 * Google Workspace administrators may block Gmail access or external
 * forwarding. Gmail and Apps Script quotas, maximum execution time, attachment
 * limits, transient Google service errors, and approximate trigger schedules
 * are controlled by Google and may change. Sheet protection limits editing but
 * does not hide rule tabs from other workbook viewers. Forwarding and saving
 * the processed ID are separate operations; a rare interruption after Gmail
 * accepts a forward but before the ID is saved can cause a duplicate retry.
 */
const CONFIG = {
  CHECK_EVERY_MINUTES: 5,
  SEARCH_LOOKBACK_DAYS: 30,
  MAX_THREADS_PER_RUN: 500,
  BATCH_FORWARD_LIMIT: 50,
  RETAIN_PROCESSED_DAYS: 60,
  INCLUDE_SPAM: true,
  MAX_RETRIES: 5,
  RETRY_DELAY_HOURS: 2,
  WATCHDOG_EVERY_HOURS: 6,

  SUMMARY_HOUR: 23,
  SUMMARY_TIME_ZONE: "Asia/Manila",

  RULE_SHEET_PREFIX: "Rules - ",
  RULE_HEADER_ROW: 6,
  RULE_DATA_ROWS: 500,
  RULE_SCHEMA_VERSION: 2,
  INFO_SHEET_NAME: "AutoForward Info",
  INFO_SHEET_VERSION: 3,

  // Optional: addresses that may edit every protected account rule tab.
  ADMIN_EMAILS: [],

  ACCOUNT_REGISTRATION_PREFIX: "AF_ACCOUNT_REGISTRATION_",
  PROCESSED_PREFIX: "AF_FORWARDED_",
  FAILED_RETRY_PREFIX: "AF_FAILED_RETRY_",
  SUMMARY_RECORD_PREFIX: "AF_DAILY_FORWARD_",
  STARTED_AT_KEY: "AF_AUTOMATION_STARTED_AT",
  LAST_CLEANUP_KEY: "AF_LAST_CLEANUP",
  LAST_TRIGGER_REPAIR_KEY: "AF_LAST_TRIGGER_REPAIR",
  SUMMARY_RECIPIENT_KEY: "AF_SUMMARY_RECIPIENT",

  ROOT_LABEL: "AutoForward",
  DETECTED_LABEL: "AutoForward/Detected",
  FORWARDED_LABEL: "AutoForward/Forwarded",
  FAILED_LABEL: "AutoForward/Failed",
  RETRY_EXHAUSTED_LABEL: "AutoForward/Retry Exhausted"
};


const RULE_COLUMNS = [
  "Enabled",
  "Sender",
  "Match Mode",
  "Keywords (one per line)",
  "Recipients (one per line)",
  "Notes"
];


let rulesCache_ = null;


/**
 * Adds the account-facing controls whenever the bound spreadsheet opens.
 */
function onOpen() {
  try {
    ensureAutoForwardInfoSheet_();
  } catch (error) {
    Logger.log(
      "AutoForward Info tab could not be prepared: " + error.message
    );
  }

  SpreadsheetApp.getUi()
    .createMenu("⚡ AutoForward")
    .addItem("📖 Open usage guide", "openAutoForwardInfo")
    .addSeparator()
    .addItem(
      "👤 Set up or open my rule tab",
      "createOrOpenMyRuleTab"
    )
    .addItem(
      "♻️ Adopt my active legacy rule tab",
      "adoptActiveLegacyRuleTab"
    )
    .addSeparator()
    .addItem("✅ Validate my rules", "validateMyRules")
    .addItem(
      "🔎 Preview matching emails",
      "previewMatchingSecEmails"
    )
    .addItem(
      "📬 Preview pending emails",
      "previewPendingEmails"
    )
    .addSeparator()
    .addItem(
      "▶️ Start or repair my automation",
      "activateMyAutoForwarding"
    )
    .addItem("📊 Show automation status", "showMyAutoForwardStatus")
    .addItem("⏸️ Pause automation", "stopSecEmailForwarding")
    .addSeparator()
    .addItem(
      "🧹 Reset my processed history",
      "confirmAndResetMyAutoForwarding"
    )
    .addToUi();
}


function openAutoForwardInfo() {
  const sheet = ensureAutoForwardInfoSheet_();

  if (!sheet) {
    throw new Error("The AutoForward Info tab could not be opened.");
  }

  const spreadsheet = sheet.getParent();
  spreadsheet.setActiveSheet(sheet);
  spreadsheet.toast(
    "The AutoForward usage guide is open.",
    "AutoForward",
    4
  );
}


/**
 * Creates a versioned, read-only-style guide inside the bound spreadsheet.
 * The managed marker prevents an unrelated user-created tab from being erased.
 */
function ensureAutoForwardInfoSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    return null;
  }

  const markerPrefix = "AUTOFORWARD_INFO_VERSION:";
  const expectedMarker =
    markerPrefix + String(CONFIG.INFO_SHEET_VERSION);
  let sheet = spreadsheet.getSheets().find(candidate =>
    String(candidate.getRange("A1").getNote() || "")
      .startsWith(markerPrefix)
  );
  let created = false;

  if (!sheet) {
    const preferredSheet = spreadsheet.getSheetByName(
      CONFIG.INFO_SHEET_NAME
    );

    if (
      preferredSheet &&
      (
        preferredSheet.getLastRow() === 0 ||
        preferredSheet.getRange("A1").getValue() ===
          "AutoForward Information"
      )
    ) {
      sheet = preferredSheet;
    } else if (!preferredSheet) {
      sheet = spreadsheet.insertSheet(CONFIG.INFO_SHEET_NAME, 0);
      created = true;
    } else {
      sheet = spreadsheet.insertSheet(
        makeUniqueInfoSheetName_(spreadsheet),
        0
      );
      created = true;
    }
  }

  if (sheet.getRange("A1").getNote() === expectedMarker) {
    return sheet;
  }

  buildAutoForwardInfoSheet_(sheet, expectedMarker);

  if (created) {
    spreadsheet.setActiveSheet(sheet);
    spreadsheet.moveActiveSheet(1);
  }

  return sheet;
}


function makeUniqueInfoSheetName_(spreadsheet) {
  const base = "AutoForward Guide";
  let candidate = base;
  let suffix = 2;

  while (spreadsheet.getSheetByName(candidate)) {
    candidate = `${base} (${suffix})`;
    suffix++;
  }

  return candidate;
}


function buildAutoForwardInfoSheet_(sheet, marker) {
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  for (const banding of sheet.getBandings()) {
    banding.remove();
  }

  sheet.getDataRange().breakApart();
  sheet.clear();
  sheet.setConditionalFormatRules([]);
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(1);
  sheet.setTabColor("#188038");

  sheet.getRange("A1:F1").merge();
  sheet.getRange("A2:F2").merge();
  sheet.getRange("A1")
    .setValue("AutoForward Information")
    .setNote(marker);
  sheet.getRange("A2").setValue(
    "One shared Google Sheet can serve multiple Gmail users. Each user " +
    "creates, authorizes, and controls their own automation."
  );

  styleInfoTitle_(sheet);

  let row = 4;
  row = writeInfoSection_(sheet, row, "GETTING STARTED", [
    ["1. Create your tab", "Open AutoForward → Create or open my rule tab while signed in to the Gmail account that will run the automation."],
    ["2. Add rules", "Enter one sender rule per row. Use the checkbox to enable only completed rules."],
    ["3. Validate", "Choose AutoForward → Validate my rules and correct every reported row."],
    ["4. Preview", "Preview matching emails before activation. Preview actions never forward or label email."],
    ["5. Start", "Choose ▶️ Start or repair my automation and accept the requested Google permissions."],
    ["6. Confirm", "Choose 📊 Show automation status and confirm Active with 3 of 3 triggers installed."]
  ]);

  row = writeInfoSection_(sheet, row + 1, "RULE TABLE", [
    ["Enabled", "Checked rows are active. Unchecked or blank rows are ignored."],
    ["Sender", "Enter one exact sender email address, such as no-reply@example.com."],
    ["Match Mode", "All messages forwards everything from the sender. Any keyword checks the subject and plain-text body."],
    ["Keywords", "Enter one keyword or phrase per line. Leave blank only when Match Mode is All messages."],
    ["Recipients", "Enter one forwarding address per line. Commas and semicolons are also accepted."],
    ["Notes", "Optional information for users. Notes do not affect matching or forwarding."]
  ]);

  row = writeInfoSection_(sheet, row + 1, "AUTOFORWARD MENU", [
    ["Create or open", "Creates one account-owned rule tab or opens the tab already registered to your Gmail account."],
    ["Adopt legacy tab", "Converts the active old six-column rule tab. Review and validate it before activation."],
    ["Validate rules", "Checks enabled rows without reading or changing Gmail."],
    ["Preview matching", "Reads recent matching emails and writes details to the Apps Script execution log."],
    ["Preview pending", "Shows only emails the next active run would attempt."],
    ["Start or repair", "Validates rules and installs or repairs your monitor, daily summary, and self-repair watchdog triggers."],
    ["Show status", "Checks trigger health, repairs missing triggers when active, and displays account, tab, last run, and last error."],
    ["Pause", "Removes all three personal triggers but keeps your rules and processed history."],
    ["Reset history", "Pauses and clears your history. Recent email may forward again after reactivation."]
  ]);

  row = writeInfoSection_(sheet, row + 1, "GMAIL LABELS & MESSAGE STATE", [
    ["Detected", "The conversation contains at least one message that matched an enabled rule."],
    ["Forwarded", "Gmail accepted a forward request for at least one message in the conversation. It does not prove recipient delivery or mean every message in the conversation was sent."],
    ["Failed", "At least one message is waiting for another retry after the configured backoff delay."],
    ["Retry Exhausted", "At least one message reached the retry limit and needs manual review or intervention."],
    ["Source of truth", "Duplicate prevention and retry decisions use per-message IDs stored for the trigger owner. Gmail labels are thread-level review indicators."]
  ]);

  row = writeInfoSection_(sheet, row + 1, "PRIVATE COPIES & SELF-REPAIR", [
    ["Distribute copies", "Give each user their own Google Sheets copy. Installable triggers are personal and must be created once by that user."],
    ["One-click start", "If a new private copy contains exactly one valid Rules tab, Start or repair automatically assigns it to the signed-in account."],
    ["Safe fallback", "If no single valid starter tab can be identified, AutoForward creates a new account-owned rule tab instead of guessing."],
    ["Trigger watchdog", "The monitor, summary, and watchdog restore missing AutoForward triggers when at least one personal trigger remains."],
    ["Recovery limit", "If all triggers are deleted or Google authorization is revoked, the user must choose Start or repair once to restore access."]
  ]);

  row = writeInfoSection_(sheet, row + 1, "ACCOUNT ISOLATION & PRIVACY", [
    ["Mailbox ownership", "A trigger searches the Gmail mailbox of the user who created and authorized that trigger."],
    ["Rule ownership", "Automation loads the permanent sheet ID registered to the effective Gmail account, not the currently selected tab."],
    ["Editing", "The assigned account, configured administrators, and spreadsheet owner may edit a protected rule tab."],
    ["Visibility", "Protection does not hide tabs. Anyone who can view this workbook may see every tab. Use separate private spreadsheets when visibility must be restricted."],
    ["Trusted users", "Editors of a bound spreadsheet may also have access to its Apps Script project. Use this shared design only for trusted internal users."]
  ]);

  writeInfoSection_(sheet, row + 1, "TROUBLESHOOTING", [
    ["Menu is missing", "Confirm you have Editor access, reload the spreadsheet, and wait a few seconds for AutoForward to appear."],
    ["Authorization fails", "Ask your Google Workspace administrator whether Gmail, Apps Script, or external forwarding is restricted."],
    ["No rule tab", "Run Create or open my rule tab from the Gmail account that will own the automation."],
    ["No forwarding", "Validate rules, check Show automation status, confirm 3 triggers, and review Apps Script Executions for the last error."],
    ["Wrong account", "Pause the automation, sign out or switch Google accounts, reopen the sheet, and activate from the intended Gmail account."],
    ["Safety", "Start with one controlled sender, keyword, and recipient before enabling production rules."]
  ]);

  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidths(2, 5, 150);
  sheet.setRowHeight(1, 44);
  sheet.setRowHeight(2, 42);
  sheet.getDataRange()
    .setFontFamily("Arial")
    .setVerticalAlignment("middle");
  SpreadsheetApp.flush();
}


function styleInfoTitle_(sheet) {
  sheet.getRange("A1:F1")
    .setBackground("#137333")
    .setFontColor("#ffffff")
    .setFontSize(17)
    .setFontWeight("bold")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");
  sheet.getRange("A2:F2")
    .setBackground("#e6f4ea")
    .setFontColor("#3c4043")
    .setFontSize(10)
    .setWrap(true)
    .setVerticalAlignment("middle");
}


function writeInfoSection_(sheet, startRow, title, entries) {
  sheet.getRange(startRow, 1, 1, 6).merge();
  sheet.getRange(startRow, 1)
    .setValue(title)
    .setBackground("#d9ead3")
    .setFontColor("#0d652d")
    .setFontWeight("bold")
    .setFontSize(10);

  entries.forEach((entry, index) => {
    const row = startRow + index + 1;
    sheet.getRange(row, 2, 1, 5).merge();
    sheet.getRange(row, 1)
      .setValue(entry[0])
      .setBackground(index % 2 === 0 ? "#f8f9fa" : "#ffffff")
      .setFontWeight("bold")
      .setFontColor("#3c4043");
    sheet.getRange(row, 2)
      .setValue(entry[1])
      .setBackground(index % 2 === 0 ? "#f8f9fa" : "#ffffff")
      .setFontColor("#5f6368")
      .setWrap(true);
    sheet.setRowHeight(row, 42);
  });

  return startRow + entries.length + 1;
}


/**
 * Creates a polished, account-owned rule tab or opens the existing one.
 */
function createOrOpenMyRuleTab() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error(
      "Another account is being registered. Please try again shortly."
    );
  }

  let newSheet = null;

  try {
    const account = getCurrentAccount_();
    const spreadsheet = getBoundSpreadsheet_();
    prepareCopiedProjectForSpreadsheet_(spreadsheet);
    let existing = getAccountRegistration_(account);

    if (existing) {
      const existingSheet = findSheetById_(
        spreadsheet,
        existing.ruleSheetId
      );

      if (!existingSheet) {
        PropertiesService.getScriptProperties().deleteProperty(
          registrationKey_(account)
        );
        existing = null;
      }

      if (existing) {
        rememberSpreadsheetForUser_(spreadsheet.getId());
        refreshRuleSheetIdentity_(
          existingSheet,
          account,
          existing.status
        );
        spreadsheet.setActiveSheet(existingSheet);
        spreadsheet.toast(
          `Opened the rule tab assigned to ${account}.`,
          "AutoForward",
          5
        );
        return existingSheet.getName();
      }
    }

    const reusableSheet = hasRegistrationsForSpreadsheet_(
      spreadsheet.getId()
    )
      ? null
      : findReusableCopiedRuleSheet_(spreadsheet);

    if (reusableSheet) {
      const expectedName = CONFIG.RULE_SHEET_PREFIX + account;
      const reusableName = reusableSheet.getName() === expectedName
        ? expectedName
        : makeUniqueRuleSheetName_(spreadsheet, account);

      if (reusableSheet.getName() !== reusableName) {
        reusableSheet.setName(reusableName);
      }
      refreshRuleSheetIdentity_(
        reusableSheet,
        account,
        "Not activated"
      );
      protectAccountRuleSheet_(reusableSheet, account, spreadsheet);
      prepareUnactivatedAccount_();

      const copiedRegistration = {
        account,
        spreadsheetId: spreadsheet.getId(),
        ruleSheetId: reusableSheet.getSheetId(),
        ruleSheetName: reusableSheet.getName(),
        status: "Not activated",
        createdAt: Date.now(),
        activatedAt: 0,
        lastRunAt: 0,
        lastError: "",
        schemaVersion: CONFIG.RULE_SCHEMA_VERSION
      };

      saveAccountRegistration_(copiedRegistration);
      rememberSpreadsheetForUser_(spreadsheet.getId());
      spreadsheet.setActiveSheet(reusableSheet);
      spreadsheet.toast(
        `Prepared the copied rules for ${account}. You can now start ` +
        "automation.",
        "AutoForward",
        10
      );
      return reusableSheet.getName();
    }

    const sheetName = makeUniqueRuleSheetName_(spreadsheet, account);
    newSheet = spreadsheet.insertSheet(sheetName);

    formatNewRuleSheet_(newSheet, account);
    protectAccountRuleSheet_(newSheet, account, spreadsheet);
    prepareUnactivatedAccount_();

    const registration = {
      account,
      spreadsheetId: spreadsheet.getId(),
      ruleSheetId: newSheet.getSheetId(),
      ruleSheetName: newSheet.getName(),
      status: "Not activated",
      createdAt: Date.now(),
      activatedAt: 0,
      lastRunAt: 0,
      lastError: "",
      schemaVersion: CONFIG.RULE_SCHEMA_VERSION
    };

    saveAccountRegistration_(registration);
    rememberSpreadsheetForUser_(spreadsheet.getId());
    spreadsheet.setActiveSheet(newSheet);
    spreadsheet.toast(
      `Created ${newSheet.getName()} for ${account}.`,
      "AutoForward",
      8
    );

    return newSheet.getName();
  } catch (error) {
    if (newSheet) {
      try {
        newSheet.getParent().deleteSheet(newSheet);
      } catch (cleanupError) {
        Logger.log(
          "Could not remove the partially created rule tab: " +
          cleanupError.message
        );
      }
    }

    showErrorToUser_(error);
    throw error;
  } finally {
    lock.releaseLock();
  }
}


/**
 * A private workbook copy commonly contains one valid rule tab belonging to
 * the template owner. Reuse it only when the choice is unambiguous.
 */
function findReusableCopiedRuleSheet_(spreadsheet) {
  const candidates = [];

  for (const sheet of spreadsheet.getSheets()) {
    if (!sheet.getName().startsWith(CONFIG.RULE_SHEET_PREFIX)) {
      continue;
    }

    try {
      loadRulesFromSheet_(sheet);
      candidates.push(sheet);
    } catch (error) {
      Logger.log(
        `Rule tab ${sheet.getName()} is not reusable: ${error.message}`
      );
    }
  }

  return candidates.length === 1 ? candidates[0] : null;
}


function hasRegistrationsForSpreadsheet_(spreadsheetId) {
  const records = PropertiesService.getScriptProperties().getProperties();

  for (const [key, value] of Object.entries(records)) {
    if (!key.startsWith(CONFIG.ACCOUNT_REGISTRATION_PREFIX)) {
      continue;
    }

    try {
      const registration = JSON.parse(value);

      if (String(registration.spreadsheetId) === String(spreadsheetId)) {
        return true;
      }
    } catch (error) {
      // Damaged registrations are cleaned during copied-project preparation.
    }
  }

  return false;
}


/**
 * Converts the active six-column legacy tab and assigns it to this account.
 * It never activates forwarding; the user must validate and activate later.
 */
function adoptActiveLegacyRuleTab() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error(
      "Another account is being registered. Please try again shortly."
    );
  }

  try {
    const account = getCurrentAccount_();
    const spreadsheet = getBoundSpreadsheet_();
    prepareCopiedProjectForSpreadsheet_(spreadsheet);
    const sheet = spreadsheet.getActiveSheet();
    const existing = getAccountRegistration_(account);

    if (existing) {
      const registeredSheet = findSheetById_(
        spreadsheet,
        existing.ruleSheetId
      );

      if (
        registeredSheet &&
        registeredSheet.getSheetId() !== sheet.getSheetId()
      ) {
        throw new Error(
          `${account} already owns ${registeredSheet.getName()}. ` +
          "Only one rule tab can be assigned to an account."
        );
      }

      if (
        registeredSheet &&
        registeredSheet.getSheetId() === sheet.getSheetId()
      ) {
        spreadsheet.toast(
          `${sheet.getName()} is already assigned to ${account}.`,
          "AutoForward",
          6
        );
        return sheet.getName();
      }
    }

    rulesCache_ = null;
    loadRulesFromSheet_(sheet);
    upgradeLegacyRuleSheet_(sheet, account, spreadsheet);
    protectAccountRuleSheet_(sheet, account, spreadsheet);
    prepareUnactivatedAccount_();

    const registration = {
      account,
      spreadsheetId: spreadsheet.getId(),
      ruleSheetId: sheet.getSheetId(),
      ruleSheetName: sheet.getName(),
      status: "Not activated",
      createdAt: existing && existing.createdAt
        ? existing.createdAt
        : Date.now(),
      activatedAt: 0,
      lastRunAt: 0,
      lastError: "",
      schemaVersion: CONFIG.RULE_SCHEMA_VERSION
    };

    saveAccountRegistration_(registration);
    rememberSpreadsheetForUser_(spreadsheet.getId());
    spreadsheet.toast(
      `Assigned ${sheet.getName()} to ${account}. Validate it before ` +
      "activation.",
      "AutoForward",
      10
    );
    return sheet.getName();
  } catch (error) {
    showErrorToUser_(error);
    throw error;
  } finally {
    lock.releaseLock();
  }
}


/**
 * Validates the signed-in account's registered rule tab.
 */
function validateMyRules() {
  try {
    rulesCache_ = null;
    const context = getAccountContext_();
    const rules = loadRulesFromSheet_(context.ruleSheet);
    const validatedAt = formatDateTime_(new Date());

    updateRuleSheetStatus_(context.ruleSheet, {
      validation: `${validatedAt} — ${rules.length} enabled rule(s)`
    });

    context.spreadsheet.toast(
      `${rules.length} enabled rule(s) are valid.`,
      "AutoForward",
      8
    );

    return {
      account: context.account,
      sheet: context.ruleSheet.getName(),
      enabledRules: rules.length,
      senders: rules.map(rule => rule.sender)
    };
  } catch (error) {
    recordCurrentAccountError_(error);
    showErrorToUser_(error);
    throw error;
  }
}


/**
 * Activates or repairs this Gmail user's own time-based triggers.
 */
function activateMyAutoForwarding() {
  try {
    ensureRegistrationForInteractiveStart_();
    validateConfiguration_();

    const context = getAccountContext_();
    const properties = getUserProperties_();

    getOrCreateLabel_(CONFIG.ROOT_LABEL);
    getOrCreateLabel_(CONFIG.DETECTED_LABEL);
    getOrCreateLabel_(CONFIG.FORWARDED_LABEL);
    getOrCreateLabel_(CONFIG.FAILED_LABEL);
    getOrCreateLabel_(CONFIG.RETRY_EXHAUSTED_LABEL);

    if (!properties.getProperty(CONFIG.STARTED_AT_KEY)) {
      properties.setProperty(
        CONFIG.STARTED_AT_KEY,
        String(Date.now())
      );
    }

    properties.setProperty(
      CONFIG.SUMMARY_RECIPIENT_KEY,
      context.account
    );

    removeExistingTriggers_();
    const triggerHealth = ensureRequiredTriggers_();

    updateAccountRegistration_(context.account, {
      status: "Active",
      activatedAt: Date.now(),
      lastError: "",
      ruleSheetName: context.ruleSheet.getName()
    });
    updateRuleSheetStatus_(context.ruleSheet, {
      status: "Active",
      validation:
        `${formatDateTime_(new Date())} — configuration valid`
    });

    context.spreadsheet.toast(
      `Automation is active for ${context.account} with ` +
      `${triggerHealth.total} of 3 triggers.`,
      "AutoForward",
      8
    );
    Logger.log(`AutoForward activated for ${context.account}.`);
  } catch (error) {
    recordCurrentAccountError_(error);
    showErrorToUser_(error);
    throw error;
  }
}


function ensureRegistrationForInteractiveStart_() {
  const account = getCurrentAccount_();
  const spreadsheet = getBoundSpreadsheet_();

  prepareCopiedProjectForSpreadsheet_(spreadsheet);

  let registration = getAccountRegistration_(account);

  if (registration) {
    const registrationBelongsHere =
      String(registration.spreadsheetId) === spreadsheet.getId();
    const registeredSheet = registrationBelongsHere
      ? findSheetById_(spreadsheet, registration.ruleSheetId)
      : null;

    if (!registeredSheet) {
      PropertiesService.getScriptProperties().deleteProperty(
        registrationKey_(account)
      );
      prepareUnactivatedAccount_();
      registration = null;

      Logger.log(
        `Removed a stale AutoForward registration for ${account}; its ` +
        "registered rule tab no longer exists in this spreadsheet."
      );
    }
  }

  if (!registration) {
    createOrOpenMyRuleTab();
  }
}


/**
 * Backwards-compatible setup entry point.
 */
function setupSecEmailForwarding() {
  return activateMyAutoForwarding();
}


/**
 * Main function called automatically by each account's own trigger.
 */
function monitorAndForwardSecEmails() {
  const lock = LockService.getUserLock();

  if (!lock.tryLock(30000)) {
    Logger.log("Another execution for this Gmail account is running.");
    return;
  }

  let context = null;
  let forwardedCount = 0;
  let failedCount = 0;
  let detectedCount = 0;

  try {
    rulesCache_ = null;
    context = getAccountContext_();

    Logger.log(
      `AutoForward run starting for ${context.account}. ` +
      `Sheet: ${context.ruleSheet.getName()}`
    );

    cleanupProcessedRecords_();

    const properties = getUserProperties_();
    const processedRecords = properties.getProperties();
    const startedAt = Number(
      properties.getProperty(CONFIG.STARTED_AT_KEY) || 0
    );

    if (!startedAt) {
      throw new Error(
        "Automation is not initialized for this Gmail account. " +
        "Use AutoForward > Start or repair my automation."
      );
    }

    maybeRepairRequiredTriggers_(properties);

    const detectedLabel = getOrCreateLabel_(CONFIG.DETECTED_LABEL);
    const forwardedLabel = getOrCreateLabel_(CONFIG.FORWARDED_LABEL);
    const failedLabel = getOrCreateLabel_(CONFIG.FAILED_LABEL);
    const retryExhaustedLabel = getOrCreateLabel_(
      CONFIG.RETRY_EXHAUSTED_LABEL
    );

    Logger.log(
      "Labels verified: Detected, Forwarded, Failed, Retry Exhausted."
    );

    const messages = getCandidateMessages_();

    if (messages.length === 0) {
      Logger.log("No candidate messages found. Nothing to process this run.");

      const completedAt = Date.now();
      updateAccountRegistration_(context.account, {
        status: "Active",
        lastRunAt: completedAt,
        lastError: ""
      });
      updateRuleSheetStatus_(context.ruleSheet, {
        status: "Active",
        lastRun: formatDateTime_(new Date(completedAt))
      });
      return;
    }

    // ── Pass 1: identify all matching messages and apply Detected label ──
    const matchingEntries = [];

    for (const message of messages) {
      const messageId = message.getId();
      const processedKey = CONFIG.PROCESSED_PREFIX + messageId;
      const retryKey = CONFIG.FAILED_RETRY_PREFIX + messageId;

      // Skip already-forwarded messages
      if (processedRecords[processedKey]) {
        continue;
      }

      // Skip messages from before automation was activated
      if (message.getDate().getTime() < startedAt) {
        continue;
      }

      // Check retry backoff: if previously failed, respect the delay
      const retryData = getRetryData_(properties, retryKey);
      if (retryData && retryData.count >= CONFIG.MAX_RETRIES) {
        continue;
      }
      if (retryData && Date.now() - retryData.lastAttempt <
          CONFIG.RETRY_DELAY_HOURS * 60 * 60 * 1000) {
        continue;
      }

      const rule = findMatchingRule_(message);

      if (!rule) {
        continue;
      }

      const thread = message.getThread();

      // Apply Detected label immediately for all matching messages
      try {
        thread.addLabel(detectedLabel);
        detectedCount++;
      } catch (labelError) {
        Logger.log(
          `Could not apply Detected label to ${messageId}: ${labelError.message}`
        );
      }

      matchingEntries.push({
        message,
        rule,
        thread,
        messageId,
        processedKey,
        retryKey,
        retryData
      });
    }

    Logger.log(
      `Pass 1 complete: ${detectedCount} Detected label(s) applied; ` +
      `${matchingEntries.length} message(s) ready to forward.`
    );

    // ── Pass 2: forward matching messages with retry tracking ──
    const batchLimit = CONFIG.BATCH_FORWARD_LIMIT || 50;
    const forwardBatch = matchingEntries.slice(0, batchLimit);

    for (const entry of forwardBatch) {
      const { message, rule, thread, messageId, processedKey, retryKey, retryData } = entry;

      try {
        const matchedKeywords = getMatchedKeywords_(message, rule);

        Logger.log(
          `Forwarding "${message.getSubject()}" from ${rule.sender}. ` +
          `Matched: ${matchedKeywords.join(", ") || "all messages"}`
        );

        message.forward(rule.recipients.join(","));

        const processedAt = String(Date.now());
        properties.setProperty(processedKey, processedAt);
        forwardedCount++;

        // Clear this message's retry state. Thread failure labels are then
        // recalculated from every message in the conversation.
        if (retryData) {
          properties.deleteProperty(retryKey);
        }

        // Apply Forwarded and recalculate both retry-state labels.
        try {
          thread.addLabel(forwardedLabel);
          syncThreadFailureLabels_(
            thread,
            properties,
            failedLabel,
            retryExhaustedLabel
          );
        } catch (labelError) {
          Logger.log(
            `Message ${messageId} was forwarded, but labels could not ` +
            `be updated: ${labelError.message}`
          );
        }

        try {
          recordForwardForDailySummary_(
            properties,
            message,
            rule,
            processedAt
          );
        } catch (summaryError) {
          Logger.log(
            `Message ${messageId} was forwarded, but its summary ` +
            `record failed: ${summaryError.message}`
          );
        }
      } catch (error) {
        Logger.log(
          `Failed forwarding ${messageId}: ${error.message}`
        );
        failedCount++;

        // Increment retry counter
        const newRetryCount = (retryData ? retryData.count : 0) + 1;
        setRetryData_(properties, retryKey, newRetryCount);

        // Recalculate retry labels for the entire conversation. Failed means
        // another retry is pending; Retry Exhausted means manual review.
        try {
          syncThreadFailureLabels_(
            thread,
            properties,
            failedLabel,
            retryExhaustedLabel
          );
        } catch (labelError) {
          Logger.log(
            `Could not synchronize failure labels for ${messageId}: ` +
            labelError.message
          );
        }

        // The retry record, rather than a successful-processed marker, keeps
        // this message out of subsequent attempts.
        if (newRetryCount >= CONFIG.MAX_RETRIES) {
          Logger.log(
            `Message ${messageId} has failed ${newRetryCount} times. ` +
            `It will not be retried again within the configured search window.`
          );
        }
      }
    }

    const completedAt = Date.now();
    updateAccountRegistration_(context.account, {
      status: "Active",
      lastRunAt: completedAt,
      lastError: ""
    });
    updateRuleSheetStatus_(context.ruleSheet, {
      status: "Active",
      lastRun: formatDateTime_(new Date(completedAt))
    });

    Logger.log(
      `Run complete: ${detectedCount} detected, ${forwardedCount} forwarded, ` +
      `${failedCount} failed, ${matchingEntries.length - forwardBatch.length} ` +
      `deferred to next run.`
    );
  } catch (error) {
    recordCurrentAccountError_(error, context);
    throw error;
  } finally {
    lock.releaseLock();
  }
}


/**
 * Searches Gmail for messages from every enabled sender in this account's tab.
 * Finds eligible inbox, configured spam, and retry-pending messages.
 * Forwarded labels are intentionally not excluded because labels belong to
 * threads while successful-processing records belong to individual messages.
 */
function getCandidateMessages_() {
  const rules = getRules_();
  const uniqueSenders = Array.from(
    new Set(rules.map(rule => normalizeEmail_(rule.sender)))
  );

  if (uniqueSenders.length === 0) {
    Logger.log("No enabled senders configured. Nothing to search.");
    return [];
  }

  // Build sender OR-group, capped to avoid hitting Gmail query length limits
  const maxSendersPerQuery = 50;
  const senderBatches = [];
  for (let i = 0; i < uniqueSenders.length; i += maxSendersPerQuery) {
    senderBatches.push(
      uniqueSenders
        .slice(i, i + maxSendersPerQuery)
        .map(sender => `from:${sender}`)
        .join(" ")
    );
  }

  const monitoredLocations = CONFIG.INCLUDE_SPAM
    ? `{in:inbox in:spam label:"${CONFIG.FAILED_LABEL}"}`
    : `{in:inbox label:"${CONFIG.FAILED_LABEL}"}`;
  const baseQuery =
    `newer_than:${CONFIG.SEARCH_LOOKBACK_DAYS}d ` +
    `${monitoredLocations}`;

  const messageMap = new Map();
  let totalThreadsFound = 0;

  for (const senderGroup of senderBatches) {
    const query = `${baseQuery} {${senderGroup}}`;
    Logger.log(`Searching Gmail: ${query}`);

    let start = 0;
    const batchSize = 50;

    while (totalThreadsFound < CONFIG.MAX_THREADS_PER_RUN) {
      const amount = Math.min(
        batchSize,
        CONFIG.MAX_THREADS_PER_RUN - totalThreadsFound
      );

      const threads = GmailApp.search(query, start, amount);

      if (threads.length === 0) {
        break;
      }

      totalThreadsFound += threads.length;

      for (const thread of threads) {
        const isFailedThread = threadHasLabel_(
          thread,
          CONFIG.FAILED_LABEL
        );
        const isMonitoredSpam =
          CONFIG.INCLUDE_SPAM && thread.isInSpam();
        const messages = thread.getMessages();

        for (const message of messages) {
          if (
            (
              message.isInInbox() ||
              isMonitoredSpam ||
              isFailedThread
            ) &&
            !message.isInTrash() &&
            !message.isDraft()
          ) {
            messageMap.set(message.getId(), message);
          }
        }
      }

      start += threads.length;

      if (threads.length < amount) {
        break;
      }
    }
  }

  const messages = Array.from(messageMap.values());
  messages.sort(
    (first, second) =>
      first.getDate().getTime() - second.getDate().getTime()
  );

  Logger.log(
    `Gmail search complete: ${totalThreadsFound} thread(s) scanned, ` +
    `${messages.length} unique candidate message(s) found.`
  );

  return messages;
}


function findMatchingRule_(message) {
  const sender = extractEmailAddress_(message.getFrom());

  for (const rule of getRules_()) {
    if (sender !== normalizeEmail_(rule.sender)) {
      continue;
    }

    if (rule.matchAll) {
      return rule;
    }

    if (getMatchedKeywords_(message, rule).length > 0) {
      return rule;
    }
  }

  return null;
}


function getMatchedKeywords_(message, rule) {
  const searchableText =
    `${message.getSubject() || ""}\n${message.getPlainBody() || ""}`;

  return rule.keywords.filter(keyword =>
    matchesKeyword_(searchableText, keyword)
  );
}


/**
 * Short uppercase codes are matched as whole words.
 */
function matchesKeyword_(text, keyword) {
  const cleanKeyword = String(keyword).trim();

  if (!cleanKeyword) {
    return false;
  }

  const isCode =
    /^[A-Z0-9]+$/.test(cleanKeyword) && cleanKeyword.length <= 10;

  if (isCode) {
    const escapedKeyword = cleanKeyword.replace(
      /[.*+?^${}()|[\]\\]/g,
      "\\$&"
    );
    return new RegExp(`\\b${escapedKeyword}\\b`, "i").test(text);
  }

  return text
    .toLowerCase()
    .includes(cleanKeyword.toLowerCase());
}


function extractEmailAddress_(fromValue) {
  const match = String(fromValue).match(/<([^>]+)>/);
  return normalizeEmail_(match ? match[1] : fromValue);
}


function normalizeEmail_(email) {
  return String(email || "").trim().toLowerCase();
}


/**
 * Returns the trigger owner/current menu user's normalized Gmail address.
 */
function getCurrentAccount_() {
  const effectiveUser = normalizeEmail_(
    Session.getEffectiveUser().getEmail()
  );

  if (!isValidEmail_(effectiveUser)) {
    throw new Error(
      "Google did not provide your signed-in email address. Open this " +
      "sheet using the Gmail account that will run AutoForward, then " +
      "authorize the requested permissions."
    );
  }

  return effectiveUser;
}


function getBoundSpreadsheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error(
      "This action must be run from the Google Sheet containing the " +
      "AutoForward Apps Script."
    );
  }

  return spreadsheet;
}


function getAccountContext_() {
  const account = getCurrentAccount_();
  const registration = getAccountRegistration_(account);

  if (!registration) {
    throw new Error(
      `No rule tab is registered for ${account}. Use AutoForward > ` +
      "Create or open my rule tab first."
    );
  }

  if (registration.account !== account) {
    throw new Error("The account registration owner does not match.");
  }

  const spreadsheet = SpreadsheetApp.openById(
    registration.spreadsheetId
  );
  const ruleSheet = findSheetById_(
    spreadsheet,
    registration.ruleSheetId
  );

  if (!ruleSheet) {
    throw new Error(
      `The registered rule tab for ${account} no longer exists. ` +
      "Open the spreadsheet and choose AutoForward > Start or repair my " +
      "automation to rebuild the local registration."
    );
  }

  return {
    account,
    registration,
    spreadsheet,
    ruleSheet
  };
}


function registrationKey_(account) {
  return CONFIG.ACCOUNT_REGISTRATION_PREFIX + normalizeEmail_(account);
}


/**
 * A copied bound script may inherit registrations that point to the source
 * workbook. Remove only those foreign registrations and reset only the current
 * user's state in this copied script project. The source workbook is untouched.
 */
function prepareCopiedProjectForSpreadsheet_(spreadsheet) {
  const spreadsheetId = spreadsheet.getId();
  const scriptProperties = PropertiesService.getScriptProperties();
  const registrations = scriptProperties.getProperties();
  let removedCount = 0;

  for (const [key, value] of Object.entries(registrations)) {
    if (!key.startsWith(CONFIG.ACCOUNT_REGISTRATION_PREFIX)) {
      continue;
    }

    try {
      const registration = JSON.parse(value);

      if (
        registration.spreadsheetId &&
        String(registration.spreadsheetId) !== spreadsheetId
      ) {
        scriptProperties.deleteProperty(key);
        removedCount++;
      }
    } catch (error) {
      scriptProperties.deleteProperty(key);
      removedCount++;
    }
  }

  if (removedCount > 0) {
    clearCurrentUserAutomationStateForCopy_();
    Logger.log(
      `Prepared copied AutoForward project: removed ${removedCount} ` +
      "registration(s) belonging to another spreadsheet."
    );
  }
}


function clearCurrentUserAutomationStateForCopy_() {
  removeExistingTriggers_();

  const properties = getUserProperties_();
  const records = properties.getProperties();

  for (const key of Object.keys(records)) {
    if (
      key.startsWith(CONFIG.PROCESSED_PREFIX) ||
      key.startsWith(CONFIG.FAILED_RETRY_PREFIX) ||
      key.startsWith(CONFIG.SUMMARY_RECORD_PREFIX) ||
      key === CONFIG.STARTED_AT_KEY ||
      key === CONFIG.LAST_CLEANUP_KEY ||
      key === CONFIG.SUMMARY_RECIPIENT_KEY ||
      key === "AF_BOUND_SPREADSHEET_ID"
    ) {
      properties.deleteProperty(key);
    }
  }
}


function getAccountRegistration_(account) {
  const value = PropertiesService
    .getScriptProperties()
    .getProperty(registrationKey_(account));

  if (!value) {
    return null;
  }

  try {
    const registration = JSON.parse(value);
    registration.account = normalizeEmail_(registration.account);
    registration.ruleSheetId = Number(registration.ruleSheetId);
    return registration;
  } catch (error) {
    throw new Error(
      `The AutoForward registration for ${account} is damaged: ` +
      error.message
    );
  }
}


function saveAccountRegistration_(registration) {
  const account = normalizeEmail_(registration.account);

  if (!isValidEmail_(account)) {
    throw new Error("Cannot save a registration without a valid account.");
  }

  const cleanRegistration = Object.assign({}, registration, {
    account,
    ruleSheetId: Number(registration.ruleSheetId),
    schemaVersion: CONFIG.RULE_SCHEMA_VERSION
  });

  PropertiesService.getScriptProperties().setProperty(
    registrationKey_(account),
    JSON.stringify(cleanRegistration)
  );

  return cleanRegistration;
}


function updateAccountRegistration_(account, updates) {
  const registration = getAccountRegistration_(account);

  if (!registration) {
    throw new Error(`No AutoForward registration exists for ${account}.`);
  }

  return saveAccountRegistration_(
    Object.assign({}, registration, updates)
  );
}


function rememberSpreadsheetForUser_(spreadsheetId) {
  getUserProperties_().setProperty(
    "AF_BOUND_SPREADSHEET_ID",
    String(spreadsheetId)
  );
}


function getUserProperties_() {
  return PropertiesService.getUserProperties();
}


function getRetryData_(properties, retryKey) {
  const raw = properties.getProperty(retryKey);

  if (!raw) {
    return null;
  }

  try {
    return JSON.parse(raw);
  } catch (error) {
    return null;
  }
}


function setRetryData_(properties, retryKey, count) {
  properties.setProperty(
    retryKey,
    JSON.stringify({
      count: Number(count),
      lastAttempt: Date.now()
    })
  );
}


/**
 * Keeps human-facing failure labels aligned with the per-message retry state.
 * A thread may legitimately have Forwarded together with Failed or
 * Retry Exhausted when different messages in the conversation have different
 * outcomes.
 */
function syncThreadFailureLabels_(
  thread,
  properties,
  failedLabel,
  retryExhaustedLabel
) {
  let hasPendingRetry = false;
  let hasExhaustedRetry = false;

  for (const threadMessage of thread.getMessages()) {
    const retryKey =
      CONFIG.FAILED_RETRY_PREFIX + threadMessage.getId();
    const retryData = getRetryData_(properties, retryKey);

    if (!retryData || !Number(retryData.count)) {
      continue;
    }

    if (Number(retryData.count) >= CONFIG.MAX_RETRIES) {
      hasExhaustedRetry = true;
    } else {
      hasPendingRetry = true;
    }
  }

  if (hasPendingRetry) {
    thread.addLabel(failedLabel);
  } else {
    thread.removeLabel(failedLabel);
  }

  if (hasExhaustedRetry) {
    thread.addLabel(retryExhaustedLabel);
  } else {
    thread.removeLabel(retryExhaustedLabel);
  }
}


function findSheetById_(spreadsheet, sheetId) {
  const numericId = Number(sheetId);
  return spreadsheet
    .getSheets()
    .find(sheet => sheet.getSheetId() === numericId) || null;
}


function makeUniqueRuleSheetName_(spreadsheet, account) {
  const cleanAccount = String(account)
    .replace(/[\\/?*\[\]:]/g, "-")
    .trim();
  const base = (CONFIG.RULE_SHEET_PREFIX + cleanAccount).slice(0, 100);
  let candidate = base;
  let suffix = 2;

  while (spreadsheet.getSheetByName(candidate)) {
    const suffixText = ` (${suffix})`;
    candidate = base.slice(0, 100 - suffixText.length) + suffixText;
    suffix++;
  }

  return candidate;
}


/**
 * Creates the visual rule editor used by every account.
 */
function formatNewRuleSheet_(sheet, account) {
  const requiredRows = CONFIG.RULE_HEADER_ROW + CONFIG.RULE_DATA_ROWS;

  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(
      sheet.getMaxRows(),
      requiredRows - sheet.getMaxRows()
    );
  }

  if (sheet.getMaxColumns() < RULE_COLUMNS.length) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      RULE_COLUMNS.length - sheet.getMaxColumns()
    );
  }

  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(CONFIG.RULE_HEADER_ROW);
  sheet.setTabColor("#1a73e8");

  sheet.getRange("A1:F1").merge();
  sheet.getRange("B2:C2").merge();
  sheet.getRange("E2:F2").merge();
  sheet.getRange("B3:C3").merge();
  sheet.getRange("E3:F3").merge();
  sheet.getRange("A4:F4").merge();

  sheet.getRange("A1").setValue("AutoForward Rules");
  sheet.getRange("A2").setValue("Assigned Gmail");
  sheet.getRange("B2").setValue(account);
  sheet.getRange("D2").setValue("Automation");
  sheet.getRange("E2").setValue("Not activated");
  sheet.getRange("A3").setValue("Last validation");
  sheet.getRange("B3").setValue("Not yet validated");
  sheet.getRange("D3").setValue("Last run");
  sheet.getRange("E3").setValue("Never");
  sheet.getRange("A4").setValue(
    "Add one rule per row. Choose All messages to forward every email " +
    "from a sender, or Any keyword to match the subject and message body."
  );
  sheet
    .getRange(CONFIG.RULE_HEADER_ROW, 1, 1, RULE_COLUMNS.length)
    .setValues([RULE_COLUMNS]);

  sheet.getRange("A1:F1")
    .setBackground("#174ea6")
    .setFontColor("#ffffff")
    .setFontFamily("Arial")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");

  sheet.getRange("A2:F3")
    .setBackground("#e8f0fe")
    .setFontFamily("Arial")
    .setFontColor("#202124")
    .setVerticalAlignment("middle");
  sheet.getRangeList(["A2", "D2", "A3", "D3"])
    .setFontWeight("bold")
    .setFontColor("#174ea6");

  sheet.getRange("A4:F4")
    .setBackground("#f8f9fa")
    .setFontColor("#5f6368")
    .setFontStyle("italic")
    .setWrap(true)
    .setVerticalAlignment("middle");

  sheet
    .getRange(CONFIG.RULE_HEADER_ROW, 1, 1, RULE_COLUMNS.length)
    .setBackground("#1a73e8")
    .setFontColor("#ffffff")
    .setFontFamily("Arial")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  const dataRange = sheet.getRange(
    CONFIG.RULE_HEADER_ROW + 1,
    1,
    CONFIG.RULE_DATA_ROWS,
    RULE_COLUMNS.length
  );
  dataRange
    .setFontFamily("Arial")
    .setFontColor("#202124")
    .setVerticalAlignment("top");
  if (sheet.getBandings().length === 0) {
    dataRange.applyRowBanding(
      SpreadsheetApp.BandingTheme.LIGHT_GREY,
      false,
      false
    );
  }
  sheet.getRange(
    CONFIG.RULE_HEADER_ROW + 1,
    1,
    CONFIG.RULE_DATA_ROWS,
    1
  ).setHorizontalAlignment("center");
  sheet.getRange(
    CONFIG.RULE_HEADER_ROW + 1,
    3,
    CONFIG.RULE_DATA_ROWS,
    1
  ).setHorizontalAlignment("center");
  sheet.getRange(
    CONFIG.RULE_HEADER_ROW + 1,
    4,
    CONFIG.RULE_DATA_ROWS,
    3
  ).setWrap(true);

  const checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .build();
  sheet.getRange(
    CONFIG.RULE_HEADER_ROW + 1,
    1,
    CONFIG.RULE_DATA_ROWS,
    1
  ).setDataValidation(checkboxRule);

  const matchModeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["All messages", "Any keyword"], true)
    .setAllowInvalid(false)
    .setHelpText(
      "All messages ignores Keywords. Any keyword searches the subject " +
      "and plain-text message body."
    )
    .build();
  sheet.getRange(
    CONFIG.RULE_HEADER_ROW + 1,
    3,
    CONFIG.RULE_DATA_ROWS,
    1
  ).setDataValidation(matchModeRule);

  const senderRule = SpreadsheetApp.newDataValidation()
    .requireTextIsEmail()
    .setAllowInvalid(true)
    .setHelpText("Enter one sender email address in each rule row.")
    .build();
  sheet.getRange(
    CONFIG.RULE_HEADER_ROW + 1,
    2,
    CONFIG.RULE_DATA_ROWS,
    1
  ).setDataValidation(senderRule);

  const dataAddress =
    `A${CONFIG.RULE_HEADER_ROW + 1}:F${requiredRows}`;
  const disabledRule = SpreadsheetApp
    .newConditionalFormatRule()
    .whenFormulaSatisfied(
      `=$A${CONFIG.RULE_HEADER_ROW + 1}=FALSE`
    )
    .setBackground("#f1f3f4")
    .setFontColor("#80868b")
    .setRanges([sheet.getRange(dataAddress)])
    .build();
  const missingSenderRule = SpreadsheetApp
    .newConditionalFormatRule()
    .whenFormulaSatisfied(
      `=AND($A${CONFIG.RULE_HEADER_ROW + 1}=TRUE,` +
      `$B${CONFIG.RULE_HEADER_ROW + 1}="")`
    )
    .setBackground("#fce8e6")
    .setRanges([
      sheet.getRange(
        CONFIG.RULE_HEADER_ROW + 1,
        2,
        CONFIG.RULE_DATA_ROWS,
        1
      )
    ])
    .build();
  sheet.setConditionalFormatRules([
    missingSenderRule,
    disabledRule
  ]);

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 240);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 260);
  sheet.setColumnWidth(5, 300);
  sheet.setColumnWidth(6, 220);
  sheet.setRowHeight(1, 42);
  sheet.setRowHeights(2, 2, 30);
  sheet.setRowHeight(4, 48);
  sheet.setRowHeight(CONFIG.RULE_HEADER_ROW, 38);

  const filterRange = sheet.getRange(
    CONFIG.RULE_HEADER_ROW,
    1,
    CONFIG.RULE_DATA_ROWS + 1,
    RULE_COLUMNS.length
  );
  if (!sheet.getFilter()) {
    filterRange.createFilter();
  }
  SpreadsheetApp.flush();
}


function upgradeLegacyRuleSheet_(sheet, account, spreadsheet) {
  let headerRow = findRuleHeaderRow_(sheet);
  let headers = sheet
    .getRange(headerRow, 1, 1, Math.max(sheet.getLastColumn(), 6))
    .getDisplayValues()[0]
    .map(normalizeRuleHeader_);
  const legacyMatchAllIndex = headers.indexOf("match all");
  const matchModeIndex = headers.indexOf("match mode");
  const expectedLeadingHeaders = [
    "enabled",
    "sender",
    legacyMatchAllIndex >= 0 ? "match all" : "match mode",
    "keywords",
    "recipients"
  ];

  if (
    !expectedLeadingHeaders.every(
      (header, index) => headers[index] === header
    )
  ) {
    throw new Error(
      "The legacy columns must be ordered as Enabled, Sender, Match " +
      "All/Match Mode, Keywords, Recipients, and optional Notes."
    );
  }

  if (headerRow !== 1 && headerRow !== CONFIG.RULE_HEADER_ROW) {
    throw new Error(
      "The legacy header must be on row 1 before it can be upgraded."
    );
  }

  if (headerRow === 1) {
    const oldLastRow = sheet.getLastRow();
    const modeColumn = legacyMatchAllIndex >= 0
      ? legacyMatchAllIndex + 1
      : matchModeIndex + 1;

    if (modeColumn < 1) {
      throw new Error(
        "The active legacy tab must contain Match All or Match Mode."
      );
    }

    const modeValues = oldLastRow > 1
      ? sheet.getRange(2, modeColumn, oldLastRow - 1, 1).getValues()
      : [];

    sheet.insertRowsBefore(1, CONFIG.RULE_HEADER_ROW - 1);
    headerRow = CONFIG.RULE_HEADER_ROW;

    if (modeValues.length > 0) {
      const convertedModes = modeValues.map((row, index) => {
        if (legacyMatchAllIndex >= 0) {
          return [
            parseRuleBoolean_(
              row[0],
              false,
              "Match All",
              index + 2
            )
              ? "All messages"
              : "Any keyword"
          ];
        }

        return [
          parseMatchMode_(row[0], index + 2)
            ? "All messages"
            : "Any keyword"
        ];
      });

      sheet
        .getRange(headerRow + 1, 3, convertedModes.length, 1)
        .setValues(convertedModes);
    }
  }

  headers = sheet
    .getRange(headerRow, 1, 1, Math.max(sheet.getLastColumn(), 6))
    .getDisplayValues()[0]
    .map(normalizeRuleHeader_);

  if (
    !headers.includes("match mode") &&
    !headers.includes("match all")
  ) {
    throw new Error("The rule tab does not contain a matching column.");
  }

  formatNewRuleSheet_(sheet, account);

  const preferredName =
    (CONFIG.RULE_SHEET_PREFIX + account).slice(0, 100);

  if (sheet.getName() !== preferredName) {
    sheet.setName(makeUniqueRuleSheetName_(spreadsheet, account));
  }
}


function protectAccountRuleSheet_(sheet, account, spreadsheet) {
  const allowedEmails = new Set(
    [account]
      .concat(CONFIG.ADMIN_EMAILS || [])
      .map(normalizeEmail_)
      .filter(isValidEmail_)
  );

  try {
    const owner = spreadsheet.getOwner();
    if (owner) {
      const ownerEmail = normalizeEmail_(owner.getEmail());
      if (isValidEmail_(ownerEmail)) {
        allowedEmails.add(ownerEmail);
      }
    }
  } catch (error) {
    Logger.log("Spreadsheet owner could not be read: " + error.message);
  }

  const description = `AutoForward rule tab for ${account}`;
  const managedProtections = sheet
    .getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .filter(item =>
      String(item.getDescription() || "")
        .startsWith("AutoForward rule tab for ")
    );
  const existingProtection = managedProtections[0] || null;

  for (const staleProtection of managedProtections.slice(1)) {
    staleProtection.remove();
  }

  const protection = (existingProtection || sheet.protect())
    .setDescription(description)
    .setWarningOnly(false);

  protection.addEditors(Array.from(allowedEmails));

  const removableEditors = protection
    .getEditors()
    .filter(user =>
      !allowedEmails.has(normalizeEmail_(user.getEmail()))
    );

  if (removableEditors.length > 0) {
    protection.removeEditors(removableEditors);
  }

  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}


function refreshRuleSheetIdentity_(sheet, account, status) {
  sheet.getRange("B2").setValue(account);
  sheet.getRange("E2").setValue(status || "Not activated");
}


function updateRuleSheetStatus_(sheet, updates) {
  if (!sheet) {
    return;
  }

  if (Object.prototype.hasOwnProperty.call(updates, "status")) {
    sheet.getRange("E2").setValue(updates.status);
  }

  if (Object.prototype.hasOwnProperty.call(updates, "validation")) {
    sheet.getRange("B3").setValue(updates.validation);
  }

  if (Object.prototype.hasOwnProperty.call(updates, "lastRun")) {
    sheet.getRange("E3").setValue(updates.lastRun);
  }
}


/**
 * Tests the current account's registered Google Sheet without email changes.
 */
function testGoogleSheetConnection() {
  return validateMyRules();
}


function getRules_() {
  if (!rulesCache_) {
    const context = getAccountContext_();
    rulesCache_ = loadRulesFromSheet_(context.ruleSheet);
  }

  return rulesCache_;
}


function loadRulesFromSheet_(sheet) {
  const headerRow = findRuleHeaderRow_(sheet);
  const lastRow = sheet.getLastRow();

  if (lastRow <= headerRow) {
    throw new Error(
      `No rule rows were found in ${sheet.getName()}.`
    );
  }

  const columnCount = Math.max(
    sheet.getLastColumn(),
    RULE_COLUMNS.length
  );
  const headers = sheet
    .getRange(headerRow, 1, 1, columnCount)
    .getDisplayValues()[0]
    .map(normalizeRuleHeader_);

  const requiredHeaders = [
    "enabled",
    "sender",
    "keywords",
    "recipients"
  ];
  const columns = {};

  for (const header of requiredHeaders) {
    const index = headers.indexOf(header);

    if (index < 0) {
      throw new Error(
        `Missing required column "${header}" in ${sheet.getName()}.`
      );
    }

    columns[header] = index;
  }

  const matchModeIndex = headers.indexOf("match mode");
  const legacyMatchAllIndex = headers.indexOf("match all");

  if (matchModeIndex < 0 && legacyMatchAllIndex < 0) {
    throw new Error(
      `Missing required column "Match Mode" in ${sheet.getName()}.`
    );
  }

  const values = sheet
    .getRange(
      headerRow + 1,
      1,
      lastRow - headerRow,
      columnCount
    )
    .getValues();
  const rules = [];

  for (let index = 0; index < values.length; index++) {
    const row = values[index];
    const rowNumber = headerRow + index + 1;

    if (row.every(value => !String(value).trim())) {
      continue;
    }

    const enabled = parseRuleBoolean_(
      row[columns.enabled],
      false,
      "Enabled",
      rowNumber
    );

    if (!enabled) {
      continue;
    }

    let matchAll;

    if (matchModeIndex >= 0) {
      matchAll = parseMatchMode_(row[matchModeIndex], rowNumber);
    } else {
      matchAll = parseRuleBoolean_(
        row[legacyMatchAllIndex],
        false,
        "Match All",
        rowNumber
      );
    }

    rules.push({
      sender: normalizeEmail_(row[columns.sender]),
      matchAll,
      keywords: splitRuleList_(row[columns.keywords], false),
      recipients: Array.from(
        new Set(
          splitRuleList_(row[columns.recipients], true)
            .map(normalizeEmail_)
        )
      ),
      sourceRow: rowNumber
    });
  }

  if (rules.length === 0) {
    throw new Error(
      `No enabled rules were found in ${sheet.getName()}.`
    );
  }

  validateRules_(rules);
  return rules;
}


function findRuleHeaderRow_(sheet) {
  const scanRows = Math.min(20, Math.max(sheet.getLastRow(), 1));
  const scanColumns = Math.max(
    RULE_COLUMNS.length,
    Math.min(sheet.getLastColumn(), 20)
  );
  const values = sheet
    .getRange(1, 1, scanRows, scanColumns)
    .getDisplayValues();

  for (let index = 0; index < values.length; index++) {
    const headers = values[index].map(normalizeRuleHeader_);

    if (
      headers.includes("enabled") &&
      headers.includes("sender") &&
      headers.includes("recipients")
    ) {
      return index + 1;
    }
  }

  throw new Error(
    `Could not find the AutoForward rule headers in ${sheet.getName()}.`
  );
}


function validateRules_(rules) {
  const matchAllSeen = new Set();

  for (const rule of rules) {
    const rowLabel = `row ${rule.sourceRow}`;

    if (!isValidEmail_(rule.sender)) {
      throw new Error(
        `Invalid sender on ${rowLabel}: ${rule.sender || "(blank)"}`
      );
    }

    if (matchAllSeen.has(rule.sender)) {
      throw new Error(
        `The All messages rule above ${rowLabel} already captures every ` +
        `email from ${rule.sender}. Move it below specific rules or ` +
        "disable the redundant row."
      );
    }

    if (!rule.matchAll && !rule.keywords.length) {
      throw new Error(
        `No keywords are configured on ${rowLabel} for ${rule.sender}. ` +
        "Choose All messages or enter at least one keyword."
      );
    }

    if (!rule.recipients.length) {
      throw new Error(
        `No recipients are configured on ${rowLabel} for ${rule.sender}.`
      );
    }

    const invalidRecipient = rule.recipients.find(
      recipient => !isValidEmail_(recipient)
    );

    if (invalidRecipient) {
      throw new Error(
        `Invalid recipient on ${rowLabel}: ${invalidRecipient}`
      );
    }

    if (rule.recipients.length > 50) {
      throw new Error(
        `Too many recipients on ${rowLabel}. Apps Script permits no more ` +
        "than 50 recipients per message."
      );
    }

    if (rule.matchAll) {
      matchAllSeen.add(rule.sender);
    }
  }
}


function normalizeRuleHeader_(value) {
  return String(value || "")
    .replace(/\([^)]*\)/g, "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}


function parseMatchMode_(value, rowNumber) {
  const cleanValue = String(value || "").trim().toLowerCase();

  if (["all messages", "all", "match all"].includes(cleanValue)) {
    return true;
  }

  if (["any keyword", "keywords", "keyword"].includes(cleanValue)) {
    return false;
  }

  throw new Error(
    `Invalid Match Mode on row ${rowNumber}: ${value || "(blank)"}`
  );
}


function parseRuleBoolean_(
  value,
  defaultValue,
  columnName,
  rowNumber
) {
  if (typeof value === "boolean") {
    return value;
  }

  const cleanValue = String(value || "").trim().toUpperCase();

  if (!cleanValue) {
    return defaultValue;
  }

  if (["TRUE", "YES", "Y", "1"].includes(cleanValue)) {
    return true;
  }

  if (["FALSE", "NO", "N", "0"].includes(cleanValue)) {
    return false;
  }

  throw new Error(
    `Invalid ${columnName} value on row ${rowNumber}: ${value}`
  );
}


function splitRuleList_(value, allowCommas) {
  const separator = allowCommas ? /[\n;,]+/ : /[\n;]+/;

  return String(value || "")
    .split(separator)
    .map(item => item.trim())
    .filter(Boolean);
}


function isValidEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(
    normalizeEmail_(value)
  );
}


function getOrCreateLabel_(labelName) {
  return GmailApp.getUserLabelByName(labelName) ||
    GmailApp.createLabel(labelName);
}


function threadHasLabel_(thread, labelName) {
  return thread.getLabels().some(
    label => label.getName() === labelName
  );
}


function recordForwardForDailySummary_(
  properties,
  message,
  rule,
  processedAt
) {
  const record = {
    forwardedAt: Number(processedAt),
    sender: extractEmailAddress_(message.getFrom()),
    subject: String(message.getSubject() || "(no subject)")
      .slice(0, 300),
    recipients: rule.recipients
  };

  properties.setProperty(
    CONFIG.SUMMARY_RECORD_PREFIX + message.getId(),
    JSON.stringify(record)
  );
}


function getSummaryRecipient_() {
  const properties = getUserProperties_();
  const savedRecipient = normalizeEmail_(
    properties.getProperty(CONFIG.SUMMARY_RECIPIENT_KEY)
  );

  if (isValidEmail_(savedRecipient)) {
    return savedRecipient;
  }

  const account = getCurrentAccount_();
  properties.setProperty(CONFIG.SUMMARY_RECIPIENT_KEY, account);
  return account;
}


/**
 * Logs every recent matching email without changing Gmail.
 */
function previewMatchingSecEmails() {
  try {
    rulesCache_ = null;
    validateConfiguration_();

    const messages = getCandidateMessages_();
    let matchCount = 0;

    for (const message of messages) {
      const rule = findMatchingRule_(message);

      if (!rule) {
        continue;
      }

      matchCount++;
      Logger.log(JSON.stringify({
        date: message.getDate(),
        sender: extractEmailAddress_(message.getFrom()),
        subject: message.getSubject(),
        matchedKeywords: getMatchedKeywords_(message, rule),
        recipients: rule.recipients
      }));
    }

    getAccountContext_().spreadsheet.toast(
      `${matchCount} matching email(s) found. See the execution log.`,
      "AutoForward preview",
      10
    );
    Logger.log(`${matchCount} matching email(s) found.`);
    return matchCount;
  } catch (error) {
    recordCurrentAccountError_(error);
    showErrorToUser_(error);
    throw error;
  }
}


/**
 * Logs only messages that the next live run would attempt.
 */
function previewPendingEmails() {
  try {
    rulesCache_ = null;
    validateConfiguration_();

    const properties = getUserProperties_();
    const records = properties.getProperties();
    const startedAt = Number(records[CONFIG.STARTED_AT_KEY] || 0);

    if (!startedAt) {
      throw new Error(
        "Automation has not been activated for this Gmail account."
      );
    }

    const messages = getCandidateMessages_();
    let pendingCount = 0;

    for (const message of messages) {
      const messageId = message.getId();
      const processedKey = CONFIG.PROCESSED_PREFIX + messageId;
      const retryKey = CONFIG.FAILED_RETRY_PREFIX + messageId;
      const retryData = getRetryData_(properties, retryKey);

      if (
        records[processedKey] ||
        message.getDate().getTime() < startedAt ||
        (
          retryData &&
          (
            retryData.count >= CONFIG.MAX_RETRIES ||
            Date.now() - retryData.lastAttempt <
              CONFIG.RETRY_DELAY_HOURS * 60 * 60 * 1000
          )
        )
      ) {
        continue;
      }

      const rule = findMatchingRule_(message);

      if (!rule) {
        continue;
      }

      pendingCount++;
      Logger.log(JSON.stringify({
        date: message.getDate(),
        sender: extractEmailAddress_(message.getFrom()),
        subject: message.getSubject(),
        matchedKeywords: getMatchedKeywords_(message, rule),
        recipients: rule.recipients,
        retryingFailed: threadHasLabel_(
          message.getThread(),
          CONFIG.FAILED_LABEL
        )
      }));
    }

    getAccountContext_().spreadsheet.toast(
      `${pendingCount} pending email(s). See the execution log.`,
      "AutoForward preview",
      10
    );
    Logger.log(`${pendingCount} pending email(s) would be attempted.`);
    return pendingCount;
  } catch (error) {
    recordCurrentAccountError_(error);
    showErrorToUser_(error);
    throw error;
  }
}


function sendDailyAutoForwardSummary() {
  const lock = LockService.getUserLock();

  if (!lock.tryLock(30000)) {
    Logger.log(
      "Daily summary skipped because this account is still processing."
    );
    return;
  }

  try {
    const context = getAccountContext_();
    const properties = getUserProperties_();

    if (isRegistrationExpectedActive_(context.registration)) {
      ensureRequiredTriggers_();
    }

    const allProperties = properties.getProperties();
    const summaryRecords = [];

    for (const [key, value] of Object.entries(allProperties)) {
      if (!key.startsWith(CONFIG.SUMMARY_RECORD_PREFIX)) {
        continue;
      }

      try {
        summaryRecords.push({
          key,
          record: JSON.parse(value)
        });
      } catch (error) {
        Logger.log(
          `Invalid daily-summary record ${key}: ${error.message}`
        );
      }
    }

    summaryRecords.sort(
      (first, second) =>
        first.record.forwardedAt - second.record.forwardedAt
    );

    const generatedAt = formatDateTime_(new Date());
    const reportDate = Utilities.formatDate(
      new Date(),
      CONFIG.SUMMARY_TIME_ZONE,
      "yyyy-MM-dd"
    );
    const lines = [
      `AutoForward daily activity summary for ${context.account}`,
      `Generated: ${generatedAt} (${CONFIG.SUMMARY_TIME_ZONE})`,
      `Successfully forwarded: ${summaryRecords.length}`,
      ""
    ];

    if (summaryRecords.length === 0) {
      lines.push("No emails were forwarded since the last summary.");
    } else {
      summaryRecords.forEach((item, index) => {
        const record = item.record;
        const forwardedAt = formatDateTime_(
          new Date(record.forwardedAt)
        );

        lines.push(
          `${index + 1}. ${record.subject}`,
          `   Time: ${forwardedAt}`,
          `   From: ${record.sender}`,
          `   To: ${record.recipients.join(", ")}`,
          ""
        );
      });
    }

    GmailApp.sendEmail(
      getSummaryRecipient_(),
      `AutoForward Daily Summary - ${reportDate}`,
      lines.join("\n")
    );

    for (const item of summaryRecords) {
      properties.deleteProperty(item.key);
    }

    Logger.log(
      `Daily summary sent with ${summaryRecords.length} record(s).`
    );
  } catch (error) {
    recordCurrentAccountError_(error);
    throw error;
  } finally {
    lock.releaseLock();
  }
}


function cleanupProcessedRecords_() {
  const properties = getUserProperties_();
  const lastCleanup = Number(
    properties.getProperty(CONFIG.LAST_CLEANUP_KEY) || 0
  );
  const oneDay = 24 * 60 * 60 * 1000;

  if (Date.now() - lastCleanup < oneDay) {
    return;
  }

  const cutoff =
    Date.now() - CONFIG.RETAIN_PROCESSED_DAYS * oneDay;
  const records = properties.getProperties();
  const retryThreadsToSync = new Map();

  for (const [key, value] of Object.entries(records)) {
    if (
      key.startsWith(CONFIG.PROCESSED_PREFIX) &&
      Number(value) < cutoff
    ) {
      properties.deleteProperty(key);
    }

    // Clean up stale retry records
    if (key.startsWith(CONFIG.FAILED_RETRY_PREFIX)) {
      try {
        const data = JSON.parse(value);
        if (data.lastAttempt && Number(data.lastAttempt) < cutoff) {
          properties.deleteProperty(key);
          rememberRetryThreadForSync_(
            retryThreadsToSync,
            key.slice(CONFIG.FAILED_RETRY_PREFIX.length)
          );
        }
      } catch (parseError) {
        properties.deleteProperty(key);
        rememberRetryThreadForSync_(
          retryThreadsToSync,
          key.slice(CONFIG.FAILED_RETRY_PREFIX.length)
        );
      }
    }
  }

  if (retryThreadsToSync.size > 0) {
    const failedLabel = getOrCreateLabel_(CONFIG.FAILED_LABEL);
    const retryExhaustedLabel = getOrCreateLabel_(
      CONFIG.RETRY_EXHAUSTED_LABEL
    );

    for (const thread of retryThreadsToSync.values()) {
      try {
        syncThreadFailureLabels_(
          thread,
          properties,
          failedLabel,
          retryExhaustedLabel
        );
      } catch (labelError) {
        Logger.log(
          `Could not clean stale retry labels from thread ${thread.getId()}: ` +
          labelError.message
        );
      }
    }
  }

  properties.setProperty(
    CONFIG.LAST_CLEANUP_KEY,
    String(Date.now())
  );
}


function rememberRetryThreadForSync_(threadMap, messageId) {
  try {
    const message = GmailApp.getMessageById(messageId);

    if (message) {
      const thread = message.getThread();
      threadMap.set(thread.getId(), thread);
    }
  } catch (error) {
    Logger.log(
      `Could not locate message ${messageId} while cleaning retry state: ` +
      error.message
    );
  }
}


function getManagedTriggerHandlers_() {
  return [
    "monitorAndForwardSecEmails",
    "sendDailyAutoForwardSummary",
    "autoRepairAutoForwarding"
  ];
}


function createManagedTrigger_(handlerFunction) {
  if (handlerFunction === "monitorAndForwardSecEmails") {
    return ScriptApp.newTrigger(handlerFunction)
      .timeBased()
      .everyMinutes(CONFIG.CHECK_EVERY_MINUTES)
      .create();
  }

  if (handlerFunction === "sendDailyAutoForwardSummary") {
    return ScriptApp.newTrigger(handlerFunction)
      .timeBased()
      .atHour(CONFIG.SUMMARY_HOUR)
      .everyDays(1)
      .inTimezone(CONFIG.SUMMARY_TIME_ZONE)
      .create();
  }

  if (handlerFunction === "autoRepairAutoForwarding") {
    return ScriptApp.newTrigger(handlerFunction)
      .timeBased()
      .everyHours(CONFIG.WATCHDOG_EVERY_HOURS)
      .create();
  }

  throw new Error(`Unknown AutoForward trigger handler: ${handlerFunction}`);
}


/**
 * Creates missing managed triggers and removes accidental duplicates.
 * The monitor and watchdog both call this, allowing either one to restore the
 * other plus the daily summary trigger.
 */
function ensureRequiredTriggers_() {
  const handlers = getManagedTriggerHandlers_();
  const handlerSet = new Set(handlers);
  const existingByHandler = new Map();
  const created = [];
  const removedDuplicates = [];

  for (const trigger of ScriptApp.getProjectTriggers()) {
    const handler = trigger.getHandlerFunction();

    if (!handlerSet.has(handler)) {
      continue;
    }

    if (!existingByHandler.has(handler)) {
      existingByHandler.set(handler, trigger);
    } else {
      ScriptApp.deleteTrigger(trigger);
      removedDuplicates.push(handler);
    }
  }

  for (const handler of handlers) {
    if (!existingByHandler.has(handler)) {
      existingByHandler.set(handler, createManagedTrigger_(handler));
      created.push(handler);
    }
  }

  getUserProperties_().setProperty(
    CONFIG.LAST_TRIGGER_REPAIR_KEY,
    String(Date.now())
  );

  if (created.length || removedDuplicates.length) {
    Logger.log(JSON.stringify({
      triggerRepair: true,
      created,
      removedDuplicates
    }));
  }

  return {
    total: existingByHandler.size,
    created,
    removedDuplicates
  };
}


function maybeRepairRequiredTriggers_(properties) {
  const lastRepair = Number(
    properties.getProperty(CONFIG.LAST_TRIGGER_REPAIR_KEY) || 0
  );
  const oneHour = 60 * 60 * 1000;

  if (Date.now() - lastRepair >= oneHour) {
    return ensureRequiredTriggers_();
  }

  return null;
}


/**
 * Redundant watchdog. If the monitor trigger is deleted while this trigger
 * remains, it recreates it. If this watchdog is deleted, the monitor recreates
 * it during its hourly health check.
 */
function autoRepairAutoForwarding() {
  const lock = LockService.getUserLock();

  if (!lock.tryLock(30000)) {
    Logger.log("AutoForward self-repair skipped because another run is active.");
    return;
  }

  try {
    const context = getAccountContext_();
    const properties = getUserProperties_();
    const startedAt = Number(
      properties.getProperty(CONFIG.STARTED_AT_KEY) || 0
    );

    if (!startedAt || !isRegistrationExpectedActive_(context.registration)) {
      Logger.log("AutoForward self-repair skipped because automation is paused.");
      return;
    }

    const health = ensureRequiredTriggers_();

    if (health.created.length || health.removedDuplicates.length) {
      updateAccountRegistration_(context.account, {
        status: "Active — self-repaired",
        lastError: ""
      });
      updateRuleSheetStatus_(context.ruleSheet, {
        status: "Active — self-repaired"
      });
    }
  } catch (error) {
    recordCurrentAccountError_(error);
    throw error;
  } finally {
    lock.releaseLock();
  }
}


function isRegistrationExpectedActive_(registration) {
  const status = String(
    registration && registration.status || ""
  ).toLowerCase();

  return !(
    status.includes("paused") ||
    status.includes("reset") ||
    status.includes("not activated")
  );
}


function removeExistingTriggers_() {
  const handlerFunctions = new Set(getManagedTriggerHandlers_());

  for (const trigger of ScriptApp.getProjectTriggers()) {
    if (handlerFunctions.has(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  getUserProperties_().deleteProperty(CONFIG.LAST_TRIGGER_REPAIR_KEY);
}


function prepareUnactivatedAccount_() {
  removeExistingTriggers_();

  const properties = getUserProperties_();
  const records = properties.getProperties();

  properties.deleteProperty(CONFIG.STARTED_AT_KEY);
  properties.deleteProperty(CONFIG.SUMMARY_RECIPIENT_KEY);

  for (const key of Object.keys(records)) {
    if (key.startsWith(CONFIG.SUMMARY_RECORD_PREFIX)) {
      properties.deleteProperty(key);
    }
  }
}


function validateConfiguration_() {
  const allowedIntervals = [1, 5, 10, 15, 30];

  if (!allowedIntervals.includes(CONFIG.CHECK_EVERY_MINUTES)) {
    throw new Error(
      "Trigger interval must be 1, 5, 10, 15, or 30 minutes."
    );
  }

  if (
    !Number.isInteger(CONFIG.SUMMARY_HOUR) ||
    CONFIG.SUMMARY_HOUR < 0 ||
    CONFIG.SUMMARY_HOUR > 23
  ) {
    throw new Error(
      "Summary hour must be a whole number from 0 to 23."
    );
  }

  if (
    !Number.isInteger(CONFIG.BATCH_FORWARD_LIMIT) ||
    CONFIG.BATCH_FORWARD_LIMIT < 1
  ) {
    throw new Error("Batch forward limit must be a positive whole number.");
  }

  if (
    !Number.isInteger(CONFIG.MAX_RETRIES) ||
    CONFIG.MAX_RETRIES < 1
  ) {
    throw new Error("Maximum retries must be a positive whole number.");
  }

  if (
    !Number.isFinite(CONFIG.RETRY_DELAY_HOURS) ||
    CONFIG.RETRY_DELAY_HOURS < 0
  ) {
    throw new Error("Retry delay hours must be zero or greater.");
  }

  if (
    !Number.isInteger(CONFIG.WATCHDOG_EVERY_HOURS) ||
    CONFIG.WATCHDOG_EVERY_HOURS < 1
  ) {
    throw new Error(
      "Watchdog interval must be a positive whole number of hours."
    );
  }

  if (CONFIG.RETAIN_PROCESSED_DAYS < CONFIG.SEARCH_LOOKBACK_DAYS) {
    throw new Error(
      "Processed-record retention must be at least as long as the Gmail " +
      "search lookback to prevent old messages from becoming eligible again."
    );
  }

  getRules_();
}


function showMyAutoForwardStatus() {
  try {
    const context = getAccountContext_();
    const properties = getUserProperties_();
    const registration = getAccountRegistration_(context.account);
    const startedAt = Number(
      properties.getProperty(CONFIG.STARTED_AT_KEY) || 0
    );

    let repairNote = "";

    if (startedAt && isRegistrationExpectedActive_(registration)) {
      const health = ensureRequiredTriggers_();

      if (health.created.length || health.removedDuplicates.length) {
        repairNote =
          `🛠️ Self-repair: restored ${health.created.length} missing ` +
          `trigger(s) and removed ${health.removedDuplicates.length} ` +
          "duplicate(s)";
      }
    }

    const triggerNames = ScriptApp.getProjectTriggers()
      .map(trigger => trigger.getHandlerFunction())
      .filter(name => getManagedTriggerHandlers_().includes(name));
    const uniqueTriggerCount = new Set(triggerNames).size;
    const isRunning =
      startedAt &&
      isRegistrationExpectedActive_(registration) &&
      uniqueTriggerCount === getManagedTriggerHandlers_().length;
    const hasError = Boolean(registration.lastError);
    const automationIcon = isRunning
      ? hasError ? "🟡" : "🟢"
      : startedAt ? "🟡" : "⏸️";
    const automationStatus = isRunning
      ? hasError ? "ACTIVE — NEEDS ATTENTION" : "ACTIVE"
      : registration.status || "INACTIVE";
    const statusLines = [
      `${automationIcon} AUTOMATION: ${automationStatus}`,
      "",
      `📧 Gmail account: ${context.account}`,
      `📄 Rule tab: ${context.ruleSheet.getName()}`,
      `⚙️ Triggers installed: ${uniqueTriggerCount} of 3`,
      `▶️ First started: ${startedAt ? formatDateTime_(new Date(startedAt)) : "No"}`,
      `🕒 Last monitor run: ${registration.lastRunAt ? formatDateTime_(new Date(registration.lastRunAt)) : "Never"}`,
      `🚨 Last error: ${registration.lastError || "None"}`
    ];

    if (repairNote) {
      statusLines.push("", repairNote);
    }

    SpreadsheetApp.getUi().alert(
      `${automationIcon} AutoForward status`,
      statusLines.join("\n"),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return statusLines;
  } catch (error) {
    showErrorToUser_(error);
    throw error;
  }
}


function stopSecEmailForwarding() {
  try {
    const context = getAccountContext_();
    removeExistingTriggers_();
    updateAccountRegistration_(context.account, {
      status: "Paused",
      lastError: ""
    });
    updateRuleSheetStatus_(context.ruleSheet, {
      status: "Paused"
    });
    context.spreadsheet.toast(
      `Automation is paused for ${context.account}.`,
      "AutoForward",
      8
    );
    Logger.log(`AutoForward paused for ${context.account}.`);
  } catch (error) {
    showErrorToUser_(error);
    throw error;
  }
}


function confirmAndResetMyAutoForwarding() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Reset AutoForward history?",
    "This pauses your automation and clears your processed-message " +
    "history. Recently received emails could be forwarded again after " +
    "you reactivate it.",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return false;
  }

  resetSecEmailForwarding();
  return true;
}


function resetSecEmailForwarding() {
  try {
    const context = getAccountContext_();
    removeExistingTriggers_();

    const properties = getUserProperties_();
    const records = properties.getProperties();

    for (const key of Object.keys(records)) {
      if (
        key.startsWith(CONFIG.PROCESSED_PREFIX) ||
        key.startsWith(CONFIG.FAILED_RETRY_PREFIX) ||
        key.startsWith(CONFIG.SUMMARY_RECORD_PREFIX) ||
        key === CONFIG.STARTED_AT_KEY ||
        key === CONFIG.LAST_CLEANUP_KEY
      ) {
        properties.deleteProperty(key);
      }
    }

    clearRetryStateLabels_();

    updateAccountRegistration_(context.account, {
      status: "Reset — not activated",
      activatedAt: 0,
      lastRunAt: 0,
      lastError: ""
    });
    updateRuleSheetStatus_(context.ruleSheet, {
      status: "Reset — not activated",
      lastRun: "Never"
    });
    context.spreadsheet.toast(
      `Processed history was reset for ${context.account}.`,
      "AutoForward",
      8
    );
    Logger.log(`AutoForward reset for ${context.account}.`);
  } catch (error) {
    showErrorToUser_(error);
    throw error;
  }
}


function clearRetryStateLabels_() {
  for (const labelName of [
    CONFIG.FAILED_LABEL,
    CONFIG.RETRY_EXHAUSTED_LABEL
  ]) {
    try {
      const label = GmailApp.getUserLabelByName(labelName);

      if (label) {
        label.deleteLabel();
      }
    } catch (error) {
      Logger.log(
        `Could not clear retry-state label ${labelName}: ${error.message}`
      );
    }
  }
}


function recordCurrentAccountError_(error, existingContext) {
  try {
    const context = existingContext || getAccountContext_();
    const message = sanitizeStatusText_(error.message || String(error));

    updateAccountRegistration_(context.account, {
      lastError: message
    });
    updateRuleSheetStatus_(context.ruleSheet, {
      status: "Needs attention"
    });
  } catch (recordError) {
    Logger.log(
      "Could not save AutoForward error status: " + recordError.message
    );
  }
}


function sanitizeStatusText_(value) {
  return String(value || "")
    .replace(/[\r\n]+/g, " ")
    .slice(0, 500);
}


function formatDateTime_(date) {
  return Utilities.formatDate(
    date,
    CONFIG.SUMMARY_TIME_ZONE,
    "yyyy-MM-dd HH:mm:ss"
  );
}


function showErrorToUser_(error) {
  Logger.log(error && error.stack ? error.stack : String(error));

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      sanitizeStatusText_(error.message || error),
      "AutoForward error",
      10
    );
  } catch (uiError) {
    // Time triggers have no active spreadsheet UI.
  }
}
