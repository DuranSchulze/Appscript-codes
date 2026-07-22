const CONFIG = {
  CHECK_EVERY_MINUTES: 5,
  SEARCH_LOOKBACK_DAYS: 30,
  MAX_THREADS_PER_RUN: 500,
  RETAIN_PROCESSED_DAYS: 60,
  INCLUDE_SPAM: true,

  SUMMARY_HOUR: 23,
  SUMMARY_TIME_ZONE: "Asia/Manila",

  RULE_SHEET_PREFIX: "Rules - ",
  RULE_HEADER_ROW: 6,
  RULE_DATA_ROWS: 500,
  RULE_SCHEMA_VERSION: 2,

  // Optional: addresses that may edit every protected account rule tab.
  ADMIN_EMAILS: [],

  ACCOUNT_REGISTRATION_PREFIX: "AF_ACCOUNT_REGISTRATION_",
  PROCESSED_PREFIX: "AF_FORWARDED_",
  SUMMARY_RECORD_PREFIX: "AF_DAILY_FORWARD_",
  STARTED_AT_KEY: "AF_AUTOMATION_STARTED_AT",
  LAST_CLEANUP_KEY: "AF_LAST_CLEANUP",
  SUMMARY_RECIPIENT_KEY: "AF_SUMMARY_RECIPIENT",

  ROOT_LABEL: "AutoForward",
  DETECTED_LABEL: "AutoForward/Detected",
  FORWARDED_LABEL: "AutoForward/Forwarded",
  FAILED_LABEL: "AutoForward/Failed"
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
  SpreadsheetApp.getUi()
    .createMenu("AutoForward")
    .addItem(
      "Create or open my rule tab",
      "createOrOpenMyRuleTab"
    )
    .addItem(
      "Adopt my active legacy rule tab",
      "adoptActiveLegacyRuleTab"
    )
    .addSeparator()
    .addItem("Validate my rules", "validateMyRules")
    .addItem(
      "Preview matching emails",
      "previewMatchingSecEmails"
    )
    .addItem(
      "Preview pending emails",
      "previewPendingEmails"
    )
    .addSeparator()
    .addItem(
      "Activate or repair my automation",
      "activateMyAutoForwarding"
    )
    .addItem("Show my status", "showMyAutoForwardStatus")
    .addItem("Pause my automation", "stopSecEmailForwarding")
    .addSeparator()
    .addItem(
      "Reset my processed history",
      "confirmAndResetMyAutoForwarding"
    )
    .addToUi();
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
    let existing = getAccountRegistration_(account);

    if (existing) {
      const existingSheet = findSheetById_(
        spreadsheet,
        existing.ruleSheetId
      );

      if (!existingSheet) {
        if (existing.spreadsheetId !== spreadsheet.getId()) {
          throw new Error(
            "Your registered rule tab is in another spreadsheet and is " +
            "no longer available. Ask an administrator to repair the " +
            "registration."
          );
        }

        PropertiesService.getScriptProperties().deleteProperty(
          registrationKey_(account)
        );
        existing = null;
      }

      if (existing && existing.spreadsheetId !== spreadsheet.getId()) {
        throw new Error(
          "Your account is registered to a different AutoForward " +
          "spreadsheet. Open that spreadsheet or ask an administrator " +
          "to move the registration."
        );
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
    validateConfiguration_();

    const context = getAccountContext_();
    const properties = getUserProperties_();

    getOrCreateLabel_(CONFIG.ROOT_LABEL);
    getOrCreateLabel_(CONFIG.DETECTED_LABEL);
    getOrCreateLabel_(CONFIG.FORWARDED_LABEL);
    getOrCreateLabel_(CONFIG.FAILED_LABEL);

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

    ScriptApp.newTrigger("monitorAndForwardSecEmails")
      .timeBased()
      .everyMinutes(CONFIG.CHECK_EVERY_MINUTES)
      .create();

    ScriptApp.newTrigger("sendDailyAutoForwardSummary")
      .timeBased()
      .atHour(CONFIG.SUMMARY_HOUR)
      .everyDays(1)
      .inTimezone(CONFIG.SUMMARY_TIME_ZONE)
      .create();

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
      `Automation is active for ${context.account}.`,
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

  try {
    rulesCache_ = null;
    context = getAccountContext_();
    cleanupProcessedRecords_();

    const properties = getUserProperties_();
    const processedRecords = properties.getProperties();
    const startedAt = Number(
      properties.getProperty(CONFIG.STARTED_AT_KEY) || 0
    );

    if (!startedAt) {
      throw new Error(
        "Automation is not initialized for this Gmail account. " +
        "Use AutoForward > Activate or repair my automation."
      );
    }

    const detectedLabel = getOrCreateLabel_(CONFIG.DETECTED_LABEL);
    const forwardedLabel = getOrCreateLabel_(CONFIG.FORWARDED_LABEL);
    const failedLabel = getOrCreateLabel_(CONFIG.FAILED_LABEL);
    const messages = getCandidateMessages_();

    for (const message of messages) {
      const messageId = message.getId();
      const processedKey = CONFIG.PROCESSED_PREFIX + messageId;

      if (processedRecords[processedKey]) {
        continue;
      }

      if (message.getDate().getTime() < startedAt) {
        continue;
      }

      const rule = findMatchingRule_(message);

      if (!rule) {
        continue;
      }

      const thread = message.getThread();
      thread.addLabel(detectedLabel);

      try {
        const matchedKeywords = getMatchedKeywords_(message, rule);

        Logger.log(
          `Forwarding "${message.getSubject()}" from ${rule.sender}. ` +
          `Matched: ${matchedKeywords.join(", ") || "all messages"}`
        );

        message.forward(rule.recipients.join(","));

        const processedAt = String(Date.now());
        properties.setProperty(processedKey, processedAt);
        processedRecords[processedKey] = processedAt;

        try {
          thread.addLabel(forwardedLabel);
          thread.removeLabel(failedLabel);
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

        try {
          thread.addLabel(failedLabel);
        } catch (labelError) {
          Logger.log(
            `Could not add the Failed label to ${messageId}: ` +
            labelError.message
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
  } catch (error) {
    recordCurrentAccountError_(error, context);
    throw error;
  } finally {
    lock.releaseLock();
  }
}


/**
 * Searches Gmail for messages from every enabled sender in this account's tab.
 */
function getCandidateMessages_() {
  const senderSearch = Array.from(
    new Set(
      getRules_().map(rule => normalizeEmail_(rule.sender))
    )
  )
    .map(sender => `from:${sender}`)
    .join(" ");

  const monitoredLocations = CONFIG.INCLUDE_SPAM
    ? `{in:inbox in:spam label:"${CONFIG.FAILED_LABEL}"}`
    : `{in:inbox label:"${CONFIG.FAILED_LABEL}"}`;

  const query =
    `newer_than:${CONFIG.SEARCH_LOOKBACK_DAYS}d ` +
    `${monitoredLocations} {${senderSearch}}`;

  const messageMap = new Map();
  let start = 0;
  const batchSize = 50;

  while (start < CONFIG.MAX_THREADS_PER_RUN) {
    const amount = Math.min(
      batchSize,
      CONFIG.MAX_THREADS_PER_RUN - start
    );
    const threads = GmailApp.search(query, start, amount);

    if (threads.length === 0) {
      break;
    }

    for (const thread of threads) {
      const isFailedThread = threadHasLabel_(
        thread,
        CONFIG.FAILED_LABEL
      );
      const isMonitoredSpam =
        CONFIG.INCLUDE_SPAM && thread.isInSpam();

      for (const message of thread.getMessages()) {
        if (
          (
            message.isInInbox() ||
            isMonitoredSpam ||
            isFailedThread
          ) &&
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

  const messages = Array.from(messageMap.values());
  messages.sort(
    (first, second) =>
      first.getDate().getTime() - second.getDate().getTime()
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
      "Ask an administrator to repair the registration."
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
  const existingProtection = sheet
    .getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .find(item => item.getDescription() === description);
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

    if (rule.recipients.length > 100) {
      throw new Error(
        `Too many recipients on ${rowLabel}. Use 100 or fewer.`
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

      if (
        records[processedKey] ||
        message.getDate().getTime() < startedAt
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

  for (const [key, value] of Object.entries(records)) {
    if (
      key.startsWith(CONFIG.PROCESSED_PREFIX) &&
      Number(value) < cutoff
    ) {
      properties.deleteProperty(key);
    }
  }

  properties.setProperty(
    CONFIG.LAST_CLEANUP_KEY,
    String(Date.now())
  );
}


function removeExistingTriggers_() {
  const handlerFunctions = new Set([
    "monitorAndForwardSecEmails",
    "sendDailyAutoForwardSummary"
  ]);

  for (const trigger of ScriptApp.getProjectTriggers()) {
    if (handlerFunctions.has(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
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

  getRules_();
}


function showMyAutoForwardStatus() {
  try {
    const context = getAccountContext_();
    const properties = getUserProperties_();
    const startedAt = Number(
      properties.getProperty(CONFIG.STARTED_AT_KEY) || 0
    );
    const triggerNames = ScriptApp.getProjectTriggers()
      .map(trigger => trigger.getHandlerFunction())
      .filter(name => [
        "monitorAndForwardSecEmails",
        "sendDailyAutoForwardSummary"
      ].includes(name));
    const registration = getAccountRegistration_(context.account);
    const statusLines = [
      `Gmail account: ${context.account}`,
      `Rule tab: ${context.ruleSheet.getName()}`,
      `Registration: ${registration.status || "Unknown"}`,
      `Activated: ${startedAt ? formatDateTime_(new Date(startedAt)) : "No"}`,
      `Triggers installed: ${triggerNames.length} of 2`,
      `Last run: ${registration.lastRunAt ? formatDateTime_(new Date(registration.lastRunAt)) : "Never"}`,
      `Last error: ${registration.lastError || "None"}`
    ];

    SpreadsheetApp.getUi().alert(
      "AutoForward status",
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
        key.startsWith(CONFIG.SUMMARY_RECORD_PREFIX) ||
        key === CONFIG.STARTED_AT_KEY ||
        key === CONFIG.LAST_CLEANUP_KEY
      ) {
        properties.deleteProperty(key);
      }
    }

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
