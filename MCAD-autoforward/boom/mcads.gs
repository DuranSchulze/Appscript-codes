const CONFIG = {
  CHECK_EVERY_MINUTES: 5,
  SEARCH_LOOKBACK_DAYS: 30,
  MAX_THREADS_PER_RUN: 500,
  RETAIN_PROCESSED_DAYS: 60,
  INCLUDE_SPAM: true,

  PROCESSED_PREFIX: "SEC_FORWARDED_",
  SUMMARY_RECORD_PREFIX: "SEC_DAILY_FORWARD_",
  STARTED_AT_KEY: "SEC_AUTOMATION_STARTED_AT",
  LAST_CLEANUP_KEY: "SEC_LAST_CLEANUP",
  SUMMARY_RECIPIENT_KEY: "SEC_SUMMARY_RECIPIENT",

  SUMMARY_HOUR: 23,
  SUMMARY_TIME_ZONE: "Asia/Manila",

  // Required: paste the ID from the shared Google Sheet URL.
  RULES_SPREADSHEET_ID: "PASTE_SPREADSHEET_ID_HERE",

  // Set this to the rules tab assigned to this Gmail account.
  RULES_SHEET_NAME: "Rules - Code.gs",

  ROOT_LABEL: "AutoForward",
  DETECTED_LABEL: "AutoForward/Detected",
  FORWARDED_LABEL: "AutoForward/Forwarded",
  FAILED_LABEL: "AutoForward/Failed"
};


/**
 * Run this function once when you are ready.
 *
 * It creates the automatic trigger and prevents older emails
 * from being forwarded when the automation starts.
 */
function setupSecEmailForwarding() {
  validateConfiguration_();

  getOrCreateLabel_(CONFIG.ROOT_LABEL);
  getOrCreateLabel_(CONFIG.DETECTED_LABEL);
  getOrCreateLabel_(CONFIG.FORWARDED_LABEL);
  getOrCreateLabel_(CONFIG.FAILED_LABEL);

  const properties = PropertiesService.getScriptProperties();

  getSummaryRecipient_();

  if (!properties.getProperty(CONFIG.STARTED_AT_KEY)) {
    properties.setProperty(
      CONFIG.STARTED_AT_KEY,
      String(Date.now())
    );
  }

  // Replace triggers only after setup checks succeed.
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

  Logger.log("SEC email-forwarding automation activated.");
}


/**
 * Main function called automatically by the trigger.
 */
function monitorAndForwardSecEmails() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    Logger.log("Another execution is already running.");
    return;
  }

  try {
    cleanupProcessedRecords_();

    const properties = PropertiesService.getScriptProperties();
    const processedRecords = properties.getProperties();

    const startedAt = Number(
      properties.getProperty(CONFIG.STARTED_AT_KEY) || 0
    );

    if (!startedAt) {
      Logger.log(
        "Automation is not initialized. Run " +
        "setupSecEmailForwarding first."
      );
      return;
    }

    const detectedLabel = getOrCreateLabel_(
      CONFIG.DETECTED_LABEL
    );
    const forwardedLabel = getOrCreateLabel_(
      CONFIG.FORWARDED_LABEL
    );
    const failedLabel = getOrCreateLabel_(
      CONFIG.FAILED_LABEL
    );

    const messages = getCandidateMessages_();

    if (messages.length === 0) {
      Logger.log("No candidate SEC emails found.");
      return;
    }

    for (const message of messages) {
      const messageId = message.getId();
      const processedKey =
        CONFIG.PROCESSED_PREFIX + messageId;

      if (processedRecords[processedKey]) {
        continue;
      }

      // Ignore messages received before the automation started.
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
        const matchedKeywords = getMatchedKeywords_(
          message,
          rule
        );

        Logger.log(
          `Forwarding "${message.getSubject()}" ` +
          `from ${rule.sender}. ` +
          `Matched: ${matchedKeywords.join(", ")}`
        );

        /*
         * This forwards the original Gmail message,
         * including its normal forwarded content and attachments.
         */
        message.forward(rule.recipients.join(","));

        const processedAt = String(Date.now());

        // Save only after forwarding succeeds.
        properties.setProperty(
          processedKey,
          processedAt
        );

        processedRecords[processedKey] = processedAt;

        try {
          thread.addLabel(forwardedLabel);
          thread.removeLabel(failedLabel);
        } catch (labelError) {
          Logger.log(
            `Message ${messageId} was forwarded, but the ` +
            `status labels could not be updated: ` +
            labelError.message
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
            `Message ${messageId} was forwarded, but it ` +
            `could not be added to the daily summary: ` +
            summaryError.message
          );
        }

        Logger.log(
          `Forwarded message ${messageId} to: ` +
          rule.recipients.join(", ")
        );
      } catch (error) {
        Logger.log(
          `Failed forwarding ${messageId}: ${error.message}`
        );

        try {
          thread.addLabel(failedLabel);
        } catch (labelError) {
          Logger.log(
            `Failed label could not be added to message ` +
            `${messageId}: ${labelError.message}`
          );
        }
      }
    }
  } finally {
    lock.releaseLock();
  }
}


/**
 * Searches for messages from every configured sender.
 */
function getCandidateMessages_() {
  const senderSearch = getRules_()
    .map(rule => `from:${normalizeEmail_(rule.sender)}`)
    .join(" ");

  // Curly braces mean OR in Gmail search.
  const monitoredLocations = CONFIG.INCLUDE_SPAM
    ? `{in:inbox in:spam label:"${CONFIG.FAILED_LABEL}"}`
    : `{in:inbox label:"${CONFIG.FAILED_LABEL}"}`;

  const query =
    `newer_than:${CONFIG.SEARCH_LOOKBACK_DAYS}d ` +
    `${monitoredLocations} ` +
    `{${senderSearch}}`;

  const messageMap = new Map();

  let start = 0;
  const batchSize = 50;

  while (start < CONFIG.MAX_THREADS_PER_RUN) {
    const amount = Math.min(
      batchSize,
      CONFIG.MAX_THREADS_PER_RUN - start
    );

    const threads = GmailApp.search(
      query,
      start,
      amount
    );

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

  // Process the oldest matching message first.
  messages.sort(
    (first, second) =>
      first.getDate().getTime() -
      second.getDate().getTime()
  );

  return messages;
}


/**
 * Finds the sender rule and confirms a keyword match.
 */
function findMatchingRule_(message) {
  const sender = extractEmailAddress_(
    message.getFrom()
  );

  for (const rule of getRules_()) {
    if (
      sender !== normalizeEmail_(rule.sender)
    ) {
      continue;
    }

    // matchAll forwards every email from this sender.
    if (rule.matchAll) {
      return rule;
    }

    const matchedKeywords = getMatchedKeywords_(
      message,
      rule
    );

    if (matchedKeywords.length > 0) {
      return rule;
    }
  }

  return null;
}


/**
 * Checks both subject and plain-text email content.
 */
function getMatchedKeywords_(message, rule) {
  const subject = message.getSubject() || "";
  const body = message.getPlainBody() || "";

  const searchableText = `${subject}\n${body}`;

  return rule.keywords.filter(keyword =>
    matchesKeyword_(searchableText, keyword)
  );
}


/**
 * Short uppercase keywords such as AFS, OTP and MC28
 * are matched as whole words to avoid accidental matches.
 */
function matchesKeyword_(text, keyword) {
  const cleanKeyword = String(keyword).trim();

  if (!cleanKeyword) {
    return false;
  }

  const isCode =
    /^[A-Z0-9]+$/.test(cleanKeyword) &&
    cleanKeyword.length <= 10;

  if (isCode) {
    const escapedKeyword = cleanKeyword.replace(
      /[.*+?^${}()|[\]\\]/g,
      "\\$&"
    );

    const pattern = new RegExp(
      `\\b${escapedKeyword}\\b`,
      "i"
    );

    return pattern.test(text);
  }

  return text
    .toLowerCase()
    .includes(cleanKeyword.toLowerCase());
}


/**
 * Extracts the address from:
 * "SEC Notification <no-reply@sec.gov.ph>"
 */
function extractEmailAddress_(fromValue) {
  const match = String(fromValue).match(
    /<([^>]+)>/
  );

  return normalizeEmail_(
    match ? match[1] : fromValue
  );
}


function normalizeEmail_(email) {
  return String(email || "")
    .trim()
    .toLowerCase();
}


let rulesCache_ = null;


/**
 * Tests the configured Google Sheet without forwarding or changing email.
 * Run this manually after setting RULES_SPREADSHEET_ID and RULES_SHEET_NAME.
 */
function testGoogleSheetConnection() {
  try {
    // Force a fresh read so this test always checks the current Sheet data.
    rulesCache_ = null;
    const rules = loadRulesFromSheet_();
    const spreadsheet = SpreadsheetApp.openById(
      String(CONFIG.RULES_SPREADSHEET_ID).trim()
    );
    const result = {
      connected: true,
      spreadsheet: spreadsheet.getName(),
      sheet: CONFIG.RULES_SHEET_NAME,
      enabledRules: rules.length,
      senders: rules.map(rule => rule.sender)
    };

    Logger.log(
      "Google Sheet connection successful:\n" +
      JSON.stringify(result, null, 2)
    );

    return result;
  } catch (error) {
    Logger.log(
      "Google Sheet connection failed: " + error.message
    );
    throw error;
  }
}


/**
 * Loads the active forwarding rules from this script's assigned tab.
 * The cache lasts only for the current Apps Script execution.
 */
function getRules_() {
  if (!rulesCache_) {
    rulesCache_ = loadRulesFromSheet_();
  }

  return rulesCache_;
}


function loadRulesFromSheet_() {
  const spreadsheetId = String(
    CONFIG.RULES_SPREADSHEET_ID || ""
  ).trim();

  if (
    !spreadsheetId ||
    spreadsheetId === "PASTE_SPREADSHEET_ID_HERE"
  ) {
    throw new Error(
      "Set CONFIG.RULES_SPREADSHEET_ID to the shared " +
      "Google Spreadsheet ID before running setup."
    );
  }

  const spreadsheet =
    SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName(
    CONFIG.RULES_SHEET_NAME
  );

  if (!sheet) {
    throw new Error(
      `Rules tab not found: ${CONFIG.RULES_SHEET_NAME}`
    );
  }

  const values = sheet.getDataRange().getDisplayValues();

  if (values.length < 2) {
    throw new Error(
      `No rule rows found in ${CONFIG.RULES_SHEET_NAME}`
    );
  }

  const headers = values[0].map(
    normalizeRuleHeader_
  );
  const requiredHeaders = [
    "enabled",
    "sender",
    "match all",
    "keywords",
    "recipients"
  ];
  const columns = {};

  for (const header of requiredHeaders) {
    const index = headers.indexOf(header);

    if (index < 0) {
      throw new Error(
        `Missing required column "${header}" in ` +
        CONFIG.RULES_SHEET_NAME
      );
    }

    columns[header] = index;
  }

  const rules = [];

  for (let index = 1; index < values.length; index++) {
    const row = values[index];
    const rowNumber = index + 1;

    if (row.every(value => !String(value).trim())) {
      continue;
    }

    const enabled = parseRuleBoolean_(
      row[columns["enabled"]],
      false,
      "Enabled",
      rowNumber
    );

    if (!enabled) {
      continue;
    }

    rules.push({
      sender: String(
        row[columns["sender"]] || ""
      ).trim(),
      matchAll: parseRuleBoolean_(
        row[columns["match all"]],
        false,
        "Match All",
        rowNumber
      ),
      keywords: splitRuleList_(
        row[columns["keywords"]],
        false
      ),
      recipients: splitRuleList_(
        row[columns["recipients"]],
        true
      ),
      sourceRow: rowNumber
    });
  }

  if (rules.length === 0) {
    throw new Error(
      `No enabled rules found in ${CONFIG.RULES_SHEET_NAME}`
    );
  }

  validateRules_(rules);
  return rules;
}


function validateRules_(rules) {
  for (const rule of rules) {
    const rowLabel = `row ${rule.sourceRow}`;

    if (!normalizeEmail_(rule.sender).includes("@")) {
      throw new Error(
        `Invalid sender on ${rowLabel}: ${rule.sender}`
      );
    }

    if (!rule.matchAll && !rule.keywords.length) {
      throw new Error(
        `No keywords configured on ${rowLabel} for ` +
        `${rule.sender}. Set Match All to TRUE to ` +
        "forward everything from this sender."
      );
    }

    if (!rule.recipients.length) {
      throw new Error(
        `No recipients configured on ${rowLabel} for ` +
        rule.sender
      );
    }

    const invalidRecipient = rule.recipients.find(
      recipient =>
        !normalizeEmail_(recipient).includes("@")
    );

    if (invalidRecipient) {
      throw new Error(
        `Invalid recipient on ${rowLabel}: ` +
        invalidRecipient
      );
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


function parseRuleBoolean_(
  value,
  defaultValue,
  columnName,
  rowNumber
) {
  const cleanValue = String(value || "")
    .trim()
    .toUpperCase();

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
    `Invalid ${columnName} value on row ${rowNumber}: ` +
    value
  );
}


function splitRuleList_(value, allowCommas) {
  const separator = allowCommas
    ? /[\n;,]+/
    : /[\n;]+/;

  return String(value || "")
    .split(separator)
    .map(item => item.trim())
    .filter(Boolean);
}



/**
 * Returns an existing Gmail label or creates it when needed.
 * A slash creates the nested label structure in Gmail.
 */
function getOrCreateLabel_(labelName) {
  return GmailApp.getUserLabelByName(labelName) ||
    GmailApp.createLabel(labelName);
}


function threadHasLabel_(thread, labelName) {
  return thread.getLabels().some(
    label => label.getName() === labelName
  );
}


/**
 * Stores one compact record until the next daily summary is sent.
 */
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


/**
 * Uses the account that ran setup as the summary recipient.
 */
function getSummaryRecipient_() {
  const properties =
    PropertiesService.getScriptProperties();

  const savedRecipient = normalizeEmail_(
    properties.getProperty(
      CONFIG.SUMMARY_RECIPIENT_KEY
    )
  );

  if (savedRecipient) {
    return savedRecipient;
  }

  const effectiveUser = normalizeEmail_(
    Session.getEffectiveUser().getEmail()
  );

  if (!effectiveUser) {
    throw new Error(
      "Could not determine the daily-summary recipient. " +
      "Run setupSecEmailForwarding manually from the owner account."
    );
  }

  properties.setProperty(
    CONFIG.SUMMARY_RECIPIENT_KEY,
    effectiveUser
  );

  return effectiveUser;
}


/**
 * Run this before setup to preview matching emails.
 * It does not forward anything.
 */
function previewMatchingSecEmails() {
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
      sender: extractEmailAddress_(
        message.getFrom()
      ),
      subject: message.getSubject(),
      matchedKeywords: getMatchedKeywords_(
        message,
        rule
      ),
      recipients: rule.recipients
    }));
  }

  Logger.log(`${matchCount} matching email(s) found.`);
}


/**
 * Shows only messages that the next live run would attempt.
 * It does not forward, label, or otherwise change any email.
 */
function previewPendingEmails() {
  validateConfiguration_();

  const properties =
    PropertiesService.getScriptProperties();
  const records = properties.getProperties();
  const startedAt = Number(
    records[CONFIG.STARTED_AT_KEY] || 0
  );

  if (!startedAt) {
    Logger.log(
      "Automation has not been set up yet. Run " +
      "setupSecEmailForwarding first."
    );
    return;
  }

  const messages = getCandidateMessages_();
  let pendingCount = 0;

  for (const message of messages) {
    const messageId = message.getId();
    const processedKey =
      CONFIG.PROCESSED_PREFIX + messageId;

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

  Logger.log(
    `${pendingCount} pending email(s) would be attempted.`
  );
}


/**
 * Sends all successful forwards recorded since the prior summary.
 * The daily trigger runs during the 11 PM hour in Manila.
 */
function sendDailyAutoForwardSummary() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    Logger.log(
      "Daily summary skipped because forwarding is still running. " +
      "Unreported records will remain for the next summary."
    );
    return;
  }

  try {
    const properties =
      PropertiesService.getScriptProperties();
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

    const generatedAt = Utilities.formatDate(
      new Date(),
      CONFIG.SUMMARY_TIME_ZONE,
      "yyyy-MM-dd HH:mm:ss"
    );
    const reportDate = Utilities.formatDate(
      new Date(),
      CONFIG.SUMMARY_TIME_ZONE,
      "yyyy-MM-dd"
    );

    const lines = [
      "AutoForward daily activity summary",
      `Generated: ${generatedAt} (${CONFIG.SUMMARY_TIME_ZONE})`,
      `Successfully forwarded: ${summaryRecords.length}`,
      ""
    ];

    if (summaryRecords.length === 0) {
      lines.push("No emails were forwarded since the last summary.");
    } else {
      summaryRecords.forEach((item, index) => {
        const record = item.record;
        const forwardedAt = Utilities.formatDate(
          new Date(record.forwardedAt),
          CONFIG.SUMMARY_TIME_ZONE,
          "yyyy-MM-dd HH:mm:ss"
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

    // Delete records only after the summary email succeeds.
    for (const item of summaryRecords) {
      properties.deleteProperty(item.key);
    }

    Logger.log(
      `Daily summary sent with ${summaryRecords.length} record(s).`
    );
  } finally {
    lock.releaseLock();
  }
}


/**
 * Removes processed records that are no longer needed.
 */
function cleanupProcessedRecords_() {
  const properties =
    PropertiesService.getScriptProperties();

  const lastCleanup = Number(
    properties.getProperty(
      CONFIG.LAST_CLEANUP_KEY
    ) || 0
  );

  const oneDay = 24 * 60 * 60 * 1000;

  if (Date.now() - lastCleanup < oneDay) {
    return;
  }

  const cutoff =
    Date.now() -
    CONFIG.RETAIN_PROCESSED_DAYS * oneDay;

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


/**
 * Prevents duplicate automatic triggers.
 */
function removeExistingTriggers_() {
  const handlerFunctions = new Set([
    "monitorAndForwardSecEmails",
    "sendDailyAutoForwardSummary"
  ]);

  for (
    const trigger of ScriptApp.getProjectTriggers()
  ) {
    if (
      handlerFunctions.has(trigger.getHandlerFunction())
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}


function validateConfiguration_() {
  const allowedIntervals = [1, 5, 10, 15, 30];

  if (
    !allowedIntervals.includes(
      CONFIG.CHECK_EVERY_MINUTES
    )
  ) {
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


/**
 * Stops the automation.
 */
function stopSecEmailForwarding() {
  removeExistingTriggers_();
  Logger.log("SEC forwarding automation stopped.");
}


/**
 * Resets the automation start date and processed history.
 * Use carefully because recent messages could be forwarded again.
 */
function resetSecEmailForwarding() {
  removeExistingTriggers_();

  const properties =
    PropertiesService.getScriptProperties();

  const records = properties.getProperties();

  for (const key of Object.keys(records)) {
    if (
      key.startsWith(CONFIG.PROCESSED_PREFIX) ||
      key.startsWith(CONFIG.SUMMARY_RECORD_PREFIX) ||
      key === CONFIG.STARTED_AT_KEY
    ) {
      properties.deleteProperty(key);
    }
  }

  Logger.log("SEC forwarding automation reset.");
}
