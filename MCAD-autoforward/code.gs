const CONFIG = {
  CHECK_EVERY_MINUTES: 5,
  SEARCH_LOOKBACK_DAYS: 30,
  MAX_THREADS_PER_RUN: 500,
  RETAIN_PROCESSED_DAYS: 60,

  PROCESSED_PREFIX: "SEC_FORWARDED_",
  SUMMARY_RECORD_PREFIX: "SEC_DAILY_FORWARD_",
  STARTED_AT_KEY: "SEC_AUTOMATION_STARTED_AT",
  LAST_CLEANUP_KEY: "SEC_LAST_CLEANUP",
  SUMMARY_RECIPIENT_KEY: "SEC_SUMMARY_RECIPIENT",

  SUMMARY_HOUR: 23,
  SUMMARY_TIME_ZONE: "Asia/Manila",

  ROOT_LABEL: "AutoForward",
  DETECTED_LABEL: "AutoForward/Detected",
  FORWARDED_LABEL: "AutoForward/Forwarded",
  FAILED_LABEL: "AutoForward/Failed",

RULES: [
  {
    sender: "noreply-cifssost@sec.gov.ph",
    keywords: [
      "GFFS",
      "email validation",
      "AFS"
    ],
    recipients: [
      "felise@duranschulze.com",
      "stephanie@duranschulze.com",
      "carlnathaniel@duranschulze.com",
      "projects@filepino.com",
      "alhyn@filepino.com",
      "fatima@filepino.com",
      "accounts@filepino.com",
      "reception@filepino.com"
    ]
  },

  {
    sender: "no-reply@sec.gov.ph",
    keywords: [
      "eAmend",
      "MC28",
      "SEC general notice",
      "SEC general notices",
      "OTP"
    ],
    recipients: [
      "felise@duranschulze.com",
      "stephanie@duranschulze.com",
      "carlnathaniel@duranschulze.com",
      "projects@filepino.com",
      "alhyn@filepino.com",
      "fatima@filepino.com"
    ]
  },

  {
    sender: "service@intl.paypal.com",
    keywords: [
      "DDS",
      "PayPal payment",
      "payment received"
    ],
    recipients: [
      "marywendy@duranschulze.com",
      "billing@duranschulze.com"
    ]
  },

  {
    sender: "bpi_cards_estatement@bpi.com.ph",
    keywords: [
      "DDS CC",
      "credit card",
      "electronic statement",
      "e-statement"
    ],
    recipients: [
      "accounts.payable1@filepino.com"
    ]
  },

  {
    sender: "e-corr@ipophl.gov.ph",

    // Forward all emails from this sender.
    matchAll: true,
    keywords: [],

    recipients: [
      "paula@duranschulze.com"
    ]
  },

  {
    sender: "msoa@metrobankcard.com",

    // Forward all emails from this sender.
    matchAll: true,
    keywords: [],

    recipients: [
      "accounts.payable1@filepino.com",
      "irish@filepino.com",
      "accounts.payable2@filepino.com"
    ]
  },

  {
    sender: "no-reply2@globe.com.ph",
    keywords: [
      "9088922337"
    ],
    recipients: [
      "accounts.payable1@filepino.com",
      "irish@filepino.com",
      "accounts.payable2@filepino.com"
    ]
  },

  {
    sender: "no-reply2@globe.com.ph",
    keywords: [
      "9178353723"
    ],
    recipients: [
      "roselyn.salazar@lifetrackmed.com"
    ]
  },

  {
    sender: "zafajardo9@gmail.com",
    keywords: [
      "TEST"
    ],
    recipients: [
      "seo@filepino.com"
    ]
  }
]
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
  const senderSearch = CONFIG.RULES
    .map(rule => `from:${normalizeEmail_(rule.sender)}`)
    .join(" ");

  // Curly braces mean OR in Gmail search.
  const query =
    `newer_than:${CONFIG.SEARCH_LOOKBACK_DAYS}d ` +
    `{in:inbox label:"${CONFIG.FAILED_LABEL}"} ` +
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

      for (const message of thread.getMessages()) {
        if (
          (message.isInInbox() || isFailedThread) &&
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

  for (const rule of CONFIG.RULES) {
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

  for (const rule of CONFIG.RULES) {
    if (!normalizeEmail_(rule.sender).includes("@")) {
      throw new Error(
        `Invalid monitored sender: ${rule.sender}`
      );
    }

    if (!rule.matchAll && !rule.keywords.length) {
      throw new Error(
        `No keywords configured for ${rule.sender}. ` +
        `Set matchAll: true to forward everything.`
      );
    }

    if (!rule.recipients.length) {
      throw new Error(
        `No recipients configured for ${rule.sender}`
      );
    }
  }
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
