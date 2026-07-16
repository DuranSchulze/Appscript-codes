const CONFIG = {
  CHECK_EVERY_MINUTES: 5,
  SEARCH_LOOKBACK_DAYS: 30,
  MAX_THREADS_PER_RUN: 500,
  RETAIN_PROCESSED_DAYS: 60,

  PROCESSED_PREFIX: "SEC_FORWARDED_",
  STARTED_AT_KEY: "SEC_AUTOMATION_STARTED_AT",
  LAST_CLEANUP_KEY: "SEC_LAST_CLEANUP",

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
      "accounts@filepino.com"
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
    keywords: [
      "524005XXXXXXXXXX"
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
  removeExistingTriggers_();

  const properties = PropertiesService.getScriptProperties();

  if (!properties.getProperty(CONFIG.STARTED_AT_KEY)) {
    properties.setProperty(
      CONFIG.STARTED_AT_KEY,
      String(Date.now())
    );
  }

  ScriptApp.newTrigger("monitorAndForwardSecEmails")
    .timeBased()
    .everyMinutes(CONFIG.CHECK_EVERY_MINUTES)
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

        Logger.log(
          `Forwarded message ${messageId} to: ` +
          rule.recipients.join(", ")
        );
      } catch (error) {
        Logger.log(
          `Failed forwarding ${messageId}: ${error.message}`
        );

        // Failed messages will be retried next time.
        message.star();
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
    `in:inbox newer_than:${CONFIG.SEARCH_LOOKBACK_DAYS}d ` +
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
      for (const message of thread.getMessages()) {
        if (
          message.isInInbox() &&
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
  for (
    const trigger of ScriptApp.getProjectTriggers()
  ) {
    if (
      trigger.getHandlerFunction() ===
      "monitorAndForwardSecEmails"
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

  for (const rule of CONFIG.RULES) {
    if (!normalizeEmail_(rule.sender).includes("@")) {
      throw new Error(
        `Invalid monitored sender: ${rule.sender}`
      );
    }

    if (!rule.keywords.length) {
      throw new Error(
        `No keywords configured for ${rule.sender}`
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
      key === CONFIG.STARTED_AT_KEY
    ) {
      properties.deleteProperty(key);
    }
  }

  Logger.log("SEC forwarding automation reset.");
}
