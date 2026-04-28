
// ═══════════════════════════════════════════════════════════════════════════
// FOUNDATION — config constants, pure utilities, properties access, sheet/row I/O
// ═══════════════════════════════════════════════════════════════════════════


// ═══════════════════════════════════════════════════════════════════════════
// ConfigConstants — HEADERS, STATUS, SEND_MODE, header aliases, column contract
// ═══════════════════════════════════════════════════════════════════════════

// System Version: April 27, 2026

var CONFIG = {
  SHEET_NAME: "VISA automation",
  LOGS_SHEET_NAME: "LOGS",
  AUTOMATION_SHEET_PROPERTY_KEY: "AUTOMATION_SHEET_NAMES",
  HEADER_ROW: 2,
  DATA_START_ROW: 3,
  TRIGGER_HOUR: 8,
  REPLY_SCAN_TRIGGER_HOUR: 9,
  SENDER_NAME: "Office",
  STATIC_REDIRECT_URL: "https://pastebin.com/8n85J6k6",
};

var HEADERS = {
  NO: "No.",
  CLIENT_NAME: "Client Name",
  CLIENT_EMAIL: "Client Email",
  DOC_TYPE: "Type of ID/Document",
  EXPIRY_DATE: "Expiry Date",
  NOTICE_DATE: "Notice Date",
  REMARKS: "Description",
  ATTACHMENTS: "Attached Files",
  STATUS: "Status",
  STAFF_NAME: "Name of Staff",
  STAFF_EMAIL: "Assigned Staff Email",
  SEND_MODE: "Send Mode",
  SENT_AT: "Sent At",
  SENT_THREAD_ID: "Sent Thread Id",
  SENT_MESSAGE_ID: "Sent Message Id",
  REPLY_STATUS: "Reply Status",
  REPLIED_AT: "Replied At",
  REPLY_KEYWORD: "Reply Keyword",
  OPEN_TOKEN: "Open Tracking Token",
  FIRST_OPENED_AT: "First Opened At",
  LAST_OPENED_AT: "Last Opened At",
  OPEN_COUNT: "Open Count",
  FINAL_NOTICE_SENT_AT: "Final Notice Sent At",
  FINAL_NOTICE_THREAD_ID: "Final Notice Thread Id",
  FINAL_NOTICE_MESSAGE_ID: "Final Notice Message Id",
};

// Team A — required user-input columns. Setup verifies all are present in
// the sheet (creating any missing) and the daily run validates per-row.
var REQUIRED_USER_COLUMNS = [
  "CLIENT_NAME",
  "CLIENT_EMAIL",
  "EXPIRY_DATE",
  "NOTICE_DATE",
  "REMARKS",
  "ATTACHMENTS",
  "STAFF_NAME",
  "STAFF_EMAIL",
];

// Team V — code-managed columns. Setup creates the header only when missing;
// existing headers (incl. user-renamed variants matched via aliases) are
// preserved. Order is the order ensure-* helpers add them.
var MANAGED_COLUMNS = [
  "STATUS",
  "REPLY_STATUS",
  "FINAL_NOTICE_SENT_AT",
  "FINAL_NOTICE_THREAD_ID",
  "FINAL_NOTICE_MESSAGE_ID",
  "SEND_MODE",
  "SENT_AT",
  "SENT_THREAD_ID",
  "SENT_MESSAGE_ID",
  "REPLIED_AT",
  "REPLY_KEYWORD",
  "OPEN_TOKEN",
  "FIRST_OPENED_AT",
  "LAST_OPENED_AT",
  "OPEN_COUNT",
];

var HEADER_ALIASES = {
  CLIENT_NAME: ["Seller", "Seller Name", "Buyer", "Buyer Name"],
  CLIENT_EMAIL: [
    "Seller Email",
    "Buyer Email",
    "Seller E-mail",
    "Buyer E-mail",
    "Seller Mail",
    "Buyer Mail",
  ],
  DOC_TYPE: ["Services", "Service"],
  EXPIRY_DATE: ["Due Date", "Renewal Date", "Expiry Date/Renewal Date"],
  NOTICE_DATE: ["Remaining Days", "Reminder Days"],
  REMARKS: [
    "Remarks",
    "Reminder (Email Content)",
    "Reminder Email Content",
    "Reminder Content",
  ],
  ATTACHMENTS: [
    "Attached File",
    "Gsheet",
    "GSheet",
    "Google Sheet",
    "Google Sheets",
  ],
  STATUS: ["Project Status"],
  STAFF_NAME: ["Staff Name", "Assigned Staff", "Staff", "Owner", "Handler"],
  STAFF_EMAIL: ["Staff Email", "Owner Email", "Handler Email"],
  SENT_THREAD_ID: ["Sent Thread ID"],
  SENT_MESSAGE_ID: ["Sent Message ID"],
  SEND_MODE: ["Send Option", "Mode"],
  OPEN_TOKEN: ["Open Token"],
  FIRST_OPENED_AT: ["First Open At"],
  LAST_OPENED_AT: ["Last Open At"],
  FINAL_NOTICE_SENT_AT: ["Final Notice Date", "Final Sent At"],
  FINAL_NOTICE_THREAD_ID: ["Final Notice Thread ID", "Final Thread ID"],
  FINAL_NOTICE_MESSAGE_ID: ["Final Notice Message ID", "Final Message ID"],
};

var STATUS = {
  ACTIVE: "Active",
  NOTICE_SENT: "Notice Sent",
  SENT: "Sent",
  ERROR: "Error",
  SKIPPED: "Skipped",
};

var SEND_MODE = {
  AUTO: "Auto",
  HOLD: "Hold",
  MANUAL_ONLY: "Manual Only",
};

var REPLY_STATUS = {
  PENDING: "Pending",
  REPLIED: "Replied",
};

var PROP_KEYS = {
  SPREADSHEET_ID: "SPREADSHEET_ID",
  REPLY_KEYWORDS: "REPLY_KEYWORDS",
  AI_ENABLED: "AI_ENABLED",
  AI_PROVIDER: "AI_PROVIDER",
  AI_API_KEY: "AI_API_KEY",
  AI_MODEL: "AI_MODEL",
  FALLBACK_TEMPLATE_MODE: "FALLBACK_TEMPLATE_MODE",
  FALLBACK_TEMPLATE: "FALLBACK_TEMPLATE",
  OPEN_TRACKING_BASE_URL: "OPEN_TRACKING_BASE_URL",
  DEFAULT_CC_EMAILS: "DEFAULT_CC_EMAILS",
  DAILY_TRIGGER_HOUR: "DAILY_TRIGGER_HOUR",
  DAILY_TRIGGER_MINUTE: "DAILY_TRIGGER_MINUTE",
};

var AI_PROVIDER = {
  GEMINI: "gemini",
};

var FALLBACK_TEMPLATE_MODE = {
  HARDCODED: "HARDCODED",
  PROPERTY: "PROPERTY",
};

var DEFAULT_REPLY_KEYWORDS = ["ACK", "RECEIVED", "OK"];
var DEFAULT_AI_MODEL = "models/gemini-1.5-flash";
var DEFAULT_NOTICE_OPTIONS = [
  "7 days before",
  "14 days before",
  "30 days before",
  "60 days before",
  "90 days before",
  "1 year before",
  "2 years before",
  "On expiry date",
];
var PROP_KEY_NOTICE_OPTIONS_PREFIX = "NOTICE_OPTIONS_";

var LOG_COL = {
  TIMESTAMP: 1,
  TAB: 2,
  CLIENT_NAME: 3,
  ACTION: 4,
  DETAIL: 5,
};

var TAB_CONFIG = {
  PREFIX: "TAB_CONFIG_",
  DEFAULT_HEADER_ROW: 2,
};

var TAB_CONFIG_KEYS = {
  COLUMN_MAP: "COLUMN_MAP",
  HEADER_ROW: "HEADER_ROW",
  DATA_START_ROW: "DATA_START_ROW",
  NOTICE_OPTIONS: "NOTICE_OPTIONS",
  STATUS_OPTIONS: "STATUS_OPTIONS",
  SEND_MODE_OPTIONS: "SEND_MODE_OPTIONS",
  LAST_SELECTED: "LAST_SELECTED",
};

var FLEXIBLE_HEADER_ALIASES = {
  NO: [
    "No.",
    "No",
    "Number",
    "#",
    "ID",
    "Ref",
    "Reference",
    "Ref No",
    "Ref. No.",
    "Reference No",
  ],
  CLIENT_NAME: [
    "Client Name",
    "Name",
    "Full Name",
    "Client",
    "Applicant Name",
    "Applicant",
    "Person",
    "Contact Name",
    "Seller",
    "Seller Name",
    "Buyer",
    "Buyer Name",
  ],
  CLIENT_EMAIL: [
    "Client Email",
    "Email",
    "Email Address",
    "E-mail",
    "E-mail Address",
    "Contact Email",
    "Mail",
    "Seller Email",
    "Buyer Email",
    "Seller E-mail",
    "Buyer E-mail",
    "Seller Mail",
    "Buyer Mail",
  ],
  DOC_TYPE: [
    "Type of ID/Document",
    "Document Type",
    "Doc Type",
    "Type",
    "ID Type",
    "Visa Type",
    "Permit Type",
    "Document",
    "Services",
    "Service",
  ],
  EXPIRY_DATE: [
    "Expiry Date",
    "Expiration Date",
    "Expires On",
    "Valid Until",
    "End Date",
    "Date of Expiration",
    "Expiry",
    "Due Date",
    "Renewal Date",
    "Expiry Date/Renewal Date",
  ],
  NOTICE_DATE: [
    "Notice Date",
    "Notice",
    "Reminder Date",
    "Send On",
    "Notify On",
    "Advance Notice",
    "Remaining Days",
    "Reminder Days",
  ],
  REMARKS: [
    "Description",
    "Remarks",
    "Notes",
    "Comments",
    "Message",
    "Body",
    "Email Body",
    "Note",
    "Reminder (Email Content)",
    "Reminder Email Content",
    "Reminder Content",
  ],
  ATTACHMENTS: [
    "Attached Files",
    "Attachments",
    "Files",
    "Docs",
    "Documents",
    "Attached",
    "Drive Links",
    "Attached File",
    "Gsheet",
    "GSheet",
    "Google Sheet",
    "Google Sheets",
  ],
  STATUS: [
    "Status",
    "State",
    "Send Status",
    "Processing Status",
    "Project Status",
  ],
  STAFF_NAME: [
    "Name of Staff",
    "Staff Name",
    "Assigned Staff",
    "Staff",
    "Owner",
    "Handler",
    "Assigned To",
  ],
  STAFF_EMAIL: [
    "Assigned Staff Email",
    "Staff Email",
    "Handler Email",
    "Owner Email",
    "Assignee Email",
  ],
  SEND_MODE: [
    "Send Mode",
    "Send Option",
    "Mode",
    "Send",
    "Auto Send",
    "Processing Mode",
  ],
  SENT_AT: ["Sent At", "Date Sent", "Sent On", "Processed At", "Last Sent"],
  SENT_THREAD_ID: [
    "Sent Thread Id",
    "Sent Thread ID",
    "Thread ID",
    "Gmail Thread",
    "Thread",
  ],
  SENT_MESSAGE_ID: [
    "Sent Message Id",
    "Sent Message ID",
    "Message ID",
    "Gmail Message",
  ],
  REPLY_STATUS: ["Reply Status", "Response Status", "Acknowledged", "Replied"],
  REPLIED_AT: ["Replied At", "Reply Date", "Response Date", "Replied On"],
  REPLY_KEYWORD: ["Reply Keyword", "Keyword", "Ack Keyword"],
  OPEN_TOKEN: ["Open Tracking Token", "Open Token", "Tracking Token", "Token"],
  FIRST_OPENED_AT: ["First Opened At", "First Open At", "First Viewed"],
  LAST_OPENED_AT: ["Last Opened At", "Last Open At", "Last Viewed"],
  OPEN_COUNT: ["Open Count", "View Count", "Times Opened"],
  FINAL_NOTICE_SENT_AT: [
    "Final Notice Sent At",
    "Final Notice Date",
    "Final Sent At",
  ],
  FINAL_NOTICE_THREAD_ID: [
    "Final Notice Thread Id",
    "Final Notice Thread ID",
    "Final Thread ID",
  ],
  FINAL_NOTICE_MESSAGE_ID: [
    "Final Notice Message Id",
    "Final Notice Message ID",
    "Final Message ID",
  ],
};

// ═══════════════════════════════════════════════════════════════════════════
// SharedUtils — pure helpers (parsing, normalization, dates)
// ═══════════════════════════════════════════════════════════════════════════

// 90 Shared Utils


function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || "").trim());
}


function normalizeEmailAddress(value) {
  var email = String(value || "").trim();
  if (!email) return "";

  var angleMatch = email.match(/<([^<>]+)>/);
  if (angleMatch && angleMatch[1]) {
    email = String(angleMatch[1]).trim();
  }

  return email.replace(/^[\s"'`<]+|[\s"'`>]+$/g, "").trim();
}


function normalizeEmailList(value) {
  var rawValues = [];

  if (Array.isArray(value)) {
    rawValues = value;
  } else {
    var text = String(value || "");
    if (!text.trim()) return [];
    rawValues = text.split(/[,;\n]/);
  }

  var seen = {};
  var result = [];

  for (var i = 0; i < rawValues.length; i++) {
    var email = normalizeEmailAddress(rawValues[i]);
    if (!email) continue;

    var key = email.toLowerCase();
    if (seen[key]) continue;
    seen[key] = true;
    result.push(email);
  }

  return result;
}


function parseClientEmails(rawValue) {
  var list = normalizeEmailList(rawValue);
  return list.length > 0 ? list : [];
}


function validateEmailList(emails) {
  var invalid = [];
  for (var i = 0; i < emails.length; i++) {
    if (!isValidEmail(emails[i])) invalid.push(emails[i]);
  }
  return invalid;
}


function mergeUniqueEmails() {
  var merged = [];
  var seen = {};

  for (var i = 0; i < arguments.length; i++) {
    var source = normalizeEmailList(arguments[i]);
    for (var j = 0; j < source.length; j++) {
      var email = source[j];
      var key = email.toLowerCase();
      if (seen[key]) continue;
      seen[key] = true;
      merged.push(email);
    }
  }

  return merged;
}

function getMidnight(date) {
  var d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

function isSameDay(dateA, dateB) {
  var a = getMidnight(dateA);
  var b = getMidnight(dateB);
  return (
    a.getFullYear() === b.getFullYear() &&
    a.getMonth() === b.getMonth() &&
    a.getDate() === b.getDate()
  );
}

function isTargetDateDue(targetDate, referenceDate) {
  return (
    getMidnight(targetDate).getTime() <= getMidnight(referenceDate).getTime()
  );
}


function getSupportedNoticeDateHint() {
  return (
    'Supported formats: "N days/weeks/months/years before", ' +
    '"N days/weeks/months/years", numeric day count (e.g. "7"), ' +
    '"On expiry date", or an explicit date (e.g. "2026-12-31").'
  );
}


function parseIsoDateString(value) {
  var text = String(value || "").trim();
  var match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!match) return null;

  var year = parseInt(match[1], 10);
  var month = parseInt(match[2], 10);
  var day = parseInt(match[3], 10);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;

  var parsed = new Date(year, month - 1, day);
  if (
    parsed.getFullYear() !== year ||
    parsed.getMonth() !== month - 1 ||
    parsed.getDate() !== day
  ) {
    return null;
  }

  return parsed;
}

function parseNoticeOffset(noticeStr) {
  var s = String(noticeStr || "")
    .toLowerCase()
    .trim();
  if (!s) return null;

  if (/^on(\s+the)?\s+expiry\s+date$/.test(s)) {
    return { unit: "days", value: 0 };
  }

  if (/^\d+$/.test(s)) {
    return { unit: "days", value: parseInt(s, 10) };
  }

  var relative = s.match(
    /^(\d+)\s*(day|days|week|weeks|month|months|year|years|yr|yrs)(?:\s+before)?$/,
  );
  if (relative) {
    var value = parseInt(relative[1], 10);
    var unitRaw = relative[2];

    if (unitRaw === "week" || unitRaw === "weeks") {
      return { unit: "days", value: value * 7 };
    }
    if (unitRaw === "month" || unitRaw === "months") {
      return { unit: "months", value: value };
    }
    if (
      unitRaw === "year" ||
      unitRaw === "years" ||
      unitRaw === "yr" ||
      unitRaw === "yrs"
    ) {
      return { unit: "years", value: value };
    }

    return { unit: "days", value: value };
  }

  var isoDate = parseIsoDateString(s);
  if (isoDate) {
    return { unit: "absolute_date", value: getMidnight(isoDate) };
  }

  var directDate = new Date(s);
  if (!isNaN(directDate.getTime())) {
    return { unit: "absolute_date", value: getMidnight(directDate) };
  }

  return null;
}

function computeTargetDate(expiryDate, offset) {
  if (offset.unit === "absolute_date") {
    return getMidnight(offset.value);
  }

  var target = new Date(expiryDate);
  if (offset.unit === "days") {
    target.setDate(target.getDate() - offset.value);
  } else if (offset.unit === "months") {
    target.setMonth(target.getMonth() - offset.value);
  } else if (offset.unit === "years") {
    target.setFullYear(target.getFullYear() - offset.value);
  }
  return getMidnight(target);
}

function formatDate(date) {
  var months = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];
  return (
    date.getDate() + " " + months[date.getMonth()] + " " + date.getFullYear()
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// PropertiesStore — typed wrappers around PropertiesService
// ═══════════════════════════════════════════════════════════════════════════

// 20 Properties

function getAutomationProperties() {
  return PropertiesService.getDocumentProperties();
}

function rememberSpreadsheetId(ss) {
  try {
    if (!ss) return;
    var id = ss.getId();
    if (id) setPropString(PROP_KEYS.SPREADSHEET_ID, id);
  } catch (e) {}
}

function getPropString(key, fallbackValue) {
  var value = getAutomationProperties().getProperty(key);
  if (value === null || value === undefined || String(value).trim() === "") {
    return fallbackValue || "";
  }
  return String(value).trim();
}

function setPropString(key, value) {
  var props = getAutomationProperties();
  var text = String(value === null || value === undefined ? "" : value).trim();
  if (!text) {
    props.deleteProperty(key);
    return;
  }
  props.setProperty(key, text);
}

function getPropBoolean(key, fallbackValue) {
  var raw = getAutomationProperties().getProperty(key);
  if (raw === null || raw === undefined || String(raw).trim() === "") {
    return !!fallbackValue;
  }
  return String(raw).toLowerCase().trim() === "true";
}

function setPropBoolean(key, value) {
  getAutomationProperties().setProperty(key, value ? "true" : "false");
}

// ═══════════════════════════════════════════════════════════════════════════
// SheetResolution — configured tab storage and spreadsheet lookup
// ═══════════════════════════════════════════════════════════════════════════

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

// ═══════════════════════════════════════════════════════════════════════════
// RowOps — row read/write, status writes, ensure-column helpers
// ═══════════════════════════════════════════════════════════════════════════

// 23 Sheet Row Ops

function getCellStr(row, colIndex) {
  if (!colIndex) return "";
  return String(row[colIndex - 1] || "").trim();
}

function isStatusActive(statusValue) {
  return (
    String(statusValue || "")
      .trim()
      .toLowerCase() === STATUS.ACTIVE.toLowerCase()
  );
}


function isStatusNoticeSent(statusValue) {
  return (
    String(statusValue || "")
      .trim()
      .toLowerCase() === STATUS.NOTICE_SENT.toLowerCase()
  );
}

function isStatusBlank(statusValue) {
  return String(statusValue || "").trim() === "";
}

function isProcessableStatus(statusValue) {
  return (
    isStatusActive(statusValue) ||
    isStatusNoticeSent(statusValue) ||
    isStatusBlank(statusValue)
  );
}


function isReplyStatusReplied(replyStatusValue) {
  return (
    String(replyStatusValue || "")
      .trim()
      .toLowerCase() === REPLY_STATUS.REPLIED.toLowerCase()
  );
}

function isSameNoValue(cellValue, inputValue) {
  var cellStr = String(
    cellValue === null || cellValue === undefined ? "" : cellValue,
  ).trim();
  var inputStr = String(
    inputValue === null || inputValue === undefined ? "" : inputValue,
  ).trim();

  if (!cellStr || !inputStr) return false;
  if (cellStr === inputStr) return true;

  var cellNum = Number(cellStr);
  var inputNum = Number(inputStr);
  if (!isNaN(cellNum) && !isNaN(inputNum)) return cellNum === inputNum;

  return cellStr.toLowerCase() === inputStr.toLowerCase();
}

function findRowNumberByNo(sheet, colMap, noValue, dataStartRow) {
  var startRow = dataStartRow || CONFIG.DATA_START_ROW;
  if (!colMap.NO) {
    return {
      rowNum: null,
      warning: "",
      error: 'Column "No." not found in the header row.',
    };
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
    return { rowNum: null, warning: "", error: "No data rows found." };
  }

  var numDataRows = lastRow - startRow + 1;
  var noValues = sheet
    .getRange(startRow, colMap.NO, numDataRows, 1)
    .getValues();

  var matches = [];
  for (var i = 0; i < noValues.length; i++) {
    if (isSameNoValue(noValues[i][0], noValue)) {
      matches.push(startRow + i);
    }
  }

  if (matches.length === 0) {
    return {
      rowNum: null,
      warning: "",
      error: 'No row found for No. "' + noValue + '".',
    };
  }

  if (matches.length > 1) {
    return {
      rowNum: matches[0],
      warning:
        'Multiple rows found for No. "' +
        noValue +
        '". Using first match at row ' +
        matches[0] +
        ".",
      error: "",
    };
  }

  return { rowNum: matches[0], warning: "", error: "" };
}

function setStatus(sheet, rowIndex, statusColIndex, statusValue) {
  if (!statusColIndex) return;
  sheet.getRange(rowIndex, statusColIndex).setValue(statusValue);
}


function getDefaultStatusOptions() {
  return [
    STATUS.ACTIVE,
    STATUS.NOTICE_SENT,
    STATUS.SENT,
    STATUS.ERROR,
    STATUS.SKIPPED,
  ];
}


function getStatusOptionsForTab(sheet, tabName, colMap) {
  if (!sheet || !colMap || !colMap.STATUS) return getDefaultStatusOptions();

  var dataStartRow = getTabDataStartRow(tabName);
  var sampleRow = Math.max(dataStartRow, 1);

  try {
    var rule = sheet.getRange(sampleRow, colMap.STATUS).getDataValidation();
    if (!rule) return getDefaultStatusOptions();

    var criteriaType = rule.getCriteriaType();
    var criteriaValues = rule.getCriteriaValues();
    if (
      criteriaType !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST ||
      !criteriaValues ||
      !criteriaValues[0]
    ) {
      return getDefaultStatusOptions();
    }

    var options = criteriaValues[0]
      .map(function (value) {
        return String(value || "").trim();
      })
      .filter(function (value) {
        return !!value;
      });

    return options.length > 0 ? options : getDefaultStatusOptions();
  } catch (e) {
    return getDefaultStatusOptions();
  }
}


function resolveStatusValueForTab(sheet, tabName, colMap, desiredStatus) {
  var target = String(desiredStatus || "").trim();
  if (!target) return target;

  var options = getStatusOptionsForTab(sheet, tabName, colMap);
  var targetLower = target.toLowerCase();

  for (var i = 0; i < options.length; i++) {
    if (
      String(options[i] || "")
        .trim()
        .toLowerCase() === targetLower
    ) {
      return options[i];
    }
  }

  return target;
}


function setResolvedStatus(sheet, rowIndex, colMap, tabName, desiredStatus) {
  if (!colMap || !colMap.STATUS) return;
  setStatus(
    sheet,
    rowIndex,
    colMap.STATUS,
    resolveStatusValueForTab(sheet, tabName, colMap, desiredStatus),
  );
}

function setStaffEmail(sheet, rowIndex, staffEmailColIndex, senderEmail) {
  if (!staffEmailColIndex || !senderEmail) return;
  sheet.getRange(rowIndex, staffEmailColIndex).setValue(senderEmail);
}

function setCellValueIfColumn(sheet, rowIndex, colIndex, value) {
  if (!colIndex) return;
  sheet.getRange(rowIndex, colIndex).setValue(value);
}

function clearCellValueIfColumn(sheet, rowIndex, colIndex) {
  if (!colIndex) return;
  sheet.getRange(rowIndex, colIndex).clearContent();
}

function writePostSendMetadata(sheet, rowIndex, colMap, meta) {
  var sentAt = meta && meta.sentAt ? meta.sentAt : new Date();
  var threadId = meta && meta.threadId ? String(meta.threadId) : "";
  var messageId = meta && meta.messageId ? String(meta.messageId) : "";
  var openToken = meta && meta.openToken ? String(meta.openToken) : "";

  setCellValueIfColumn(sheet, rowIndex, colMap.SENT_AT, sentAt);
  setCellValueIfColumn(sheet, rowIndex, colMap.SENT_THREAD_ID, threadId);
  setCellValueIfColumn(sheet, rowIndex, colMap.SENT_MESSAGE_ID, messageId);
  setCellValueIfColumn(sheet, rowIndex, colMap.OPEN_TOKEN, openToken);

  if (colMap.REPLY_STATUS) {
    setCellValueIfColumn(
      sheet,
      rowIndex,
      colMap.REPLY_STATUS,
      REPLY_STATUS.PENDING,
    );
  }
  clearCellValueIfColumn(sheet, rowIndex, colMap.REPLIED_AT);
  clearCellValueIfColumn(sheet, rowIndex, colMap.REPLY_KEYWORD);
}


function applyReplyStatusValidation(sheet, tabName, colMap) {
  if (!sheet || !colMap || !colMap.REPLY_STATUS) return;

  var dataStartRow = getTabDataStartRow(tabName);
  var lastRow = sheet.getLastRow();
  var dataLastRow = Math.max(lastRow, dataStartRow + 100);
  var replyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([REPLY_STATUS.PENDING, REPLY_STATUS.REPLIED], true)
    .setAllowInvalid(true)
    .build();

  sheet
    .getRange(
      dataStartRow,
      colMap.REPLY_STATUS,
      dataLastRow - dataStartRow + 1,
      1,
    )
    .setDataValidation(replyRule);
}


function ensureReplyStatusColumn(sheet, tabName, colMap) {
  if (!sheet || !colMap || !colMap.STATUS) return colMap || {};
  if (colMap.REPLY_STATUS) {
    applyReplyStatusValidation(sheet, tabName, colMap);
    return colMap;
  }

  var headerRow = getTabHeaderRow(tabName);
  sheet.insertColumnAfter(colMap.STATUS);
  sheet.getRange(headerRow, colMap.STATUS + 1).setValue(HEADERS.REPLY_STATUS);

  var updatedMap = buildColumnMap(sheet, tabName);
  applyReplyStatusValidation(sheet, tabName, updatedMap);
  return updatedMap;
}


function ensureColumnExists(sheet, tabName, logicalKey) {
  if (logicalKey && buildColumnMap(sheet, tabName)[logicalKey]) {
    return buildColumnMap(sheet, tabName)[logicalKey];
  }

  var headerRow = getTabHeaderRow(tabName);
  var colIndex = sheet.getLastColumn() + 1;
  sheet.getRange(headerRow, colIndex).setValue(HEADERS[logicalKey]);
  return buildColumnMap(sheet, tabName)[logicalKey] || colIndex;
}


function ensureReplyMetadataColumns(sheet, tabName, colMap) {
  var updatedMap = colMap || {};

  updatedMap = ensureReplyStatusColumn(sheet, tabName, updatedMap);

  if (!updatedMap.REPLIED_AT) {
    updatedMap.REPLIED_AT = ensureColumnExists(sheet, tabName, "REPLIED_AT");
  }

  return buildColumnMap(sheet, tabName);
}


function ensureFinalNoticeColumns(sheet, tabName, colMap) {
  var keys = [
    "FINAL_NOTICE_SENT_AT",
    "FINAL_NOTICE_THREAD_ID",
    "FINAL_NOTICE_MESSAGE_ID",
  ];

  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    if (!colMap[key]) {
      colMap[key] = ensureColumnExists(sheet, tabName, key);
    }
  }

  return colMap;
}


function ensureOpenTrackingColumns(sheet, tabName, colMap) {
  var keys = ["OPEN_TOKEN", "FIRST_OPENED_AT", "LAST_OPENED_AT", "OPEN_COUNT"];
  var updatedMap = colMap || {};

  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    if (!updatedMap[key]) {
      updatedMap[key] = ensureColumnExists(sheet, tabName, key);
    }
  }

  return buildColumnMap(sheet, tabName);
}


function ensureSetupAutomationColumns(sheet, tabName, colMap) {
  var updatedMap = colMap || buildColumnMap(sheet, tabName);

  // MANAGED_COLUMNS lists every code-owned column. Header is added only
  // when missing (existing user-renamed variants are matched via aliases).
  for (var i = 0; i < MANAGED_COLUMNS.length; i++) {
    var key = MANAGED_COLUMNS[i];
    if (!updatedMap[key]) {
      updatedMap[key] = ensureColumnExists(sheet, tabName, key);
    }
  }

  // Refresh map then apply REPLY_STATUS validation rule.
  updatedMap = buildColumnMap(sheet, tabName);
  if (updatedMap.REPLY_STATUS) {
    applyReplyStatusValidation(sheet, tabName, updatedMap);
  }

  return updatedMap;
}

// Adds any Team A user-input columns that are missing. Used by the setup
// wizard when the user opts in to creating them.
function ensureUserInputColumns(sheet, tabName, colMap) {
  var updatedMap = colMap || buildColumnMap(sheet, tabName);

  for (var i = 0; i < REQUIRED_USER_COLUMNS.length; i++) {
    var key = REQUIRED_USER_COLUMNS[i];
    if (!updatedMap[key]) {
      updatedMap[key] = ensureColumnExists(sheet, tabName, key);
    }
  }

  return buildColumnMap(sheet, tabName);
}


function writeFinalNoticeMetadata(sheet, rowIndex, colMap, meta) {
  var finalAt = meta && meta.sentAt ? meta.sentAt : new Date();
  var threadId = meta && meta.threadId ? String(meta.threadId) : "";
  var messageId = meta && meta.messageId ? String(meta.messageId) : "";

  setCellValueIfColumn(sheet, rowIndex, colMap.FINAL_NOTICE_SENT_AT, finalAt);
  setCellValueIfColumn(
    sheet,
    rowIndex,
    colMap.FINAL_NOTICE_THREAD_ID,
    threadId,
  );
  setCellValueIfColumn(
    sheet,
    rowIndex,
    colMap.FINAL_NOTICE_MESSAGE_ID,
    messageId,
  );
}
