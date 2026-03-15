/**
 * CAR Special Projects Monitoring System
 * Google Apps Script Implementation
 *
 * Features:
 * - Deadline tracking (DST and CGT/DOD calculations)
 * - Automated reminder generation
 * - Workflow activity logging
 * - Email notifications for approaching deadlines
 * - Multi-sheet support across all service types
 * - Flexible sheet/header names via Script Properties (survives renames)
 */

// ============================================================================
// CONFIGURATION DEFAULTS
// These are fallback values only. All runtime code uses getConfig().
// To change a name without editing code: CAR Monitoring → Settings
// ============================================================================

const CONFIG_DEFAULTS = {
  SHEETS: {
    CAR: "CAR -- Transfer of Shares",
    REAL_PROPERTY: "Real Property Services",
    TITLE_TRANSFER: "Title Transfer",
    APOSTILLE: "Apostille & Special Projects",
    COMPLETED: "Completed Projects",
    DASHBOARD: "Dashboard",
    ACTIVITY_LOG: "Activity Log",
  },
  HEADERS: {
    NO: "No.",
    DATE: "Date",
    CLIENT_NAME: "Company / Client Name",
    SELLER_DONOR: "Seller / Donor",
    BUYER_DONEE: "Buyer / Donee",
    SERVICE_TYPE: "Service Type",
    NOTARY_DATE: "Notary Date",
    DST_DUE_DATE: "DST Due Date",
    CGT_DOD_DUE_DATE: "CGT / DOD Due Date",
    DST_REMAINING: "DST Remaining Days",
    CGT_REMAINING: "CGT Remaining Days",
    STATUS: "Project Status",
    REMINDER: "Reminder",
    REMARKS: "Remarks",
    STAFF_EMAIL: "Staff Email",
    CLIENT_EMAIL: "Client Email",
    REVISIT_DATE: "Revisit Date",
    REVISIT_STATUS: "Revisit Status",
    REVISIT_NOTES: "Revisit Notes",
  },
  STATUS: {
    ONGOING: "On Going",
    COMPLETED: "Completed",
  },
  SERVICE_TYPES: [
    "Sale",
    "Donation",
    "Estate",
    "Transfer",
    "Apostille",
    "Other",
  ],
  REMINDER_DAYS: [7, 3, 1, 0, -1],
  DST_RULE: { day: 5, monthOffset: 1 },
  CGT_RULE: { days: 30 },
};

// Keep CONFIG as alias to defaults so any legacy direct references don't hard-break.
const CONFIG = CONFIG_DEFAULTS;
const EMAIL_RECIPIENT_MODES = {
  STAFF_ONLY: "staff_only",
  STAFF_AND_CLIENT: "staff_and_client",
};
const REVISIT_OPEN_STATUSES = ["open", "pending", "for revisit", "revisit"];
const DEFAULT_AUTOMATION_HOUR = 8;
const DEFAULT_AUTOMATION_MINUTE = 0;

// ============================================================================
// RUNTIME CONFIG — reads Script Properties, falls back to defaults
// ============================================================================

let _configCache = null;

function getConfig() {
  if (_configCache) return _configCache;

  migrateLegacySettings();
  const props = PropertiesService.getScriptProperties().getProperties();

  const sheets = {};
  Object.keys(CONFIG_DEFAULTS.SHEETS).forEach((key) => {
    const propKey = "SHEET_" + key;
    sheets[key] = props[propKey] || CONFIG_DEFAULTS.SHEETS[key];
  });

  const headers = {};
  Object.keys(CONFIG_DEFAULTS.HEADERS).forEach((key) => {
    const propKey = "HEADER_" + key;
    headers[key] = props[propKey] || CONFIG_DEFAULTS.HEADERS[key];
  });

  const statusOngoing =
    props["STATUS_ONGOING"] || CONFIG_DEFAULTS.STATUS.ONGOING;
  const statusCompleted =
    props["STATUS_COMPLETED"] || CONFIG_DEFAULTS.STATUS.COMPLETED;

  const headerRowProp = parseInt(props["HEADER_ROW"], 10);
  const dataStartRowProp = parseInt(props["DATA_START_ROW"], 10);
  const automationHourProp = parseInt(props["AUTOMATION_HOUR"], 10);
  const automationMinuteProp = parseInt(props["AUTOMATION_MINUTE"], 10);
  const workingSheets = parseJsonArrayProperty(props["WORKING_SHEETS"]);

  _configCache = {
    SHEETS: sheets,
    HEADERS: headers,
    STATUS: { ONGOING: statusOngoing, COMPLETED: statusCompleted },
    SERVICE_TYPES: CONFIG_DEFAULTS.SERVICE_TYPES,
    REMINDER_DAYS: CONFIG_DEFAULTS.REMINDER_DAYS,
    DST_RULE: CONFIG_DEFAULTS.DST_RULE,
    CGT_RULE: CONFIG_DEFAULTS.CGT_RULE,
    HEADER_ROW: isNaN(headerRowProp) ? null : headerRowProp,
    DATA_START_ROW: isNaN(dataStartRowProp) ? null : dataStartRowProp,
    WORKING_SHEETS: workingSheets,
    AUTOMATION_HOUR: isNaN(automationHourProp)
      ? DEFAULT_AUTOMATION_HOUR
      : automationHourProp,
    AUTOMATION_MINUTE: isNaN(automationMinuteProp)
      ? DEFAULT_AUTOMATION_MINUTE
      : automationMinuteProp,
    EMAIL_RECIPIENT_MODE:
      props["EMAIL_RECIPIENT_MODE"] || EMAIL_RECIPIENT_MODES.STAFF_ONLY,
  };

  return _configCache;
}

function invalidateConfigCache() {
  _configCache = null;
}

function parseJsonArrayProperty(raw) {
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function migrateLegacySettings() {
  const props = PropertiesService.getScriptProperties();
  const legacyWorkingSheet = props.getProperty("WORKING_SHEET");
  const workingSheets = props.getProperty("WORKING_SHEETS");

  if (legacyWorkingSheet && !workingSheets) {
    props.setProperty("WORKING_SHEETS", JSON.stringify([legacyWorkingSheet]));
    props.deleteProperty("WORKING_SHEET");
  }
}

function getWorkingSheets() {
  return getConfig().WORKING_SHEETS || [];
}

function setWorkingSheets(sheetNames) {
  const cleaned = Array.from(
    new Set(
      (sheetNames || [])
        .map((name) => (name || "").toString().trim())
        .filter(Boolean),
    ),
  );
  const props = PropertiesService.getScriptProperties();
  if (cleaned.length === 0) props.deleteProperty("WORKING_SHEETS");
  else props.setProperty("WORKING_SHEETS", JSON.stringify(cleaned));
  props.deleteProperty("WORKING_SHEET");
  invalidateConfigCache();
}

function getConfiguredAutomationTime() {
  const cfg = getConfig();
  return { hour: cfg.AUTOMATION_HOUR, minute: cfg.AUTOMATION_MINUTE };
}

function setConfiguredAutomationTime(hour, minute) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("AUTOMATION_HOUR", String(hour));
  props.setProperty("AUTOMATION_MINUTE", String(minute));
  invalidateConfigCache();
}

function getEmailRecipientMode() {
  return getConfig().EMAIL_RECIPIENT_MODE || EMAIL_RECIPIENT_MODES.STAFF_ONLY;
}

function setEmailRecipientMode(mode) {
  PropertiesService.getScriptProperties().setProperty(
    "EMAIL_RECIPIENT_MODE",
    mode,
  );
  invalidateConfigCache();
}

function formatTimeLabel(hour, minute) {
  const safeHour = Number(hour) || 0;
  const safeMinute = Number(minute) || 0;
  const date = new Date(2000, 0, 1, safeHour, safeMinute, 0, 0);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "hh:mm a");
}

// ============================================================================
// NOTIFICATION EMAIL LIST — global CC recipients for all reminder emails
// ============================================================================

function getNotificationEmails() {
  const raw = PropertiesService.getScriptProperties().getProperty(
    "NOTIFICATION_EMAILS",
  );
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function setNotificationEmails(arr) {
  PropertiesService.getScriptProperties().setProperty(
    "NOTIFICATION_EMAILS",
    JSON.stringify(arr),
  );
}

function viewNotificationEmails() {
  const ui = SpreadsheetApp.getUi();
  const emails = getNotificationEmails();
  if (emails.length === 0) {
    ui.alert(
      "Notification Emails",
      "No notification emails configured.",
      ui.ButtonSet.OK,
    );
    return;
  }
  const list = emails.map((e, i) => `  ${i + 1}. ${e}`).join("\n");
  ui.alert(
    "Notification Emails",
    `Current global CC recipients:\n\n${list}`,
    ui.ButtonSet.OK,
  );
}

function addNotificationEmail() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Add Notification Email",
    "Enter the email address to add to the global CC list:",
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const email = response.getResponseText().trim().toLowerCase();
  if (!email) return;
  if (!email.includes("@")) {
    ui.alert("Invalid email address. Please include an '@' sign.");
    return;
  }

  const emails = getNotificationEmails();
  if (emails.includes(email)) {
    ui.alert(`"${email}" is already in the notification list.`);
    return;
  }

  emails.push(email);
  setNotificationEmails(emails);
  logActivity("SYSTEM", "Notification Email Added", email);
  ui.alert(
    `"${email}" has been added to the notification list.\n\nTotal: ${emails.length} email(s).`,
  );
}

function removeNotificationEmail() {
  const ui = SpreadsheetApp.getUi();
  const emails = getNotificationEmails();

  if (emails.length === 0) {
    ui.alert(
      "Notification Emails",
      "No notification emails to remove.",
      ui.ButtonSet.OK,
    );
    return;
  }

  const list = emails.map((e, i) => `  ${i + 1}. ${e}`).join("\n");
  const response = ui.prompt(
    "Remove Notification Email",
    `Current list:\n${list}\n\nEnter the number of the email to remove:`,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();
  const index = parseInt(input, 10) - 1;

  if (isNaN(index) || index < 0 || index >= emails.length) {
    ui.alert(
      `Invalid selection. Please enter a number between 1 and ${emails.length}.`,
    );
    return;
  }

  const removed = emails.splice(index, 1)[0];
  setNotificationEmails(emails);
  logActivity("SYSTEM", "Notification Email Removed", removed);
  ui.alert(
    `"${removed}" has been removed.\n\nRemaining: ${emails.length} email(s).`,
  );
}

function clearNotificationEmails() {
  const ui = SpreadsheetApp.getUi();
  const emails = getNotificationEmails();

  if (emails.length === 0) {
    ui.alert(
      "Notification Emails",
      "The notification list is already empty.",
      ui.ButtonSet.OK,
    );
    return;
  }

  const response = ui.alert(
    "Clear Notification Emails",
    `This will remove all ${emails.length} email(s) from the global CC list. Continue?`,
    ui.ButtonSet.YES_NO,
  );
  if (response !== ui.Button.YES) return;

  setNotificationEmails([]);
  logActivity(
    "SYSTEM",
    "Notification Emails Cleared",
    `Removed ${emails.length} email(s)`,
  );
  ui.alert("All notification emails have been cleared.");
}

function getWorkingSheet() {
  const sheets = getWorkingSheets();
  return sheets.length === 1 ? sheets[0] : null;
}

function setWorkingSheet(sheetName) {
  setWorkingSheets(sheetName ? [sheetName] : []);
}

function getEligibleAutomationSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  const systemSheets = [
    cfg.SHEETS.COMPLETED,
    cfg.SHEETS.DASHBOARD,
    cfg.SHEETS.ACTIVITY_LOG,
  ];

  return ss
    .getSheets()
    .map((sheet) => sheet.getName())
    .filter((name) => !systemSheets.includes(name));
}

function selectWorkingSheets() {
  const ui = SpreadsheetApp.getUi();
  const allSheets = getEligibleAutomationSheets();

  if (allSheets.length === 0) {
    ui.alert("No sheets found. Please add a data sheet first.");
    return;
  }

  const current = getWorkingSheets();
  const currentLabel =
    current.length > 0
      ? `Current: ${current.join(", ")}`
      : "Current: (auto-detect eligible sheets)";
  const list = allSheets.map((name, i) => `  ${i + 1}. ${name}`).join("\n");

  const response = ui.prompt(
    "Quick Setup — Automation Tabs",
    `${currentLabel}\n\nAvailable sheets:\n${list}\n\nEnter one or more numbers separated by commas (example: 1,3,4).\nLeave blank to use all eligible sheets automatically:`,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();

  if (input === "") {
    setWorkingSheets([]);
    logActivity(
      "SYSTEM",
      "Automation Tabs",
      "Reset to auto-detect all eligible sheets",
    );
    ui.alert(
      "Automation tabs cleared. The script will use all eligible project sheets automatically.",
    );
    return;
  }

  const indexes = Array.from(
    new Set(
      input
        .split(",")
        .map((part) => parseInt(part.trim(), 10) - 1)
        .filter((value) => !isNaN(value)),
    ),
  );
  if (
    indexes.length === 0 ||
    indexes.some((index) => index < 0 || index >= allSheets.length)
  ) {
    ui.alert(
      `Invalid selection. Enter one or more numbers between 1 and ${allSheets.length}.`,
    );
    return;
  }

  const selected = indexes.map((index) => allSheets[index]);
  setWorkingSheets(selected);
  logActivity("SYSTEM", "Automation Tabs", `Set to ${selected.join(", ")}`);
  ui.alert(
    `Automation tabs set to:\n\n${selected.join("\n")}\n\nNext, confirm the header row and data start row.`,
  );
}

function selectWorkingSheet() {
  selectWorkingSheets();
}

function getProjectSheets() {
  const selected = getWorkingSheets();
  if (selected.length > 0) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return selected.filter((sheetName) => !!ss.getSheetByName(sheetName));
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  const configured = [
    cfg.SHEETS.CAR,
    cfg.SHEETS.REAL_PROPERTY,
    cfg.SHEETS.TITLE_TRANSFER,
    cfg.SHEETS.APOSTILLE,
  ];
  const detected = ss
    .getSheets()
    .map((sheet) => sheet.getName())
    .filter((name) => isProjectSheetName(name));

  return Array.from(new Set(configured.concat(detected))).filter((name) =>
    isProjectSheetName(name),
  );
}

function getHeaderAliases() {
  const cfg = getConfig();
  return {
    [cfg.HEADERS.CLIENT_NAME]: [
      "Company",
      "Company / Client Name",
      "Company Name",
      "Client",
      "Client Name",
    ],
    [cfg.HEADERS.SERVICE_TYPE]: ["Services", "Service", "Service Type"],
    [cfg.HEADERS.NOTARY_DATE]: [
      "NOTARY DATE OF DOCUMENT",
      "Notary Date of Document",
      "Notary Date",
    ],
    [cfg.HEADERS.STATUS]: ["Current Status", "Status", "Project Status"],
    [cfg.HEADERS.REMARKS]: ["Remarks", "Notes"],
    [cfg.HEADERS.STAFF_EMAIL]: [
      "Email Address",
      "Staff Email",
      "Assignee Email",
      "Email",
    ],
    [cfg.HEADERS.CLIENT_EMAIL]: [
      "Email Address",
      "Client Email",
      "Contact Email",
      "Email",
    ],
    [cfg.HEADERS.DST_DUE_DATE]: ["DST Due Date"],
    [cfg.HEADERS.CGT_DOD_DUE_DATE]: ["CGT / DOD Due Date", "CGT/ DOD Due Date"],
    [cfg.HEADERS.DST_REMAINING]: ["Remaining Days", "DST Remaining Days"],
    [cfg.HEADERS.CGT_REMAINING]: ["Remaining Days", "CGT Remaining Days"],
    [cfg.HEADERS.REMINDER]: ["Reminder", "Alert"],
    [cfg.HEADERS.REVISIT_DATE]: ["Revisit Date", "Follow-up Date"],
    [cfg.HEADERS.REVISIT_STATUS]: ["Revisit Status", "Follow-up Status"],
    [cfg.HEADERS.REVISIT_NOTES]: ["Revisit Notes", "Follow-up Notes"],
  };
}

function normalizeHeaderName(header) {
  return (header || "")
    .toString()
    .toLowerCase()
    .replace(/["']/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function getHeaderRow(sheet) {
  return detectHeaderRow(sheet);
}

function getHeaders(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (!lastColumn) return [];
  const row = getHeaderRow(sheet);
  return sheet.getRange(row, 1, 1, lastColumn).getValues()[0];
}

function findHeaderIndexByName(headers, headerName) {
  if (!headers || !headers.length) return null;

  const aliases = getHeaderAliases();
  const targetNames = [headerName].concat(aliases[headerName] || []);
  const normalizedTargets = targetNames.map((name) =>
    normalizeHeaderName(name),
  );

  for (let i = 0; i < headers.length; i++) {
    const normalizedHeader = normalizeHeaderName(headers[i]);
    if (!normalizedHeader) continue;

    if (normalizedTargets.includes(normalizedHeader)) return i + 1;

    if (
      headerName === getConfig().HEADERS.DST_DUE_DATE &&
      normalizedHeader.indexOf("dst due date") === 0
    ) {
      return i + 1;
    }

    if (
      headerName === getConfig().HEADERS.CGT_DOD_DUE_DATE &&
      (normalizedHeader.indexOf("cgt / dod due date") === 0 ||
        normalizedHeader.indexOf("cgt/ dod due date") === 0)
    ) {
      return i + 1;
    }
  }

  return null;
}

function isDueDateHeader(header) {
  return normalizeHeaderName(header).indexOf("due date") >= 0;
}

function isProjectSheetName(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return false;

  const cfg = getConfig();
  const systemSheets = [
    cfg.SHEETS.COMPLETED,
    cfg.SHEETS.DASHBOARD,
    cfg.SHEETS.ACTIVITY_LOG,
  ];
  if (systemSheets.includes(sheetName)) return false;

  const headers = getHeaders(sheet);
  if (!headers.length) return false;

  const hasNotaryDate = !!findHeaderIndexByName(
    headers,
    cfg.HEADERS.NOTARY_DATE,
  );
  const hasDueDate = headers.some((header) => isDueDateHeader(header));
  const hasReminder = headers.some(
    (header) => normalizeHeaderName(header) === "reminder",
  );

  return hasNotaryDate || hasDueDate || hasReminder;
}

function getDeadlineGroups(sheet) {
  const headers = getHeaders(sheet);
  const groups = [];

  for (let i = 0; i < headers.length; i++) {
    if (!isDueDateHeader(headers[i])) continue;

    const group = {
      label: headers[i],
      dueDateCol: i + 1,
      remainingCol: null,
      statusCol: null,
      reminderCol: null,
    };

    for (let j = i + 1; j < headers.length; j++) {
      if (isDueDateHeader(headers[j])) break;

      const normalized = normalizeHeaderName(headers[j]);
      if (!group.remainingCol && normalized === "remaining days") {
        group.remainingCol = j + 1;
        continue;
      }

      if (!group.statusCol && normalized === "status") {
        group.statusCol = j + 1;
        continue;
      }

      if (!group.reminderCol && normalized === "reminder") {
        group.reminderCol = j + 1;
      }
    }

    groups.push(group);
  }

  return groups;
}

function parseDueDateRule(headerText) {
  const normalized = normalizeHeaderName(headerText);
  if (!normalized) return null;

  if (normalized.indexOf("every 5th of the following month") >= 0) {
    return { type: "nextMonthFixedDay", day: 5, base: "notary" };
  }

  const dayMatch = normalized.match(
    /(\d+)\s+days?\s+after(?:\s+the)?\s+notary date/,
  );
  if (dayMatch) {
    return { type: "daysAfter", days: Number(dayMatch[1]), base: "notary" };
  }

  const yearMatch = normalized.match(
    /(\d+)\s+years?\s+from(?:\s+the)?\s+decedent(?:'s)?\s+date of death/,
  );
  if (yearMatch) {
    return { type: "yearsAfter", years: Number(yearMatch[1]), base: "death" };
  }

  return null;
}

function calculateDueDateFromRule(baseDate, rule) {
  if (!baseDate || !(baseDate instanceof Date) || !rule) return null;

  const date = new Date(baseDate);
  if (rule.type === "nextMonthFixedDay") {
    date.setMonth(date.getMonth() + 1);
    date.setDate(rule.day);
    return date;
  }

  if (rule.type === "daysAfter") {
    date.setDate(date.getDate() + rule.days);
    return date;
  }

  if (rule.type === "yearsAfter") {
    date.setFullYear(date.getFullYear() + rule.years);
    return date;
  }

  return null;
}

function getBaseDateForRule(sheet, row, rule) {
  if (!rule) return null;

  const cfg = getConfig();
  if (rule.base === "notary") {
    const notaryCol = getColumnIndex(sheet, cfg.HEADERS.NOTARY_DATE);
    return notaryCol ? sheet.getRange(row, notaryCol).getValue() : null;
  }

  if (rule.base === "death") {
    const headers = getHeaders(sheet);
    const deathIndex = headers.findIndex(
      (header) => normalizeHeaderName(header).indexOf("date of death") >= 0,
    );
    return deathIndex >= 0
      ? sheet.getRange(row, deathIndex + 1).getValue()
      : null;
  }

  return null;
}

function buildReminderText(label, remainingDays) {
  if (remainingDays === null || remainingDays === "" || isNaN(remainingDays)) {
    return "";
  }

  const cleanLabel = (label || "Deadline")
    .toString()
    .replace(/\s+/g, " ")
    .trim();

  if (remainingDays < 0) {
    return `${cleanLabel} OVERDUE by ${Math.abs(remainingDays)} days`;
  }

  if (remainingDays <= 7) {
    return `${cleanLabel} due in ${remainingDays} days`;
  }

  return "";
}

// ============================================================================
// SETTINGS — view and edit sheet/header names via menu dialog
// ============================================================================

function showSettings() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties().getProperties();
  const cfg = getConfig();
  const workingSheets = getWorkingSheets();
  const automationTime = formatTimeLabel(
    cfg.AUTOMATION_HOUR,
    cfg.AUTOMATION_MINUTE,
  );

  let msg = "CAR Monitoring — Current Settings\n";
  msg += "=".repeat(44) + "\n\n";
  msg += "SHEET NAMES (Script Property → Current Value)\n";
  Object.keys(CONFIG_DEFAULTS.SHEETS).forEach((key) => {
    const propKey = "SHEET_" + key;
    const isCustom = !!props[propKey];
    msg += `  ${propKey}: "${cfg.SHEETS[key]}"${isCustom ? " (custom)" : " (default)"}\n`;
  });
  msg += "\nCOLUMN HEADERS (Script Property → Current Value)\n";
  Object.keys(CONFIG_DEFAULTS.HEADERS).forEach((key) => {
    const propKey = "HEADER_" + key;
    const isCustom = !!props[propKey];
    msg += `  ${propKey}: "${cfg.HEADERS[key]}"${isCustom ? " (custom)" : " (default)"}\n`;
  });
  msg += "\nSTATUS VALUES\n";
  msg += `  STATUS_ONGOING: "${cfg.STATUS.ONGOING}"${props["STATUS_ONGOING"] ? " (custom)" : " (default)"}\n`;
  msg += `  STATUS_COMPLETED: "${cfg.STATUS.COMPLETED}"${props["STATUS_COMPLETED"] ? " (custom)" : " (default)"}\n`;
  msg += "\nAUTOMATION TARGET\n";
  msg += `  WORKING_SHEETS: ${workingSheets.length > 0 ? workingSheets.join(", ") : "(auto-detect eligible sheets)"}${workingSheets.length > 0 ? " (custom)" : ""}\n`;
  msg += "\nROW LAYOUT\n";
  msg += `  HEADER_ROW: ${cfg.HEADER_ROW || "(auto-detect)"}${props["HEADER_ROW"] ? " (custom)" : ""}\n`;
  msg += `  DATA_START_ROW: ${cfg.DATA_START_ROW || "(header row + 1)"}${props["DATA_START_ROW"] ? " (custom)" : ""}\n`;
  msg += "\nEMAIL\n";
  msg += `  EMAIL_RECIPIENT_MODE: ${cfg.EMAIL_RECIPIENT_MODE}\n`;
  msg += `  NOTIFICATION_EMAILS: ${getNotificationEmails().length > 0 ? getNotificationEmails().join(", ") : "(none)"}\n`;
  msg += "\nAUTOMATION SCHEDULE\n";
  msg += `  Daily Run Time: ${automationTime}\n`;
  msg += "\n— Use Quick Setup for guided configuration\n";
  msg += "— Use Settings / Automation / Email menus for individual updates\n";
  msg += "— Use Reset to Defaults to clear custom settings";

  ui.alert("Settings", msg, ui.ButtonSet.OK);
}

function configureNotificationEmails() {
  const ui = SpreadsheetApp.getUi();
  const current = getNotificationEmails();
  const response = ui.prompt(
    "Notification CC Emails",
    `Current CC emails: ${current.length ? current.join(", ") : "(none)"}\n\nEnter one or more email addresses separated by commas.\nLeave blank to clear the CC list:`,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();
  const emails = input
    ? Array.from(
        new Set(
          input
            .split(",")
            .map((email) => email.trim().toLowerCase())
            .filter((email) => email && email.includes("@")),
        ),
      )
    : [];

  setNotificationEmails(emails);
  logActivity(
    "SYSTEM",
    "Notification Emails Updated",
    emails.length ? emails.join(", ") : "Cleared",
  );
  ui.alert(
    "Notification CC Emails",
    emails.length
      ? `Saved ${emails.length} CC email(s).`
      : "The CC email list has been cleared.",
    ui.ButtonSet.OK,
  );
}

function configureDailyAutomationTime() {
  const ui = SpreadsheetApp.getUi();
  const current = getConfiguredAutomationTime();
  const response = ui.prompt(
    "Automation Schedule",
    `Current daily run time: ${formatTimeLabel(current.hour, current.minute)}\n\nEnter the daily run time in 24-hour format HH:MM (example: 08:00 or 14:30):`,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();
  const match = input.match(/^(\d{1,2}):(\d{2})$/);
  if (!match) {
    ui.alert("Invalid time. Please use HH:MM in 24-hour format.");
    return;
  }

  const hour = Number(match[1]);
  const minute = Number(match[2]);
  if (hour < 0 || hour > 23 || minute < 0 || minute > 59) {
    ui.alert("Invalid time. Hour must be 0-23 and minute must be 0-59.");
    return;
  }

  setConfiguredAutomationTime(hour, minute);
  syncDailyTrigger(false);
  logActivity(
    "SYSTEM",
    "Automation Schedule Updated",
    formatTimeLabel(hour, minute),
  );
  ui.alert(
    "Automation Schedule",
    `Daily automation time saved as ${formatTimeLabel(hour, minute)}.`,
    ui.ButtonSet.OK,
  );
}

function configureEmailRecipientMode() {
  const ui = SpreadsheetApp.getUi();
  const current = getEmailRecipientMode();
  const response = ui.prompt(
    "Reminder Recipient Mode",
    `Current mode: ${current}\n\nEnter:\n  1 = Staff Email only\n  2 = Staff Email and Client Email`,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();
  let mode = null;
  if (input === "1") mode = EMAIL_RECIPIENT_MODES.STAFF_ONLY;
  if (input === "2") mode = EMAIL_RECIPIENT_MODES.STAFF_AND_CLIENT;
  if (!mode) {
    ui.alert("Invalid selection. Enter 1 or 2.");
    return;
  }

  setEmailRecipientMode(mode);
  logActivity("SYSTEM", "Email Recipient Mode Updated", mode);
  ui.alert(
    "Reminder Recipient Mode",
    `Recipient mode saved as "${mode}".`,
    ui.ButtonSet.OK,
  );
}

function resetSettingsToDefaults() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Reset Settings",
    "Reset ALL sheet names and column headers to their default values?",
    ui.ButtonSet.YES_NO,
  );
  if (response !== ui.Button.YES) return;

  const scriptProps = PropertiesService.getScriptProperties();
  Object.keys(CONFIG_DEFAULTS.SHEETS).forEach((key) =>
    scriptProps.deleteProperty("SHEET_" + key),
  );
  Object.keys(CONFIG_DEFAULTS.HEADERS).forEach((key) =>
    scriptProps.deleteProperty("HEADER_" + key),
  );
  scriptProps.deleteProperty("STATUS_ONGOING");
  scriptProps.deleteProperty("STATUS_COMPLETED");
  scriptProps.deleteProperty("HEADER_ROW");
  scriptProps.deleteProperty("DATA_START_ROW");
  scriptProps.deleteProperty("WORKING_SHEET");
  scriptProps.deleteProperty("WORKING_SHEETS");
  scriptProps.deleteProperty("AUTOMATION_HOUR");
  scriptProps.deleteProperty("AUTOMATION_MINUTE");
  scriptProps.deleteProperty("EMAIL_RECIPIENT_MODE");
  scriptProps.deleteProperty("NOTIFICATION_EMAILS");
  invalidateConfigCache();

  ui.alert("All settings have been reset to defaults.");
  logActivity(
    "SYSTEM",
    "Settings Reset",
    "All guided settings restored to defaults",
  );
}

function setHeaderRow() {
  const ui = SpreadsheetApp.getUi();
  const current = getConfig().HEADER_ROW;
  const response = ui.prompt(
    "Set Header Row",
    `The header row is the row that contains your column names (Company, Seller/Donor, etc.).

Current value: ${current || "auto-detect"}

Enter the row number (e.g. 3), or leave blank to use auto-detection:`,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();
  const scriptProps = PropertiesService.getScriptProperties();

  if (input === "") {
    scriptProps.deleteProperty("HEADER_ROW");
    invalidateConfigCache();
    ui.alert("Header row reset to auto-detection.");
    logActivity("SYSTEM", "Header Row", "Reset to auto-detect");
    return;
  }

  const num = parseInt(input, 10);
  if (isNaN(num) || num < 1) {
    ui.alert("Invalid input. Please enter a positive row number.");
    return;
  }

  scriptProps.setProperty("HEADER_ROW", String(num));
  invalidateConfigCache();
  logActivity("SYSTEM", "Header Row", `Set to row ${num}`);
  ui.alert(
    `Header row set to row ${num}.`,
  );
}

function setDataStartRow() {
  const ui = SpreadsheetApp.getUi();
  const current = getConfig().DATA_START_ROW;
  const response = ui.prompt(
    "Set Data Start Row",
    `The data start row is where your first data record begins (below the header row).

Current value: ${current || "header row + 1"}

Enter the row number (e.g. 4), or leave blank to use header row + 1:`,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();
  const scriptProps = PropertiesService.getScriptProperties();

  if (input === "") {
    scriptProps.deleteProperty("DATA_START_ROW");
    invalidateConfigCache();
    ui.alert("Data start row reset to header row + 1.");
    logActivity("SYSTEM", "Data Start Row", "Reset to header row + 1");
    return;
  }

  const num = parseInt(input, 10);
  if (isNaN(num) || num < 1) {
    ui.alert("Invalid input. Please enter a positive row number.");
    return;
  }

  scriptProps.setProperty("DATA_START_ROW", String(num));
  invalidateConfigCache();
  logActivity("SYSTEM", "Data Start Row", `Set to row ${num}`);
  ui.alert(
    `Data start row set to row ${num}.`,
  );
}

// ============================================================================
// SETUP WIZARD WRAPPERS
// ============================================================================

function runRowSetup() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    "Step 2 — Header & Data Rows",
    "You will be prompted twice:\n\n  1. Header Row — the row containing your column names (e.g. Company, Seller/Donor)\n  2. Data Start Row — the first row where actual project records begin\n\nLeave either blank to use auto-detection.",
    ui.ButtonSet.OK,
  );
  setHeaderRow();
  setDataStartRow();
}

function manageNotificationEmails() {
  configureNotificationEmails();
}

function setupSheetWithSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const selectedSheets = getProjectSheets();
  const sheetName = selectedSheets[0] || ss.getActiveSheet().getName();
  const sheet = ss.getSheetByName(sheetName) || ss.getActiveSheet();

  selectedSheets.forEach((name) => setupSheet(name));
  ensureActivityLogSheet();
  ensureCompletedSheet();
  ensureDashboardSheet();

  const cfg = getConfig();
  const hRow = getHeaderRow(sheet);
  const dRow = getDataStartRow(sheet);
  const emails = getNotificationEmails();

  SpreadsheetApp.getUi().alert(
    "Setup Applied",
    `Automation tabs: ${selectedSheets.length > 0 ? selectedSheets.join(", ") : sheet.getName()}\nHeader row: ${hRow}\nData start row: ${dRow}\nNotification emails: ${emails.length > 0 ? emails.join(", ") : "(none)"}\nDaily automation: ${formatTimeLabel(cfg.AUTOMATION_HOUR, cfg.AUTOMATION_MINUTE)}`,
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

function quickSetup() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    "Quick Setup",
    "This guided setup will ask you which tabs should be automated, where your header/data rows are, which CC emails to use, what time daily automation should run, and who should receive reminders.",
    ui.ButtonSet.OK,
  );

  selectWorkingSheets();
  runRowSetup();
  configureNotificationEmails();
  configureDailyAutomationTime();
  configureEmailRecipientMode();
  setupSheetWithSummary();
  showValidationReport();
}

// ============================================================================
// ON OPEN - CUSTOM MENU
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("CAR Monitoring")
    .addItem("Quick Setup", "quickSetup")
    .addItem("Run Validation", "showValidationReport")
    .addSeparator()
    .addItem("Update All Deadlines", "updateAllDeadlines")
    .addItem("Generate Reminders", "generateAllReminders")
    .addItem("Send Reminder Emails", "sendReminderEmails")
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("Automation Settings")
        .addItem("Run Daily Check", "runDailyCheck")
        .addItem("Set Automated Tabs", "selectWorkingSheets")
        .addItem("Set Header & Data Rows", "runRowSetup")
        .addItem("Set Daily Run Time", "configureDailyAutomationTime")
        .addItem("Create / Refresh Trigger", "createDailyTrigger")
        .addItem("Remove Triggers", "removeTriggers"),
    )
    .addSubMenu(
      ui
        .createMenu("Email Settings")
        .addItem("Set Recipient Mode", "configureEmailRecipientMode")
        .addItem("Set CC Emails", "configureNotificationEmails")
        .addItem("View CC Emails", "viewNotificationEmails")
        .addItem("Add CC Email", "addNotificationEmail")
        .addItem("Remove CC Email", "removeNotificationEmail")
        .addItem("Clear CC Emails", "clearNotificationEmails"),
    )
    .addSubMenu(
      ui
        .createMenu("Settings")
        .addItem("View Current Settings", "showSettings")
        .addItem("Reset to Defaults", "resetSettingsToDefaults")
    )
    .addSubMenu(
      ui
        .createMenu("Advanced Settings")
        .addItem("Setup All Service Sheets", "setupAllSheets")
        .addItem("Validate Headers", "validateAllHeaders")
        .addItem("View Activity Log", "viewActivityLog")
        .addItem("Clear Old Logs", "clearOldLogs")
        .addItem("Diagnostics", "showDiagnostics"),
    )
    .addToUi();
}

// ============================================================================
// SHEET SETUP & VALIDATION
// ============================================================================

function detectHeaderRow(sheet) {
  const configured = getConfig().HEADER_ROW;
  if (configured && configured >= 1) return configured;

  const maxScan = Math.min(5, sheet.getLastRow());
  if (maxScan < 1) return 1;

  const scanData = sheet
    .getRange(1, 1, maxScan, sheet.getLastColumn())
    .getValues();
  let bestRow = 1;
  let bestCount = 0;

  for (let i = 0; i < scanData.length; i++) {
    const nonEmpty = scanData[i].filter((v) => v !== "" && v !== null).length;
    if (nonEmpty > bestCount) {
      bestCount = nonEmpty;
      bestRow = i + 1;
    }
  }

  return bestRow;
}

function getDataStartRow(sheet) {
  const configured = getConfig().DATA_START_ROW;
  if (configured && configured >= 1) return configured;
  return getHeaderRow(sheet) + 1;
}

function setupSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  const name = sheetName || ss.getActiveSheet().getName();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.getActiveSheet();
  }

  const lastCol = sheet.getLastColumn();
  const hasExistingHeaders = lastCol > 0 && sheet.getLastRow() > 0;

  if (!hasExistingHeaders) {
    const defaultHeaders = [
      cfg.HEADERS.NO,
      cfg.HEADERS.DATE,
      cfg.HEADERS.CLIENT_NAME,
      cfg.HEADERS.SELLER_DONOR,
      cfg.HEADERS.BUYER_DONEE,
      cfg.HEADERS.SERVICE_TYPE,
      cfg.HEADERS.NOTARY_DATE,
      cfg.HEADERS.DST_DUE_DATE,
      cfg.HEADERS.CGT_DOD_DUE_DATE,
      cfg.HEADERS.DST_REMAINING,
      cfg.HEADERS.CGT_REMAINING,
      cfg.HEADERS.STATUS,
      cfg.HEADERS.REMINDER,
      cfg.HEADERS.REMARKS,
      cfg.HEADERS.STAFF_EMAIL,
      cfg.HEADERS.CLIENT_EMAIL,
      cfg.HEADERS.REVISIT_DATE,
      cfg.HEADERS.REVISIT_STATUS,
      cfg.HEADERS.REVISIT_NOTES,
    ];
    sheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
    sheet
      .getRange(1, 1, 1, defaultHeaders.length)
      .setFontWeight("bold")
      .setBackground("#4285f4")
      .setFontColor("white");
    sheet.getRange(1, 1, 1, sheet.getMaxColumns()).breakApart();
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, defaultHeaders.length);

    applyDataValidations(sheet, 1);
    applyConditionalFormatting(sheet, 1);

    logActivity(
      "SYSTEM",
      "Sheet Setup",
      `Initialized ${sheet.getName()} with default headers`,
    );
    SpreadsheetApp.getActive().toast(
      `Sheet "${sheet.getName()}" setup complete`,
      "Success",
    );
    return sheet;
  }

  const headerRow = detectHeaderRow(sheet);
  const numCols = sheet.getLastColumn();

  sheet.getRange(headerRow, 1, 1, sheet.getMaxColumns()).breakApart();
  sheet.setFrozenRows(headerRow);
  sheet.autoResizeColumns(1, numCols);

  applyDataValidations(sheet, headerRow);
  applyConditionalFormatting(sheet, headerRow);

  logActivity(
    "SYSTEM",
    "Sheet Setup",
    `Prepared existing sheet ${sheet.getName()} (header row ${headerRow})`,
  );
  SpreadsheetApp.getActive().toast(
    `Sheet "${sheet.getName()}" setup complete`,
    "Success",
  );
  return sheet;
}

function setupAllSheets() {
  getProjectSheets().forEach((name) => setupSheet(name));
  ensureActivityLogSheet();
  ensureCompletedSheet();
  ensureDashboardSheet();

  SpreadsheetApp.getActive().toast(
    "All service sheets initialized",
    "Complete",
  );
}

function getStatusValuesFromSheet(sheet) {
  const cfg = getConfig();
  const statusCol = getColumnIndex(sheet, cfg.HEADERS.STATUS);
  if (!statusCol) return [];

  const lastRow = sheet.getLastRow();
  const dataStart = getDataStartRow(sheet);
  if (lastRow < dataStart) return [];

  const values = sheet
    .getRange(dataStart, statusCol, lastRow - dataStart + 1)
    .getValues()
    .flat()
    .map((v) => v.toString().trim())
    .filter((v) => v !== "");

  const unique = Array.from(new Set(values));
  return unique;
}

function applyDataValidations(sheet, headerRow) {
  const hRow = headerRow || getHeaderRow(sheet);
  const dataStart = headerRow ? hRow + 1 : getDataStartRow(sheet);
  const lastRow = Math.max(sheet.getLastRow(), hRow + 100);
  const cfg = getConfig();

  const serviceTypeCol = getColumnIndex(sheet, cfg.HEADERS.SERVICE_TYPE);
  if (serviceTypeCol) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(cfg.SERVICE_TYPES, true)
      .setAllowInvalid(false)
      .build();
    sheet
      .getRange(dataStart, serviceTypeCol, lastRow - hRow)
      .setDataValidation(rule);
  }

  const statusCol = getColumnIndex(sheet, cfg.HEADERS.STATUS);
  if (statusCol) {
    const sheetStatusValues = getStatusValuesFromSheet(sheet);
    const statusList =
      sheetStatusValues.length >= 1
        ? sheetStatusValues
        : [cfg.STATUS.ONGOING, cfg.STATUS.COMPLETED];
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(statusList, true)
      .setAllowInvalid(true)
      .build();
    sheet
      .getRange(dataStart, statusCol, lastRow - hRow)
      .setDataValidation(rule);
  }

  const revisitStatusCol = getColumnIndex(sheet, cfg.HEADERS.REVISIT_STATUS);
  if (revisitStatusCol) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Open", "Pending", "Completed"], true)
      .setAllowInvalid(true)
      .build();
    sheet
      .getRange(dataStart, revisitStatusCol, lastRow - hRow)
      .setDataValidation(rule);
  }
}

function applyConditionalFormatting(sheet, headerRow) {
  const hRow = headerRow || getHeaderRow(sheet);
  const dataStart = headerRow ? hRow + 1 : getDataStartRow(sheet);
  const cfg = getConfig();
  const groupedRemainingCols = getDeadlineGroups(sheet)
    .map((group) => group.remainingCol)
    .filter(Boolean);
  const remainingCols = Array.from(
    new Set(
      groupedRemainingCols.concat([
        getColumnIndex(sheet, cfg.HEADERS.DST_REMAINING),
        getColumnIndex(sheet, cfg.HEADERS.CGT_REMAINING),
      ]),
    ),
  ).filter(Boolean);

  const lastRow = Math.max(sheet.getLastRow(), hRow + 100);

  remainingCols.forEach((col) => {
    if (!col) return;
    const range = sheet.getRange(dataStart, col, lastRow - hRow);

    const overdueRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground("#ffcccc")
      .setFontColor("#cc0000")
      .setRanges([range])
      .build();

    const warningRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0, 7)
      .setBackground("#ffeb9c")
      .setFontColor("#9c6500")
      .setRanges([range])
      .build();

    const goodRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(7)
      .setBackground("#c6efce")
      .setFontColor("#006100")
      .setRanges([range])
      .build();

    const rules = sheet.getConditionalFormatRules();
    rules.push(overdueRule, warningRule, goodRule);
    sheet.setConditionalFormatRules(rules);
  });
}

function getColumnIndex(sheet, headerName) {
  return findHeaderIndexByName(getHeaders(sheet), headerName);
}

function validateAllHeaders() {
  const results = validateSystem();
  const issues = results.issues.concat(results.warnings);
  if (issues.length > 0) {
    SpreadsheetApp.getUi().alert(
      "Header Validation Issues:\n\n" + issues.join("\n"),
    );
  } else {
    SpreadsheetApp.getActive().toast(
      "All headers validated successfully",
      "Success",
    );
  }
  return issues;
}

function getOpenRevisitInfo(sheet, row) {
  const cfg = getConfig();
  const revisitDateCol = getColumnIndex(sheet, cfg.HEADERS.REVISIT_DATE);
  const revisitStatusCol = getColumnIndex(sheet, cfg.HEADERS.REVISIT_STATUS);
  const revisitNotesCol = getColumnIndex(sheet, cfg.HEADERS.REVISIT_NOTES);

  if (!revisitDateCol) return null;

  const revisitDate = sheet.getRange(row, revisitDateCol).getValue();
  if (!(revisitDate instanceof Date)) return null;

  const status = revisitStatusCol
    ? sheet.getRange(row, revisitStatusCol).getValue()
    : "";
  const normalizedStatus = normalizeHeaderName(status);
  if (normalizedStatus && !REVISIT_OPEN_STATUSES.includes(normalizedStatus)) {
    return null;
  }

  return {
    date: revisitDate,
    status: status,
    notes: revisitNotesCol ? sheet.getRange(row, revisitNotesCol).getValue() : "",
    remainingDays: calculateRemainingDays(revisitDate),
  };
}

function buildRevisitReminder(info) {
  if (!info) return "";
  return buildReminderText("Revisit", info.remainingDays);
}

function validateSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  const issues = [];
  const warnings = [];
  const notes = [];
  const selectedSheets = getProjectSheets();

  notes.push(`Spreadsheet: ${ss.getName()}`);
  notes.push(
    `Daily automation time: ${formatTimeLabel(
      cfg.AUTOMATION_HOUR,
      cfg.AUTOMATION_MINUTE,
    )}`,
  );
  notes.push(`Email recipient mode: ${cfg.EMAIL_RECIPIENT_MODE}`);

  if (selectedSheets.length === 0) {
    issues.push("No automation tabs are configured or detectable.");
  }

  selectedSheets.forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      issues.push(`Selected automation tab is missing: ${sheetName}`);
      return;
    }

    const headers = getHeaders(sheet);
    const headerRow = getHeaderRow(sheet);
    const dataStartRow = getDataStartRow(sheet);
    if (headerRow >= dataStartRow) {
      issues.push(
        `${sheetName}: Data start row (${dataStartRow}) must be below header row (${headerRow}).`,
      );
    }
    const required = [
      cfg.HEADERS.CLIENT_NAME,
      cfg.HEADERS.NOTARY_DATE,
      cfg.HEADERS.STATUS,
      cfg.HEADERS.REMARKS,
      cfg.HEADERS.STAFF_EMAIL,
    ];

    required.forEach((req) => {
      if (!findHeaderIndexByName(headers, req)) {
        issues.push(`${sheetName}: Missing required header "${req}"`);
      }
    });

    if (!findHeaderIndexByName(headers, cfg.HEADERS.REVISIT_DATE)) {
      warnings.push(
        `${sheetName}: Revisit Date column not found. Revisit automation will be skipped on this tab.`,
      );
    }
  });

  const quota = MailApp.getRemainingDailyQuota();
  if (quota <= 0) issues.push("Mail quota is exhausted.");
  else if (quota < 10) warnings.push(`Mail quota is low (${quota} remaining).`);

  const triggers = ScriptApp.getProjectTriggers().filter(
    (t) => t.getHandlerFunction() === "runDailyCheck",
  );
  if (triggers.length === 0) {
    warnings.push("Daily automation trigger is not installed.");
  } else {
    notes.push(`Daily trigger installed: ${triggers.length}`);
  }

  return {
    passed: issues.length === 0,
    issues: issues,
    warnings: warnings,
    notes: notes,
  };
}

// ============================================================================
// DEADLINE CALCULATION ENGINE
// ============================================================================

function calculateDSTDueDate(notaryDate) {
  if (!notaryDate || !(notaryDate instanceof Date)) return null;
  const rule = getConfig().DST_RULE;
  const date = new Date(notaryDate);
  date.setMonth(date.getMonth() + rule.monthOffset);
  date.setDate(rule.day);
  return date;
}

function calculateCGTDueDate(notaryDate) {
  if (!notaryDate || !(notaryDate instanceof Date)) return null;
  const date = new Date(notaryDate);
  date.setDate(date.getDate() + getConfig().CGT_RULE.days);
  return date;
}

function calculateRemainingDays(dueDate) {
  if (!dueDate || !(dueDate instanceof Date)) return null;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const due = new Date(dueDate);
  due.setHours(0, 0, 0, 0);

  const diffTime = due - today;
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

  return diffDays;
}

function updateRowDeadlines(sheet, row) {
  const cfg = getConfig();
  const groups = getDeadlineGroups(sheet);

  if (groups.length > 0) {
    let updated = false;

    groups.forEach((group) => {
      const rule = parseDueDateRule(group.label);
      const baseDate = getBaseDateForRule(sheet, row, rule);
      let dueDate = sheet.getRange(row, group.dueDateCol).getValue();

      if (rule && baseDate instanceof Date) {
        dueDate = calculateDueDateFromRule(baseDate, rule);
        sheet.getRange(row, group.dueDateCol).setValue(dueDate);
        updated = true;
      }

      if (group.remainingCol && dueDate instanceof Date) {
        sheet
          .getRange(row, group.remainingCol)
          .setValue(calculateRemainingDays(dueDate));
        updated = true;
      }
    });

    if (updated) return true;
  }

  const notaryCol = getColumnIndex(sheet, cfg.HEADERS.NOTARY_DATE);
  const dstDueCol = getColumnIndex(sheet, cfg.HEADERS.DST_DUE_DATE);
  const cgtDueCol = getColumnIndex(sheet, cfg.HEADERS.CGT_DOD_DUE_DATE);
  const dstRemainCol = getColumnIndex(sheet, cfg.HEADERS.DST_REMAINING);
  const cgtRemainCol = getColumnIndex(sheet, cfg.HEADERS.CGT_REMAINING);

  if (!notaryCol) return false;

  const notaryDate = sheet.getRange(row, notaryCol).getValue();
  if (!notaryDate || !(notaryDate instanceof Date)) return false;

  const dstDue = calculateDSTDueDate(notaryDate);
  const cgtDue = calculateCGTDueDate(notaryDate);

  if (dstDueCol) sheet.getRange(row, dstDueCol).setValue(dstDue);
  if (cgtDueCol) sheet.getRange(row, cgtDueCol).setValue(cgtDue);

  if (dstRemainCol) {
    sheet.getRange(row, dstRemainCol).setValue(calculateRemainingDays(dstDue));
  }

  if (cgtRemainCol) {
    sheet.getRange(row, cgtRemainCol).setValue(calculateRemainingDays(cgtDue));
  }

  return true;
}

function updateAllDeadlines() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getProjectSheets();

  let updatedCount = 0;

  sheets.forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    const dataStart = getDataStartRow(sheet);
    if (lastRow < dataStart) return;

    for (let row = dataStart; row <= lastRow; row++) {
      if (updateRowDeadlines(sheet, row)) {
        updatedCount++;
      }
    }
  });

  logActivity(
    "SYSTEM",
    "Update Deadlines",
    `Updated ${updatedCount} rows across all sheets`,
  );
  SpreadsheetApp.getActive().toast(
    `Updated ${updatedCount} records`,
    "Deadlines Updated",
  );
}

// ============================================================================
// STATUS & REMINDER SYSTEM
// ============================================================================

function generateReminderMessage(dstRemaining, cgtRemaining, clientName) {
  const reminders = [];

  if (dstRemaining !== null) {
    if (dstRemaining < 0) {
      reminders.push(`DST payment OVERDUE by ${Math.abs(dstRemaining)} days`);
    } else if (dstRemaining <= 7) {
      reminders.push(`DST payment due in ${dstRemaining} days`);
    }
  }

  if (cgtRemaining !== null) {
    if (cgtRemaining < 0) {
      reminders.push(
        `CGT/DOD filing OVERDUE by ${Math.abs(cgtRemaining)} days`,
      );
    } else if (cgtRemaining <= 7) {
      reminders.push(`CGT/DOD filing due in ${cgtRemaining} days`);
    }
  }

  if (reminders.length === 0) return "";

  return "⚠️ " + reminders.join("; ");
}

function updateRowReminder(sheet, row) {
  const cfg = getConfig();
  const statusCol = getColumnIndex(sheet, cfg.HEADERS.STATUS);
  const groups = getDeadlineGroups(sheet);
  const revisitInfo = getOpenRevisitInfo(sheet, row);
  const revisitReminder = buildRevisitReminder(revisitInfo);

  const currentStatus = statusCol
    ? sheet.getRange(row, statusCol).getValue()
    : "";

  if (groups.length > 0) {
    let hasReminder = false;

    groups.forEach((group) => {
      if (!group.reminderCol) return;

      if (currentStatus === cfg.STATUS.COMPLETED) {
        sheet.getRange(row, group.reminderCol).setValue("");
        return;
      }

      const remainingValue = group.remainingCol
        ? sheet.getRange(row, group.remainingCol).getValue()
        : null;
      const reminders = [];
      const primaryReminder = buildReminderText(
        group.label,
        Number(remainingValue),
      );
      if (primaryReminder) reminders.push(primaryReminder);
      if (revisitReminder) reminders.push(revisitReminder);
      const reminder = reminders.join(" | ");
      sheet.getRange(row, group.reminderCol).setValue(reminder);

      if (reminder) hasReminder = true;
    });

    return hasReminder;
  }

  const clientCol = getColumnIndex(sheet, cfg.HEADERS.CLIENT_NAME);
  const dstRemainCol = getColumnIndex(sheet, cfg.HEADERS.DST_REMAINING);
  const cgtRemainCol = getColumnIndex(sheet, cfg.HEADERS.CGT_REMAINING);
  const reminderCol = getColumnIndex(sheet, cfg.HEADERS.REMINDER);

  if (!reminderCol) return false;

  const clientName = clientCol ? sheet.getRange(row, clientCol).getValue() : "";
  const dstRemaining = dstRemainCol
    ? sheet.getRange(row, dstRemainCol).getValue()
    : null;
  const cgtRemaining = cgtRemainCol
    ? sheet.getRange(row, cgtRemainCol).getValue()
    : null;

  if (currentStatus === cfg.STATUS.COMPLETED) {
    sheet.getRange(row, reminderCol).setValue("");
    return true;
  }

  const reminder = generateReminderMessage(
    dstRemaining,
    cgtRemaining,
    clientName,
  );
  const finalReminder = [reminder, revisitReminder].filter(Boolean).join(" | ");
  sheet.getRange(row, reminderCol).setValue(finalReminder);

  return finalReminder !== "";
}

function generateAllReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getProjectSheets();

  let reminderCount = 0;

  sheets.forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    const dataStart = getDataStartRow(sheet);
    if (lastRow < dataStart) return;

    for (let row = dataStart; row <= lastRow; row++) {
      if (updateRowReminder(sheet, row)) {
        reminderCount++;
      }
    }
  });

  logActivity(
    "SYSTEM",
    "Generate Reminders",
    `Generated ${reminderCount} reminders`,
  );
  SpreadsheetApp.getActive().toast(
    `${reminderCount} reminders generated`,
    "Reminders",
  );
}

function autoUpdateStatus(sheet, row) {
  const cfg = getConfig();
  const statusCol = getColumnIndex(sheet, cfg.HEADERS.STATUS);
  const remarksCol = getColumnIndex(sheet, cfg.HEADERS.REMARKS);

  if (!statusCol || !remarksCol) return;

  const remarks = sheet.getRange(row, remarksCol).getValue();
  const currentStatus = sheet.getRange(row, statusCol).getValue();

  if (!remarks) return;

  const completionKeywords = [
    "completed",
    "done",
    "finished",
    "released",
    "ecar issued",
  ];
  const lowerRemarks = remarks.toString().toLowerCase();

  const isComplete = completionKeywords.some((kw) => lowerRemarks.includes(kw));

  if (isComplete && currentStatus !== cfg.STATUS.COMPLETED) {
    sheet.getRange(row, statusCol).setValue(cfg.STATUS.COMPLETED);

    const clientCol = getColumnIndex(sheet, cfg.HEADERS.CLIENT_NAME);
    const clientName = clientCol
      ? sheet.getRange(row, clientCol).getValue()
      : "Unknown";
    logActivity(
      clientName,
      "Status Auto-Update",
      "Marked as Completed based on remarks",
    );
  }
}

// ============================================================================
// ACTIVITY LOGGING
// ============================================================================

function ensureCompletedSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  let sheet = ss.getSheetByName(cfg.SHEETS.COMPLETED);

  if (!sheet) {
    sheet = ss.insertSheet(cfg.SHEETS.COMPLETED);
    const headers = [
      cfg.HEADERS.NO,
      cfg.HEADERS.DATE,
      cfg.HEADERS.CLIENT_NAME,
      cfg.HEADERS.SELLER_DONOR,
      cfg.HEADERS.BUYER_DONEE,
      cfg.HEADERS.SERVICE_TYPE,
      cfg.HEADERS.NOTARY_DATE,
      cfg.HEADERS.DST_DUE_DATE,
      cfg.HEADERS.CGT_DOD_DUE_DATE,
      cfg.HEADERS.STATUS,
      cfg.HEADERS.REMARKS,
      "Archived Date",
    ];
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#34a853");
    headerRange.setFontColor("white");
    sheet.setFrozenRows(1);
    Logger.log("Created Completed Projects sheet");
  }

  return sheet;
}

function ensureDashboardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  let sheet = ss.getSheetByName(cfg.SHEETS.DASHBOARD);

  if (!sheet) {
    sheet = ss.insertSheet(cfg.SHEETS.DASHBOARD);
    Logger.log("Created Dashboard sheet");
  }

  updateDashboard(ss, sheet);
  return sheet;
}

function updateDashboard(ss, dashSheet) {
  const cfg = getConfig();
  const sheet =
    dashSheet ||
    (ss || SpreadsheetApp.getActiveSpreadsheet()).getSheetByName(
      cfg.SHEETS.DASHBOARD,
    );
  if (!sheet) return;

  sheet.clearContents();

  const title = [["CAR Special Projects — Dashboard"]];
  sheet.getRange(1, 1).setValues(title).setFontWeight("bold").setFontSize(14);

  const generated = [["Generated:", new Date().toLocaleString()]];
  sheet.getRange(2, 1, 1, 2).setValues(generated);

  const summaryHeaders = [
    ["Sheet", "Total", "On Going", "Completed", "Overdue DST", "Overdue CGT"],
  ];
  sheet
    .getRange(4, 1, 1, 6)
    .setValues(summaryHeaders)
    .setFontWeight("bold")
    .setBackground("#4285f4")
    .setFontColor("white");

  const sourceSheets = ss ? ss : SpreadsheetApp.getActiveSpreadsheet();
  let summaryRow = 5;

  getProjectSheets().forEach((sheetName) => {
    const src = sourceSheets.getSheetByName(sheetName);
    const dataStart = src ? getDataStartRow(src) : 2;
    if (!src || src.getLastRow() < dataStart) {
      sheet
        .getRange(summaryRow, 1, 1, 6)
        .setValues([[sheetName, 0, 0, 0, 0, 0]]);
      summaryRow++;
      return;
    }

    const headers = getHeaders(src);
    const data = src
      .getRange(dataStart, 1, src.getLastRow() - dataStart + 1, headers.length)
      .getValues();
    const colMap = {};
    headers.forEach((h, i) => (colMap[h] = i));

    let total = 0,
      ongoing = 0,
      completed = 0,
      overdueDst = 0,
      overdueCgt = 0;

    data.forEach((row) => {
      const client =
        colMap[cfg.HEADERS.CLIENT_NAME] !== undefined
          ? row[colMap[cfg.HEADERS.CLIENT_NAME]]
          : "";
      if (!client) return;
      total++;
      const status =
        colMap[cfg.HEADERS.STATUS] !== undefined
          ? row[colMap[cfg.HEADERS.STATUS]]
          : "";
      if (status === cfg.STATUS.COMPLETED) completed++;
      else ongoing++;
      const dst =
        colMap[cfg.HEADERS.DST_REMAINING] !== undefined
          ? row[colMap[cfg.HEADERS.DST_REMAINING]]
          : null;
      const cgt =
        colMap[cfg.HEADERS.CGT_REMAINING] !== undefined
          ? row[colMap[cfg.HEADERS.CGT_REMAINING]]
          : null;
      if (typeof dst === "number" && dst < 0) overdueDst++;
      if (typeof cgt === "number" && cgt < 0) overdueCgt++;
    });

    sheet
      .getRange(summaryRow, 1, 1, 6)
      .setValues([
        [sheetName, total, ongoing, completed, overdueDst, overdueCgt],
      ]);
    summaryRow++;
  });

  sheet.autoResizeColumns(1, 6);
}

function ensureActivityLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  let sheet = ss.getSheetByName(cfg.SHEETS.ACTIVITY_LOG);

  if (!sheet) {
    sheet = ss.insertSheet(cfg.SHEETS.ACTIVITY_LOG);
    const headers = ["Timestamp", "Client Name", "Action", "Detail", "User"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function logActivity(clientName, action, detail) {
  try {
    const sheet = ensureActivityLogSheet();
    const user = Session.getEffectiveUser().getEmail();
    const timestamp = new Date();

    sheet.appendRow([timestamp, clientName, action, detail, user]);
  } catch (e) {
    Logger.log("Failed to log activity: " + e.message);
  }
}

function viewActivityLog() {
  const sheet = ensureActivityLogSheet();
  SpreadsheetApp.setActiveSheet(sheet);
}

function clearOldLogs(daysToKeep = 90) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Clear Old Logs",
    `Delete activity logs older than ${daysToKeep} days?`,
    ui.ButtonSet.YES_NO,
  );

  if (response !== ui.Button.YES) return;

  const sheet = ensureActivityLogSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);

  const timestamps = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let deletedCount = 0;

  for (let i = timestamps.length - 1; i >= 0; i--) {
    const rowDate = timestamps[i][0];
    if (rowDate instanceof Date && rowDate < cutoffDate) {
      sheet.deleteRow(i + 2);
      deletedCount++;
    }
  }

  SpreadsheetApp.getActive().toast(
    `Deleted ${deletedCount} old log entries`,
    "Logs Cleared",
  );
}

function getProjectHistory(clientName) {
  const sheet = ensureActivityLogSheet();
  const data = sheet.getDataRange().getValues();
  const history = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === clientName) {
      history.push({
        timestamp: data[i][0],
        action: data[i][2],
        detail: data[i][3],
        user: data[i][4],
      });
    }
  }

  return history.sort((a, b) => b.timestamp - a.timestamp);
}

// ============================================================================
// EMAIL NOTIFICATIONS
// ============================================================================

function sendReminderEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  const sheets = getProjectSheets();

  let emailCount = 0;
  const quota = MailApp.getRemainingDailyQuota();

  if (quota < 10) {
    SpreadsheetApp.getUi().alert(
      "Email quota too low (" + quota + "). Try again tomorrow.",
    );
    return;
  }

  sheets.forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    const dataStart = getDataStartRow(sheet);
    if (lastRow < dataStart) return;

    const headers = getHeaders(sheet);
    const data = sheet
      .getRange(dataStart, 1, lastRow - dataStart + 1, headers.length)
      .getValues();
    const deadlineGroups = getDeadlineGroups(sheet);

    const genericReminderIndex = findHeaderIndexByName(
      headers,
      cfg.HEADERS.REMINDER,
    );
    const serviceTypeIndex = findHeaderIndexByName(
      headers,
      cfg.HEADERS.SERVICE_TYPE,
    );
    const clientNameIndex = findHeaderIndexByName(
      headers,
      cfg.HEADERS.CLIENT_NAME,
    );
    const staffEmailIndex = findHeaderIndexByName(
      headers,
      cfg.HEADERS.STAFF_EMAIL,
    );
    const clientEmailIndex = findHeaderIndexByName(
      headers,
      cfg.HEADERS.CLIENT_EMAIL,
    );
    const remarksIndex = findHeaderIndexByName(headers, cfg.HEADERS.REMARKS);
    const revisitDateIndex = findHeaderIndexByName(
      headers,
      cfg.HEADERS.REVISIT_DATE,
    );
    const revisitStatusIndex = findHeaderIndexByName(
      headers,
      cfg.HEADERS.REVISIT_STATUS,
    );
    const revisitNotesIndex = findHeaderIndexByName(
      headers,
      cfg.HEADERS.REVISIT_NOTES,
    );

    data.forEach((row, idx) => {
      const actualRow = idx + dataStart;
      const reminders = [];
      const deadlineSummaries = [];

      deadlineGroups.forEach((group) => {
        if (group.reminderCol) {
          const reminderValue = row[group.reminderCol - 1];
          if (reminderValue) reminders.push(reminderValue.toString());
        }

        if (group.remainingCol) {
          const remainingValue = row[group.remainingCol - 1];
          const remainingDays =
            remainingValue === "" || remainingValue === null
              ? null
              : Number(remainingValue);

          if (remainingDays !== null && !isNaN(remainingDays)) {
            deadlineSummaries.push({
              label: group.label,
              remainingDays: remainingDays,
            });
          }
        }
      });

      if (
        genericReminderIndex &&
        (!deadlineGroups.length || !reminders.length) &&
        row[genericReminderIndex - 1]
      ) {
        reminders.push(row[genericReminderIndex - 1].toString());
      }

      if (!reminders.length) return;

      const clientName = clientNameIndex ? row[clientNameIndex - 1] : "";
      const clientEmail = clientEmailIndex ? row[clientEmailIndex - 1] : "";
      const staffEmail = staffEmailIndex ? row[staffEmailIndex - 1] : "";
      const serviceType = serviceTypeIndex ? row[serviceTypeIndex - 1] : "";
      const remarks = remarksIndex ? row[remarksIndex - 1] : "";
      const revisitInfo =
        revisitDateIndex && row[revisitDateIndex - 1] instanceof Date
          ? {
              date: row[revisitDateIndex - 1],
              status: revisitStatusIndex ? row[revisitStatusIndex - 1] : "",
              notes: revisitNotesIndex ? row[revisitNotesIndex - 1] : "",
              remainingDays: calculateRemainingDays(row[revisitDateIndex - 1]),
            }
          : null;

      const recipients = [];
      const mode = getEmailRecipientMode();
      if (staffEmail && staffEmail.toString().includes("@")) {
        recipients.push(staffEmail);
      }
      if (
        mode === EMAIL_RECIPIENT_MODES.STAFF_AND_CLIENT &&
        clientEmail &&
        clientEmail.toString().includes("@") &&
        clientEmail.toString() !== staffEmail.toString()
      ) {
        recipients.push(clientEmail);
      }

      if (recipients.length === 0) return;

      try {
        sendReminderEmail({
          recipients: recipients,
          clientName: clientName,
          serviceType: serviceType,
          reminders: reminders,
          deadlines: deadlineSummaries,
          sheetName: sheetName,
          rowNumber: actualRow,
          remarks: remarks,
          revisit: revisitInfo,
        });

        emailCount++;
        logActivity(
          clientName,
          "Reminder Email Sent",
          `To: ${recipients.join(", ")}`,
        );
      } catch (e) {
        logActivity(clientName, "Email Failed", e.message);
      }
    });
  });

  SpreadsheetApp.getActive().toast(
    `Sent ${emailCount} reminder emails`,
    "Emails Sent",
  );
}

function sendReminderEmail(data) {
  const subject = `Reminder: ${data.clientName} - ${data.serviceType || "Project"}`;

  let body = `Dear Team,\n\n`;
  body += `This is an automated reminder regarding the following project:\n\n`;
  body += `Client: ${data.clientName}\n`;
  body += `Service: ${data.serviceType || "N/A"}\n`;
  body += `Sheet: ${data.sheetName}\n\n`;
  if (data.rowNumber) {
    body += `Row: ${data.rowNumber}\n\n`;
  }
  body += `ALERT:\n- ${data.reminders.join("\n- ")}\n\n`;

  if (data.deadlines && data.deadlines.length) {
    body += `Deadlines:\n`;
    data.deadlines.forEach((deadline) => {
      const message =
        deadline.remainingDays < 0
          ? `OVERDUE by ${Math.abs(deadline.remainingDays)} days`
          : `${deadline.remainingDays} days remaining`;
      body += `- ${deadline.label}: ${message}\n`;
    });
  }

  if (data.revisit && data.revisit.remainingDays !== null) {
    const revisitMessage =
      data.revisit.remainingDays < 0
        ? `OVERDUE by ${Math.abs(data.revisit.remainingDays)} days`
        : `${data.revisit.remainingDays} days remaining`;
    body += `\nRevisit:\n- Date: ${formatDate(data.revisit.date)}\n- Status: ${data.revisit.status || "Open"}\n- Timing: ${revisitMessage}\n`;
    if (data.revisit.notes) {
      body += `- Notes: ${data.revisit.notes}\n`;
    }
  }

  if (data.remarks) {
    body += `\nRemarks:\n${data.remarks}\n`;
  }

  body += `\nPlease take appropriate action to ensure compliance.\n\n`;
  body += `---\n`;
  body += `CAR Special Projects Monitoring System\n`;
  body += `This is an automated message. Please do not reply directly to this email.`;

  const notificationEmails = getNotificationEmails();
  const toSet = new Set(data.recipients.map((e) => e.toString().toLowerCase()));
  const ccList = notificationEmails.filter((e) => !toSet.has(e.toLowerCase()));

  const mailOptions = {
    to: data.recipients.join(","),
    subject: subject,
    body: body,
    name: "CAR Monitoring System",
  };
  if (ccList.length > 0) mailOptions.cc = ccList.join(",");

  MailApp.sendEmail(mailOptions);
}

// ============================================================================
// AUTOMATION TRIGGERS
// ============================================================================

function onEdit(e) {
  if (!e) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const cfg = getConfig();

  if (!getProjectSheets().includes(sheetName)) return;

  const row = e.range.getRow();
  if (row < getDataStartRow(sheet)) return;

  const notaryCol = getColumnIndex(sheet, cfg.HEADERS.NOTARY_DATE);
  const remarksCol = getColumnIndex(sheet, cfg.HEADERS.REMARKS);
  const statusCol = getColumnIndex(sheet, cfg.HEADERS.STATUS);
  const revisitDateCol = getColumnIndex(sheet, cfg.HEADERS.REVISIT_DATE);
  const revisitStatusCol = getColumnIndex(sheet, cfg.HEADERS.REVISIT_STATUS);

  if (notaryCol && e.range.getColumn() === notaryCol) {
    updateRowDeadlines(sheet, row);
    updateRowReminder(sheet, row);

    const clientCol = getColumnIndex(sheet, cfg.HEADERS.CLIENT_NAME);
    const clientName = clientCol
      ? sheet.getRange(row, clientCol).getValue()
      : "Unknown";
    logActivity(
      clientName,
      "Notary Date Updated",
      `Row ${row} in ${sheetName}`,
    );
  }

  if (remarksCol && e.range.getColumn() === remarksCol) {
    autoUpdateStatus(sheet, row);
  }

  if (statusCol && e.range.getColumn() === statusCol) {
    updateRowReminder(sheet, row);
  }

  if (
    (revisitDateCol && e.range.getColumn() === revisitDateCol) ||
    (revisitStatusCol && e.range.getColumn() === revisitStatusCol)
  ) {
    updateRowReminder(sheet, row);
  }
}

function runDailyCheck() {
  const startTime = new Date();
  Logger.log("Starting daily check at " + startTime);

  updateAllDeadlines();
  generateAllReminders();
  sendReminderEmails();
  updateDashboard();

  const endTime = new Date();
  const duration = (endTime - startTime) / 1000;

  logActivity("SYSTEM", "Daily Check Complete", `Duration: ${duration}s`);
  Logger.log("Daily check completed in " + duration + " seconds");
}

function syncDailyTrigger(showToast) {
  const triggers = ScriptApp.getProjectTriggers().filter(
    (t) => t.getHandlerFunction() === "runDailyCheck",
  );
  triggers.forEach((trigger) => ScriptApp.deleteTrigger(trigger));

  const time = getConfiguredAutomationTime();
  let builder = ScriptApp.newTrigger("runDailyCheck")
    .timeBased()
    .everyDays(1)
    .atHour(time.hour);

  if (typeof builder.nearMinute === "function") {
    builder = builder.nearMinute(time.minute);
  }
  builder.create();

  if (showToast) {
    SpreadsheetApp.getActive().toast(
      `Daily trigger created (${formatTimeLabel(time.hour, time.minute)})`,
      "Trigger Created",
    );
  }
}

function createDailyTrigger(silent) {
  const triggers = ScriptApp.getProjectTriggers();
  const existing = triggers.find(
    (t) => t.getHandlerFunction() === "runDailyCheck",
  );

  if (existing && !silent) {
    const response = SpreadsheetApp.getUi().alert(
      "Create / Refresh Trigger",
      "A daily trigger already exists. Replace it with the currently saved schedule?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO,
    );
    if (response !== SpreadsheetApp.getUi().Button.YES) return;
  }

  syncDailyTrigger(!silent);
  const time = getConfiguredAutomationTime();
  logActivity(
    "SYSTEM",
    existing ? "Trigger Updated" : "Trigger Created",
    `Daily check at ${formatTimeLabel(time.hour, time.minute)}`,
  );
}

function removeTriggers() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Remove Triggers",
    "Remove all automation triggers?",
    ui.ButtonSet.YES_NO,
  );

  if (response !== ui.Button.YES) return;

  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((t) => ScriptApp.deleteTrigger(t));

  SpreadsheetApp.getActive().toast(
    `Removed ${triggers.length} triggers`,
    "Triggers Removed",
  );
  logActivity("SYSTEM", "Triggers Removed", `Count: ${triggers.length}`);
}

// ============================================================================
// DIAGNOSTICS
// ============================================================================

function showDiagnostics() {
  showValidationReport();
  logActivity("SYSTEM", "Diagnostics Run", "Full system check performed");
}

function showValidationReport() {
  const ui = SpreadsheetApp.getUi();
  const results = validateSystem();
  let report = "CAR Monitoring System Validation\n";
  report += "=".repeat(40) + "\n\n";

  report += `Overall Status: ${results.passed ? "READY" : "ACTION REQUIRED"}\n\n`;

  report += "Checks:\n";
  if (results.notes.length === 0) report += "  (none)\n";
  else results.notes.forEach((note) => (report += `  - ${note}\n`));

  report += "\nIssues:\n";
  if (results.issues.length === 0) report += "  None\n";
  else results.issues.forEach((issue) => (report += `  - ${issue}\n`));

  report += "\nWarnings:\n";
  if (results.warnings.length === 0) report += "  None\n";
  else results.warnings.forEach((warning) => (report += `  - ${warning}\n`));

  ui.alert(report);
  return results;
}

// ============================================================================
// UTILITIES
// ============================================================================

function formatDate(date) {
  if (!date || !(date instanceof Date)) return "";
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "MMM dd, yyyy",
  );
}

function getSheetData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1).map((row) => {
    const obj = {};
    headers.forEach((h, i) => (obj[h] = row[i]));
    return obj;
  });

  return { headers, rows };
}

// Quick setup function for initial run
function initializeSystem() {
  setupAllSheets();
  updateAllDeadlines();
  generateAllReminders();
  SpreadsheetApp.getUi().alert(
    "System initialized successfully!\n\nNext steps:\n1. Review sheets and add your data\n2. Set up email triggers if needed\n3. Run diagnostics to verify",
  );
}
