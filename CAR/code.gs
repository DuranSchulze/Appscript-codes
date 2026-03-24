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
    REMINDER_SEND_LOG: "Reminder Send Log",
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
const PROP_DST_CGT_TABS = "FORMULA_DST_CGT_TABS";
const PROP_TRANSFER_TAX_TABS = "FORMULA_TRANSFER_TAX_TABS";
const DEADLINE_GROUP_TYPES = {
  DST: "DST",
  CGT_DOD: "CGT / DOD",
  TRANSFER_TAX: "Transfer Tax",
  ESTATE_TAX: "Estate Tax",
};

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
  const dstCgtTabs = parseJsonArrayProperty(props[PROP_DST_CGT_TABS]);
  const transferTaxTabs = parseJsonArrayProperty(props[PROP_TRANSFER_TAX_TABS]);

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
    DST_CGT_TABS: dstCgtTabs,
    TRANSFER_TAX_TABS: transferTaxTabs,
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

function getConfiguredDstCgtTabs() {
  return getConfig().DST_CGT_TABS || [];
}

function setConfiguredDstCgtTabs(sheetNames) {
  setJsonArrayProperty(PROP_DST_CGT_TABS, sheetNames);
}

function getConfiguredTransferTaxTabs() {
  return getConfig().TRANSFER_TAX_TABS || [];
}

function setConfiguredTransferTaxTabs(sheetNames) {
  setJsonArrayProperty(PROP_TRANSFER_TAX_TABS, sheetNames);
}

function setJsonArrayProperty(key, values) {
  const cleaned = Array.from(
    new Set(
      (values || [])
        .map((name) => (name || "").toString().trim())
        .filter(Boolean),
    ),
  );
  const props = PropertiesService.getScriptProperties();
  if (cleaned.length === 0) props.deleteProperty(key);
  else props.setProperty(key, JSON.stringify(cleaned));
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

function isValidEmailAddress(value) {
  return !!value && value.toString().trim().includes("@");
}

function getDailyCheckTriggers() {
  return ScriptApp.getProjectTriggers().filter(
    (t) => t.getHandlerFunction() === "runDailyCheck",
  );
}

function getAutomaticSendingStatus() {
  const time = getConfiguredAutomationTime();
  const triggers = getDailyCheckTriggers();
  return {
    enabled: triggers.length > 0,
    count: triggers.length,
    hour: time.hour,
    minute: time.minute,
    timeLabel: formatTimeLabel(time.hour, time.minute),
  };
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
    cfg.SHEETS.REMINDER_SEND_LOG,
  ];

  return ss
    .getSheets()
    .map((sheet) => sheet.getName())
    .filter((name) => !systemSheets.includes(name));
}

function promptForSheetSelection(title, promptText, currentSelection) {
  const ui = SpreadsheetApp.getUi();
  const allSheets = getEligibleAutomationSheets();

  if (allSheets.length === 0) {
    ui.alert(title, "No eligible data tabs were found in this spreadsheet.", ui.ButtonSet.OK);
    return null;
  }

  const currentLabel =
    currentSelection && currentSelection.length
      ? `Current: ${currentSelection.join(", ")}`
      : "Current: (none selected)";
  const dstCgtTabs = getConfiguredDstCgtTabs();
  const transferTabs = getConfiguredTransferTaxTabs();
  const list = allSheets
    .map((name, i) => {
      const markers = [];
      if (currentSelection && currentSelection.includes(name)) {
        markers.push("already applied to this group");
      }
      if (dstCgtTabs.includes(name)) markers.push("assigned to DST/CGT");
      if (transferTabs.includes(name)) markers.push("assigned to Transfer Tax");
      return `  ${i + 1}. ${name}${markers.length ? ` [${markers.join("; ")}]` : ""}`;
    })
    .join("\n");
  const response = ui.prompt(
    title,
    `${promptText}\n\n${currentLabel}\n\nAvailable tabs:\n${list}\n\nEnter one or more numbers separated by commas.\nLeave blank if none apply:`,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return null;

  const input = response.getResponseText().trim();
  if (!input) return [];

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
      title,
      `Invalid selection. Enter one or more numbers between 1 and ${allSheets.length}.`,
      ui.ButtonSet.OK,
    );
    return null;
  }

  return indexes.map((index) => allSheets[index]);
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
    "Select Automation Tabs",
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
    [cfg.HEADERS.SELLER_DONOR]: [
      "Seller / Donor",
      "Seller/Donor",
      "Seller / Doner",
      "Seller/Doner",
      "Seller Donor",
      "Seller Doner",
    ],
    [cfg.HEADERS.BUYER_DONEE]: [
      "Buyer / Donee",
      "Buyer/Donee",
      "Buyer Donee",
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
      "Staff Email",
      "Assignee Email",
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
    .replace(/[()]/g, " ")
    .replace(/[/:_-]+/g, " ")
    .replace(/[.,]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function headerContainsPhrase(header, phrases) {
  const normalizedHeader = normalizeHeaderName(header);
  return (phrases || []).some((phrase) =>
    normalizedHeader.indexOf(normalizeHeaderName(phrase)) >= 0,
  );
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

    if (headerContainsPhrase(headers[i], targetNames)) return i + 1;
  }

  return null;
}

function isDueDateHeader(header) {
  return headerContainsPhrase(header, [
    "dst due date",
    "cgt dod due date",
    "cgt / dod due date",
    "cgt/dod due date",
    "transfer tax due date",
    "estate tax due date",
  ]);
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
    cfg.SHEETS.REMINDER_SEND_LOG,
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
    (header) => headerContainsPhrase(header, ["reminder"]),
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

      if (!group.remainingCol && headerContainsPhrase(headers[j], ["remaining days"])) {
        group.remainingCol = j + 1;
        continue;
      }

      if (!group.statusCol && headerContainsPhrase(headers[j], ["status"])) {
        group.statusCol = j + 1;
        continue;
      }

      if (!group.reminderCol && headerContainsPhrase(headers[j], ["reminder"])) {
        group.reminderCol = j + 1;
      }
    }

    groups.push(group);
  }

  return groups;
}

function getFormulaDrivenDeadlineGroups(sheet) {
  return getDeadlineGroups(sheet)
    .map((group) => {
      let type = null;
      if (headerContainsPhrase(group.label, ["dst due date"])) {
        type = DEADLINE_GROUP_TYPES.DST;
      } else if (
        headerContainsPhrase(group.label, [
          "cgt dod due date",
          "cgt / dod due date",
          "cgt/dod due date",
        ])
      ) {
        type = DEADLINE_GROUP_TYPES.CGT_DOD;
      } else if (headerContainsPhrase(group.label, ["transfer tax due date"])) {
        type = DEADLINE_GROUP_TYPES.TRANSFER_TAX;
      } else if (headerContainsPhrase(group.label, ["estate tax due date"])) {
        type = DEADLINE_GROUP_TYPES.ESTATE_TAX;
      }
      if (
        type === DEADLINE_GROUP_TYPES.ESTATE_TAX &&
        (!group.remainingCol || !group.statusCol || !group.reminderCol)
      ) {
        return null;
      }
      return type ? Object.assign({ type: type }, group) : null;
    })
    .filter(Boolean);
}

function getExpectedGroupTypesForSheet(sheetName) {
  const expected = [];
  if (getConfiguredDstCgtTabs().includes(sheetName)) {
    expected.push(DEADLINE_GROUP_TYPES.DST, DEADLINE_GROUP_TYPES.CGT_DOD);
  }
  if (getConfiguredTransferTaxTabs().includes(sheetName)) {
    expected.push(DEADLINE_GROUP_TYPES.TRANSFER_TAX);
  }
  return Array.from(new Set(expected));
}

function describeConfiguredGroupsForSheet(sheetName) {
  const groups = [];
  if (getConfiguredDstCgtTabs().includes(sheetName)) groups.push("DST/CGT");
  if (getConfiguredTransferTaxTabs().includes(sheetName)) groups.push("Transfer Tax");
  return groups;
}

function getSendableGroupTypesForSheet(sheet) {
  const available = getFormulaDrivenDeadlineGroups(sheet).map((group) => group.type);
  const expected = getExpectedGroupTypesForSheet(sheet.getName());
  const allowed = expected.filter((type) => available.includes(type));
  if (
    getConfiguredTransferTaxTabs().includes(sheet.getName()) &&
    available.includes(DEADLINE_GROUP_TYPES.ESTATE_TAX)
  ) {
    allowed.push(DEADLINE_GROUP_TYPES.ESTATE_TAX);
  }
  return Array.from(new Set(allowed));
}

function getTrackedDeadlineGroups(sheet) {
  const allowedTypes = getSendableGroupTypesForSheet(sheet);
  return getFormulaDrivenDeadlineGroups(sheet).filter((group) =>
    allowedTypes.includes(group.type),
  );
}

function hasFormulaAtDataStart(sheet, column) {
  if (!column) return false;
  const row = getDataStartRow(sheet);
  return !!sheet.getRange(row, column).getFormula();
}

function verifyFormulaDrivenGroup(sheet, group) {
  const missing = [];
  if (!group.remainingCol) missing.push("Remaining Days column");
  if (!group.statusCol) missing.push("Status column");
  if (!group.reminderCol) missing.push("Reminder column");

  const formulaMissing = [];
  ["dueDateCol", "remainingCol", "statusCol", "reminderCol"].forEach((key) => {
    const col = group[key];
    if (col && !hasFormulaAtDataStart(sheet, col)) {
      const labelMap = {
        dueDateCol: "Due Date formula",
        remainingCol: "Remaining Days formula",
        statusCol: "Status formula",
        reminderCol: "Reminder formula",
      };
      formulaMissing.push(labelMap[key]);
    }
  });

  return {
    missingColumns: missing,
    missingFormulas: formulaMissing,
  };
}

function columnToLetter(column) {
  let temp = "";
  let letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function findDateOfDeathColumn(headers) {
  for (let i = 0; i < headers.length; i++) {
    if (
      headerContainsPhrase(headers[i], [
        "decedents date of death",
        "decedent date of death",
        "date of death",
      ])
    ) {
      return i + 1;
    }
  }
  return null;
}

function buildArrayFormulaForGroup(groupType, refs) {
  const companyRange = `${refs.companyCol}${refs.startRow}:${refs.companyCol}`;
  const baseRange = `${refs.baseCol}${refs.startRow}:${refs.baseCol}`;
  const dueRange = `${refs.dueCol}${refs.startRow}:${refs.dueCol}`;
  const remainingRange = `${refs.remainingCol}${refs.startRow}:${refs.remainingCol}`;
  const statusRange = `${refs.statusCol}${refs.startRow}:${refs.statusCol}`;
  let dueFormula = "";
  let reminderSuffix = "";

  if (groupType === DEADLINE_GROUP_TYPES.DST) {
    dueFormula = `=ARRAYFORMULA(IF(ISNUMBER(${baseRange}),EOMONTH(${baseRange},0)+5,""))`;
    reminderSuffix = "DST Payment for CAR Transfer of Shares";
  } else if (groupType === DEADLINE_GROUP_TYPES.CGT_DOD) {
    dueFormula = `=ARRAYFORMULA(IF(ISNUMBER(${baseRange}),${baseRange}+30,""))`;
    reminderSuffix = "CGT / DOD Payment for CAR Transfer of Shares";
  } else if (groupType === DEADLINE_GROUP_TYPES.TRANSFER_TAX) {
    dueFormula = `=ARRAYFORMULA(IF(ISNUMBER(${baseRange}),${baseRange}+60,""))`;
    reminderSuffix = "Transfer Tax Payment";
  } else if (groupType === DEADLINE_GROUP_TYPES.ESTATE_TAX) {
    dueFormula = `=ARRAYFORMULA(IF(ISNUMBER(${baseRange}),EDATE(${baseRange},12),""))`;
    reminderSuffix = "Estate Tax Payment";
  } else {
    return null;
  }

  const remainingFormula = `=ARRAYFORMULA(IF((${baseRange}="")+(${dueRange}=""),"",DATEVALUE(${dueRange})-TODAY()))`;
  const statusFormula = `=ARRAYFORMULA(IF(${dueRange}="","",IF(((${remainingRange}=15)+(${remainingRange}=10)+(${remainingRange}=5)+(${remainingRange}=2)+(${remainingRange}=0))>0,"SEND","WAIT")))`;
  const reminderFormula = `=ARRAYFORMULA(IF(${statusRange}="WAIT","",IF(${dueRange}="","",(IF(${remainingRange}=0,"🔴 DUE TODAY: ","⚠️ REMINDER ("&${remainingRange}&" days left): ")&${companyRange}&" - ${reminderSuffix}"))))`;

  return {
    dueFormula: dueFormula,
    remainingFormula: remainingFormula,
    statusFormula: statusFormula,
    reminderFormula: reminderFormula,
  };
}

function applyFormulasToConfiguredSheets() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  const sheetNames = Array.from(
    new Set(getConfiguredDstCgtTabs().concat(getConfiguredTransferTaxTabs())),
  );

  if (!sheetNames.length) {
    ui.alert("Apply Formulas", "No configured tabs found. Run Quick Setup first.", ui.ButtonSet.OK);
    return;
  }

  let applied = 0;
  const skipped = [];

  sheetNames.forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      skipped.push(`${sheetName}: tab not found`);
      return;
    }

    const headers = getHeaders(sheet);
    const companyCol = getColumnIndex(sheet, cfg.HEADERS.CLIENT_NAME);
    const notaryCol = getColumnIndex(sheet, cfg.HEADERS.NOTARY_DATE);
    const dateOfDeathCol = findDateOfDeathColumn(headers);
    const startRow = getDataStartRow(sheet);
    const groups = getTrackedDeadlineGroups(sheet);

    if (!companyCol) {
      skipped.push(`${sheetName}: Company column not found`);
      return;
    }

    groups.forEach((group) => {
      if (!group.remainingCol || !group.statusCol || !group.reminderCol) {
        skipped.push(`${sheetName}: ${group.type} group is incomplete`);
        return;
      }

      let baseCol = notaryCol;
      if (group.type === DEADLINE_GROUP_TYPES.ESTATE_TAX) {
        baseCol = dateOfDeathCol || notaryCol;
      }
      if (!baseCol) {
        skipped.push(`${sheetName}: ${group.type} base date column not found`);
        return;
      }

      const formulas = buildArrayFormulaForGroup(group.type, {
        startRow: startRow,
        baseCol: columnToLetter(baseCol),
        companyCol: columnToLetter(companyCol),
        dueCol: columnToLetter(group.dueDateCol),
        remainingCol: columnToLetter(group.remainingCol),
        statusCol: columnToLetter(group.statusCol),
      });
      if (!formulas) return;

      sheet.getRange(startRow, group.dueDateCol).setFormula(formulas.dueFormula);
      sheet.getRange(startRow, group.remainingCol).setFormula(formulas.remainingFormula);
      sheet.getRange(startRow, group.statusCol).setFormula(formulas.statusFormula);
      sheet.getRange(startRow, group.reminderCol).setFormula(formulas.reminderFormula);
      applied++;
    });
  });

  const message = `Applied formulas to ${applied} deadline group(s)${
    skipped.length ? `\n\nSkipped:\n- ${skipped.join("\n- ")}` : ""
  }`;
  logActivity("SYSTEM", "Apply Formulas", message.replace(/\n/g, " | "));
  ui.alert("Apply Formulas", message, ui.ButtonSet.OK);
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
  const automation = getAutomaticSendingStatus();
  const dstCgtTabs = getConfiguredDstCgtTabs();
  const transferTabs = getConfiguredTransferTaxTabs();
  const cfg = getConfig();

  let msg = "CAR Monitoring — Formula Reminder Settings\n";
  msg += "=".repeat(44) + "\n\n";
  msg += "FORMULA-DRIVEN TAB GROUPS\n";
  msg += `  DST / CGT Tabs: ${dstCgtTabs.length ? dstCgtTabs.join(", ") : "(none selected)"}\n`;
  msg += `  Transfer Tax Tabs: ${transferTabs.length ? transferTabs.join(", ") : "(none selected)"}\n`;
  msg += "\nEMAIL RULE\n";
  msg += "  Recipient Column: Staff Email\n";
  msg += "  Send Condition: Status = SEND and Reminder is not blank\n";
  msg += "  Duplicate Rule: One email per row/group/day\n";
  msg += "\nROW SAFEGUARDS\n";
  msg += `  Header Row: ${cfg.HEADER_ROW || "(auto-detect)"}\n`;
  msg += `  Data Start Row: ${cfg.DATA_START_ROW || "(auto-detect first formula/data row)"}\n`;
  msg += "\nAUTOMATIC SENDING\n";
  msg += `  Status: ${automation.enabled ? "ACTIVE" : "NOT ACTIVE"}\n`;
  msg += `  Daily Run Time: ${automation.timeLabel}\n`;
  msg += `  Installed Triggers: ${automation.count}\n`;
  msg += "\n— Use Quick Setup to select tabs and verify formulas\n";
  msg += "— Use Run Validation to review missing groups or formulas\n";
  msg += "— Use Send Reminder Emails to send reminders from existing sheet formulas";

  ui.alert("Settings", msg, ui.ButtonSet.OK);
}

function configureNotificationEmails() {
  const ui = SpreadsheetApp.getUi();
  const current = getNotificationEmails();
  const response = ui.prompt(
    "2.4 Set Notification Emails",
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
    "2.4 Set Notification Emails",
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
    "Set Automatic Sending Time",
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
  const automationWasActive = getDailyCheckTriggers().length > 0;
  if (automationWasActive) syncDailyTrigger(false);
  logActivity(
    "SYSTEM",
    "Automation Schedule Updated",
    `${formatTimeLabel(hour, minute)}${automationWasActive ? " (trigger refreshed)" : " (saved only)"}`,
  );
  ui.alert(
    "Set Automatic Sending Time",
    automationWasActive
      ? `Automatic Sending time saved as ${formatTimeLabel(hour, minute)}.\n\nThe active daily trigger was refreshed to match the new schedule.`
      : `Automatic Sending time saved as ${formatTimeLabel(hour, minute)}.\n\nAutomatic Sending is not active yet. Use 5.1 Start Automatic Sending when you are ready.`,
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
    "Reset Formula Reminder Settings",
    "Reset the selected DST/CGT tabs, Transfer Tax tabs, schedule, and reminder settings?",
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
  scriptProps.deleteProperty(PROP_DST_CGT_TABS);
  scriptProps.deleteProperty(PROP_TRANSFER_TAX_TABS);
  scriptProps.deleteProperty("AUTOMATION_HOUR");
  scriptProps.deleteProperty("AUTOMATION_MINUTE");
  scriptProps.deleteProperty("EMAIL_RECIPIENT_MODE");
  scriptProps.deleteProperty("NOTIFICATION_EMAILS");
  invalidateConfigCache();

  ui.alert("Formula reminder settings have been reset.");
  logActivity(
    "SYSTEM",
    "Settings Reset",
    "Formula reminder setup restored to defaults",
  );
}

function setHeaderRow() {
  const ui = SpreadsheetApp.getUi();
  const current = getConfig().HEADER_ROW;
  const response = ui.prompt(
    "Set Header Row",
    `The header row is the row that contains your column names (Company, Seller/Donor, etc.).
Merged title or label rows above it should not be selected.

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
    `The data start row is where your first actual data entry begins (below the header row).
If some rows are merged and used only as labels or section titles, skip those and enter the first real input row.
The first real input row often contains formula cells in the due-date group columns.

Current value: ${current || "auto-detect first usable row"}

Enter the row number (e.g. 5), or leave blank to use auto-detection:`,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();
  const scriptProps = PropertiesService.getScriptProperties();

  if (input === "") {
    scriptProps.deleteProperty("DATA_START_ROW");
    invalidateConfigCache();
    ui.alert("Data start row reset to auto-detection.");
    logActivity("SYSTEM", "Data Start Row", "Reset to auto-detect");
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

  const automation = getAutomaticSendingStatus();
  SpreadsheetApp.getUi().alert(
    "Setup Applied",
    `Automation tabs: ${selectedSheets.length > 0 ? selectedSheets.join(", ") : sheet.getName()}\nHeader row: ${hRow}\nData start row: ${dRow}\nNotification emails: ${emails.length > 0 ? emails.join(", ") : "(none)"}\nAutomatic Sending time: ${formatTimeLabel(cfg.AUTOMATION_HOUR, cfg.AUTOMATION_MINUTE)}\nAutomatic Sending status: ${automation.enabled ? "ACTIVE" : "NOT ACTIVE"}`,
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

function quickSetup() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    "Quick Setup",
    "This setup will ask you for:\n\n1. Tabs that contain DST and CGT / DOD deadline groups\n2. Tabs that contain Transfer Tax deadline groups\n\nThe spreadsheet formulas remain the source of truth. This setup only verifies the required due-date columns and formulas, then saves the tabs for reminder sending.",
    ui.ButtonSet.OK,
  );

  const dstCgtTabs = promptForSheetSelection(
    "Quick Setup — DST / CGT Tabs",
    "Select the tabs that contain both the DST Due Date group and the CGT / DOD Due Date group.",
    getConfiguredDstCgtTabs(),
  );
  if (dstCgtTabs === null) return;

  const transferTabs = promptForSheetSelection(
    "Quick Setup — Transfer Tax Tabs",
    "Select the tabs that contain the Transfer Tax Due Date group. Estate Tax will be auto-detected on these tabs when present.",
    getConfiguredTransferTaxTabs(),
  );
  if (transferTabs === null) return;

  runRowSetup();

  setConfiguredDstCgtTabs(dstCgtTabs);
  setConfiguredTransferTaxTabs(transferTabs);

  const results = validateSystem();
  let msg = "Quick Setup Saved\n";
  msg += "=".repeat(30) + "\n\n";
  msg += `DST / CGT Tabs: ${dstCgtTabs.length ? dstCgtTabs.join(", ") : "(none selected)"}\n`;
  msg += `Transfer Tax Tabs: ${transferTabs.length ? transferTabs.join(", ") : "(none selected)"}\n`;
  msg += `Header Row: ${getConfig().HEADER_ROW || "auto-detect"}\n`;
  msg += `Data Start Row: ${getConfig().DATA_START_ROW || "auto-detect"}\n`;
  msg += `Automatic Sending Time: ${getAutomaticSendingStatus().timeLabel}\n\n`;
  msg += `Validation Status: ${results.issues.length ? "ACTION REQUIRED" : results.warnings.length ? "READY WITH WARNINGS" : "READY"}\n`;
  if (results.issues.length) {
    msg += `\nIssues Found: ${results.issues.length}\nUse Run Validation to review the exact tabs, groups, and formulas that need attention.`;
  } else if (results.warnings.length) {
    msg += `\nWarnings Found: ${results.warnings.length}\nUse Run Validation to review them.`;
  } else {
    msg += "\nAll selected tabs passed validation.";
  }
  ui.alert("Quick Setup", msg, ui.ButtonSet.OK);
  logActivity(
    "SYSTEM",
    "Quick Setup Saved",
    `DST/CGT: ${dstCgtTabs.join(", ") || "(none)"} | Transfer: ${transferTabs.join(", ") || "(none)"}`,
  );
}

function showAbout() {
  const ui = SpreadsheetApp.getUi();
  const dstCgtTabs = getConfiguredDstCgtTabs();
  const transferTabs = getConfiguredTransferTaxTabs();
  let msg = "CAR Monitoring — About\n";
  msg += "=".repeat(30) + "\n\n";
  msg += "What this Apps Script does:\n";
  msg += "- It does not compute deadlines itself.\n";
  msg += "- It reads the formulas already present in your sheet.\n";
  msg += '- It scans each due-date group and sends an email only when that group\'s Status is "SEND" and Reminder is not blank.\n';
  msg += "- It sends reminders to the Staff Email column only.\n";
  msg += "- It sends separate emails per deadline group.\n";
  msg += "- It prevents duplicate same-day sends for the same row, tab, deadline group, and recipient.\n";
  msg += "\nPer tab group:\n";
  msg += `- DST / CGT Tabs: ${dstCgtTabs.length ? dstCgtTabs.join(", ") : "(none selected)"}\n`;
  msg += "  Expected groups: DST Due Date and CGT / DOD Due Date\n";
  msg += "  The script validates that these due-date groups, formula cells, Reminder, Status, and Staff Email exist.\n";
  msg += `- Transfer Tax Tabs: ${transferTabs.length ? transferTabs.join(", ") : "(none selected)"}\n`;
  msg += "  Expected group: Transfer Tax Due Date\n";
  msg += "  Estate Tax is also included automatically on these tabs when detected.\n";
  msg += "\nAutomatic Sending:\n";
  msg += "- If enabled, the daily trigger runs the same reminder scan automatically.\n";
  msg += "- If disabled, you can still use Send Reminder Emails manually.\n";
  msg += "\nQuick Setup:\n";
  msg += "- Lets you choose which tabs belong to DST/CGT and Transfer Tax groups.\n";
  msg += "- Shows already-assigned tabs during selection.\n";
  msg += "- Saves your selections and validates whether the selected tabs match the expected deadline groups.\n";
  msg += "\nApply Formulas:\n";
  msg += "- Writes the ARRAYFORMULA formulas into the configured Due Date, Remaining Days, Status, and Reminder columns.\n";
  msg += "- Uses the detected column positions for each tab instead of fixed letters.\n";

  ui.alert("About", msg, ui.ButtonSet.OK);
}

// ============================================================================
// ON OPEN - CUSTOM MENU
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("CAR Monitoring")
    .addItem("Quick Setup", "quickSetup")
    .addItem("Apply Formulas to Configured Sheets", "applyFormulasToConfiguredSheets")
    .addItem("Run Validation", "showValidationReport")
    .addItem("Send Reminder Emails", "sendReminderEmails")
    .addItem("About", "showAbout")
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("Automatic Sending")
        .addItem("Start Automatic Sending", "startAutomaticSending")
        .addItem("Set Automatic Sending Time", "configureDailyAutomationTime")
        .addItem("Stop Automatic Sending", "stopAutomaticSending")
        .addItem("Check Automatic Sending Status", "showAutomaticSendingStatus"),
    )
    .addSubMenu(
      ui
        .createMenu("Settings")
        .addItem("View Current Settings", "showSettings")
        .addItem("Set Header Row", "setHeaderRow")
        .addItem("Set Data Start Row", "setDataStartRow")
        .addItem("Reset Formula Reminder Settings", "resetSettingsToDefaults")
        .addItem("View Activity Log", "viewActivityLog"),
    )
    .addSubMenu(
      ui
        .createMenu("Advanced Tools")
        .addItem("Run Daily Check Now", "runDailyCheck")
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

  const maxScan = Math.min(10, sheet.getLastRow());
  if (maxScan < 1) return 1;

  const scanData = sheet.getRange(1, 1, maxScan, sheet.getLastColumn()).getValues();
  let bestRow = 1;
  let bestScore = -1;

  for (let i = 0; i < scanData.length; i++) {
    const rowNumber = i + 1;
    const rowValues = scanData[i];
    const nonEmpty = rowValues.filter((v) => v !== "" && v !== null).length;
    const headerHits = rowValues.filter((value) =>
      headerContainsPhrase(value, [
        "company",
        "seller donor",
        "buyer donee",
        "services",
        "staff email",
        "notary date",
        "dst due date",
        "cgt dod due date",
        "cgt / dod due date",
        "transfer tax due date",
        "estate tax due date",
        "remaining days",
        "status",
        "reminder",
      ]),
    ).length;
    const isLikelyMergedLabel =
      sheet.getRange(rowNumber, 1, 1, Math.max(1, sheet.getLastColumn())).isPartOfMerge() &&
      nonEmpty <= 2;
    const score = headerHits * 10 + nonEmpty - (isLikelyMergedLabel ? 100 : 0);

    if (score > bestScore) {
      bestScore = score;
      bestRow = i + 1;
    }
  }

  return bestRow;
}

function getDataStartRow(sheet) {
  const configured = getConfig().DATA_START_ROW;
  if (configured && configured >= 1) return configured;
  const headerRow = getHeaderRow(sheet);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow <= headerRow || lastColumn === 0) return headerRow + 1;

  const groups = getFormulaDrivenDeadlineGroups(sheet);
  for (let row = headerRow + 1; row <= lastRow; row++) {
    const rowRange = sheet.getRange(row, 1, 1, lastColumn);
    const values = rowRange.getValues()[0];
    const formulas = rowRange.getFormulas()[0];
    const hasAnyData = values.some((value) => value !== "" && value !== null);
    if (!hasAnyData) continue;

    const hasTrackedFormula = groups.some((group) =>
      [group.dueDateCol, group.remainingCol, group.statusCol, group.reminderCol]
        .filter(Boolean)
        .some((col) => !!formulas[col - 1]),
    );
    const isLikelyMergedLabel = rowRange.isPartOfMerge() && !hasTrackedFormula;
    if (isLikelyMergedLabel) continue;
    if (hasTrackedFormula) return row;

    const staffEmailCol = findHeaderIndexByName(getHeaders(sheet), getConfig().HEADERS.STAFF_EMAIL);
    const companyCol = findHeaderIndexByName(getHeaders(sheet), getConfig().HEADERS.CLIENT_NAME);
    if (
      (staffEmailCol && values[staffEmailCol - 1] !== "") ||
      (companyCol && values[companyCol - 1] !== "")
    ) {
      return row;
    }
  }

  return headerRow + 1;
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

function createValidationEntry(message, fix) {
  return { message: message, fix: fix };
}

function validateSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const issues = [];
  const warnings = [];
  const notes = [];
  const dstCgtTabs = getConfiguredDstCgtTabs();
  const transferTabs = getConfiguredTransferTaxTabs();
  const allConfiguredTabs = Array.from(new Set(dstCgtTabs.concat(transferTabs)));

  notes.push(`Spreadsheet: ${ss.getName()}`);
  notes.push(`Automatic sending time: ${getAutomaticSendingStatus().timeLabel}`);
  notes.push("Recipient column: Staff Email");

  if (dstCgtTabs.length === 0 && transferTabs.length === 0) {
    issues.push(
      createValidationEntry(
        "No tabs are configured yet for formula-driven reminders.",
        "Fix: Run Quick Setup and select at least one DST/CGT tab or one Transfer Tax tab.",
      ),
    );
  }

  allConfiguredTabs.forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      issues.push(
        createValidationEntry(
          `Configured tab is missing: ${sheetName}`,
          "Fix: Run Quick Setup again and select only tabs that still exist in the spreadsheet.",
        ),
      );
      return;
    }

    const headers = getHeaders(sheet);
    if (!headers.length) {
      issues.push(
        createValidationEntry(
          `${sheetName}: No header row could be detected.`,
          "Fix: Make sure the tab has a visible header row before the formula columns.",
        ),
      );
      return;
    }

    const headerRow = getHeaderRow(sheet);
    const dataStartRow = getDataStartRow(sheet);
    if (dataStartRow <= headerRow) {
      issues.push(
        createValidationEntry(
          `${sheetName}: Data start row (${dataStartRow}) must be below the header row (${headerRow}).`,
          "Fix: Use Settings > Set Header Row or Set Data Start Row and choose the first real input row below the headers.",
        ),
      );
    } else {
      const rowRange = sheet.getRange(dataStartRow, 1, 1, Math.max(1, sheet.getLastColumn()));
      const rowValues = rowRange.getValues()[0];
      const rowFormulas = rowRange.getFormulas()[0];
      const hasAnyFormula = rowFormulas.some(Boolean);
      const hasAnyData = rowValues.some((value) => value !== "" && value !== null);
      if (rowRange.isPartOfMerge() && !hasAnyFormula) {
        warnings.push(
          createValidationEntry(
            `${sheetName}: The selected data start row (${dataStartRow}) looks like a merged label row, not a real input row.`,
            "Fix: Set Data Start Row to the first non-label row where actual data begins and formula cells are present.",
          ),
        );
      } else if (!hasAnyData) {
        warnings.push(
          createValidationEntry(
            `${sheetName}: The selected data start row (${dataStartRow}) is empty.`,
            "Fix: Set Data Start Row to the first actual row used for data entry and formulas.",
          ),
        );
      }
    }

    if (!findHeaderIndexByName(headers, getConfig().HEADERS.NOTARY_DATE)) {
      issues.push(
        createValidationEntry(
          `${sheetName}: Missing required header "NOTARY DATE OF DOCUMENT".`,
          'Fix: Make sure the tab contains the "NOTARY DATE OF DOCUMENT" column.',
        ),
      );
    }

    const staffEmailCol = findHeaderIndexByName(headers, getConfig().HEADERS.STAFF_EMAIL);
    if (!staffEmailCol) {
      issues.push(
        createValidationEntry(
          `${sheetName}: Missing required header "Staff Email".`,
          'Fix: Add or restore the "Staff Email" column on this tab.',
        ),
      );
    }

    const groups = getFormulaDrivenDeadlineGroups(sheet);
    const groupsByType = {};
    groups.forEach((group) => {
      groupsByType[group.type] = group;
    });

    getExpectedGroupTypesForSheet(sheetName).forEach((type) => {
      const group = groupsByType[type];
      if (!group) {
        issues.push(
          createValidationEntry(
            `${sheetName}: This tab is assigned to ${describeConfiguredGroupsForSheet(sheetName).join(" + ")}, but the required deadline group "${type}" was not detected.`,
            `Fix: Restore the "${type}" Due Date / Remaining Days / Status / Reminder columns on this tab, or remove this tab from the wrong Quick Setup group.`,
          ),
        );
        return;
      }

      const verification = verifyFormulaDrivenGroup(sheet, group);
      if (verification.missingColumns.length) {
        issues.push(
          createValidationEntry(
            `${sheetName}: "${type}" group was detected, but it is incomplete (${verification.missingColumns.join(", ")} missing).`,
            `Fix: Restore the full "${type}" group with Due Date, Remaining Days, Status, and Reminder columns.`,
          ),
        );
      }
      if (verification.missingFormulas.length) {
        warnings.push(
          createValidationEntry(
            `${sheetName}: "${type}" group is missing formula cells at the first data row (${verification.missingFormulas.join(", ")}).`,
            `Fix: Restore the existing sheet formulas for the "${type}" group. Apps Script will not overwrite them.`,
          ),
        );
      }
    });

    const estateGroup = groupsByType[DEADLINE_GROUP_TYPES.ESTATE_TAX];
    if (estateGroup && getConfiguredTransferTaxTabs().includes(sheetName)) {
      const estateVerification = verifyFormulaDrivenGroup(sheet, estateGroup);
      if (estateVerification.missingColumns.length) {
        issues.push(
          createValidationEntry(
            `${sheetName}: "Estate Tax" group was auto-detected, but it is incomplete (${estateVerification.missingColumns.join(", ")} missing).`,
            'Fix: Restore the full "Estate Tax" Due Date / Remaining Days / Status / Reminder columns or remove the partial group.',
          ),
        );
      }
      if (estateVerification.missingFormulas.length) {
        warnings.push(
          createValidationEntry(
            `${sheetName}: "Estate Tax" group is missing formula cells at the first data row (${estateVerification.missingFormulas.join(", ")}).`,
            'Fix: Restore the existing sheet formulas for the "Estate Tax" group. Apps Script will not overwrite them.',
          ),
        );
      }
      notes.push(`${sheetName}: Estate Tax group auto-detected.`);
    }
  });

  let quota = null;
  try {
    quota = MailApp.getRemainingDailyQuota();
  } catch (e) {
    issues.push(
      createValidationEntry(
        "Unable to read the MailApp quota. Authorization may be incomplete.",
        "Fix: Run 4.3 Send Reminder Emails or 5.1 Start Automatic Sending once and accept the required Apps Script permissions.",
      ),
    );
  }
  if (quota !== null) {
    if (quota <= 0) {
      issues.push(
        createValidationEntry(
          "Mail quota is exhausted.",
          "Fix: Wait for the Google Apps Script mail quota to reset, then run 4.3 Send Reminder Emails again.",
        ),
      );
    } else if (quota < 10) {
      warnings.push(
        createValidationEntry(
          `Mail quota is low (${quota} remaining).`,
          "Fix: Limit manual test sends or wait for the quota reset before sending reminders in bulk.",
        ),
      );
    }
  }

  const triggers = getDailyCheckTriggers();
  if (triggers.length === 0) {
    warnings.push(
      createValidationEntry(
        "Automatic Sending is not enabled yet.",
        "Fix: 5. Automatic Sending > 5.1 Start Automatic Sending to install the daily trigger.",
      ),
    );
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

function ensureReminderSendLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  let sheet = ss.getSheetByName(cfg.SHEETS.REMINDER_SEND_LOG);

  if (!sheet) {
    sheet = ss.insertSheet(cfg.SHEETS.REMINDER_SEND_LOG);
    const headers = [
      "Date",
      "Sheet",
      "Row",
      "Deadline Group",
      "Recipient",
      "Trigger Type",
      "Status",
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function getTodayKey() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function wasReminderSentToday(sheetName, rowNumber, deadlineGroup, recipientLogKey) {
  const sheet = ensureReminderSendLogSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const today = getTodayKey();
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  return data.some((entry) => {
    const dateKey = entry[0] instanceof Date
      ? Utilities.formatDate(entry[0], Session.getScriptTimeZone(), "yyyy-MM-dd")
      : entry[0].toString();
    return (
      dateKey === today &&
      entry[1] === sheetName &&
      Number(entry[2]) === Number(rowNumber) &&
      entry[3] === deadlineGroup &&
      entry[4] === recipientLogKey &&
      entry[6] === "SENT"
    );
  });
}

function logReminderSend(sheetName, rowNumber, deadlineGroup, recipientLogKey, triggerType, status) {
  ensureReminderSendLogSheet().appendRow([
    new Date(),
    sheetName,
    rowNumber,
    deadlineGroup,
    recipientLogKey,
    triggerType,
    status,
  ]);
}

function markReminderCellAsSent(sheet, rowNumber, columnNumber) {
  if (!sheet || !rowNumber || !columnNumber) return;
  sheet
    .getRange(rowNumber, columnNumber)
    .setBackground("#93c47d")
    .setFontColor("#ffffff")
    .setFontWeight("bold");
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

function parseEmailAddresses(value) {
  if (value === null || value === undefined || value === "") return [];

  const seen = {};
  return value
    .toString()
    .split(/[,\n;]+/)
    .map((email) => email.trim().toLowerCase())
    .filter((email) => {
      if (!isValidEmailAddress(email) || seen[email]) return false;
      seen[email] = true;
      return true;
    });
}

function buildReminderRecipients(staffEmail, clientEmail) {
  return parseEmailAddresses(staffEmail);
}

function buildReminderCcRecipients(includeNotificationEmails, recipients) {
  if (includeNotificationEmails === false) return [];

  const recipientLookup = {};
  (recipients || []).forEach((email) => {
    recipientLookup[email] = true;
  });

  return getNotificationEmails().filter((email) => !recipientLookup[email]);
}

function buildRecipientLogKey(toRecipients, ccRecipients) {
  const toList = (toRecipients || []).join(", ");
  const ccList = (ccRecipients || []).join(", ");
  return ccList ? `TO: ${toList} | CC: ${ccList}` : `TO: ${toList}`;
}

function getPartyDetailsForReminder(headers, row) {
  const cfg = getConfig();
  const sellerDonorIndex = findHeaderIndexByName(headers, cfg.HEADERS.SELLER_DONOR);
  const buyerDoneeIndex = findHeaderIndexByName(headers, cfg.HEADERS.BUYER_DONEE);

  return {
    sellerDonor: sellerDonorIndex ? row[sellerDonorIndex - 1] : "",
    buyerDonee: buyerDoneeIndex ? row[buyerDoneeIndex - 1] : "",
  };
}

function getReminderEmailPayloadForRow(sheet, rowNumber) {
  const cfg = getConfig();
  const dataStart = getDataStartRow(sheet);
  if (rowNumber < dataStart) {
    return {
      error: `Row ${rowNumber} is above the data start row (${dataStart}).`,
    };
  }

  autoUpdateStatus(sheet, rowNumber);
  updateRowDeadlines(sheet, rowNumber);
  updateRowReminder(sheet, rowNumber);

  const headers = getHeaders(sheet);
  const row = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  const deadlineGroups = getDeadlineGroups(sheet);

  const genericReminderIndex = findHeaderIndexByName(headers, cfg.HEADERS.REMINDER);
  const serviceTypeIndex = findHeaderIndexByName(headers, cfg.HEADERS.SERVICE_TYPE);
  const clientNameIndex = findHeaderIndexByName(headers, cfg.HEADERS.CLIENT_NAME);
  const staffEmailIndex = findHeaderIndexByName(headers, cfg.HEADERS.STAFF_EMAIL);
  const clientEmailIndex = findHeaderIndexByName(headers, cfg.HEADERS.CLIENT_EMAIL);
  const remarksIndex = findHeaderIndexByName(headers, cfg.HEADERS.REMARKS);
  const revisitDateIndex = findHeaderIndexByName(headers, cfg.HEADERS.REVISIT_DATE);
  const revisitStatusIndex = findHeaderIndexByName(
    headers,
    cfg.HEADERS.REVISIT_STATUS,
  );
  const revisitNotesIndex = findHeaderIndexByName(headers, cfg.HEADERS.REVISIT_NOTES);

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

  if (!reminders.length) {
    return {
      error:
        "No reminder message is currently generated for this row. Update the due dates or choose a row that is near or past due.",
    };
  }

  const clientName = clientNameIndex ? row[clientNameIndex - 1] : "";
  const staffEmail = staffEmailIndex ? row[staffEmailIndex - 1] : "";
  const clientEmail = clientEmailIndex ? row[clientEmailIndex - 1] : "";
  const partyDetails = getPartyDetailsForReminder(headers, row);
  const recipients = buildReminderRecipients(staffEmail, clientEmail);
  if (recipients.length === 0) {
    return {
      error:
        'No valid Staff Email was found for this row. Add a valid email address in the "Staff Email" column first.',
    };
  }

  const revisitInfo =
    revisitDateIndex && row[revisitDateIndex - 1] instanceof Date
      ? {
          date: row[revisitDateIndex - 1],
          status: revisitStatusIndex ? row[revisitStatusIndex - 1] : "",
          notes: revisitNotesIndex ? row[revisitNotesIndex - 1] : "",
          remainingDays: calculateRemainingDays(row[revisitDateIndex - 1]),
        }
      : null;

  return {
    data: {
      recipients: recipients,
      clientName: clientName || "Unknown Client",
      sellerDonor: partyDetails.sellerDonor || "",
      buyerDonee: partyDetails.buyerDonee || "",
      serviceType: serviceTypeIndex ? row[serviceTypeIndex - 1] : "",
      reminders: reminders,
      deadlines: deadlineSummaries,
      sheetName: sheet.getName(),
      rowNumber: rowNumber,
      remarks: remarksIndex ? row[remarksIndex - 1] : "",
      revisit: revisitInfo,
    },
  };
}

function sendReminderEmails() {
  return sendReminderEmailsByTriggerType("manual");
}

function sendReminderEmailsByTriggerType(triggerType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  const interactive = triggerType !== "automatic";
  const sheets = Array.from(
    new Set(getConfiguredDstCgtTabs().concat(getConfiguredTransferTaxTabs())),
  );

  let emailCount = 0;
  const skipped = [];
  const quota = MailApp.getRemainingDailyQuota();

  if (quota <= 0) {
    if (interactive) {
      SpreadsheetApp.getUi().alert(
        "Email quota is exhausted. Try again tomorrow.",
      );
    }
    logActivity("SYSTEM", "Reminder Sending Skipped", "Mail quota exhausted");
    return;
  }

  if (sheets.length === 0) {
    if (interactive) {
      SpreadsheetApp.getUi().alert(
        "No tabs are configured yet. Run Quick Setup first.",
      );
    }
    logActivity("SYSTEM", "Reminder Sending Skipped", "No configured tabs");
    return;
  }

  sheets.forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      skipped.push(`${sheetName}: tab not found`);
      return;
    }

    const lastRow = sheet.getLastRow();
    const dataStart = getDataStartRow(sheet);
    if (lastRow < dataStart) return;

    const headers = getHeaders(sheet);
    const data = sheet
      .getRange(dataStart, 1, lastRow - dataStart + 1, headers.length)
      .getValues();
    const deadlineGroups = getTrackedDeadlineGroups(sheet);
    if (!deadlineGroups.length) {
      skipped.push(`${sheetName}: no configured deadline groups were detected`);
      return;
    }

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
    const remarksIndex = findHeaderIndexByName(headers, cfg.HEADERS.REMARKS);

    data.forEach((row, idx) => {
      const actualRow = idx + dataStart;
      const clientName = clientNameIndex ? row[clientNameIndex - 1] : "";
      const staffEmail = staffEmailIndex ? row[staffEmailIndex - 1] : "";
      const partyDetails = getPartyDetailsForReminder(headers, row);
      const serviceType = serviceTypeIndex ? row[serviceTypeIndex - 1] : "";
      const remarks = remarksIndex ? row[remarksIndex - 1] : "";
      const recipients = buildReminderRecipients(staffEmail);
      const ccRecipients = buildReminderCcRecipients(true, recipients);

      if (!recipients.length) {
        skipped.push(`${sheetName} row ${actualRow}: missing valid Staff Email`);
        return;
      }

      deadlineGroups.forEach((group) => {
        const statusValue = group.statusCol ? row[group.statusCol - 1] : "";
        const reminderValue = group.reminderCol ? row[group.reminderCol - 1] : "";
        const remainingValue = group.remainingCol ? row[group.remainingCol - 1] : "";
        const dueDateValue = row[group.dueDateCol - 1];
        const shouldSend =
          statusValue &&
          statusValue.toString().trim().toUpperCase() === "SEND" &&
          reminderValue &&
          reminderValue.toString().trim();

        if (!shouldSend) return;

        const recipientLogKey = buildRecipientLogKey(recipients, ccRecipients);
        if (wasReminderSentToday(sheetName, actualRow, group.type, recipientLogKey)) {
          skipped.push(`${sheetName} row ${actualRow}: ${group.type} already sent today`);
          return;
        }

        const remainingDays =
          remainingValue === "" || remainingValue === null ? null : Number(remainingValue);
        const deadlineSummary =
          remainingDays === null || isNaN(remainingDays)
            ? []
            : [{ label: group.label, remainingDays: remainingDays }];

        try {
          sendReminderEmail({
            recipients: recipients,
            ccRecipients: ccRecipients,
            clientName: clientName || "Unknown Client",
            sellerDonor: partyDetails.sellerDonor || "",
            buyerDonee: partyDetails.buyerDonee || "",
            serviceType: serviceType,
            reminders: [reminderValue.toString()],
            deadlines: deadlineSummary,
            sheetName: sheetName,
            rowNumber: actualRow,
            remarks: remarks,
            subjectPrefix: `[${group.type}]`,
          });
          emailCount++;
          logReminderSend(
            sheetName,
            actualRow,
            group.type,
            recipientLogKey,
            triggerType,
            "SENT",
          );
          logActivity(
            clientName || "Unknown Client",
            "Reminder Email Sent",
            `${group.type} to ${recipientLogKey} (${triggerType})`,
          );
          markReminderCellAsSent(sheet, actualRow, group.reminderCol);
        } catch (e) {
          logReminderSend(
            sheetName,
            actualRow,
            group.type,
            recipientLogKey,
            triggerType,
            "FAILED",
          );
          skipped.push(`${sheetName} row ${actualRow}: ${group.type} failed (${e.message})`);
          logActivity(clientName || "Unknown Client", "Email Failed", e.message);
        }
      });
    });
  });

  const summary = `Sent ${emailCount} reminder email(s)${
    skipped.length ? `\n\nSkipped / Issues:\n- ${skipped.join("\n- ")}` : ""
  }`;
  if (interactive) {
    SpreadsheetApp.getActive().toast(`Sent ${emailCount} reminder email(s)`, "Emails Sent");
  }
  logActivity("SYSTEM", "Reminder Sending Complete", summary.replace(/\n/g, " | "));
  if (interactive) {
    SpreadsheetApp.getUi().alert(
      "Send Reminder Emails",
      summary,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

function sendReminderEmail(data) {
  const subjectPrefix = data.subjectPrefix ? data.subjectPrefix + " " : "";
  const subject = `${subjectPrefix}Reminder: ${data.clientName} - ${data.serviceType || "Project"}`;
  const toRecipients = data.recipients || [];
  const ccRecipients =
    data.ccRecipients || buildReminderCcRecipients(data.includeNotificationEmails, toRecipients);
  const recipientLogKey = buildRecipientLogKey(toRecipients, ccRecipients);
  const escapeHtml = (value) =>
    String(value === null || value === undefined ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  const formatTimingMessage = (remainingDays) =>
    remainingDays < 0
      ? `Overdue by ${Math.abs(remainingDays)} day${Math.abs(remainingDays) === 1 ? "" : "s"}`
      : `${remainingDays} day${remainingDays === 1 ? "" : "s"} remaining`;

  let body = `Dear Team,\n\n`;
  body += `This is an automated reminder regarding the following project:\n\n`;
  body += `Client: ${data.clientName}\n`;
  if (data.sellerDonor) {
    body += `Seller/Donor: ${data.sellerDonor}\n`;
  }
  if (data.buyerDonee) {
    body += `Buyer/Donee: ${data.buyerDonee}\n`;
  }
  body += `Service: ${data.serviceType || "N/A"}\n`;
  body += `Sheet: ${data.sheetName}\n\n`;
  if (data.rowNumber) {
    body += `Row: ${data.rowNumber}\n\n`;
  }
  body += `Sent To:\n`;
  body += `- To: ${toRecipients.join(", ") || "N/A"}\n`;
  if (ccRecipients.length) {
    body += `- CC: ${ccRecipients.join(", ")}\n`;
  }
  body += `\n`;
  body += `ALERT:\n- ${data.reminders.join("\n- ")}\n\n`;

  if (data.deadlines && data.deadlines.length) {
    body += `Deadlines:\n`;
    data.deadlines.forEach((deadline) => {
      body += `- ${deadline.label}: ${formatTimingMessage(deadline.remainingDays)}\n`;
    });
  }

  if (data.revisit && data.revisit.remainingDays !== null) {
    const revisitMessage = formatTimingMessage(data.revisit.remainingDays);
    body += `\nRevisit:\n- Date: ${formatDate(data.revisit.date)}\n- Status: ${data.revisit.status || "Open"}\n- Timing: ${revisitMessage}\n`;
    if (data.revisit.notes) {
      body += `- Notes: ${data.revisit.notes}\n`;
    }
  }

  body += `\nPlease take appropriate action to ensure compliance.\n\n`;
  body += `---\n`;
  body += `Projects Monitoring System\n`;
  body += `This is an automated message. Please do not reply directly to this email.`;

  const reminderItemsHtml = data.reminders
    .map(
      (reminder) =>
        `<li style="margin:0 0 8px 0;">${escapeHtml(reminder)}</li>`,
    )
    .join("");
  const deadlineItemsHtml =
    data.deadlines && data.deadlines.length
      ? data.deadlines
          .map((deadline) => {
            const statusColor = deadline.remainingDays < 0 ? "#b91c1c" : "#1d4ed8";
            return `<tr>
              <td style="padding:10px 0;border-bottom:1px solid #dbeafe;color:#0f172a;font-size:14px;">${escapeHtml(deadline.label)}</td>
              <td style="padding:10px 0;border-bottom:1px solid #dbeafe;color:${statusColor};font-size:14px;font-weight:600;text-align:right;">${escapeHtml(formatTimingMessage(deadline.remainingDays))}</td>
            </tr>`;
          })
          .join("")
      : "";
  const metaRows = [
    { label: "Client", value: data.clientName },
    ...(data.sellerDonor
      ? [{ label: "Seller/Donor", value: data.sellerDonor }]
      : []),
    ...(data.buyerDonee
      ? [{ label: "Buyer/Donee", value: data.buyerDonee }]
      : []),
    { label: "Service", value: data.serviceType || "N/A" },
    { label: "Sheet", value: data.sheetName },
    { label: "Sent To", value: recipientLogKey },
  ];
  if (data.rowNumber) {
    metaRows.push({ label: "Row", value: data.rowNumber });
  }
  const metaHtml = metaRows
    .map(
      (item) => `<tr>
        <td style="padding:6px 0;color:#475569;font-size:13px;width:88px;">${escapeHtml(item.label)}</td>
        <td style="padding:6px 0;color:#0f172a;font-size:13px;font-weight:600;">${escapeHtml(item.value)}</td>
      </tr>`,
    )
    .join("");
  const recipientsHtml = `<div style="margin-top:20px;padding:18px;border:1px solid #dbeafe;border-radius:12px;background:#f8fbff;">
          <div style="font-size:13px;font-weight:700;letter-spacing:0.04em;text-transform:uppercase;color:#2563eb;margin-bottom:10px;">Sent To</div>
          <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%;border-collapse:collapse;">
            <tr>
              <td style="padding:5px 0;color:#475569;font-size:13px;width:88px;">To</td>
              <td style="padding:5px 0;color:#0f172a;font-size:13px;font-weight:600;">${escapeHtml(toRecipients.join(", ") || "N/A")}</td>
            </tr>
            ${
              ccRecipients.length
                ? `<tr>
                    <td style="padding:5px 0;color:#475569;font-size:13px;">CC</td>
                    <td style="padding:5px 0;color:#0f172a;font-size:13px;font-weight:600;">${escapeHtml(ccRecipients.join(", "))}</td>
                  </tr>`
                : ""
            }
          </table>
        </div>`;
  const revisitHtml =
    data.revisit && data.revisit.remainingDays !== null
      ? `<div style="margin-top:20px;padding:18px;border:1px solid #dbeafe;border-radius:12px;background:#f8fbff;">
          <div style="font-size:13px;font-weight:700;letter-spacing:0.04em;text-transform:uppercase;color:#2563eb;margin-bottom:10px;">Revisit</div>
          <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%;border-collapse:collapse;">
            <tr>
              <td style="padding:5px 0;color:#475569;font-size:13px;width:88px;">Date</td>
              <td style="padding:5px 0;color:#0f172a;font-size:13px;font-weight:600;">${escapeHtml(formatDate(data.revisit.date))}</td>
            </tr>
            <tr>
              <td style="padding:5px 0;color:#475569;font-size:13px;">Status</td>
              <td style="padding:5px 0;color:#0f172a;font-size:13px;font-weight:600;">${escapeHtml(data.revisit.status || "Open")}</td>
            </tr>
            <tr>
              <td style="padding:5px 0;color:#475569;font-size:13px;">Timing</td>
              <td style="padding:5px 0;color:#1d4ed8;font-size:13px;font-weight:600;">${escapeHtml(formatTimingMessage(data.revisit.remainingDays))}</td>
            </tr>
            ${
              data.revisit.notes
                ? `<tr>
                    <td style="padding:5px 0;color:#475569;font-size:13px;">Notes</td>
                    <td style="padding:5px 0;color:#0f172a;font-size:13px;">${escapeHtml(data.revisit.notes)}</td>
                  </tr>`
                : ""
            }
          </table>
        </div>`
      : "";
  const htmlBody = `<!DOCTYPE html>
<html>
  <body style="margin:0;padding:24px;background:#eff6ff;font-family:Arial,sans-serif;color:#0f172a;">
    <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width:640px;margin:0 auto;border-collapse:collapse;">
      <tr>
        <td style="padding:0;">
          <div style="background:#ffffff;border:1px solid #dbeafe;border-radius:18px;overflow:hidden;">
            <div style="padding:24px 28px;background:#1d4ed8;color:#ffffff;">
              <div style="font-size:12px;letter-spacing:0.08em;text-transform:uppercase;opacity:0.9;">Projects Monitoring System</div>
              <div style="margin-top:8px;font-size:24px;font-weight:700;line-height:1.3;">Reminder Notice</div>
              <div style="margin-top:8px;font-size:14px;line-height:1.6;color:#dbeafe;">Please review the project details below and take the needed action.</div>
            </div>
            <div style="padding:24px 28px;">
              <div style="padding:18px;border:1px solid #dbeafe;border-radius:12px;background:#f8fbff;">
                <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%;border-collapse:collapse;">${metaHtml}</table>
              </div>
              ${recipientsHtml}
              <div style="margin-top:20px;">
                <div style="font-size:13px;font-weight:700;letter-spacing:0.04em;text-transform:uppercase;color:#2563eb;margin-bottom:10px;">Reminder</div>
                <ul style="margin:0;padding-left:20px;color:#0f172a;font-size:14px;line-height:1.6;">${reminderItemsHtml}</ul>
              </div>
              ${
                deadlineItemsHtml
                  ? `<div style="margin-top:20px;">
                      <div style="font-size:13px;font-weight:700;letter-spacing:0.04em;text-transform:uppercase;color:#2563eb;margin-bottom:10px;">Deadlines</div>
                      <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%;border-collapse:collapse;">${deadlineItemsHtml}</table>
                    </div>`
                  : ""
              }
              ${revisitHtml}
              <div style="margin-top:24px;font-size:14px;line-height:1.6;color:#334155;">This is an automated reminder. Please do not reply directly to this email.</div>
            </div>
          </div>
        </td>
      </tr>
    </table>
  </body>
</html>`;

  const mailOptions = {
    to: toRecipients.join(","),
    subject: subject,
    body: body,
    htmlBody: htmlBody,
    name: "Projects Monitoring System",
  };
  if (ccRecipients.length) {
    mailOptions.cc = ccRecipients.join(",");
  }

  MailApp.sendEmail(mailOptions);
}

function testReminderForOneRow() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();

  if (!getProjectSheets().includes(sheetName)) {
    ui.alert(
      "Test Reminder for One Row",
      "Open one of the configured automation tabs first, then run this action again.",
      ui.ButtonSet.OK,
    );
    return;
  }

  const currentRow = sheet.getActiveRange()
    ? sheet.getActiveRange().getRow()
    : getDataStartRow(sheet);
  const response = ui.prompt(
    "Test Reminder for One Row",
    `Active tab: ${sheetName}\nCurrent selected row: ${currentRow}\nData starts at row: ${getDataStartRow(sheet)}\n\nEnter the row number to test:`,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const rowNumber = parseInt(response.getResponseText().trim(), 10);
  if (isNaN(rowNumber) || rowNumber < 1) {
    ui.alert("Invalid row number. Please enter a positive row number.");
    return;
  }

  const payload = getReminderEmailPayloadForRow(sheet, rowNumber);
  if (!payload.data) {
    ui.alert("Test Reminder for One Row", payload.error, ui.ButtonSet.OK);
    return;
  }

  try {
    sendReminderEmail({
      recipients: payload.data.recipients,
      clientName: payload.data.clientName,
      sellerDonor: payload.data.sellerDonor,
      buyerDonee: payload.data.buyerDonee,
      serviceType: payload.data.serviceType,
      reminders: payload.data.reminders,
      deadlines: payload.data.deadlines,
      sheetName: payload.data.sheetName,
      rowNumber: payload.data.rowNumber,
      revisit: payload.data.revisit,
      subjectPrefix: "[TEST ROW]",
    });
  } catch (e) {
    ui.alert("Test Reminder for One Row", `Email failed: ${e.message}`, ui.ButtonSet.OK);
    return;
  }

  logActivity(
    payload.data.clientName,
    "Test Reminder Email Sent",
    `Row ${rowNumber} in ${sheetName} to ${payload.data.recipients.join(", ")}`,
  );
  ui.alert(
    "Test Reminder for One Row",
    `Test reminder sent for row ${rowNumber}.\n\nRecipients: ${payload.data.recipients.join(", ")}`,
    ui.ButtonSet.OK,
  );
}

function sendSampleReminderEmail() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Send Test Email to a Chosen Address",
    "Enter the email address that should receive the sample reminder email:",
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const email = response.getResponseText().trim().toLowerCase();
  if (!isValidEmailAddress(email)) {
    ui.alert("Invalid email address. Please include a valid email.");
    return;
  }

  try {
    sendReminderEmail({
      recipients: [email],
      clientName: "Sample Client",
      sellerDonor: "Sample Seller / Donor",
      buyerDonee: "Sample Buyer / Donee",
      serviceType: "Sample Reminder",
      reminders: [
        "This is a sample reminder email used to confirm that reminder sending is working.",
      ],
      deadlines: [
        { label: "DST Due Date", remainingDays: 3 },
        { label: "CGT / DOD Due Date", remainingDays: 10 },
      ],
      sheetName: "Settings Test",
      rowNumber: null,
      subjectPrefix: "[TEST]",
      includeNotificationEmails: false,
    });
  } catch (e) {
    ui.alert(
      "Send Test Email to a Chosen Address",
      `Email failed: ${e.message}`,
      ui.ButtonSet.OK,
    );
    return;
  }

  logActivity("SYSTEM", "Sample Reminder Email Sent", `To: ${email}`);
  ui.alert(
    "Send Test Email to a Chosen Address",
    `Sample reminder email sent to ${email}.`,
    ui.ButtonSet.OK,
  );
}

// ============================================================================
// AUTOMATION TRIGGERS
// ============================================================================

function onEdit(e) {
  if (!e) return;
  // Formula-driven restart: Google Sheet formulas are the source of truth.
  // Apps Script intentionally does not overwrite due-date, remaining-days,
  // status, or reminder columns on edit.
}

function runDailyCheck() {
  const startTime = new Date();
  Logger.log("Starting daily check at " + startTime);

  sendReminderEmailsByTriggerType("automatic");

  const endTime = new Date();
  const duration = (endTime - startTime) / 1000;

  logActivity("SYSTEM", "Daily Check Complete", `Duration: ${duration}s`);
  Logger.log("Daily check completed in " + duration + " seconds");
}

function showAutomaticSendingStatus() {
  const ui = SpreadsheetApp.getUi();
  const status = getAutomaticSendingStatus();
  const dstCgtTabs = getConfiguredDstCgtTabs();
  const transferTabs = getConfiguredTransferTaxTabs();
  ui.alert(
    "Automatic Sending Status",
    `Status: ${status.enabled ? "ACTIVE" : "NOT ACTIVE"}\nDaily run time: ${status.timeLabel}\nInstalled triggers: ${status.count}\nDST / CGT Tabs: ${dstCgtTabs.length ? dstCgtTabs.join(", ") : "(none selected)"}\nTransfer Tax Tabs: ${transferTabs.length ? transferTabs.join(", ") : "(none selected)"}\n\n${status.enabled ? "Automatic Sending is ready and will scan the configured tabs using the sheet formulas." : "Automatic Sending is not enabled yet. Use Start Automatic Sending to install the daily trigger."}`,
    ui.ButtonSet.OK,
  );
}

function startAutomaticSending() {
  createDailyTrigger(false);
}

function stopAutomaticSending() {
  removeTriggers();
}

function syncDailyTrigger(showToast) {
  const triggers = getDailyCheckTriggers();
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
      "Start Automatic Sending",
      "Automatic Sending is already active. Replace the current daily trigger with the saved schedule?",
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
  if (!silent) showAutomaticSendingStatus();
}

function removeTriggers() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Stop Automatic Sending",
    "Stop Automatic Sending and remove all daily triggers?",
    ui.ButtonSet.YES_NO,
  );

  if (response !== ui.Button.YES) return;

  const triggers = getDailyCheckTriggers();
  triggers.forEach((t) => ScriptApp.deleteTrigger(t));

  SpreadsheetApp.getActive().toast(
    `Removed ${triggers.length} triggers`,
    "Triggers Removed",
  );
  logActivity("SYSTEM", "Triggers Removed", `Count: ${triggers.length}`);
  ui.alert(
    "Stop Automatic Sending",
    triggers.length > 0
      ? "Automatic Sending has been stopped and the daily triggers were removed."
      : "No Automatic Sending trigger was installed.",
    ui.ButtonSet.OK,
  );
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
  const overallStatus = results.issues.length
    ? "ACTION REQUIRED"
    : results.warnings.length
      ? "READY WITH WARNINGS"
      : "READY";
  report += `Overall Status: ${overallStatus}\n\n`;

  report += "Checks:\n";
  if (results.notes.length === 0) report += "  (none)\n";
  else results.notes.forEach((note) => (report += `  - ${note}\n`));

  report += "\nIssues:\n";
  if (results.issues.length === 0) report += "  None\n";
  else {
    results.issues.forEach((issue) => {
      report += `  - ${issue.message}\n`;
      report += `    ${issue.fix}\n`;
    });
  }

  report += "\nWarnings:\n";
  if (results.warnings.length === 0) report += "  None\n";
  else {
    results.warnings.forEach((warning) => {
      report += `  - ${warning.message}\n`;
      report += `    ${warning.fix}\n`;
    });
  }

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
