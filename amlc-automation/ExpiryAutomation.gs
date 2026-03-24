const CONFIG_SHEET_NAME = 'Automation_Config';
const MENU_NAME = 'Expiry Automation';
const DAILY_TRIGGER_FUNCTION = 'runDailyExpiryCheck';
const DAILY_TRIGGER_HOUR = 8;
const REQUIRED_HEADERS = [
  'Registration',
  'Registration Type',
  'Control No.',
  'Issue Date',
  'Expiry Date',
  'Days Remaining Before Expiry',
  'Status',
  'GD Link',
  'Remarks',
];
const CONFIG_KEYS = {
  SELECTED_TABS: 'selected_tabs',
  RECIPIENT_EMAILS: 'recipient_emails',
  LAST_SETUP_AT: 'last_setup_at',
  LAST_RUN_AT: 'last_run_at',
  LAST_SENT_AT: 'last_sent_at',
  LAST_SENT_COUNT: 'last_sent_count',
  LAST_ERROR: 'last_error',
  SENT_LOG: 'sent_log',
};

function onOpen() {
  syncConfiguredTabs_();
  SpreadsheetApp.getUi()
    .createMenu(MENU_NAME)
    .addItem('Automate This sheet', 'setupAutomation')
    .addItem('Remove automated tabs', 'removeAutomatedTabs')
    .addItem('List of emails for sending', 'setupRecipientEmails')
    .addItem('Remove emails for automation', 'removeRecipientEmails')
    .addItem('TESTING', 'sendTestEmail')
    .addToUi();
}

function setupAutomation() {
  const ui = SpreadsheetApp.getUi();
  const availableTabs = getAvailableTabs_();
  ensureConfigSheet_();
  const currentSelection = getSelectedTabs_();

  if (availableTabs.length === 0) {
    ui.alert('No data tabs found. Create at least one sheet tab with your registration data first.');
    return;
  }

  const response = ui.prompt(
    'Automate This sheet',
    'Enter the tab numbers to automate, separated by commas.\n\nAvailable tabs:\n' +
      buildNumberedList_(availableTabs) +
      (currentSelection.length
        ? '\n\nCurrently automated tabs:\n' + buildSelectedNumberedList_(availableTabs, currentSelection)
        : '\n\nCurrently automated tabs:\n(none)'),
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const selectionResult = parseTabSelectionByNumber_(response.getResponseText(), availableTabs);
  if (selectionResult.selectedTabs.length === 0) {
    ui.alert('Please enter at least one valid tab number.');
    return;
  }

  if (selectionResult.invalidEntries.length > 0) {
    ui.alert('These tab numbers are invalid: ' + selectionResult.invalidEntries.join(', '));
    return;
  }

  const selectedTabs = selectionResult.selectedTabs;
  setConfigValue_(CONFIG_KEYS.SELECTED_TABS, JSON.stringify(selectedTabs));
  setConfigValue_(CONFIG_KEYS.LAST_SETUP_AT, new Date().toISOString());
  clearConfigValue_(CONFIG_KEYS.LAST_ERROR);
  refreshDailyTrigger_();

  ui.alert(
    'Automation saved.\n\nSelected tabs:\n' +
      selectedTabs.join(', ') +
      '\n\nDaily trigger time: 8:00 AM'
  );
}

function removeAutomatedTabs() {
  const ui = SpreadsheetApp.getUi();
  const availableTabs = getAvailableTabs_();
  const currentSelection = getSelectedTabs_();

  if (currentSelection.length === 0) {
    ui.alert('There are no automated tabs to remove.');
    return;
  }

  const response = ui.prompt(
    'Remove automated tabs',
    'Enter the numbers of the automated tabs you want to remove, separated by commas.\n\nAutomated tabs:\n' +
      buildSelectedNumberedList_(availableTabs, currentSelection),
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const removalResult = parseTabSelectionByNumber_(response.getResponseText(), currentSelection);
  if (removalResult.selectedTabs.length === 0) {
    ui.alert('Please enter at least one valid tab number to remove.');
    return;
  }

  if (removalResult.invalidEntries.length > 0) {
    ui.alert('These tab numbers are invalid: ' + removalResult.invalidEntries.join(', '));
    return;
  }

  const tabsToRemove = removalResult.selectedTabs;
  const nextSelection = currentSelection.filter(function (tab) {
    return tabsToRemove.indexOf(tab) === -1;
  });

  setConfigValue_(CONFIG_KEYS.SELECTED_TABS, JSON.stringify(nextSelection));
  setConfigValue_(CONFIG_KEYS.LAST_SETUP_AT, new Date().toISOString());
  clearConfigValue_(CONFIG_KEYS.LAST_ERROR);
  if (nextSelection.length === 0) {
    deleteDailyTrigger_();
  }

  ui.alert(
    nextSelection.length > 0
      ? 'Removed tabs:\n' + tabsToRemove.join(', ') + '\n\nStill automated:\n' + nextSelection.join(', ')
      : 'Removed tabs:\n' + tabsToRemove.join(', ') + '\n\nNo tabs are automated now.'
  );
}

function setupRecipientEmails() {
  const ui = SpreadsheetApp.getUi();
  const currentEmails = getRecipientEmails_();
  const response = ui.prompt(
    'List of emails for sending',
    'Enter recipient emails separated by commas.\n\nCurrent emails:\n' +
      (currentEmails.length ? currentEmails.join(', ') : '(none)'),
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const emails = parseEmails_(response.getResponseText());
  if (emails.length === 0) {
    ui.alert('Please enter at least one valid email address.');
    return;
  }

  setConfigValue_(CONFIG_KEYS.RECIPIENT_EMAILS, JSON.stringify(emails));
  clearConfigValue_(CONFIG_KEYS.LAST_ERROR);
  ui.alert('Saved recipient emails:\n' + emails.join(', '));
}

function removeRecipientEmails() {
  const ui = SpreadsheetApp.getUi();
  const currentEmails = getRecipientEmails_();

  if (currentEmails.length === 0) {
    ui.alert('There are no saved recipient emails to remove.');
    return;
  }

  const response = ui.prompt(
    'Remove emails for automation',
    'Enter the numbers of the recipient emails you want to remove, separated by commas.\n\nSaved emails:\n' +
      buildNumberedList_(currentEmails),
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const removalResult = parseTabSelectionByNumber_(response.getResponseText(), currentEmails);
  if (removalResult.selectedTabs.length === 0) {
    ui.alert('Please enter at least one valid email number to remove.');
    return;
  }

  if (removalResult.invalidEntries.length > 0) {
    ui.alert('These email numbers are invalid: ' + removalResult.invalidEntries.join(', '));
    return;
  }

  const emailsToRemove = removalResult.selectedTabs;
  const nextEmails = currentEmails.filter(function (email) {
    return emailsToRemove.indexOf(email) === -1;
  });

  setConfigValue_(CONFIG_KEYS.RECIPIENT_EMAILS, JSON.stringify(nextEmails));
  clearConfigValue_(CONFIG_KEYS.LAST_ERROR);

  ui.alert(
    nextEmails.length > 0
      ? 'Removed emails:\n' + emailsToRemove.join(', ') + '\n\nRemaining emails:\n' + nextEmails.join(', ')
      : 'Removed emails:\n' + emailsToRemove.join(', ') + '\n\nNo recipient emails are saved now.'
  );
}

function sendTestEmail() {
  const ui = SpreadsheetApp.getUi();

  try {
    const recipients = getRecipientEmails_();
    if (recipients.length === 0) {
      ui.alert('Add recipient emails first using "List of emails for sending".');
      return;
    }

    const payload = getTestPayload_();
    sendNotificationEmail_(payload, recipients, true);
    ui.alert('Test email sent to:\n' + recipients.join(', '));
  } catch (error) {
    recordError_(error);
    ui.alert('Unable to send test email: ' + error.message);
  }
}

function runDailyExpiryCheck() {
  try {
    const selectedTabs = getSelectedTabs_();
    const recipients = getRecipientEmails_();
    const timezone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    const todayKey = formatDateKey_(new Date(), timezone);
    const sentLog = getSentLog_();
    const sentToday = sentLog[todayKey] || [];

    if (selectedTabs.length === 0) {
      throw new Error('No tabs configured. Use "Automate This sheet" first.');
    }

    if (recipients.length === 0) {
      throw new Error('No recipient emails configured. Use "List of emails for sending" first.');
    }

    let sentCount = 0;
    selectedTabs.forEach(function (sheetName) {
      const rows = getExpiringRowsForSheet_(sheetName, todayKey, timezone);
      rows.forEach(function (rowData) {
        const rowKey = buildRowKey_(rowData);
        if (sentToday.indexOf(rowKey) !== -1) {
          return;
        }

        sendNotificationEmail_(rowData, recipients, false);
        sentToday.push(rowKey);
        sentCount += 1;
      });
    });

    sentLog[todayKey] = sentToday;
    setSentLog_(sentLog);
    setConfigValue_(CONFIG_KEYS.LAST_RUN_AT, new Date().toISOString());
    setConfigValue_(CONFIG_KEYS.LAST_SENT_AT, sentCount > 0 ? new Date().toISOString() : '');
    setConfigValue_(CONFIG_KEYS.LAST_SENT_COUNT, String(sentCount));
    clearConfigValue_(CONFIG_KEYS.LAST_ERROR);
  } catch (error) {
    recordError_(error);
    throw error;
  }
}

function getAvailableTabs_() {
  return SpreadsheetApp.getActive()
    .getSheets()
    .map(function (sheet) {
      return sheet.getName();
    })
    .filter(function (sheetName) {
      return sheetName !== CONFIG_SHEET_NAME;
    });
}

function getSelectedTabs_() {
  return syncConfiguredTabs_();
}

function getRecipientEmails_() {
  return readJsonArrayConfig_(CONFIG_KEYS.RECIPIENT_EMAILS);
}

function getExpiringRowsForSheet_(sheetName, todayKey, timezone) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Configured tab not found: ' + sheetName);
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return [];
  }

  const headerMap = buildHeaderMap_(values[0]);
  validateHeaders_(headerMap, sheetName);

  const expiringRows = [];
  for (let rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    const row = values[rowIndex];
    if (isBlankRow_(row)) {
      continue;
    }

    const expiryValue = getCellByHeader_(row, headerMap, 'Expiry Date');
    if (!expiryValue) {
      continue;
    }

    const expiryDate = normalizeToDate_(expiryValue);
    if (!expiryDate) {
      continue;
    }

    if (formatDateKey_(expiryDate, timezone) !== todayKey) {
      continue;
    }

    expiringRows.push(buildRowPayload_(sheetName, row, headerMap, timezone));
  }

  return expiringRows;
}

function buildRowPayload_(sheetName, row, headerMap, timezone) {
  return {
    sheetName: sheetName,
    registration: stringifyCell_(getCellByHeader_(row, headerMap, 'Registration')),
    registrationType: stringifyCell_(getCellByHeader_(row, headerMap, 'Registration Type')),
    controlNo: stringifyCell_(getCellByHeader_(row, headerMap, 'Control No.')),
    issueDate: formatCellForEmail_(getCellByHeader_(row, headerMap, 'Issue Date'), timezone),
    expiryDate: formatCellForEmail_(getCellByHeader_(row, headerMap, 'Expiry Date'), timezone),
    daysRemaining: stringifyCell_(getCellByHeader_(row, headerMap, 'Days Remaining Before Expiry')),
    status: stringifyCell_(getCellByHeader_(row, headerMap, 'Status')),
    gdLink: stringifyCell_(getCellByHeader_(row, headerMap, 'GD Link')),
    remarks: stringifyCell_(getCellByHeader_(row, headerMap, 'Remarks')),
  };
}

function getTestPayload_() {
  const selectedTabs = getSelectedTabs_();
  const timezone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();

  for (let index = 0; index < selectedTabs.length; index += 1) {
    const sheetName = selectedTabs[index];
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) {
      continue;
    }

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) {
      continue;
    }

    const headerMap = buildHeaderMap_(values[0]);
    const firstDataRow = values[1];
    if (!isBlankRow_(firstDataRow)) {
      const payload = buildRowPayload_(sheetName, firstDataRow, headerMap, timezone);
      payload.remarks = payload.remarks || 'This is a test email using the first available row in your configured tab.';
      if (!payload.expiryDate) {
        payload.expiryDate = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd');
      }
      return payload;
    }
  }

  return {
    sheetName: 'TEST',
    registration: 'Sample Registration',
    registrationType: 'Sample Type',
    controlNo: 'CTRL-TEST-001',
    issueDate: '2026-03-24',
    expiryDate: '2026-03-24',
    daysRemaining: '0',
    status: 'Expiring Today',
    gdLink: 'https://docs.google.com/',
    remarks: 'This is a test notification from the Google Sheets expiry automation.',
  };
}

function sendNotificationEmail_(rowData, recipients, isTest) {
  const subjectPrefix = isTest ? '[TEST] ' : '';
  const subject = subjectPrefix + 'Expiry Alert: ' +
    [rowData.registration, rowData.registrationType, rowData.controlNo]
      .filter(Boolean)
      .join(' | ');

  const bodyLines = [
    rowData.remarks || 'No remarks provided.',
    '',
    'Sheet Tab: ' + (rowData.sheetName || ''),
    'Registration: ' + (rowData.registration || ''),
    'Registration Type: ' + (rowData.registrationType || ''),
    'Control No.: ' + (rowData.controlNo || ''),
    'Issue Date: ' + (rowData.issueDate || ''),
    'Expiry Date: ' + (rowData.expiryDate || ''),
    'Days Remaining Before Expiry: ' + (rowData.daysRemaining || ''),
    'Status: ' + (rowData.status || ''),
    'GD Link: ' + (rowData.gdLink || ''),
  ];
  const htmlBody = buildEmailHtml_(rowData, isTest);

  MailApp.sendEmail({
    to: recipients.join(','),
    subject: subject,
    body: bodyLines.join('\n'),
    htmlBody: htmlBody,
  });
}

function refreshDailyTrigger_() {
  deleteDailyTrigger_();

  ScriptApp.newTrigger(DAILY_TRIGGER_FUNCTION)
    .timeBased()
    .everyDays(1)
    .atHour(DAILY_TRIGGER_HOUR)
    .create();
}

function deleteDailyTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === DAILY_TRIGGER_FUNCTION) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function syncConfiguredTabs_() {
  ensureConfigSheet_();
  const availableTabs = getAvailableTabs_();
  const storedTabs = readJsonArrayConfig_(CONFIG_KEYS.SELECTED_TABS);
  const validTabs = storedTabs.filter(function (tab) {
    return availableTabs.indexOf(tab) !== -1;
  });

  if (storedTabs.length !== validTabs.length) {
    setConfigValue_(CONFIG_KEYS.SELECTED_TABS, JSON.stringify(validTabs));
  }

  return validTabs;
}

function ensureConfigSheet_() {
  const spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG_SHEET_NAME);
    sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function readJsonArrayConfig_(key) {
  const rawValue = getConfigValue_(key);
  if (!rawValue) {
    return [];
  }

  try {
    const parsed = JSON.parse(rawValue);
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    return [];
  }
}

function getSentLog_() {
  const rawValue = getConfigValue_(CONFIG_KEYS.SENT_LOG);
  if (!rawValue) {
    return {};
  }

  try {
    const parsed = JSON.parse(rawValue);
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch (error) {
    return {};
  }
}

function setSentLog_(sentLog) {
  const todayKey = formatDateKey_(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone());
  const trimmedLog = {};
  trimmedLog[todayKey] = Array.isArray(sentLog[todayKey]) ? sentLog[todayKey] : [];
  setConfigValue_(CONFIG_KEYS.SENT_LOG, JSON.stringify(trimmedLog));
}

function getConfigValue_(key) {
  const sheet = ensureConfigSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return '';
  }

  const entries = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let index = 0; index < entries.length; index += 1) {
    if (entries[index][0] === key) {
      return entries[index][1];
    }
  }

  return '';
}

function setConfigValue_(key, value) {
  const sheet = ensureConfigSheet_();
  const lastRow = sheet.getLastRow();
  const entries = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 2).getValues() : [];

  for (let index = 0; index < entries.length; index += 1) {
    if (entries[index][0] === key) {
      sheet.getRange(index + 2, 2).setValue(value);
      return;
    }
  }

  sheet.appendRow([key, value]);
}

function clearConfigValue_(key) {
  setConfigValue_(key, '');
}

function buildHeaderMap_(headerRow) {
  const headerMap = {};
  headerRow.forEach(function (header, index) {
    if (!header) {
      return;
    }
    headerMap[normalizeHeader_(header)] = index;
  });
  return headerMap;
}

function validateHeaders_(headerMap, sheetName) {
  const missingHeaders = REQUIRED_HEADERS.filter(function (header) {
    return typeof headerMap[normalizeHeader_(header)] === 'undefined';
  });

  if (missingHeaders.length > 0) {
    throw new Error('Sheet "' + sheetName + '" is missing required headers: ' + missingHeaders.join(', '));
  }
}

function getCellByHeader_(row, headerMap, headerName) {
  const index = headerMap[normalizeHeader_(headerName)];
  return typeof index === 'undefined' ? '' : row[index];
}

function normalizeHeader_(header) {
  return String(header).trim().toLowerCase().replace(/\s+/g, ' ');
}

function normalizeToDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value;
  }

  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function formatDateKey_(date, timezone) {
  return Utilities.formatDate(date, timezone, 'yyyy-MM-dd');
}

function isBlankRow_(row) {
  return row.every(function (cell) {
    return cell === '' || cell === null;
  });
}

function stringifyCell_(value) {
  if (value === null || typeof value === 'undefined') {
    return '';
  }
  return String(value);
}

function formatCellForEmail_(value, timezone) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, timezone, 'yyyy-MM-dd');
  }

  return stringifyCell_(value);
}

function parseEmails_(input) {
  return String(input || '')
    .split(/[,\n;]/)
    .map(function (email) {
      return email.trim();
    })
    .filter(function (email) {
      return email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
    })
    .filter(function (email, index, arr) {
      return arr.indexOf(email) === index;
    });
}

function parseCsvList_(input) {
  return String(input || '')
    .split(',')
    .map(function (value) {
      return value.trim();
    })
    .filter(function (value, index, arr) {
      return value && arr.indexOf(value) === index;
    });
}

function parseTabSelectionByNumber_(input, tabList) {
  const invalidEntries = [];
  const selectedTabs = [];

  String(input || '')
    .split(',')
    .map(function (value) {
      return value.trim();
    })
    .filter(function (value) {
      return value;
    })
    .forEach(function (value) {
      if (!/^\d+$/.test(value)) {
        invalidEntries.push(value);
        return;
      }

      const index = Number(value) - 1;
      if (index < 0 || index >= tabList.length) {
        invalidEntries.push(value);
        return;
      }

      const tabName = tabList[index];
      if (selectedTabs.indexOf(tabName) === -1) {
        selectedTabs.push(tabName);
      }
    });

  return {
    selectedTabs: selectedTabs,
    invalidEntries: invalidEntries,
  };
}

function buildNumberedList_(tabList) {
  return tabList.map(function (tabName, index) {
    return (index + 1) + '. ' + tabName;
  }).join('\n');
}

function buildSelectedNumberedList_(availableTabs, selectedTabs) {
  return selectedTabs.map(function (tabName, index) {
    const tabNumber = availableTabs.indexOf(tabName);
    const label = tabNumber === -1 ? '(deleted)' : String(tabNumber + 1);
    return label + '. ' + tabName;
  }).join('\n');
}

function buildRowKey_(rowData) {
  return [
    rowData.sheetName || '',
    rowData.registration || '',
    rowData.controlNo || '',
    rowData.expiryDate || '',
  ].join('||');
}

function recordError_(error) {
  setConfigValue_(CONFIG_KEYS.LAST_ERROR, error && error.message ? error.message : String(error));
}

function buildEmailHtml_(rowData, isTest) {
  const badgeLabel = isTest ? 'TEST EMAIL' : 'EXPIRY NOTICE';
  const gdLinkHtml = rowData.gdLink
    ? '<a href="' + escapeHtml_(rowData.gdLink) + '" style="color:#0f766e;text-decoration:none;font-weight:600;">Open GD Link</a>'
    : '<span style="color:#6b7280;">No GD Link provided</span>';

  return (
    '<div style="margin:0;padding:24px;background:#f3f7f6;font-family:Arial,sans-serif;color:#16302b;">' +
      '<div style="max-width:720px;margin:0 auto;background:#ffffff;border:1px solid #d9e5e2;border-radius:18px;overflow:hidden;">' +
        '<div style="padding:24px 28px;background:linear-gradient(135deg,#0f766e,#155e75);color:#ffffff;">' +
          '<div style="display:inline-block;padding:6px 10px;border-radius:999px;background:rgba(255,255,255,0.18);font-size:11px;font-weight:700;letter-spacing:0.08em;">' +
            escapeHtml_(badgeLabel) +
          '</div>' +
          '<h1 style="margin:14px 0 8px;font-size:26px;line-height:1.2;">Registration Expiry Alert</h1>' +
          '<p style="margin:0;font-size:14px;line-height:1.6;opacity:0.92;">' +
            'A registration record has reached its expiry date and needs attention today.' +
          '</p>' +
        '</div>' +
        '<div style="padding:28px;">' +
          '<div style="margin-bottom:22px;padding:18px;border-radius:14px;background:#f8fafc;border-left:5px solid #0f766e;">' +
            '<div style="font-size:12px;font-weight:700;letter-spacing:0.08em;color:#0f766e;margin-bottom:8px;">REMARKS</div>' +
            '<div style="font-size:15px;line-height:1.7;color:#1f2937;">' +
              nl2brHtml_(rowData.remarks || 'No remarks provided.') +
            '</div>' +
          '</div>' +
          '<table style="width:100%;border-collapse:collapse;font-size:14px;">' +
            buildEmailRowHtml_('Sheet Tab', rowData.sheetName) +
            buildEmailRowHtml_('Registration', rowData.registration) +
            buildEmailRowHtml_('Registration Type', rowData.registrationType) +
            buildEmailRowHtml_('Control No.', rowData.controlNo) +
            buildEmailRowHtml_('Issue Date', rowData.issueDate) +
            buildEmailRowHtml_('Expiry Date', rowData.expiryDate) +
            buildEmailRowHtml_('Days Remaining Before Expiry', rowData.daysRemaining) +
            buildEmailRowHtml_('Status', rowData.status) +
            '<tr>' +
              '<td style="padding:12px 0;border-bottom:1px solid #e5e7eb;width:220px;font-weight:700;color:#374151;vertical-align:top;">GD Link</td>' +
              '<td style="padding:12px 0;border-bottom:1px solid #e5e7eb;color:#111827;">' + gdLinkHtml + '</td>' +
            '</tr>' +
          '</table>' +
        '</div>' +
        '<div style="padding:18px 28px;background:#f8fafc;border-top:1px solid #e5e7eb;font-size:12px;line-height:1.6;color:#6b7280;">' +
          'This email was generated automatically by your Google Sheets expiry automation.' +
        '</div>' +
      '</div>' +
    '</div>'
  );
}

function buildEmailRowHtml_(label, value) {
  return (
    '<tr>' +
      '<td style="padding:12px 0;border-bottom:1px solid #e5e7eb;width:220px;font-weight:700;color:#374151;vertical-align:top;">' +
        escapeHtml_(label) +
      '</td>' +
      '<td style="padding:12px 0;border-bottom:1px solid #e5e7eb;color:#111827;">' +
        escapeHtml_(value || '') +
      '</td>' +
    '</tr>'
  );
}

function nl2brHtml_(value) {
  return escapeHtml_(value).replace(/\n/g, '<br>');
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
