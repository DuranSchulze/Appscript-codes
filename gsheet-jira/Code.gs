var SPREADSHEET_ID = '1K0c-r44qLJz4bOPrbRsCW1sOvHOYZ9SOvaBNQKkSa9o';
var SHEET_NAMES = [
  'List of Task',
  'Completed Tasks/For Monitoring',
  'Cancelled'
];
var BUSINESS_HEADERS = [
  'COMPANY',
  'DESCRIPTION',
  'STATUS',
  'REMARKS',
  'NOTES',
  'Links',
  'Progress'
];
var INTERNAL_ID_HEADER = '_TASK_ID';
var STATUS_COMPLETED_SHEET = 'Completed Tasks/For Monitoring';
var STATUS_CANCELLED_SHEET = 'Cancelled';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Project List')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Project List')
    .addItem('Open Project List', 'openTaskBoard')
    .addItem('Initialize Board Columns', 'initializeBoard')
    .addToUi();
}

function openTaskBoard() {
  var url = getWebAppUrl_();
  if (!url) {
    SpreadsheetApp.getUi().alert(
      'Deploy the script as a web app first, then set the WEB_APP_URL script property or re-open after deployment.'
    );
    return;
  }

  var html = HtmlService.createHtmlOutput(
    '<script>window.open("' + sanitizeForHtmlAttribute_(url) + '", "_blank");google.script.host.close();</script>'
  )
    .setWidth(120)
    .setHeight(40);

  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Project List');
}

function initializeBoard() {
  ensureSpreadsheetStructure_();
}

function getBoardData() {
  ensureSpreadsheetStructure_();

  var spreadsheet = getSpreadsheet_();
  var metadata = getHeaderMetadata_(spreadsheet);
  var lanes = SHEET_NAMES.map(function(sheetName) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    var tasks = getTasksForSheet_(sheet, metadata);

    return {
      id: toLaneId_(sheetName),
      title: sheetName,
      sheetName: sheetName,
      tasks: tasks
    };
  });

  return {
    spreadsheetId: spreadsheet.getId(),
    spreadsheetUrl: spreadsheet.getUrl(),
    lanes: lanes,
    statusOptions: getStatusOptions_(spreadsheet, metadata)
  };
}

function moveTask(taskId, fromSheet, toSheet, toIndex) {
  if (SHEET_NAMES.indexOf(fromSheet) === -1 || SHEET_NAMES.indexOf(toSheet) === -1) {
    throw new Error('Invalid source or destination sheet.');
  }

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    var spreadsheet = getSpreadsheet_();
    var metadata = getHeaderMetadata_(spreadsheet);
    var sourceSheet = spreadsheet.getSheetByName(fromSheet);
    var destinationSheet = spreadsheet.getSheetByName(toSheet);
    var sourceTask = findTaskRow_(sourceSheet, metadata, taskId);

    if (!sourceTask) {
      throw new Error('Task not found in source sheet.');
    }

    var rowValues = sourceTask.rowValues.slice();
    var targetStatus = getDefaultStatusForSheet_(toSheet, rowValues[metadata.indexByHeader.STATUS]);
    rowValues[metadata.indexByHeader.STATUS] = targetStatus;

    appendTaskRow_(destinationSheet, metadata, rowValues, toIndex);
    sourceSheet.deleteRow(sourceTask.rowNumber);

    return {
      success: true,
      movedTask: mapRowToTask_(
        rowValues,
        buildPlainRichTextRow_(rowValues),
        metadata,
        toSheet,
        destinationSheet.getLastRow()
      )
    };
  } finally {
    lock.releaseLock();
  }
}

function updateTask(taskId, sheetName, updates) {
  if (SHEET_NAMES.indexOf(sheetName) === -1) {
    throw new Error('Invalid sheet name.');
  }

  var allowedFields = {
    description: 'DESCRIPTION',
    remarks: 'REMARKS',
    notes: 'NOTES',
    links: 'Links',
    progress: 'Progress',
    status: 'STATUS'
  };

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    var spreadsheet = getSpreadsheet_();
    var metadata = getHeaderMetadata_(spreadsheet);
    var sheet = spreadsheet.getSheetByName(sheetName);
    var taskRow = findTaskRow_(sheet, metadata, taskId);

    if (!taskRow) {
      throw new Error('Task not found.');
    }

    Object.keys(updates || {}).forEach(function(field) {
      if (!allowedFields[field]) {
        return;
      }

      var header = allowedFields[field];
      taskRow.rowValues[metadata.indexByHeader[header]] = normalizeCellValue_(updates[field]);
    });

    var nextStatus = String(taskRow.rowValues[metadata.indexByHeader.STATUS] || '').trim();
    var moveTarget = getDestinationSheetForStatus_(sheetName, nextStatus);

    if (moveTarget && moveTarget !== sheetName) {
      var destinationSheet = spreadsheet.getSheetByName(moveTarget);
      appendTaskRow_(destinationSheet, metadata, taskRow.rowValues.slice());
      sheet.deleteRow(taskRow.rowNumber);

      return {
        success: true,
        task: mapRowToTask_(
          taskRow.rowValues,
          buildPlainRichTextRow_(taskRow.rowValues),
          metadata,
          moveTarget,
          destinationSheet.getLastRow()
        ),
        movedToSheet: moveTarget
      };
    }

    sheet.getRange(taskRow.rowNumber, 1, 1, metadata.totalColumns).setValues([taskRow.rowValues]);

    return {
      success: true,
      task: mapRowToTask_(
        taskRow.rowValues,
        buildPlainRichTextRow_(taskRow.rowValues),
        metadata,
        sheetName,
        taskRow.rowNumber
      )
    };
  } finally {
    lock.releaseLock();
  }
}

function ensureSpreadsheetStructure_() {
  var spreadsheet = getSpreadsheet_();
  var sheets = SHEET_NAMES.map(function(sheetName) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    }
    return sheet;
  });

  sheets.forEach(function(sheet) {
    ensureHeaders_(sheet);
  });

  sheets.forEach(function(sheet) {
    var metadata = getSheetHeaderMetadata_(sheet);
    assignMissingTaskIds_(sheet, metadata);
    formatLinksColumn_(sheet, metadata);
    hideInternalIdColumn_(sheet, metadata);
  });

  validateHeadersAcrossSheets_(sheets);
}

function getSpreadsheet_() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (error) {
    throw new Error(
      'Unable to open the spreadsheet. Make sure you are signed into the correct Google account and that this account has access to the sheet.'
    );
  }
}

function getHeaderMetadata_(spreadsheet) {
  var sheets = SHEET_NAMES.map(function(sheetName) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('Missing required sheet: ' + sheetName);
    }
    return sheet;
  });

  validateHeadersAcrossSheets_(sheets);
  return getSheetHeaderMetadata_(sheets[0]);
}

function getSheetHeaderMetadata_(sheet) {
  var headers = normalizeHeaders_(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
  var indexByHeader = {};

  headers.forEach(function(header, index) {
    indexByHeader[header] = index;
  });

  return {
    headers: headers,
    indexByHeader: indexByHeader,
    totalColumns: headers.length,
    idColumnNumber: indexByHeader[INTERNAL_ID_HEADER] + 1,
    statusColumnNumber: indexByHeader.STATUS + 1
  };
}

function validateHeadersAcrossSheets_(sheets) {
  sheets.forEach(function(sheet) {
    var headers = normalizeHeaders_(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
    var visibleHeaders = headers.filter(function(header) {
      return header !== INTERNAL_ID_HEADER && header !== '';
    });

    if (visibleHeaders.slice(0, BUSINESS_HEADERS.length).join('|') !== BUSINESS_HEADERS.join('|')) {
      throw new Error(
        'Sheet "' + sheet.getName() + '" must use these business headers in order: ' +
        BUSINESS_HEADERS.join(', ')
      );
    }

    if (headers.indexOf(INTERNAL_ID_HEADER) === -1) {
      throw new Error(
        'Sheet "' + sheet.getName() + '" is missing the internal helper column. Run Initialize Board Columns once.'
      );
    }
  });
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    var newHeaders = BUSINESS_HEADERS.concat([INTERNAL_ID_HEADER]);
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    sheet.setFrozenRows(1);
    sheet.hideColumns(newHeaders.length);
    return;
  }

  var lastColumn = sheet.getLastColumn();
  var headers = normalizeHeaders_(sheet.getRange(1, 1, 1, lastColumn).getValues()[0]);
  var businessHeaders = headers.slice(0, BUSINESS_HEADERS.length);

  if (businessHeaders.join('|') !== BUSINESS_HEADERS.join('|')) {
    throw new Error(
      'Sheet "' + sheet.getName() + '" has unexpected headers. Expected first columns: ' +
      BUSINESS_HEADERS.join(', ')
    );
  }

  var internalIndex = headers.indexOf(INTERNAL_ID_HEADER);
  if (internalIndex === -1) {
    var newColumn = lastColumn + 1;
    sheet.getRange(1, newColumn).setValue(INTERNAL_ID_HEADER);
    sheet.hideColumns(newColumn);
  } else {
    sheet.hideColumns(internalIndex + 1);
  }

  sheet.setFrozenRows(1);
}

function assignMissingTaskIds_(sheet, metadata) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  var idRange = sheet.getRange(2, metadata.idColumnNumber, lastRow - 1, 1);
  var idValues = idRange.getValues();
  var changed = false;

  for (var index = 0; index < idValues.length; index += 1) {
    if (!String(idValues[index][0] || '').trim()) {
      idValues[index][0] = Utilities.getUuid();
      changed = true;
    }
  }

  if (changed) {
    idRange.setValues(idValues);
  }
}

function formatLinksColumn_(sheet, metadata) {
  var linkColumn = metadata.indexByHeader.Links + 1;
  sheet.getRange(1, linkColumn, Math.max(sheet.getMaxRows(), 2), 1).setWrap(true);
}

function hideInternalIdColumn_(sheet, metadata) {
  if (metadata.idColumnNumber > 0) {
    sheet.hideColumns(metadata.idColumnNumber);
  }
}

function getTasksForSheet_(sheet, metadata) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  var values = sheet.getRange(2, 1, lastRow - 1, metadata.totalColumns).getValues();
  var richTextValues = sheet.getRange(2, 1, lastRow - 1, metadata.totalColumns).getRichTextValues();
  var tasks = [];

  values.forEach(function(rowValues, index) {
    var isBlank = BUSINESS_HEADERS.every(function(header) {
      return String(rowValues[metadata.indexByHeader[header]] || '').trim() === '';
    });

    if (!isBlank) {
      tasks.push(mapRowToTask_(rowValues, richTextValues[index], metadata, sheet.getName(), index + 2));
    }
  });

  return tasks;
}

function mapRowToTask_(rowValues, richTextRowValues, metadata, sheetName, rowNumber) {
  var descriptionIndex = metadata.indexByHeader.DESCRIPTION;
  var remarksIndex = metadata.indexByHeader.REMARKS;
  var notesIndex = metadata.indexByHeader.NOTES;
  var linksIndex = metadata.indexByHeader.Links;

  return {
    taskId: String(rowValues[metadata.indexByHeader[INTERNAL_ID_HEADER]] || ''),
    sheetName: sheetName,
    rowIndex: rowNumber,
    company: String(rowValues[metadata.indexByHeader.COMPANY] || ''),
    description: String(rowValues[metadata.indexByHeader.DESCRIPTION] || ''),
    descriptionHtml: richTextValueToHtml_(richTextRowValues[descriptionIndex]),
    status: String(rowValues[metadata.indexByHeader.STATUS] || ''),
    remarks: String(rowValues[metadata.indexByHeader.REMARKS] || ''),
    remarksHtml: richTextValueToHtml_(richTextRowValues[remarksIndex]),
    notes: String(rowValues[metadata.indexByHeader.NOTES] || ''),
    notesHtml: richTextValueToHtml_(richTextRowValues[notesIndex]),
    links: String(rowValues[metadata.indexByHeader.Links] || ''),
    linksHtml: richTextValueToHtml_(richTextRowValues[linksIndex]),
    progress: String(rowValues[metadata.indexByHeader.Progress] || '')
  };
}

function findTaskRow_(sheet, metadata, taskId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }

  var values = sheet.getRange(2, 1, lastRow - 1, metadata.totalColumns).getValues();
  for (var index = 0; index < values.length; index += 1) {
    if (String(values[index][metadata.indexByHeader[INTERNAL_ID_HEADER]]) === String(taskId)) {
      return {
        rowNumber: index + 2,
        rowValues: values[index]
      };
    }
  }

  return null;
}

function appendTaskRow_(sheet, metadata, rowValues, toIndex) {
  var insertionRow = sheet.getLastRow() + 1;
  if (typeof toIndex === 'number' && !isNaN(toIndex)) {
    insertionRow = Math.max(2, Math.min(Math.floor(toIndex), sheet.getLastRow() + 1));
    sheet.insertRowBefore(insertionRow);
  }

  sheet.getRange(insertionRow, 1, 1, metadata.totalColumns).setValues([rowValues]);
}

function getStatusOptions_(spreadsheet, metadata) {
  var options = [];
  var seen = {};
  var sourceSheet = spreadsheet.getSheetByName(SHEET_NAMES[0]);
  var validations = [];

  if (sourceSheet.getLastRow() >= 2) {
    validations = sourceSheet.getRange(2, metadata.statusColumnNumber, sourceSheet.getLastRow() - 1, 1).getDataValidations();
  } else {
    validations = [[sourceSheet.getRange(2, metadata.statusColumnNumber).getDataValidation()]];
  }

  validations.forEach(function(row) {
    row.forEach(function(rule) {
      extractStatusOptionsFromRule_(rule).forEach(function(option) {
        var key = option.toLowerCase();
        if (option && !seen[key]) {
          seen[key] = true;
          options.push(option);
        }
      });
    });
  });

  if (!options.length) {
    SHEET_NAMES.forEach(function(sheetName) {
      var sheet = spreadsheet.getSheetByName(sheetName);
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        return;
      }

      var statusValues = sheet.getRange(2, metadata.statusColumnNumber, lastRow - 1, 1).getDisplayValues();
      statusValues.forEach(function(row) {
        var value = String(row[0] || '').trim();
        var key = value.toLowerCase();
        if (value && !seen[key]) {
          seen[key] = true;
          options.push(value);
        }
      });
    });
  }

  ['Done', 'Cancelled'].forEach(function(option) {
    var key = option.toLowerCase();
    if (!seen[key]) {
      seen[key] = true;
      options.push(option);
    }
  });

  return options;
}

function extractStatusOptionsFromRule_(rule) {
  if (!rule) {
    return [];
  }

  var criteriaType = rule.getCriteriaType();
  var criteriaValues = rule.getCriteriaValues();

  if (criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    return (criteriaValues[0] || []).map(function(value) {
      return String(value || '').trim();
    }).filter(Boolean);
  }

  if (criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    var range = criteriaValues[0];
    if (!range) {
      return [];
    }

    return range.getDisplayValues().reduce(function(result, row) {
      row.forEach(function(value) {
        var normalized = String(value || '').trim();
        if (normalized) {
          result.push(normalized);
        }
      });
      return result;
    }, []);
  }

  return [];
}

function getDestinationSheetForStatus_(currentSheetName, statusValue) {
  var normalizedStatus = String(statusValue || '').trim().toLowerCase();

  if (currentSheetName === 'List of Task' && normalizedStatus === 'done') {
    return STATUS_COMPLETED_SHEET;
  }

  if (currentSheetName === 'List of Task' && normalizedStatus === 'cancelled') {
    return STATUS_CANCELLED_SHEET;
  }

  return currentSheetName;
}

function getDefaultStatusForSheet_(sheetName, currentStatus) {
  var normalizedCurrent = String(currentStatus || '').trim();
  if (sheetName === STATUS_COMPLETED_SHEET && !normalizedCurrent) {
    return 'Done';
  }
  if (sheetName === STATUS_CANCELLED_SHEET && !normalizedCurrent) {
    return 'Cancelled';
  }
  return normalizedCurrent;
}

function getWebAppUrl_() {
  var storedUrl = PropertiesService.getScriptProperties().getProperty('WEB_APP_URL');
  if (storedUrl) {
    return storedUrl;
  }

  try {
    return ScriptApp.getService().getUrl();
  } catch (error) {
    return '';
  }
}

function toLaneId_(sheetName) {
  return sheetName.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
}

function normalizeHeaders_(headers) {
  return headers.map(function(header) {
    return String(header || '').trim();
  });
}

function normalizeCellValue_(value) {
  if (value === null || typeof value === 'undefined') {
    return '';
  }
  return String(value);
}

function sanitizeForHtmlAttribute_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function richTextValueToHtml_(richTextValue) {
  if (!richTextValue) {
    return '';
  }

  var text = richTextValue.getText();
  if (!text) {
    return '';
  }

  var runs = richTextValue.getRuns();
  if (!runs || !runs.length) {
    return escapeHtml_(text).replace(/\n/g, '<br>');
  }

  return runs.map(function(run) {
    var rendered = escapeHtml_(run.getText()).replace(/\n/g, '<br>');
    var style = run.getTextStyle();
    var linkUrl = run.getLinkUrl();

    if (style) {
      if (style.isBold()) {
        rendered = '<strong>' + rendered + '</strong>';
      }
      if (style.isItalic()) {
        rendered = '<em>' + rendered + '</em>';
      }
      if (style.isUnderline()) {
        rendered = '<u>' + rendered + '</u>';
      }
      if (style.isStrikethrough()) {
        rendered = '<s>' + rendered + '</s>';
      }
    }

    if (linkUrl) {
      rendered = '<a class="task-link" href="' + sanitizeForHtmlAttribute_(linkUrl) +
        '" target="_blank" rel="noopener noreferrer">' + rendered + '</a>';
    }

    return rendered;
  }).join('');
}

function buildPlainRichTextRow_(rowValues) {
  return rowValues.map(function(value) {
    return SpreadsheetApp.newRichTextValue()
      .setText(String(value || ''))
      .build();
  });
}
