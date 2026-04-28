// 15 Validate User Columns
//
// Team A user-input columns must all be present in a configured tab. The
// setup wizard calls this to surface what's missing and offer to add the
// headers automatically. Aliases are honored — a sheet that already has
// "Staff Email" or "Remarks" passes without renaming.

function findMissingUserInputColumns(sheet, tabName) {
  var flexMap = buildFlexibleColumnMap(sheet, tabName);
  var missing = [];
  for (var i = 0; i < REQUIRED_USER_COLUMNS.length; i++) {
    var key = REQUIRED_USER_COLUMNS[i];
    if (!flexMap.map[key]) {
      missing.push({ key: key, label: HEADERS[key] });
    }
  }
  return missing;
}

// Returns true iff every Team A column is present (directly or via alias).
function userInputColumnsAreComplete(sheet, tabName) {
  return findMissingUserInputColumns(sheet, tabName).length === 0;
}

// Lists distinct, non-empty Assigned Staff Email values across the data
// rows of the given tabs. Used by the wizard to warn about staff emails
// that aren't verified Gmail aliases.
function listDistinctStaffEmailsForTabs(ss, tabNames) {
  var seen = {};
  var ordered = [];

  for (var t = 0; t < tabNames.length; t++) {
    var tabName = tabNames[t];
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) continue;

    var colMap = buildColumnMap(sheet, tabName);
    if (!colMap.STAFF_EMAIL) continue;

    var dataStartRow = getTabDataStartRow(tabName);
    var lastRow = sheet.getLastRow();
    if (lastRow < dataStartRow) continue;

    var values = sheet
      .getRange(dataStartRow, colMap.STAFF_EMAIL, lastRow - dataStartRow + 1, 1)
      .getValues();

    for (var i = 0; i < values.length; i++) {
      var raw = String(values[i][0] || "").trim().toLowerCase();
      if (!raw) continue;
      if (!seen[raw]) {
        seen[raw] = true;
        ordered.push(raw);
      }
    }
  }

  return ordered;
}

// Splits a list of staff email addresses into verified (usable as From)
// and unverified (would mis-attribute or fail). The returned object is
// passed into the wizard's pre-flight summary.
function classifyStaffEmailsByAliasVerification(staffEmails) {
  resetVerifiedSenderAliasCache();
  var verified = [];
  var unverified = [];
  for (var i = 0; i < staffEmails.length; i++) {
    if (canSendAs(staffEmails[i])) {
      verified.push(staffEmails[i]);
    } else {
      unverified.push(staffEmails[i]);
    }
  }
  return { verified: verified, unverified: unverified };
}
