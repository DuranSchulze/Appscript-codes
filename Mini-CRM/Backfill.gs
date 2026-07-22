/**
 * Menu entry point for an inclusive custom Gmail backfill date range.
 */
function menuBackfillEmailsByDateRange() {
  const ui = SpreadsheetApp.getUi();
  const startResponse = ui.prompt(
    'Backfill emails',
    'Enter the first date to include (YYYY-MM-DD).',
    ui.ButtonSet.OK_CANCEL
  );
  if (startResponse.getSelectedButton() !== ui.Button.OK) return;

  const endResponse = ui.prompt(
    'Backfill emails',
    'Enter the last date to include (YYYY-MM-DD).',
    ui.ButtonSet.OK_CANCEL
  );
  if (endResponse.getSelectedButton() !== ui.Button.OK) return;

  try {
    const result = backfillEmailsForDateRange(
      startResponse.getResponseText(),
      endResponse.getResponseText()
    );
    const limitNote = result.limitReached
      ? '\n\n⚠️ The ' + result.threadLimit + '-thread safety limit was reached. Run a smaller date range to capture the remainder.'
      : '';
    ui.alert(
      '✅ Backfill complete',
      result.newCount + ' new inbound emails added for ' + result.startLabel + ' through ' + result.endLabel + '.\n\n' +
      result.outboundSkipped + ' non-inbound messages skipped.\n' +
      result.duplicateSkipped + ' duplicate messages skipped.' + limitNote,
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert('Backfill could not run', error.message || String(error), ui.ButtonSet.OK);
  }
}

/**
 * Backfills an inclusive date range. Inputs may be YYYY-MM-DD strings or Dates.
 * Returns counters so triggers/tests can call it without depending on UI dialogs.
 */
function backfillEmailsForDateRange(startDateInput, endDateInput) {
  assertMonitoredMailboxAccount_('custom Gmail backfill');
  const startDate = normalizeBackfillDate_(startDateInput, false);
  const endDate = normalizeBackfillDate_(endDateInput, true);
  if (startDate > endDate) throw new Error('The start date must be on or before the end date.');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error('Another email sync or backfill is already running. Try again shortly.');

  try {
    return backfillEmailsForDateRangeUnlocked_(startDate, endDate);
  } finally {
    lock.releaseLock();
  }
}

function backfillEmailsForDateRangeUnlocked_(startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const query = buildBackfillGmailQuery_(startDate, endDate);
  const identityIndex = getExistingEmailIdentityIndex_(ss);
  const mailboxContext = getMonitoredMailboxContext_();
  const userEmail = mailboxContext.primaryEmail;
  const conversionSheet = getOrCreateConversionTrackingSheet(ss);
  const monthSheets = {};
  const qualificationBudget = { remaining: AI_CONFIG.MAX_QUALIFICATIONS_PER_SYNC };
  const batchSize = CONFIG.BACKFILL_BATCH_SIZE;
  const threadLimit = CONFIG.MAX_BACKFILL_THREADS;

  let threadOffset = 0;
  let newCount = 0;
  let outboundSkipped = 0;
  let duplicateSkipped = 0;
  let limitReached = false;

  while (threadOffset < threadLimit) {
    const fetchSize = Math.min(batchSize, threadLimit - threadOffset);
    const threads = GmailApp.search(query, threadOffset, fetchSize);
    if (!threads.length) break;

    for (const thread of threads) {
      for (const msg of thread.getMessages()) {
        const date = safeGetDate(msg.getDate());
        if (date < startDate || date > endDate) continue;

        const inbound = inspectInboundMessage_(msg, mailboxContext);
        if (!inbound.isInbound) {
          outboundSkipped++;
          continue;
        }

        const senderEmail = extractEmailAddress(msg.getFrom());
        const subject = msg.getSubject() || '(No Subject)';
        const identity = inspectDuplicateMessage_(msg, senderEmail, date, subject, identityIndex);
        if (identity.isDuplicate) {
          duplicateSkipped++;
          continue;
        }

        let emailData = processEmailMessage(msg, thread, userEmail, identityIndex.fallbackIds);
        emailData = qualifyEmailMessageForCrm_(msg, thread, emailData, qualificationBudget);
        const month = Utilities.formatDate(date, CONFIG.TIMEZONE, 'MMM-yyyy');

        if (!monthSheets[month]) {
          let sheet = ss.getSheetByName(month);
          if (!sheet) {
            sheet = ss.insertSheet(month);
            sheet.getRange(1, 1, 1, MONTHLY_EMAIL_HEADERS.length).setValues([MONTHLY_EMAIL_HEADERS]);
          }
          formatMonthlySheet_(sheet);
          monthSheets[month] = sheet;
        }

        monthSheets[month].appendRow(emailData);
        recordMessageIdentity_(identityIndex, identity);

        const status = emailData[8];
        if (status === 'Trash' || status === 'N/A') {
          monthSheets[month].getRange(monthSheets[month].getLastRow(), 1, 1, emailData.length)
            .setBackground('#D3D3D3')
            .setFontColor('#666666');
        }

        if (conversionSheet) updateConversionTracking(conversionSheet, emailData);
        newCount++;
      }
    }

    threadOffset += threads.length;
    if (threads.length < fetchSize) break;
    if (threadOffset >= threadLimit) limitReached = true;
  }

  // Refresh the review queue; reviewers explicitly promote qualified candidates.
  syncPotentialClientsFromMonthlySheets_(ss);

  return {
    newCount: newCount,
    outboundSkipped: outboundSkipped,
    duplicateSkipped: duplicateSkipped,
    threadLimit: threadLimit,
    limitReached: limitReached,
    startLabel: Utilities.formatDate(startDate, CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    endLabel: Utilities.formatDate(endDate, CONFIG.TIMEZONE, 'yyyy-MM-dd')
  };
}

function normalizeBackfillDate_(value, endOfDay) {
  let text = '';
  if (value instanceof Date && !isNaN(value)) {
    text = Utilities.formatDate(value, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  } else {
    text = String(value || '').trim();
  }

  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) throw new Error('Dates must use YYYY-MM-DD format.');

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = endOfDay
    ? new Date(year, month - 1, day, 23, 59, 59, 999)
    : new Date(year, month - 1, day, 0, 0, 0, 0);

  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    throw new Error('Enter a valid calendar date in YYYY-MM-DD format.');
  }
  return date;
}

function buildBackfillGmailQuery_(startDate, endDate) {
  // Gmail's after:/before: date operators are exclusive, so expand by one day
  // to make the requested spreadsheet range inclusive at both ends.
  const dayBeforeStart = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate() - 1);
  const dayAfterEnd = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate() + 1);
  return 'in:inbox after:' + Utilities.formatDate(dayBeforeStart, CONFIG.TIMEZONE, 'yyyy/MM/dd') +
    ' before:' + Utilities.formatDate(dayAfterEnd, CONFIG.TIMEZONE, 'yyyy/MM/dd');
}
