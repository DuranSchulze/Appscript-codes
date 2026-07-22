/**
 * AI-assisted potential-client qualification and review queue.
 */

const POTENTIAL_CLIENT_HEADERS = [
  'Candidate ID',
  'First Received',
  'Last Activity',
  'Name / Company',
  'Email Address',
  'Subject',
  'Latest Summary',
  'Source Month',
  'Intent Category',
  'AI Confidence',
  'Qualification Status',
  'Qualification Reason',
  'Assigned Team Member',
  'Next Follow-Up Date',
  'Source Message ID',
  'Source Thread ID',
  'Decision Source',
  'Manual Notes',
  'Promoted At',
  'Engagement Row'
];

const POTENTIAL_CLIENT_COLUMNS = {
  CANDIDATE_ID: 1,
  FIRST_RECEIVED: 2,
  LAST_ACTIVITY: 3,
  NAME: 4,
  EMAIL: 5,
  SUBJECT: 6,
  SUMMARY: 7,
  SOURCE_MONTH: 8,
  INTENT_CATEGORY: 9,
  AI_CONFIDENCE: 10,
  STATUS: 11,
  REASON: 12,
  ASSIGNED_TEAM_MEMBER: 13,
  NEXT_FOLLOW_UP: 14,
  MESSAGE_ID: 15,
  THREAD_ID: 16,
  DECISION_SOURCE: 17,
  MANUAL_NOTES: 18,
  PROMOTED_AT: 19,
  ENGAGEMENT_ROW: 20
};

const MONTHLY_QUALIFICATION_COLUMNS = {
  STATUS: 11,
  CONFIDENCE: 12,
  INTENT_CATEGORY: 13,
  REASON: 14,
  MODEL: 15,
  MESSAGE_ID: 16,
  THREAD_ID: 17,
  CLASSIFIED_AT: 18
};

function ensurePotentialClientsSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.POTENTIAL_CLIENTS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.POTENTIAL_CLIENTS_SHEET_NAME);

  const current = sheet.getRange(1, 1, 1, POTENTIAL_CLIENT_HEADERS.length).getValues()[0];
  const blank = current.every(function(value) { return !String(value || '').trim(); });
  const valid = POTENTIAL_CLIENT_HEADERS.every(function(header, index) {
    return String(current[index] || '').trim() === header;
  });

  if (blank) {
    sheet.getRange(1, 1, 1, POTENTIAL_CLIENT_HEADERS.length).setValues([POTENTIAL_CLIENT_HEADERS]);
  } else if (!valid) {
    throw new Error('Potential Clients headers do not match the required schema. Existing data was preserved.');
  }

  formatPotentialClientsSheet_(sheet);
  return sheet;
}

function formatPotentialClientsSheet_(sheet) {
  const columns = POTENTIAL_CLIENT_HEADERS.length;
  sheet.getRange(1, 1, 1, columns)
    .setFontWeight('bold')
    .setBackground('#6B3FA0')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(1, 42);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(5);
  sheet.setTabColor('#6B3FA0');

  [145, 125, 125, 180, 210, 270, 320, 105, 135, 95, 125, 300, 150, 125, 180, 180, 210, 240, 135, 110]
    .forEach(function(width, index) { sheet.setColumnWidth(index + 1, width); });

  if (sheet.getLastRow() > 1) {
    const rows = sheet.getLastRow() - 1;
    [2, 3, 14, 19].forEach(function(column) {
      sheet.getRange(2, column, rows, 1).setNumberFormat('mmm d, yyyy h:mm AM/PM');
    });
    sheet.getRange(2, 10, rows, 1).setNumberFormat('0');
    sheet.getRange(2, 1, rows, columns).setVerticalAlignment('middle');
  }

  if (!sheet.getFilter()) {
    sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 2), columns).createFilter();
  }

  if (sheet.getConditionalFormatRules().length === 0) {
    const statusRange = sheet.getRange(2, POTENTIAL_CLIENT_COLUMNS.STATUS, Math.max(sheet.getMaxRows() - 1, 1), 1);
    const rules = [
      ['Qualified', '#E6F4EA', '#137333'],
      ['Review', '#FEF7E0', '#8A4B00'],
      ['Promoted', '#E8F0FE', '#174EA6'],
      ['Rejected', '#F1F3F4', '#5F6368']
    ].map(function(config) {
      return SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(config[0])
        .setBackground(config[1])
        .setFontColor(config[2])
        .setRanges([statusRange])
        .build();
    });
    sheet.setConditionalFormatRules(rules);
  }

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Review', 'Qualified', 'Rejected'], true)
    .setAllowInvalid(false)
    .setHelpText('Choose Review, Qualified, or Rejected. Use the menu to promote a Qualified candidate.')
    .build();
  sheet.getRange(2, POTENTIAL_CLIENT_COLUMNS.STATUS, Math.max(sheet.getMaxRows() - 1, 1), 1)
    .setDataValidation(statusRule);
}

/**
 * Applies hard exclusions, then Gemini qualification, and returns a complete
 * monthly-sheet row including qualification metadata.
 */
function qualifyEmailMessageForCrm_(msg, thread, emailData, budget) {
  const status = String(emailData[8] || '');
  const messageId = safeCallId_(msg);
  const threadId = safeCallId_(thread);

  if (status === 'Trash' || status === 'N/A' || status === 'E-sign' || status === 'Client') {
    const reason = status === 'Client'
      ? 'Existing conversion/client signal detected by deterministic rules.'
      : 'Excluded by deterministic address or system-email rules: ' + status + '.';
    return appendQualificationMetadata_(emailData, {
      status: 'Rejected', confidence: 100, category: status,
      reason: reason, model: 'Deterministic rules', messageId: messageId, threadId: threadId
    });
  }

  const automatedReason = getAutomatedMessageReason_(msg);
  if (automatedReason) {
    emailData[8] = 'Non-Sales';
    return appendQualificationMetadata_(emailData, {
      status: 'Rejected', confidence: 100, category: 'Automated',
      reason: automatedReason, model: 'Deterministic Gmail headers',
      messageId: messageId, threadId: threadId
    });
  }

  if (!isAiProcessingEnabled_()) {
    emailData[8] = 'Review';
    return appendQualificationMetadata_(emailData, {
      status: 'Review', confidence: '', category: 'Unknown',
      reason: 'AI processing is paused; manual review is required.',
      model: '', messageId: messageId, threadId: threadId
    });
  }

  if (!getGeminiApiKey()) {
    emailData[8] = 'Review';
    return appendQualificationMetadata_(emailData, {
      status: 'Review', confidence: '', category: 'Unknown',
      reason: 'Gemini is not configured; manual review is required.',
      model: '', messageId: messageId, threadId: threadId
    });
  }

  if (!budget || budget.remaining <= 0) {
    emailData[8] = 'Review';
    return appendQualificationMetadata_(emailData, {
      status: 'Pending', confidence: '', category: 'Unknown',
      reason: 'Queued for deferred AI qualification to protect runtime and API quota.',
      model: '', messageId: messageId, threadId: threadId
    });
  }

  budget.remaining--;
  try {
    const result = classifyPotentialClientWithGemini_(
      emailData[5],
      msg.getPlainBody() || '',
      emailData[4]
    );
    emailData[8] = qualificationStatusToMonthlyStatus_(result.status);
    return appendQualificationMetadata_(emailData, {
      status: result.status,
      confidence: result.confidence,
      category: result.intentCategory,
      reason: result.reason,
      model: result.model,
      messageId: messageId,
      threadId: threadId
    });
  } catch (error) {
    emailData[8] = 'Review';
    if (error && error.aiBlocked) {
      if (budget) budget.remaining = 0;
      return appendQualificationMetadata_(emailData, {
        status: 'Pending', confidence: '', category: 'Unknown',
        reason: 'AI qualification deferred by the shared usage guard: ' + error.message,
        model: '', messageId: messageId, threadId: threadId
      });
    }
    return appendQualificationMetadata_(emailData, {
      status: 'Review', confidence: '', category: 'Unknown',
      reason: 'AI qualification failed: ' + error.message,
      model: '', messageId: messageId, threadId: threadId
    });
  }
}

function getAutomatedMessageReason_(msg) {
  try {
    if (!msg || typeof msg.getHeader !== 'function') return '';
    const listUnsubscribe = String(msg.getHeader('List-Unsubscribe') || '').trim();
    if (listUnsubscribe) return 'Rejected before AI: mailing-list message has a List-Unsubscribe header.';
    const autoSubmitted = String(msg.getHeader('Auto-Submitted') || '').trim().toLowerCase();
    if (autoSubmitted && autoSubmitted !== 'no') return 'Rejected before AI: Auto-Submitted header is ' + autoSubmitted + '.';
    const precedence = String(msg.getHeader('Precedence') || '').trim().toLowerCase();
    if (['bulk', 'list', 'junk'].indexOf(precedence) !== -1) {
      return 'Rejected before AI: Precedence header identifies ' + precedence + ' mail.';
    }
  } catch (error) {
    Logger.log('Automated-message header inspection skipped: ' + error.message);
  }
  return '';
}

function appendQualificationMetadata_(emailData, metadata) {
  const row = emailData.slice(0, 10);
  return row.concat([
    metadata.status || 'Pending',
    metadata.confidence === 0 ? 0 : (metadata.confidence || ''),
    metadata.category || 'Unknown',
    metadata.reason || '',
    metadata.model || '',
    metadata.messageId || '',
    metadata.threadId || '',
    metadata.status === 'Pending' ? '' : new Date()
  ]);
}

function classifyPotentialClientWithGemini_(subject, body, senderEmail) {
  const cleanBody = sanitizeEmailForQualification_(body);
  const prompt = [
    'You qualify inbound email for Duran Schulze Law, a Philippines-based professional-services firm.',
    'Decide whether this is a NEW potential sales/client inquiry for legal, immigration/visa, business registration, corporate, trademark, tax, accounting, or consultation services.',
    'Do not qualify newsletters, marketing, vendors, job applications, automated notices, spam, internal operations, payment/e-sign system notices, or routine work for an existing client.',
    'Use both the subject and body. When evidence is ambiguous, require manual review.',
    'Return JSON only with: isPotentialClient (boolean), confidence (integer 0-100), intentCategory (one of Visa, Legal, Business Formation, Trademark, Accounting, Tax, Consultation, General, Non-Sales, Unknown), reason (one concise sentence), existingClientLikely (boolean), requiresManualReview (boolean).',
    '',
    'Sender: ' + String(senderEmail || ''),
    'Subject: ' + String(subject || '(No Subject)'),
    'Body:',
    cleanBody
  ].join('\n');

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json',
      maxOutputTokens: 300
    }
  };

  const route = [getGeminiModel()];
  const fallback = getGeminiFallbackModel();
  if (fallback && fallback !== route[0]) route.push(fallback);

  const cacheKey = createAiCacheKey_('qualification:v1', [
    route.join('>'), AI_CONFIG.QUALIFIED_LEAD_THRESHOLD,
    AI_CONFIG.MANUAL_REVIEW_THRESHOLD, senderEmail, subject, cleanBody
  ]);
  const cached = getCachedAiJson_(cacheKey);
  if (cached && cached.status && cached.model) {
    cached.cacheHit = true;
    return cached;
  }

  let lastError = null;
  for (let i = 0; i < route.length; i++) {
    try {
      const raw = requestGeminiJson_(getGeminiApiKey(), route[i], payload, { purpose: 'qualification' });
      const normalized = normalizeQualificationResult_(raw);
      normalized.model = route[i];
      normalized.usedFallback = i > 0;
      putCachedAiJson_(cacheKey, normalized, 21600);
      return normalized;
    } catch (error) {
      lastError = error;
      Logger.log('Qualification model ' + route[i] + ' failed: ' + error.message);
      if (shouldStopAiRoute_(error)) break;
    }
  }
  throw lastError || new Error('No Gemini model was available for qualification.');
}

function normalizeQualificationResult_(raw) {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) {
    throw new Error('Gemini qualification response was not a JSON object.');
  }
  if (typeof raw.isPotentialClient !== 'boolean' ||
      typeof raw.existingClientLikely !== 'boolean' ||
      typeof raw.requiresManualReview !== 'boolean' ||
      typeof raw.intentCategory !== 'string' ||
      typeof raw.reason !== 'string' ||
      typeof raw.confidence !== 'number' || !isFinite(raw.confidence)) {
    throw new Error('Gemini qualification response did not match the required contract.');
  }
  const confidence = Math.max(0, Math.min(100, Math.round(Number(raw.confidence) || 0)));
  const isPotential = raw.isPotentialClient === true;
  const existingClientLikely = raw.existingClientLikely === true;
  const requiresReview = raw.requiresManualReview === true;
  let status = 'Rejected';

  if (isPotential && !existingClientLikely && !requiresReview && confidence >= AI_CONFIG.QUALIFIED_LEAD_THRESHOLD) {
    status = 'Qualified';
  } else if (!existingClientLikely && (isPotential || requiresReview || confidence >= AI_CONFIG.MANUAL_REVIEW_THRESHOLD)) {
    status = 'Review';
  } else if (existingClientLikely && requiresReview) {
    status = 'Review';
  }

  const allowedCategories = [
    'Visa', 'Legal', 'Business Formation', 'Trademark', 'Accounting',
    'Tax', 'Consultation', 'General', 'Non-Sales', 'Unknown'
  ];
  const category = allowedCategories.indexOf(raw.intentCategory) !== -1
    ? raw.intentCategory
    : (status === 'Rejected' ? 'Non-Sales' : 'Unknown');

  return {
    status: status,
    confidence: confidence,
    intentCategory: category,
    reason: String(raw.reason || 'No qualification reason was provided.').trim().substring(0, 500),
    existingClientLikely: existingClientLikely,
    requiresManualReview: requiresReview
  };
}

function sanitizeEmailForQualification_(body) {
  let text = String(body || '').replace(/\r/g, '');
  text = text.split(/\nOn .{0,300}wrote:\s*\n/i)[0];
  text = text.split(/\nFrom:\s.+\nSent:\s.+\nTo:\s.+/i)[0];
  text = text.split(/\n--\s*\n/)[0];
  text = text.split('\n')
    .filter(function(line) { return !/^\s*>/.test(line); })
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
  return text.substring(0, AI_CONFIG.MAX_QUALIFICATION_BODY_CHARS);
}

function qualificationStatusToMonthlyStatus_(status) {
  if (status === 'Qualified') return 'Prospect';
  if (status === 'Review' || status === 'Pending') return 'Review';
  return 'Non-Sales';
}

function safeCallId_(object) {
  try {
    return object && typeof object.getId === 'function' ? object.getId() : '';
  } catch (ignore) {
    return '';
  }
}

/**
 * Deferred processor for rows not classified during Gmail sync because the
 * per-run AI budget was reached.
 */
function processPendingAiQualifications() {
  assertMonitoredMailboxAccount_('pending AI qualification processing');
  if (!isAiProcessingEnabled_()) return { processed: 0, reason: 'AI processing is paused' };
  if (!getGeminiApiKey()) return { processed: 0, reason: 'Gemini not configured' };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return { processed: 0, reason: 'Another qualification run is active' };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let processed = 0;
    let guardReason = '';
    const monthlySheets = ss.getSheets().filter(function(sheet) {
      return /^[A-Za-z]{3}-\d{4}$/.test(sheet.getName());
    }).sort(function(a, b) {
      return parseMonthSheetName(b.getName()) - parseMonthSheetName(a.getName());
    });

    for (let s = 0; s < monthlySheets.length && !guardReason && processed < AI_CONFIG.MAX_PENDING_QUALIFICATIONS_PER_RUN; s++) {
      const sheet = monthlySheets[s];
      formatMonthlySheet_(sheet);
      if (sheet.getLastRow() < 2) continue;
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, MONTHLY_EMAIL_HEADERS.length).getValues();

      for (let i = 0; i < data.length && !guardReason && processed < AI_CONFIG.MAX_PENDING_QUALIFICATIONS_PER_RUN; i++) {
        const row = data[i];
        const qualificationStatus = String(row[MONTHLY_QUALIFICATION_COLUMNS.STATUS - 1] || 'Pending');
        const baseStatus = String(row[8] || '');
        if (qualificationStatus !== 'Pending' && qualificationStatus !== '') continue;
        if (['Trash', 'N/A', 'E-sign', 'Client'].indexOf(baseStatus) !== -1) continue;

        const workIdentity = row[15] || (sheet.getName() + ':' + (i + 2) + ':' + row[4] + ':' + row[5]);
        if (!canAttemptAiWork_('qualification', workIdentity)) continue;

        try {
          const result = classifyPotentialClientWithGemini_(row[5], row[9], row[4]);
          const sheetRow = i + 2;
          sheet.getRange(sheetRow, 9).setValue(qualificationStatusToMonthlyStatus_(result.status));
          sheet.getRange(sheetRow, 11, 1, 8).setValues([[
            result.status, result.confidence, result.intentCategory, result.reason,
            result.model, row[15] || '', row[16] || '', new Date()
          ]]);
          clearAiWorkFailure_('qualification', workIdentity);
          processed++;
        } catch (error) {
          recordAiWorkFailure_('qualification', workIdentity);
          Logger.log('Pending qualification failed in ' + sheet.getName() + ' row ' + (i + 2) + ': ' + error.message);
          if (error && error.aiBlocked) guardReason = error.message;
        }
      }
    }

    if (processed > 0) {
      syncPotentialClientsFromMonthlySheets_(ss);
      buildEnhancedDashboard();
    }
    return { processed: processed, reason: guardReason || '' };
  } finally {
    lock.releaseLock();
  }
}

function menuProcessPendingAiQualifications() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const result = processPendingAiQualifications();
    const message = 'Processed ' + result.processed + ' pending qualification(s).' +
      (result.reason ? '\nStopped safely: ' + result.reason : '');
    ss.toast(message, 'Potential Clients', 6);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Qualification Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function syncPotentialClientsFromMonthlySheets_(ss) {
  const target = ensurePotentialClientsSheet_(ss);
  const existing = target.getLastRow() > 1
    ? target.getRange(2, 1, target.getLastRow() - 1, POTENTIAL_CLIENT_HEADERS.length).getValues()
    : [];
  const byEmail = {};
  existing.forEach(function(row, index) {
    const email = String(row[POTENTIAL_CLIENT_COLUMNS.EMAIL - 1] || '').toLowerCase().trim();
    if (email) byEmail[email] = { row: row, sheetRow: index + 2 };
  });

  const candidates = {};
  ss.getSheets().forEach(function(sheet) {
    if (!/^[A-Za-z]{3}-\d{4}$/.test(sheet.getName()) || sheet.getLastRow() < 2) return;
    formatMonthlySheet_(sheet);
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, MONTHLY_EMAIL_HEADERS.length).getValues();
    rows.forEach(function(row) {
      const status = String(row[MONTHLY_QUALIFICATION_COLUMNS.STATUS - 1] || '');
      if (status !== 'Qualified' && status !== 'Review') return;
      const email = String(row[4] || '').toLowerCase().trim();
      if (!email) return;
      const received = safeGetDate(row[0]);
      const current = candidates[email];
      if (!current) {
        candidates[email] = { latest: row, firstReceived: received };
      } else {
        if (received < current.firstReceived) current.firstReceived = received;
        if (received > safeGetDate(current.latest[0])) current.latest = row;
      }
    });
  });

  Object.keys(candidates).forEach(function(email) {
    const candidateSource = candidates[email];
    const source = candidateSource.latest;
    const sourceStatus = source[MONTHLY_QUALIFICATION_COLUMNS.STATUS - 1];
    const existingEntry = byEmail[email];

    if (existingEntry) {
      const currentStatus = String(existingEntry.row[POTENTIAL_CLIENT_COLUMNS.STATUS - 1] || '');
      const decisionSource = String(existingEntry.row[POTENTIAL_CLIENT_COLUMNS.DECISION_SOURCE - 1] || 'AI');
      const preserveDecision = decisionSource.indexOf('Manual') === 0 || currentStatus === 'Promoted' || currentStatus === 'Rejected';
      const updates = existingEntry.row.slice();
      const recordedFirst = safeGetDate(updates[POTENTIAL_CLIENT_COLUMNS.FIRST_RECEIVED - 1]);
      if (!updates[POTENTIAL_CLIENT_COLUMNS.FIRST_RECEIVED - 1] || candidateSource.firstReceived < recordedFirst) {
        updates[POTENTIAL_CLIENT_COLUMNS.FIRST_RECEIVED - 1] = candidateSource.firstReceived;
      }
      updates[POTENTIAL_CLIENT_COLUMNS.LAST_ACTIVITY - 1] = source[0];
      updates[POTENTIAL_CLIENT_COLUMNS.SUBJECT - 1] = source[5];
      updates[POTENTIAL_CLIENT_COLUMNS.SUMMARY - 1] = source[9];
      updates[POTENTIAL_CLIENT_COLUMNS.SOURCE_MONTH - 1] = Utilities.formatDate(safeGetDate(source[0]), CONFIG.TIMEZONE, 'MMM-yyyy');
      updates[POTENTIAL_CLIENT_COLUMNS.INTENT_CATEGORY - 1] = source[12];
      updates[POTENTIAL_CLIENT_COLUMNS.AI_CONFIDENCE - 1] = source[11];
      updates[POTENTIAL_CLIENT_COLUMNS.REASON - 1] = source[13];
      updates[POTENTIAL_CLIENT_COLUMNS.MESSAGE_ID - 1] = source[15];
      updates[POTENTIAL_CLIENT_COLUMNS.THREAD_ID - 1] = source[16];
      if (!preserveDecision) updates[POTENTIAL_CLIENT_COLUMNS.STATUS - 1] = sourceStatus;
      target.getRange(existingEntry.sheetRow, 1, 1, POTENTIAL_CLIENT_HEADERS.length).setValues([updates]);
    } else {
      const firstDate = candidateSource.firstReceived;
      const lastActivity = safeGetDate(source[0]);
      const newRow = new Array(POTENTIAL_CLIENT_HEADERS.length).fill('');
      newRow[POTENTIAL_CLIENT_COLUMNS.CANDIDATE_ID - 1] = createCandidateId_(email);
      newRow[POTENTIAL_CLIENT_COLUMNS.FIRST_RECEIVED - 1] = firstDate;
      newRow[POTENTIAL_CLIENT_COLUMNS.LAST_ACTIVITY - 1] = lastActivity;
      newRow[POTENTIAL_CLIENT_COLUMNS.NAME - 1] = extractSenderName(source[3]);
      newRow[POTENTIAL_CLIENT_COLUMNS.EMAIL - 1] = email;
      newRow[POTENTIAL_CLIENT_COLUMNS.SUBJECT - 1] = source[5];
      newRow[POTENTIAL_CLIENT_COLUMNS.SUMMARY - 1] = source[9];
      newRow[POTENTIAL_CLIENT_COLUMNS.SOURCE_MONTH - 1] = Utilities.formatDate(firstDate, CONFIG.TIMEZONE, 'MMM-yyyy');
      newRow[POTENTIAL_CLIENT_COLUMNS.INTENT_CATEGORY - 1] = source[12];
      newRow[POTENTIAL_CLIENT_COLUMNS.AI_CONFIDENCE - 1] = source[11];
      newRow[POTENTIAL_CLIENT_COLUMNS.STATUS - 1] = sourceStatus;
      newRow[POTENTIAL_CLIENT_COLUMNS.REASON - 1] = source[13];
      newRow[POTENTIAL_CLIENT_COLUMNS.MESSAGE_ID - 1] = source[15];
      newRow[POTENTIAL_CLIENT_COLUMNS.THREAD_ID - 1] = source[16];
      newRow[POTENTIAL_CLIENT_COLUMNS.DECISION_SOURCE - 1] = 'AI';
      target.appendRow(newRow);
    }
  });

  formatPotentialClientsSheet_(target);
  return { candidates: Object.keys(candidates).length };
}

function createCandidateId_(email) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(email || '').toLowerCase().trim()
  );
  return 'PC-' + Utilities.base64EncodeWebSafe(digest).replace(/=+$/, '').substring(0, 16);
}

function menuPromoteSelectedPotentialClient() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = promoteSelectedPotentialClient_();
    ui.alert('✅ Candidate Promoted', 'Engagement row ' + result.engagementRow + ' is ready.', ui.ButtonSet.OK);
    buildEnhancedDashboard();
  } catch (error) {
    ui.alert('Unable to Promote', error.message, ui.ButtonSet.OK);
  }
}

function menuApproveSelectedPotentialClient() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getName() !== CONFIG.POTENTIAL_CLIENTS_SHEET_NAME) throw new Error('Select a row in Potential Clients.');
    const row = sheet.getActiveRange().getRow();
    if (row < 2) throw new Error('Select a candidate data row.');
    sheet.getRange(row, POTENTIAL_CLIENT_COLUMNS.STATUS).setValue('Qualified');
    sheet.getRange(row, POTENTIAL_CLIENT_COLUMNS.DECISION_SOURCE).setValue(getManualDecisionSource_());
    ui.alert('Candidate approved as Qualified. Use “Promote Selected to Engagement” when ready.');
    buildEnhancedDashboard();
  } catch (error) {
    ui.alert('Unable to Approve', error.message, ui.ButtonSet.OK);
  }
}

function promoteSelectedPotentialClient_() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) throw new Error('Another candidate promotion is currently running. Please try again.');
  try {
    return promoteSelectedPotentialClientUnlocked_();
  } finally {
    lock.releaseLock();
  }
}

function promoteSelectedPotentialClientUnlocked_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (sheet.getName() !== CONFIG.POTENTIAL_CLIENTS_SHEET_NAME) {
    throw new Error('Select a candidate row in the Potential Clients sheet.');
  }
  const rowNumber = sheet.getActiveRange().getRow();
  if (rowNumber < 2) throw new Error('Select a candidate data row.');

  const candidate = sheet.getRange(rowNumber, 1, 1, POTENTIAL_CLIENT_HEADERS.length).getValues()[0];
  const status = String(candidate[POTENTIAL_CLIENT_COLUMNS.STATUS - 1] || '');
  if (status !== 'Qualified' && status !== 'Promoted') {
    throw new Error('Only a Qualified candidate can be promoted. Review and qualify the candidate first.');
  }

  const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);
  if (!infoSheet) throw new Error('Engagement Information Sheet was not found.');
  if (!ensureEngagementSheetSchema_(infoSheet)) throw new Error('Engagement sheet schema is invalid.');

  const email = String(candidate[POTENTIAL_CLIENT_COLUMNS.EMAIL - 1] || '').toLowerCase().trim();
  if (!email) throw new Error('The candidate has no email address.');
  let engagementRow = Number(candidate[POTENTIAL_CLIENT_COLUMNS.ENGAGEMENT_ROW - 1]) || 0;

  if (status === 'Promoted' && engagementRow) {
    return { engagementRow: engagementRow, alreadyPromoted: true };
  }

  if (!engagementRow && infoSheet.getLastRow() > 1) {
    const emails = infoSheet.getRange(2, 5, infoSheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < emails.length; i++) {
      if (String(emails[i][0] || '').toLowerCase().trim() === email) {
        engagementRow = i + 2;
        break;
      }
    }
  }

  if (!engagementRow) {
    const engagement = new Array(ENGAGEMENT_INFO_HEADERS.length).fill('');
    engagement[0] = candidate[POTENTIAL_CLIENT_COLUMNS.FIRST_RECEIVED - 1];
    engagement[1] = candidate[POTENTIAL_CLIENT_COLUMNS.NAME - 1];
    engagement[2] = candidate[POTENTIAL_CLIENT_COLUMNS.INTENT_CATEGORY - 1];
    engagement[4] = email;
    engagement[7] = candidate[POTENTIAL_CLIENT_COLUMNS.SUMMARY - 1] || candidate[POTENTIAL_CLIENT_COLUMNS.SUBJECT - 1];
    engagement[ENGAGEMENT_COLUMNS.SOURCE_MONTH - 1] = candidate[POTENTIAL_CLIENT_COLUMNS.SOURCE_MONTH - 1];
    engagement[ENGAGEMENT_COLUMNS.AI_SCORE - 1] = candidate[POTENTIAL_CLIENT_COLUMNS.AI_CONFIDENCE - 1];
    engagement[ENGAGEMENT_COLUMNS.INTENT_CATEGORY - 1] = candidate[POTENTIAL_CLIENT_COLUMNS.INTENT_CATEGORY - 1];
    engagement[ENGAGEMENT_COLUMNS.ASSIGNED_TEAM_MEMBER - 1] = candidate[POTENTIAL_CLIENT_COLUMNS.ASSIGNED_TEAM_MEMBER - 1];
    infoSheet.appendRow(engagement);
    engagementRow = infoSheet.getLastRow();
  }

  sheet.getRange(rowNumber, POTENTIAL_CLIENT_COLUMNS.STATUS).setValue('Promoted');
  sheet.getRange(rowNumber, POTENTIAL_CLIENT_COLUMNS.DECISION_SOURCE).setValue(getManualDecisionSource_());
  sheet.getRange(rowNumber, POTENTIAL_CLIENT_COLUMNS.PROMOTED_AT).setValue(new Date());
  sheet.getRange(rowNumber, POTENTIAL_CLIENT_COLUMNS.ENGAGEMENT_ROW).setValue(engagementRow);
  setupFinancialFormulas(infoSheet);
  formatEngagementSheet_(infoSheet);
  return { engagementRow: engagementRow };
}

function menuRejectSelectedPotentialClient() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getName() !== CONFIG.POTENTIAL_CLIENTS_SHEET_NAME) throw new Error('Select a row in Potential Clients.');
    const row = sheet.getActiveRange().getRow();
    if (row < 2) throw new Error('Select a candidate data row.');
    sheet.getRange(row, POTENTIAL_CLIENT_COLUMNS.STATUS).setValue('Rejected');
    sheet.getRange(row, POTENTIAL_CLIENT_COLUMNS.DECISION_SOURCE).setValue(getManualDecisionSource_());
    ui.alert('Candidate rejected. The row remains available for audit and can be reopened later.');
    buildEnhancedDashboard();
  } catch (error) {
    ui.alert('Unable to Reject', error.message, ui.ButtonSet.OK);
  }
}

function menuReclassifySelectedPotentialClient() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getName() !== CONFIG.POTENTIAL_CLIENTS_SHEET_NAME) throw new Error('Select a row in Potential Clients.');
    const row = sheet.getActiveRange().getRow();
    if (row < 2) throw new Error('Select a candidate data row.');
    const values = sheet.getRange(row, 1, 1, POTENTIAL_CLIENT_HEADERS.length).getValues()[0];
    const result = classifyPotentialClientWithGemini_(
      values[POTENTIAL_CLIENT_COLUMNS.SUBJECT - 1],
      values[POTENTIAL_CLIENT_COLUMNS.SUMMARY - 1],
      values[POTENTIAL_CLIENT_COLUMNS.EMAIL - 1]
    );
    sheet.getRange(row, POTENTIAL_CLIENT_COLUMNS.INTENT_CATEGORY).setValue(result.intentCategory);
    sheet.getRange(row, POTENTIAL_CLIENT_COLUMNS.AI_CONFIDENCE).setValue(result.confidence);
    sheet.getRange(row, POTENTIAL_CLIENT_COLUMNS.STATUS).setValue(result.status);
    sheet.getRange(row, POTENTIAL_CLIENT_COLUMNS.REASON).setValue(result.reason);
    sheet.getRange(row, POTENTIAL_CLIENT_COLUMNS.DECISION_SOURCE).setValue('AI');
    ui.alert('✅ Reclassified', result.status + ' (' + result.confidence + '%): ' + result.reason, ui.ButtonSet.OK);
    buildEnhancedDashboard();
  } catch (error) {
    ui.alert('Unable to Reclassify', error.message, ui.ButtonSet.OK);
  }
}

function getManualDecisionSource_() {
  try {
    const email = Session.getActiveUser().getEmail();
    return email ? 'Manual: ' + email : 'Manual';
  } catch (ignore) {
    return 'Manual';
  }
}

function calculatePotentialMetrics_(rows) {
  const metrics = { total: 0, qualified: 0, review: 0, promoted: 0, rejected: 0, promotionRate: '0.0%' };
  rows.forEach(function(row) {
    const status = String(row[POTENTIAL_CLIENT_COLUMNS.STATUS - 1] || '');
    if (!status) return;
    metrics.total++;
    if (status === 'Qualified') metrics.qualified++;
    else if (status === 'Review') metrics.review++;
    else if (status === 'Promoted') metrics.promoted++;
    else if (status === 'Rejected') metrics.rejected++;
  });
  const promotable = metrics.qualified + metrics.promoted;
  metrics.promotionRate = promotable > 0 ? ((metrics.promoted / promotable) * 100).toFixed(1) + '%' : '0.0%';
  return metrics;
}

function displayPotentialCandidateSummary_(dashboard, ss) {
  dashboard.getRange('G5').setValue('🎯 Potential Client Funnel').setFontWeight('bold').setFontSize(14);
  const sheet = ss.getSheetByName(CONFIG.POTENTIAL_CLIENTS_SHEET_NAME);
  const rows = sheet && sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, POTENTIAL_CLIENT_HEADERS.length).getValues()
    : [];
  const metrics = calculatePotentialMetrics_(rows);
  dashboard.getRange(6, 7, 6, 2).setValues([
    ['Candidates', metrics.total],
    ['Qualified', metrics.qualified],
    ['Needs Review', metrics.review],
    ['Promoted', metrics.promoted],
    ['Rejected', metrics.rejected],
    ['Promotion Rate', metrics.promotionRate]
  ]);
  dashboard.getRange(6, 7, 6, 1).setFontWeight('bold');
}
