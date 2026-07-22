/**
 * Manual regression checks for the Gmail ingestion correctness layer.
 * Run runEmailIngestionRegressionTests() from the Apps Script editor.
 */
function runEmailIngestionRegressionTests() {
  const mailbox = {
    primaryEmail: 'owner@example.com',
    addresses: new Set(['owner@example.com', 'alias@example.com'])
  };

  function fakeMessage(id, from, inbox) {
    return {
      getId: function() { return id; },
      getFrom: function() { return from; },
      isDraft: function() { return false; },
      isInTrash: function() { return false; },
      isInInbox: function() { return inbox !== false; }
    };
  }

  assertIngestionTest_(
    !inspectInboundMessage_(fakeMessage('out-1', 'Owner <owner@example.com>', true), mailbox).isInbound,
    'Primary-mailbox replies must be skipped.'
  );
  assertIngestionTest_(
    !inspectInboundMessage_(fakeMessage('out-2', 'Alias <alias@example.com>', true), mailbox).isInbound,
    'Alias replies must be skipped.'
  );
  assertIngestionTest_(
    inspectInboundMessage_(fakeMessage('in-1', 'Prospect <prospect@example.net>', true), mailbox).isInbound,
    'External Inbox messages must be accepted.'
  );
  assertIngestionTest_(
    !inspectInboundMessage_(fakeMessage('sent-1', 'Other <other@example.net>', false), mailbox).isInbound,
    'Messages outside Inbox must be skipped.'
  );

  const date = new Date('2026-07-22T01:00:00Z');
  const sender = 'prospect@example.net';
  const subject = 'Visa consultation';
  const fallbackId = createUniqueEmailId(sender, date, subject);

  let index = { messageIds: new Set(['gmail-1']), fallbackIds: new Set() };
  assertIngestionTest_(
    inspectDuplicateMessage_(fakeMessage('gmail-1', sender, true), sender, date, subject, index).isDuplicate,
    'The same Gmail message ID must be rejected as a duplicate.'
  );
  assertIngestionTest_(
    !inspectDuplicateMessage_(fakeMessage('gmail-2', sender, true), sender, date, subject, index).isDuplicate,
    'A different Gmail ID must remain authoritative even when legacy fields match.'
  );

  index = { messageIds: new Set(), fallbackIds: new Set([fallbackId]) };
  assertIngestionTest_(
    inspectDuplicateMessage_(fakeMessage('gmail-legacy', sender, true), sender, date, subject, index).isDuplicate,
    'The fallback key must recognize a pre-message-ID legacy row.'
  );

  index = { messageIds: new Set(), fallbackIds: new Set() };
  recordMessageIdentity_(index, { messageId: 'gmail-new', fallbackId: fallbackId });
  assertIngestionTest_(index.messageIds.has('gmail-new'), 'New Gmail IDs must be recorded in-run.');
  assertIngestionTest_(!index.fallbackIds.has(fallbackId), 'New ID-backed messages must not populate the legacy index.');

  const idBackedRow = new Array(MONTHLY_EMAIL_HEADERS.length).fill('');
  idBackedRow[0] = date;
  idBackedRow[4] = sender;
  idBackedRow[5] = subject;
  idBackedRow[15] = 'stored-gmail-id';
  const legacyRow = new Array(10).fill('');
  legacyRow[0] = new Date('2026-07-21T01:00:00Z');
  legacyRow[4] = 'legacy@example.net';
  legacyRow[5] = 'Legacy inquiry';
  const workbook = {
    getSheets: function() {
      return [
        fakeMonthlySheetForIdentityTest_('Jul-2026', [idBackedRow], 18),
        fakeMonthlySheetForIdentityTest_('Jun-2026', [legacyRow], 10),
        fakeMonthlySheetForIdentityTest_('Dashboard', [], 10)
      ];
    }
  };
  index = getExistingEmailIdentityIndex_(workbook);
  assertIngestionTest_(index.messageIds.has('stored-gmail-id'), 'Workbook index must load stored Gmail IDs.');
  assertIngestionTest_(index.fallbackIds.size === 1, 'Only pre-ID workbook rows should enter the legacy index.');

  Logger.log('✓ Gmail ingestion regression tests passed.');
  return true;
}

function fakeMonthlySheetForIdentityTest_(name, rows, maxColumns) {
  return {
    getName: function() { return name; },
    getLastRow: function() { return rows.length + 1; },
    getMaxColumns: function() { return maxColumns; },
    getRange: function(row, column, rowCount, columnCount) {
      return {
        getValues: function() {
          return rows.slice(0, rowCount).map(function(values) {
            return values.slice(0, columnCount);
          });
        }
      };
    }
  };
}

function assertIngestionTest_(condition, message) {
  if (!condition) throw new Error('Ingestion regression failed: ' + message);
}
