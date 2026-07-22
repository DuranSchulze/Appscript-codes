/**
 * Returns the account whose OAuth authority is being used for this execution.
 * For installable triggers this is the account that created the trigger.
 */
function getEffectiveAutomationAccount_() {
  try {
    return String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  } catch (error) {
    return '';
  }
}

function getConfiguredMonitoredMailbox_() {
  return String(CONFIG.MONITORED_MAILBOX_EMAIL || '').trim().toLowerCase();
}

/**
 * Fails closed before Gmail is read or modified under the wrong Google account.
 */
function assertMonitoredMailboxAccount_(operation) {
  const expected = getConfiguredMonitoredMailbox_();
  const effective = getEffectiveAutomationAccount_();
  if (!expected) throw new Error('CONFIG.MONITORED_MAILBOX_EMAIL is not configured.');
  if (!effective) {
    throw new Error('Could not verify the Google account running ' + operation + '. Mailbox access was blocked.');
  }
  if (effective !== expected) {
    throw new Error(
      'Mailbox access blocked for ' + operation + '. This execution is authorized as ' + effective +
      ', but Mini-CRM is locked to ' + expected + '. Sign in as ' + expected +
      ' and run CRM Tracker > AI Auto-Reply > Setup AI Triggers.'
    );
  }
  return expected;
}

function menuVerifyAutomationAccount() {
  const ui = SpreadsheetApp.getUi();
  const expected = getConfiguredMonitoredMailbox_();
  const effective = getEffectiveAutomationAccount_();
  const allowed = Boolean(expected && effective && expected === effective);
  ui.alert(
    allowed ? '✅ Automation account verified' : '⛔ Wrong automation account',
    'Required mailbox: ' + (expected || '(not configured)') + '\n' +
    'Current effective account: ' + (effective || '(unavailable)') + '\n\n' +
    (allowed
      ? 'This account may create triggers and access the CRM mailbox.'
      : 'Gmail sync, drafts, sends, and mailbox automation are blocked. Sign in with the required mailbox account before setting up triggers.'),
    ui.ButtonSet.OK
  );
  return { allowed: allowed, expected: expected, effective: effective };
}
