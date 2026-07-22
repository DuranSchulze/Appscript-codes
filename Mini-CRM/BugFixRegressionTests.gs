/**
 * Safe manual checks for the PLANS.md correctness fixes.
 * Does not send/delete drafts, search Gmail, or write spreadsheet rows.
 * Run runBugFixRegressionTests() from the Apps Script editor.
 */
function runBugFixRegressionTests() {
  testSignedWebAppActions_();
  testParameterizedBackfillDates_();

  const expectedYear = Number(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy'));
  assertBugFixTest_(getCurrentArchiveYear_() === expectedYear, 'Archive cutoff must use the current CRM year.');

  Logger.log('✓ PLANS.md bug-fix regression tests passed.');
  return true;
}

function testSignedWebAppActions_() {
  const draftId = 'regression-draft-' + Utilities.getUuid();
  const expires = Date.now() + (5 * 60 * 1000);
  const nonce = Utilities.getUuid();
  const signature = computeWebAppActionSignature_('approve', draftId, expires, nonce);
  const params = {
    action: 'approve',
    draftId: draftId,
    expires: String(expires),
    nonce: nonce,
    signature: signature
  };
  let claim = null;

  try {
    claim = validateAndClaimWebAppAction_(params);
    assertBugFixTest_(claim.action === 'approve', 'A valid signed action must be accepted.');
    assertBugFixTest_(throwsBugFixTest_(function() {
      validateAndClaimWebAppAction_(params);
    }), 'A signed action must be single-use.');
  } finally {
    if (claim) releaseWebAppActionClaim_(claim.tokenKey);
  }

  const tampered = Object.assign({}, params, { action: 'reject' });
  assertBugFixTest_(throwsBugFixTest_(function() {
    validateAndClaimWebAppAction_(tampered);
  }), 'Changing the action must invalidate the signature.');

  const expiredAt = Date.now() - 1000;
  const expired = {
    action: 'approve',
    draftId: draftId,
    expires: String(expiredAt),
    nonce: nonce,
    signature: computeWebAppActionSignature_('approve', draftId, expiredAt, nonce)
  };
  assertBugFixTest_(throwsBugFixTest_(function() {
    validateAndClaimWebAppAction_(expired);
  }), 'Expired action links must be rejected.');
}

function testParameterizedBackfillDates_() {
  const start = normalizeBackfillDate_('2024-02-29', false);
  const end = normalizeBackfillDate_('2024-03-01', true);
  assertBugFixTest_(start < end, 'A valid custom date range must parse.');
  assertBugFixTest_(throwsBugFixTest_(function() {
    normalizeBackfillDate_('2026-02-30', false);
  }), 'Invalid calendar dates must be rejected.');

  const query = buildBackfillGmailQuery_(start, end);
  assertBugFixTest_(query.indexOf('after:2024/02/28') !== -1, 'Backfill query must include the first requested day.');
  assertBugFixTest_(query.indexOf('before:2024/03/02') !== -1, 'Backfill query must include the last requested day.');
}

function throwsBugFixTest_(callback) {
  try {
    callback();
    return false;
  } catch (error) {
    return true;
  }
}

function assertBugFixTest_(condition, message) {
  if (!condition) throw new Error('Bug-fix regression failed: ' + message);
}
