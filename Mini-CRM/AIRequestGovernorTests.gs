/**
 * Safe manual checks for AI cache/backoff helpers.
 * This does not consume Gemini quota. Run runAiGovernorRegressionTests().
 */
function runAiGovernorRegressionTests() {
  const blocked = createAiBlockedError_('test');
  assertAiGovernorTest_(blocked.aiBlocked === true, 'Blocked errors must be identifiable.');
  assertAiGovernorTest_(shouldStopAiRoute_(blocked), 'Blocked requests must stop fallback routing.');

  const fakeProps = {
    getProperty: function() { return JSON.stringify({ bucket: 'old', count: 99 }); }
  };
  const resetCounter = readAiCounter_(fakeProps, 'unused', 'new');
  assertAiGovernorTest_(resetCounter.count === 0, 'Usage counters must reset when their time bucket changes.');

  const identity = 'governor-test-' + new Date().getTime();
  assertAiGovernorTest_(canAttemptAiWork_('test', identity), 'Fresh work must be eligible.');
  recordAiWorkFailure_('test', identity);
  assertAiGovernorTest_(!canAttemptAiWork_('test', identity), 'Failed work must enter backoff.');
  clearAiWorkFailure_('test', identity);
  assertAiGovernorTest_(canAttemptAiWork_('test', identity), 'Cleared work must become eligible again.');

  const cacheKey = createAiCacheKey_('governor-test', [identity]);
  putCachedAiJson_(cacheKey, { ok: true }, 60);
  const cached = getCachedAiJson_(cacheKey);
  assertAiGovernorTest_(cached && cached.ok === true, 'AI JSON cache must round-trip valid data.');
  CacheService.getScriptCache().remove(cacheKey);

  const messageId = 'governor-test-message-' + new Date().getTime();
  saveAiDraftCheckpoint_(messageId, { draftId: 'draft-test', category: 'General' });
  const checkpoint = getAiDraftCheckpoint_(messageId);
  assertAiGovernorTest_(checkpoint && checkpoint.draftId === 'draft-test', 'Draft checkpoint must be recoverable.');
  clearAiDraftCheckpoint_(messageId);
  assertAiGovernorTest_(!getAiDraftCheckpoint_(messageId), 'Draft checkpoint must be removable after completion.');

  const automatedReason = getAutomatedMessageReason_({
    getHeader: function(name) { return name === 'List-Unsubscribe' ? '<mailto:unsubscribe@example.com>' : ''; }
  });
  assertAiGovernorTest_(Boolean(automatedReason), 'Bulk/list mail must be rejected before an AI request.');

  Logger.log('✓ AI request governor regression tests passed without calling Gemini.');
  return true;
}

function assertAiGovernorTest_(condition, message) {
  if (!condition) throw new Error('AI governor regression failed: ' + message);
}
