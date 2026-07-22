/**
 * Shared Gemini request governor.
 * All generateContent calls must reserve capacity here before using UrlFetchApp.
 */

function reserveAiRequest_(purpose, model) {
  purpose = String(purpose || 'general').toLowerCase();
  if (purpose !== 'configuration' && !isAiProcessingEnabled_()) {
    throw createAiBlockedError_('AI processing is disabled in the CRM menu.');
  }
  const now = new Date();
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getDocumentLock();
  if (!lock || !lock.tryLock(3000)) throw createAiBlockedError_('AI usage counter is busy. Try again shortly.');

  try {
    const cooldownUntil = Number(props.getProperty('AI_COOLDOWN_UNTIL') || 0);
    if (cooldownUntil > now.getTime()) {
      throw createAiBlockedError_('Gemini is cooling down until ' + new Date(cooldownUntil).toLocaleString() + '.');
    }

    const quotaTimezone = AI_CONFIG.AI_QUOTA_TIMEZONE || CONFIG.TIMEZONE;
    const minuteBucket = Utilities.formatDate(now, quotaTimezone, 'yyyyMMddHHmm');
    const dayBucket = Utilities.formatDate(now, quotaTimezone, 'yyyyMMdd');
    const minuteState = readAiCounter_(props, 'AI_USAGE_MINUTE', minuteBucket);
    const dayState = readAiCounter_(props, 'AI_USAGE_DAY', dayBucket);
    const purposeKey = 'AI_USAGE_DAY_' + purpose.toUpperCase().replace(/[^A-Z0-9_]/g, '_');
    const purposeState = readAiCounter_(props, purposeKey, dayBucket);

    const minuteLimit = getAiLimit_('AI_MAX_REQUESTS_PER_MINUTE', AI_CONFIG.MAX_AI_REQUESTS_PER_MINUTE);
    const dailyLimit = getAiLimit_('AI_MAX_REQUESTS_PER_DAY', AI_CONFIG.MAX_AI_REQUESTS_PER_DAY);
    const purposeLimit = getAiPurposeDailyLimit_(purpose);
    if (minuteState.count >= minuteLimit) throw createAiBlockedError_('Local Gemini per-minute limit reached.');
    if (dayState.count >= dailyLimit) throw createAiBlockedError_('Local Gemini daily limit reached.');
    if (purposeState.count >= purposeLimit) throw createAiBlockedError_('Daily Gemini limit reached for ' + purpose + '.');

    minuteState.count++;
    dayState.count++;
    purposeState.count++;
    props.setProperties({
      AI_USAGE_MINUTE: JSON.stringify(minuteState),
      AI_USAGE_DAY: JSON.stringify(dayState),
      [purposeKey]: JSON.stringify(purposeState),
      AI_LAST_REQUEST_AT: now.toISOString(),
      AI_LAST_REQUEST_PURPOSE: purpose,
      AI_LAST_REQUEST_MODEL: String(model || '')
    });
    return { minuteCount: minuteState.count, dayCount: dayState.count, purposeCount: purposeState.count };
  } finally {
    lock.releaseLock();
  }
}

function isAiProcessingEnabled_() {
  return PropertiesService.getScriptProperties().getProperty('AI_PROCESSING_ENABLED') !== 'false';
}

function menuToggleAiProcessing() {
  const ui = SpreadsheetApp.getUi();
  const enabled = isAiProcessingEnabled_();
  const response = ui.alert(
    enabled ? 'Disable AI Processing?' : 'Enable AI Processing?',
    enabled
      ? 'This pauses Gemini qualification and AI draft generation. Local FAQ matching and manual CRM work remain available.'
      : 'This resumes Gemini qualification and AI draft generation under the configured usage limits.',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;
  PropertiesService.getScriptProperties().setProperty('AI_PROCESSING_ENABLED', enabled ? 'false' : 'true');
  SpreadsheetApp.getActiveSpreadsheet().toast(
    enabled ? 'Gemini processing paused.' : 'Gemini processing enabled with usage limits.',
    'AI Usage Guard',
    6
  );
}

function readAiCounter_(props, key, bucket) {
  try {
    const parsed = JSON.parse(props.getProperty(key) || '{}');
    if (parsed.bucket === bucket) return { bucket: bucket, count: Math.max(0, Number(parsed.count) || 0) };
  } catch (ignore) {}
  return { bucket: bucket, count: 0 };
}

function getAiLimit_(propertyName, defaultValue) {
  const configured = Number(PropertiesService.getScriptProperties().getProperty(propertyName));
  return configured > 0 ? Math.floor(configured) : defaultValue;
}

function getAiPurposeDailyLimit_(purpose) {
  const defaults = {
    qualification: AI_CONFIG.MAX_AI_QUALIFICATION_REQUESTS_PER_DAY,
    auto_reply: AI_CONFIG.MAX_AI_AUTO_REPLY_REQUESTS_PER_DAY,
    configuration: AI_CONFIG.MAX_AI_CONFIGURATION_REQUESTS_PER_DAY,
    general: AI_CONFIG.MAX_AI_REQUESTS_PER_DAY
  };
  const key = 'AI_MAX_' + purpose.toUpperCase().replace(/[^A-Z0-9_]/g, '_') + '_REQUESTS_PER_DAY';
  return getAiLimit_(key, defaults[purpose] || AI_CONFIG.MAX_AI_REQUESTS_PER_DAY);
}

function createAiBlockedError_(message) {
  const error = new Error(message);
  error.aiBlocked = true;
  error.aiStopRoute = true;
  return error;
}

function recordAiHttpFailure_(statusCode) {
  const props = PropertiesService.getScriptProperties();
  const status = Number(statusCode) || 0;
  let delaySeconds = 0;
  let failures = Number(props.getProperty('AI_BACKOFF_FAILURES') || 0);

  if (status === 429) {
    failures++;
    delaySeconds = Math.min(15 * 60, 30 * Math.pow(2, Math.min(failures - 1, 5)));
  } else if (status === 401 || status === 403) {
    failures++;
    delaySeconds = 6 * 60 * 60;
  }

  if (delaySeconds > 0) {
    const jitterMs = Math.floor(Math.random() * 5000);
    props.setProperties({
      AI_BACKOFF_FAILURES: String(failures),
      AI_COOLDOWN_UNTIL: String(Date.now() + delaySeconds * 1000 + jitterMs),
      AI_LAST_RATE_ERROR: String(status),
      AI_LAST_RATE_ERROR_AT: new Date().toISOString()
    });
  }
}

function recordAiRequestSuccess_() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('AI_BACKOFF_FAILURES');
  props.deleteProperty('AI_COOLDOWN_UNTIL');
  props.setProperty('AI_LAST_SUCCESSFUL_REQUEST_AT', new Date().toISOString());
}

function shouldStopAiRoute_(error) {
  return Boolean(error && error.aiStopRoute);
}

function createAiCacheKey_(namespace, parts) {
  const input = String(namespace || '') + '\n' + (parts || []).map(function(part) {
    return String(part || '');
  }).join('\n---\n');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input);
  return 'AI_' + Utilities.base64EncodeWebSafe(digest).replace(/=+$/, '').substring(0, 40);
}

function getCachedAiJson_(key) {
  try {
    const value = CacheService.getScriptCache().get(key);
    return value ? JSON.parse(value) : null;
  } catch (ignore) {
    return null;
  }
}

function putCachedAiJson_(key, value, ttlSeconds) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(value), Math.min(Number(ttlSeconds) || 3600, 21600));
  } catch (error) {
    Logger.log('AI response cache write skipped: ' + error.message);
  }
}

function getAiWorkKey_(purpose, identity) {
  return createAiCacheKey_('work:' + purpose, [identity]);
}

function canAttemptAiWork_(purpose, identity) {
  const state = getCachedAiJson_(getAiWorkKey_(purpose, identity));
  return !state || Number(state.nextAttemptAt || 0) <= Date.now();
}

function recordAiWorkFailure_(purpose, identity) {
  const key = getAiWorkKey_(purpose, identity);
  const previous = getCachedAiJson_(key) || {};
  const attempts = Math.min(6, (Number(previous.attempts) || 0) + 1);
  const delays = [5, 15, 60, 180, 360, 360];
  putCachedAiJson_(key, {
    attempts: attempts,
    nextAttemptAt: Date.now() + delays[attempts - 1] * 60 * 1000
  }, 21600);
}

function clearAiWorkFailure_(purpose, identity) {
  try { CacheService.getScriptCache().remove(getAiWorkKey_(purpose, identity)); } catch (ignore) {}
}

function getAiDraftCheckpoint_(messageId) {
  const raw = PropertiesService.getScriptProperties().getProperty('AI_DRAFT_' + messageId);
  try { return raw ? JSON.parse(raw) : null; } catch (ignore) { return null; }
}

function saveAiDraftCheckpoint_(messageId, checkpoint) {
  PropertiesService.getScriptProperties().setProperty('AI_DRAFT_' + messageId, JSON.stringify(checkpoint));
}

function clearAiDraftCheckpoint_(messageId) {
  PropertiesService.getScriptProperties().deleteProperty('AI_DRAFT_' + messageId);
}

function getAiUsageSummary_() {
  const props = PropertiesService.getScriptProperties();
  const now = new Date();
  const quotaTimezone = AI_CONFIG.AI_QUOTA_TIMEZONE || CONFIG.TIMEZONE;
  const minute = readAiCounter_(props, 'AI_USAGE_MINUTE', Utilities.formatDate(now, quotaTimezone, 'yyyyMMddHHmm'));
  const dayBucket = Utilities.formatDate(now, quotaTimezone, 'yyyyMMdd');
  const day = readAiCounter_(props, 'AI_USAGE_DAY', dayBucket);
  const qualification = readAiCounter_(props, 'AI_USAGE_DAY_QUALIFICATION', dayBucket);
  const autoReply = readAiCounter_(props, 'AI_USAGE_DAY_AUTO_REPLY', dayBucket);
  const cooldown = Number(props.getProperty('AI_COOLDOWN_UNTIL') || 0);
  return {
    enabled: isAiProcessingEnabled_(),
    minute: minute.count,
    minuteLimit: getAiLimit_('AI_MAX_REQUESTS_PER_MINUTE', AI_CONFIG.MAX_AI_REQUESTS_PER_MINUTE),
    day: day.count,
    dayLimit: getAiLimit_('AI_MAX_REQUESTS_PER_DAY', AI_CONFIG.MAX_AI_REQUESTS_PER_DAY),
    qualification: qualification.count,
    qualificationLimit: getAiPurposeDailyLimit_('qualification'),
    autoReply: autoReply.count,
    autoReplyLimit: getAiPurposeDailyLimit_('auto_reply'),
    cooldownUntil: cooldown > Date.now() ? new Date(cooldown) : null,
    lastPurpose: props.getProperty('AI_LAST_REQUEST_PURPOSE') || 'None',
    lastModel: props.getProperty('AI_LAST_REQUEST_MODEL') || 'None'
  };
}

function menuShowAiUsage() {
  const usage = getAiUsageSummary_();
  SpreadsheetApp.getUi().alert(
    '🤖 AI Usage Guard',
    'This minute: ' + usage.minute + ' / ' + usage.minuteLimit + '\n' +
    'Processing: ' + (usage.enabled ? 'Enabled' : 'Paused') + '\n' +
    'Today: ' + usage.day + ' / ' + usage.dayLimit + '\n' +
    'Qualification today: ' + usage.qualification + ' / ' + usage.qualificationLimit + '\n' +
    'Auto-reply today: ' + usage.autoReply + ' / ' + usage.autoReplyLimit + '\n' +
    'Cooldown: ' + (usage.cooldownUntil ? usage.cooldownUntil.toLocaleString() : 'No') + '\n' +
    'Last purpose: ' + usage.lastPurpose + '\n' +
    'Last model: ' + usage.lastModel,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
