const WEB_APP_ACTION_SECRET_PROPERTY = 'WEB_APP_ACTION_SECRET';
const WEB_APP_USED_TOKEN_PREFIX = 'WEB_APP_USED_TOKEN_';
const WEB_APP_ALLOWED_ACTIONS = ['approve', 'reject'];

/**
 * Validates the signed link and shows a confirmation page. GET never mutates a
 * Gmail draft, which prevents link-preview scanners from sending email.
 */
function doGet(e) {
  try {
    const action = validateWebAppActionRequest_((e && e.parameter) || {});
    return createWebAppConfirmation_(action);
  } catch (error) {
    console.error('Web app draft link validation failed: ' + (error && error.stack ? error.stack : error));
    return createUnavailableWebAppMessage_();
  }
}

/**
 * Claims the single-use token and performs the confirmed Gmail mutation.
 */
function doPost(e) {
  let claim = null;
  try {
    assertMonitoredMailboxAccount_('web app draft action');
    claim = validateAndClaimWebAppAction_((e && e.parameter) || {});
    const output = claim.action === 'approve'
      ? handleApprove(claim.draftId)
      : handleReject(claim.draftId);
    return output;
  } catch (error) {
    // Let a valid link be retried if Gmail failed after the token was claimed.
    if (claim && claim.tokenKey) releaseWebAppActionClaim_(claim.tokenKey);
    console.error('Web app draft action failed: ' + (error && error.stack ? error.stack : error));
    return createUnavailableWebAppMessage_();
  }
}

function handleApprove(draftId) {
  assertMonitoredMailboxAccount_('draft approval');
  const draft = GmailApp.getDraft(draftId);
  if (!draft) throw new Error('Draft not found.');

  // Sending the draft itself preserves HTML, attachments, CC, and BCC.
  draft.send();
  return createWebAppMessage_('Email sent', 'The reviewed Gmail draft was sent successfully.', true);
}

function handleReject(draftId) {
  assertMonitoredMailboxAccount_('draft rejection');
  const draft = GmailApp.getDraft(draftId);
  if (!draft) throw new Error('Draft not found.');
  draft.deleteDraft();
  return createWebAppMessage_('Draft discarded', 'The Gmail draft was discarded.', true);
}

function createSecureDraftActionUrl_(action, draftId) {
  if (WEB_APP_ALLOWED_ACTIONS.indexOf(action) === -1) throw new Error('Unsupported draft action.');
  if (!draftId) throw new Error('A Gmail draft ID is required.');

  const serviceUrl = ScriptApp.getService().getUrl();
  if (!serviceUrl) {
    throw new Error('Deploy the Apps Script project as a web app before sending approval cards.');
  }

  const expires = Date.now() + (CONFIG.WEB_APP_ACTION_TOKEN_TTL_HOURS * 60 * 60 * 1000);
  const nonce = Utilities.getUuid();
  const signature = computeWebAppActionSignature_(action, String(draftId), expires, nonce);
  const params = {
    action: action,
    draftId: String(draftId),
    expires: String(expires),
    nonce: nonce,
    signature: signature
  };
  const query = Object.keys(params).map(function(key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
  }).join('&');
  return serviceUrl + '?' + query;
}

function validateAndClaimWebAppAction_(params) {
  const validated = validateWebAppActionRequest_(params);
  const tokenKey = getWebAppUsedTokenKey_(validated.signature);
  claimWebAppActionToken_(tokenKey, validated.expires);
  return { action: validated.action, draftId: validated.draftId, tokenKey: tokenKey };
}

function validateWebAppActionRequest_(params) {
  const action = String(params.action || '');
  const draftId = String(params.draftId || '');
  const expiresText = String(params.expires || '');
  const nonce = String(params.nonce || '');
  const signature = String(params.signature || '');
  const expires = Number(expiresText);
  const now = Date.now();
  const maxFuture = now + (CONFIG.WEB_APP_ACTION_TOKEN_TTL_HOURS * 60 * 60 * 1000) + (5 * 60 * 1000);

  if (WEB_APP_ALLOWED_ACTIONS.indexOf(action) === -1 || !draftId || !nonce || !signature) {
    throw new Error('Missing or invalid action parameters.');
  }
  if (!/^\d+$/.test(expiresText) || !isFinite(expires) || expires <= now || expires > maxFuture) {
    throw new Error('The action link has expired or has an invalid expiry.');
  }

  const expected = computeWebAppActionSignature_(action, draftId, expires, nonce);
  if (!constantTimeStringEquals_(signature, expected)) {
    throw new Error('Invalid action signature.');
  }

  return {
    action: action,
    draftId: draftId,
    expires: expires,
    expiresText: expiresText,
    nonce: nonce,
    signature: signature
  };
}

function computeWebAppActionSignature_(action, draftId, expires, nonce) {
  const payload = [action, draftId, String(expires), nonce].join('\n');
  const bytes = Utilities.computeHmacSha256Signature(payload, getOrCreateWebAppActionSecret_());
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

function getOrCreateWebAppActionSecret_() {
  const properties = PropertiesService.getScriptProperties();
  let secret = properties.getProperty(WEB_APP_ACTION_SECRET_PROPERTY);
  if (secret) return secret;

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    secret = properties.getProperty(WEB_APP_ACTION_SECRET_PROPERTY);
    if (!secret) {
      secret = Utilities.getUuid() + Utilities.getUuid();
      properties.setProperty(WEB_APP_ACTION_SECRET_PROPERTY, secret);
    }
    return secret;
  } finally {
    lock.releaseLock();
  }
}

function claimWebAppActionToken_(tokenKey, expires) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const properties = PropertiesService.getScriptProperties();
    cleanupExpiredWebAppActionTokens_(properties, Date.now());
    if (properties.getProperty(tokenKey)) throw new Error('This action link has already been used.');
    properties.setProperty(tokenKey, String(expires));
  } finally {
    lock.releaseLock();
  }
}

function releaseWebAppActionClaim_(tokenKey) {
  PropertiesService.getScriptProperties().deleteProperty(tokenKey);
}

function cleanupExpiredWebAppActionTokens_(properties, now) {
  const allProperties = properties.getProperties();
  Object.keys(allProperties).forEach(function(key) {
    if (key.indexOf(WEB_APP_USED_TOKEN_PREFIX) === 0 && Number(allProperties[key]) <= now) {
      properties.deleteProperty(key);
    }
  });
}

function getWebAppUsedTokenKey_(signature) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, signature);
  return WEB_APP_USED_TOKEN_PREFIX + Utilities.base64EncodeWebSafe(digest).replace(/=+$/, '');
}

function constantTimeStringEquals_(left, right) {
  const a = String(left || '');
  const b = String(right || '');
  let mismatch = a.length ^ b.length;
  const length = Math.max(a.length, b.length);
  for (let i = 0; i < length; i++) {
    mismatch |= (a.charCodeAt(i % Math.max(a.length, 1)) || 0) ^
      (b.charCodeAt(i % Math.max(b.length, 1)) || 0);
  }
  return mismatch === 0;
}

function createWebAppMessage_(title, message, success) {
  const accent = success ? '#188038' : '#b3261e';
  const safeTitle = escapeWebAppHtml_(title);
  const safeMessage = escapeWebAppHtml_(message);
  return HtmlService.createHtmlOutput(
    '<!doctype html><html><head><base target="_top"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>body{font:16px Arial,sans-serif;background:#f6f8fc;margin:0;padding:40px 20px;color:#202124}' +
    '.card{max-width:520px;margin:auto;background:white;border-radius:16px;padding:28px;box-shadow:0 4px 20px #0002}' +
    'h2{color:' + accent + ';margin-top:0}p{line-height:1.5;margin-bottom:0}</style></head>' +
    '<body><main class="card"><h2>' + safeTitle + '</h2><p>' + safeMessage + '</p></main></body></html>'
  ).setTitle(title);
}

function createWebAppConfirmation_(actionRequest) {
  const isApprove = actionRequest.action === 'approve';
  const title = isApprove ? 'Confirm email send' : 'Confirm draft discard';
  const message = isApprove
    ? 'Send this reviewed Gmail draft now?'
    : 'Permanently discard this Gmail draft?';
  const buttonText = isApprove ? 'Send email' : 'Discard draft';
  const accent = isApprove ? '#188038' : '#b3261e';
  const serviceUrl = ScriptApp.getService().getUrl();
  const fields = {
    action: actionRequest.action,
    draftId: actionRequest.draftId,
    expires: actionRequest.expiresText,
    nonce: actionRequest.nonce,
    signature: actionRequest.signature
  };
  const hiddenInputs = Object.keys(fields).map(function(key) {
    return '<input type="hidden" name="' + escapeWebAppHtml_(key) + '" value="' +
      escapeWebAppHtml_(fields[key]) + '">';
  }).join('');

  return HtmlService.createHtmlOutput(
    '<!doctype html><html><head><base target="_top"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>body{font:16px Arial,sans-serif;background:#f6f8fc;margin:0;padding:40px 20px;color:#202124}' +
    '.card{max-width:520px;margin:auto;background:white;border-radius:16px;padding:28px;box-shadow:0 4px 20px #0002}' +
    'h2{margin-top:0}p{line-height:1.5}.button{border:0;border-radius:8px;background:' + accent +
    ';color:white;font-size:16px;font-weight:600;padding:12px 20px;cursor:pointer}</style></head>' +
    '<body><main class="card"><h2>' + escapeWebAppHtml_(title) + '</h2><p>' + escapeWebAppHtml_(message) + '</p>' +
    '<form method="post" action="' + escapeWebAppHtml_(serviceUrl) + '">' + hiddenInputs +
    '<button class="button" type="submit">' + escapeWebAppHtml_(buttonText) + '</button></form></main></body></html>'
  ).setTitle(title);
}

function createUnavailableWebAppMessage_() {
  return createWebAppMessage_(
    'Link unavailable',
    'This approval link is invalid, expired, already used, or the draft is no longer available.',
    false
  );
}

function escapeWebAppHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
