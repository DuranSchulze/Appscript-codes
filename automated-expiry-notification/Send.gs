// ═══════════════════════════════════════════════════════════════════════════
// SEND — outbound email pipeline (rules → compose → AI → alias → deliver)
// ═══════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════
// SendRules — send-mode (Auto/Hold/Manual Only) helpers
// ═══════════════════════════════════════════════════════════════════════════

// 31 Notification Rules

function normalizeSendMode(modeValue) {
  var text = String(modeValue || "")
    .trim()
    .toLowerCase();
  if (!text) return SEND_MODE.AUTO;
  if (text === "auto") return SEND_MODE.AUTO;
  if (text === "hold") return SEND_MODE.HOLD;
  if (text === "manual only" || text === "manual-only" || text === "manual") {
    return SEND_MODE.MANUAL_ONLY;
  }
  return SEND_MODE.AUTO;
}

function getRowSendMode(row, colMap) {
  if (!colMap.SEND_MODE) return SEND_MODE.AUTO;
  return normalizeSendMode(getCellStr(row, colMap.SEND_MODE));
}

function getSendModeSkipReason(sendMode) {
  if (sendMode === SEND_MODE.HOLD) {
    return "Skipped by Send Mode: Hold";
  }
  if (sendMode === SEND_MODE.MANUAL_ONLY) {
    return "Skipped by Send Mode: Manual Only";
  }
  return "";
}

// ═══════════════════════════════════════════════════════════════════════════
// AliasResolver — verifies row Assigned Staff Email is a Gmail Send-As alias
// ═══════════════════════════════════════════════════════════════════════════

// 36 Send Alias Resolver
//
// Per-row outgoing emails set the From address to the row's Assigned Staff
// Email so the client sees the staff member as the sender. Gmail only allows
// this when the address is registered as a "Send mail as" alias on the
// script-runner's account. This module gates the send on that check.

var __ALIAS_CACHE = null;

function getVerifiedSenderAliases() {
  if (__ALIAS_CACHE) return __ALIAS_CACHE;

  var lookup = {};

  function add(email) {
    var normalized = String(email || "")
      .trim()
      .toLowerCase();
    if (normalized) lookup[normalized] = true;
  }

  try {
    add(Session.getEffectiveUser().getEmail());
  } catch (e) {}
  try {
    add(Session.getActiveUser().getEmail());
  } catch (e) {}

  try {
    var aliases = GmailApp.getAliases() || [];
    for (var i = 0; i < aliases.length; i++) add(aliases[i]);
  } catch (e) {}

  __ALIAS_CACHE = lookup;
  return lookup;
}

function canSendAs(email) {
  var normalized = String(email || "")
    .trim()
    .toLowerCase();
  if (!normalized) return false;
  var lookup = getVerifiedSenderAliases();
  return !!lookup[normalized];
}

function listVerifiedSenderAliases() {
  var lookup = getVerifiedSenderAliases();
  var list = [];
  for (var key in lookup) list.push(key);
  return list;
}

function resetVerifiedSenderAliasCache() {
  __ALIAS_CACHE = null;
}

// ═══════════════════════════════════════════════════════════════════════════
// AiGeneration — Gemini settings, fetch, fallback template controls
// ═══════════════════════════════════════════════════════════════════════════

// 35 Ai Generation

function isAiGenerationEnabled() {
  return getPropBoolean(PROP_KEYS.AI_ENABLED, false);
}

function getAiModel() {
  return getPropString(PROP_KEYS.AI_MODEL, DEFAULT_AI_MODEL);
}

function getAiApiKey() {
  return getPropString(PROP_KEYS.AI_API_KEY, "");
}

function getFallbackTemplateMode() {
  var mode = getPropString(
    PROP_KEYS.FALLBACK_TEMPLATE_MODE,
    FALLBACK_TEMPLATE_MODE.HARDCODED,
  ).toUpperCase();
  return mode === FALLBACK_TEMPLATE_MODE.PROPERTY
    ? FALLBACK_TEMPLATE_MODE.PROPERTY
    : FALLBACK_TEMPLATE_MODE.HARDCODED;
}

function getConfiguredFallbackTemplate() {
  return getPropString(PROP_KEYS.FALLBACK_TEMPLATE, "");
}

function maskSecret(value) {
  var text = String(value || "");
  if (!text) return "(not set)";
  if (text.length <= 8) return "********";
  return text.substring(0, 4) + "..." + text.substring(text.length - 4);
}

function tryGenerateAiEmailBody(clientName, expiryDate, docTypeText) {
  var apiKey = getAiApiKey();
  var model = getAiModel();
  if (!apiKey || !model) return null;

  var modelPath = model;
  if (modelPath.indexOf("models/") !== 0) {
    modelPath = "models/" + modelPath;
  }

  var prompt = [
    "Write a concise but professional visa/document expiry reminder email body.",
    "Tone: courteous, formal, and actionable.",
    "Do not include a subject line.",
    "Use plain text with short paragraphs.",
    "Client Name: " + clientName,
    "Document Type: " + docTypeText,
    "Expiry Date: " + formatDate(expiryDate),
  ].join("\n");

  var endpoint =
    "https://generativelanguage.googleapis.com/v1beta/" +
    modelPath +
    ":generateContent?key=" +
    encodeURIComponent(apiKey);

  var payload = {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.5,
      maxOutputTokens: 400,
    },
  };

  for (var attempt = 1; attempt <= 2; attempt++) {
    try {
      var response = UrlFetchApp.fetch(endpoint, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      });

      var statusCode = response.getResponseCode();
      var body = response.getContentText() || "";
      if (statusCode < 200 || statusCode >= 300) {
        if (attempt < 2) {
          Utilities.sleep(350);
          continue;
        }
        return null;
      }

      var parsed = JSON.parse(body);
      var text = extractGeminiText(parsed);
      if (!text) {
        if (attempt < 2) {
          Utilities.sleep(350);
          continue;
        }
        return null;
      }

      return {
        text: text.trim(),
        model: modelPath,
      };
    } catch (e) {
      if (attempt < 2) {
        Utilities.sleep(350);
        continue;
      }
      return null;
    }
  }

  return null;
}

function extractGeminiText(payload) {
  if (!payload || !payload.candidates || payload.candidates.length === 0) {
    return "";
  }

  var candidate = payload.candidates[0];
  var parts =
    candidate && candidate.content && candidate.content.parts
      ? candidate.content.parts
      : [];

  var textChunks = [];
  for (var i = 0; i < parts.length; i++) {
    var partText = String(parts[i].text || "").trim();
    if (partText) textChunks.push(partText);
  }
  return textChunks.join("\n\n");
}

function toggleAiGeneration() {
  var enabled = !isAiGenerationEnabled();
  setPropBoolean(PROP_KEYS.AI_ENABLED, enabled);
  SpreadsheetApp.getUi().alert(
    "AI Integration",
    "AI generation is now " + (enabled ? "ENABLED" : "DISABLED") + ".",
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

function setGeminiApiKey() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    "Set Gemini API Key",
    "Enter Gemini API key. Leave blank to clear.",
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var key = response.getResponseText().trim();
  setPropString(PROP_KEYS.AI_API_KEY, key);
  setPropString(PROP_KEYS.AI_PROVIDER, AI_PROVIDER.GEMINI);
  ui.alert(
    "AI Integration",
    key ? "Gemini key saved." : "Gemini key cleared.",
    ui.ButtonSet.OK,
  );
}

function selectGeminiModel() {
  var ui = SpreadsheetApp.getUi();
  var current = getAiModel();
  var apiKey = getAiApiKey();
  var models = apiKey ? listGeminiModels(apiKey) : [];

  var promptText =
    "Enter Gemini model name (e.g., models/gemini-1.5-flash).\n\nCurrent: " +
    current;
  if (models.length > 0) {
    promptText += "\n\nAvailable:\n- " + models.slice(0, 10).join("\n- ");
  }

  var response = ui.prompt(
    "Select Gemini Model",
    promptText,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var model = response.getResponseText().trim();
  if (!model) {
    ui.alert("Model cannot be empty.");
    return;
  }
  setPropString(PROP_KEYS.AI_MODEL, model);
  ui.alert("AI model saved: " + model);
}

function listGeminiModels(apiKey) {
  try {
    var response = UrlFetchApp.fetch(
      "https://generativelanguage.googleapis.com/v1beta/models?key=" +
        encodeURIComponent(apiKey),
      { muteHttpExceptions: true },
    );
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
      return [];
    }
    var payload = JSON.parse(response.getContentText() || "{}");
    var models = payload.models || [];
    return models
      .map(function (item) {
        return String(item.name || "").trim();
      })
      .filter(function (name) {
        return name.indexOf("models/gemini") === 0;
      });
  } catch (e) {
    return [];
  }
}

function testAiConnection() {
  var ui = SpreadsheetApp.getUi();
  var result = tryGenerateAiEmailBody("Test Client", new Date(), "Visa/Permit");
  if (result && result.text) {
    ui.alert(
      "AI Integration",
      "AI connection successful. Model: " + result.model,
      ui.ButtonSet.OK,
    );
  } else {
    ui.alert(
      "AI Integration",
      "AI test failed. Check API key/model and try again.",
      ui.ButtonSet.OK,
    );
  }
}

function showAiStatus() {
  var ui = SpreadsheetApp.getUi();
  var msg = [
    "AI Enabled: " + (isAiGenerationEnabled() ? "YES" : "NO"),
    "Provider: " + getPropString(PROP_KEYS.AI_PROVIDER, AI_PROVIDER.GEMINI),
    "Model: " + getAiModel(),
    "API Key: " + maskSecret(getAiApiKey()),
    "Fallback Source: " + getFallbackTemplateMode(),
    "Open Tracking URL: " + (getOpenTrackingBaseUrl() || "(not set)"),
  ].join("\n");
  ui.alert("AI Status", msg, ui.ButtonSet.OK);
}

function setFallbackTemplate() {
  var ui = SpreadsheetApp.getUi();
  var current = getConfiguredFallbackTemplate();
  var response = ui.prompt(
    "Set Fallback Template",
    "Enter fallback body template text.\n\nCurrent value (first 300 chars):\n" +
      (current ? current.substring(0, 300) : "(empty)"),
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  setPropString(PROP_KEYS.FALLBACK_TEMPLATE, response.getResponseText());
  ui.alert("Fallback template saved.");
}

function toggleFallbackTemplateSource() {
  var current = getFallbackTemplateMode();
  var next =
    current === FALLBACK_TEMPLATE_MODE.HARDCODED
      ? FALLBACK_TEMPLATE_MODE.PROPERTY
      : FALLBACK_TEMPLATE_MODE.HARDCODED;
  setPropString(PROP_KEYS.FALLBACK_TEMPLATE_MODE, next);
  SpreadsheetApp.getUi().alert(
    "Fallback Source",
    "Fallback template mode is now: " + next,
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// Attachments — Drive attachment parsing + fallback link handling
// ═══════════════════════════════════════════════════════════════════════════

// 34 Attachments

function extractDriveFileId(urlOrId) {
  if (!urlOrId) return null;
  var s = urlOrId.trim();

  // /file/d/FILE_ID/
  var m = s.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];

  // ?id=FILE_ID or &id=FILE_ID
  m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];

  // Assume raw ID if it looks like one (alphanumeric + _ -)
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s)) return s;

  return null;
}

function splitAttachmentEntries(rawField) {
  if (!rawField || rawField.trim() === "") return [];
  // Normalize: treat newlines as delimiters alongside commas
  var normalized = String(rawField).replace(/[\r\n]+/g, ",");
  return normalized
    .split(",")
    .map(function (e) {
      return e.trim();
    })
    .filter(function (e) {
      return !!e;
    });
}

function resolveAttachments(rawField) {
  if (!rawField || rawField.trim() === "") {
    return { blobs: [], warnings: [], failedLinks: [], fatalError: null };
  }

  var entries = splitAttachmentEntries(rawField);
  if (entries.length === 0) {
    return { blobs: [], warnings: [], failedLinks: [], fatalError: null };
  }

  var blobs = [];
  var warnings = [];
  var failedLinks = [];

  for (var i = 0; i < entries.length; i++) {
    var entry = entries[i];
    var fileId = extractDriveFileId(entry);

    if (!fileId) {
      warnings.push('Cannot parse Drive file ID from: "' + entry + '"');
      failedLinks.push({ label: entry, url: null });
      continue;
    }

    try {
      var file = DriveApp.getFileById(fileId);
      blobs.push({
        blob: file.getBlob(),
        name: file.getName(),
        fileId: fileId,
      });
    } catch (e) {
      warnings.push(
        "File not found or inaccessible (ID: " + fileId + "): " + e.message,
      );
      var originalUrl =
        entry.indexOf("drive.google.com") >= 0
          ? entry
          : "https://drive.google.com/file/d/" + fileId + "/view";
      failedLinks.push({ label: fileId, url: originalUrl });
    }
  }

  return {
    blobs: blobs,
    warnings: warnings,
    failedLinks: failedLinks,
    fatalError: null,
  };
}

function buildFallbackLinksHtml(failedLinks) {
  if (!failedLinks || failedLinks.length === 0) return "";

  var items = [];
  for (var i = 0; i < failedLinks.length; i++) {
    var fl = failedLinks[i];
    if (fl.url) {
      items.push(
        '<li><a href="' +
          sanitizeHtmlAttribute(fl.url) +
          '" target="_blank" rel="noopener noreferrer">' +
          sanitizeHtmlContent(fl.label) +
          "</a></li>",
      );
    } else {
      items.push("<li>" + sanitizeHtmlContent(fl.label) + "</li>");
    }
  }

  return (
    '<div style="margin-top:16px;padding:12px;background:#fff8e1;border-left:3px solid #f9a825;font-size:13px;">' +
    '<p style="margin:0 0 8px 0;font-weight:bold;color:#7a5c00;">Some files could not be attached. You can access them via the links below:</p>' +
    '<ul style="margin:0;padding-left:20px;">' +
    items.join("") +
    "</ul></div>"
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// EmailCompose — subject/body composition + token replacement
// ═══════════════════════════════════════════════════════════════════════════

// 32 Email Content

function getDefaultFallbackTemplate() {
  return (
    "Good day, [Client Name],\n\n" +
    "This is a reminder that your [Document Type] is expiring on [Expiry Date].\n\n" +
    "Please take the necessary steps before the expiry date.\n\n" +
    "Thank you."
  );
}

function resolveFallbackTemplateText() {
  var mode = getFallbackTemplateMode();
  if (mode === FALLBACK_TEMPLATE_MODE.PROPERTY) {
    var configured = getConfiguredFallbackTemplate();
    if (configured) {
      return { text: configured, source: "Fallback(Property)" };
    }
  }
  return { text: getDefaultFallbackTemplate(), source: "Fallback(Hardcoded)" };
}

function buildEmailContent(
  remarks,
  clientName,
  expiryDate,
  docType,
  openToken,
  templateContext,
) {
  var expiryStr = formatDate(expiryDate);
  var docTypeText = docType ? docType : "Visa/Permit";
  var bodyText = "";
  var source = "";

  if (remarks) {
    bodyText = applyTemplatePlaceholders(
      remarks,
      clientName,
      expiryStr,
      docTypeText,
      templateContext,
    );
    source = "Remarks";
  } else {
    var fallback = resolveFallbackTemplateText();
    bodyText = applyTemplatePlaceholders(
      fallback.text,
      clientName,
      expiryStr,
      docTypeText,
      templateContext,
    );
    source = fallback.source;
  }

  var htmlBody = String(bodyText || "").replace(/\n/g, "<br>");
  htmlBody = injectOpenTrackingPixel(htmlBody, openToken);

  // Neutral-gray card wrapper. Inline styles only (Gmail/Outlook safe).
  // Outer page background + centered white card with thin border.
  htmlBody = [
    '<div style="background:#f5f5f5;padding:24px 16px;font-family:Arial,Helvetica,sans-serif;color:#333333;">',
    '<div style="max-width:600px;margin:0 auto;background:#ffffff;border:1px solid #e5e7eb;border-radius:6px;padding:32px;font-size:14px;line-height:1.6;color:#333333;">',
    htmlBody,
    '<div style="border-top:1px solid #eeeeee;margin:24px 0 0 0;padding-top:16px;font-size:11px;color:#999999;">',
    "This is an automated reminder. Email opens may be tracked for delivery visibility.",
    "</div>",
    "</div>",
    "</div>",
  ].join("");

  return {
    htmlBody: htmlBody,
    textBody: bodyText,
    source: source || "Unknown",
  };
}

function buildEmailBody(remarks, clientName, expiryDate, docType) {
  return buildEmailContent(remarks, clientName, expiryDate, docType, "", null)
    .htmlBody;
}

function buildStageSubject(baseSubject, stage) {
  if (stage === "final") {
    return "Final Reminder: " + baseSubject + " (Expires Today)";
  }
  return baseSubject;
}

function buildStageEmailContent(
  remarks,
  clientName,
  expiryDate,
  docType,
  openToken,
  templateContext,
  stage,
) {
  var content = buildEmailContent(
    remarks,
    clientName,
    expiryDate,
    docType,
    openToken,
    templateContext,
  );

  if (stage === "final") {
    content.htmlBody =
      "<p><strong>This is your final reminder. The document expires today.</strong></p>" +
      content.htmlBody;
  }

  return content;
}

function applyTemplatePlaceholders(
  templateText,
  clientName,
  expiryStr,
  docType,
  templateContext,
) {
  var rendered = String(templateText || "")
    .replace(/\[\s*client\s*name\s*\]/gi, clientName || "")
    .replace(/\[\s*date\s*of\s*(expiration|expiry)\s*\]/gi, expiryStr || "")
    .replace(/\[\s*expiry\s*date\s*\]/gi, expiryStr || "")
    .replace(/\[\s*document\s*type\s*\]/gi, docType || "Visa/Permit");

  return replaceGenericTemplateTokens(rendered, templateContext);
}

function formatTemplateValue(value) {
  if (value instanceof Date) {
    if (!isNaN(value.getTime())) return formatDate(value);
    return "";
  }
  return String(value === null || value === undefined ? "" : value).trim();
}

function buildRowTemplateContext(sheet, tabName, rowValues) {
  var context = {};
  if (!sheet || !rowValues) return context;

  var headerInfo = resolveEffectiveHeaderRow(sheet, tabName);
  var headers = headerInfo.headers || [];
  var lastCol = sheet.getLastColumn();
  if (headers.length === 0 && lastCol > 0) {
    headers = getRowValues(sheet, headerInfo.headerRow, lastCol);
  }

  for (var i = 0; i < headers.length && i < rowValues.length; i++) {
    var headerText = String(headers[i] || "").trim();
    if (!headerText) continue;

    var rawValue = rowValues[i];
    var formattedValue = formatTemplateValue(rawValue);
    if (!formattedValue && formattedValue !== "0") continue;

    context[normalizeHeaderName(headerText)] = formattedValue;
  }

  return context;
}

function replaceGenericTemplateTokens(templateText, templateContext) {
  var context = templateContext || {};
  return String(templateText || "").replace(
    /\[([^\]]+)\]/g,
    function (match, token) {
      var normalized = normalizeHeaderName(token);
      if (!normalized) return match;
      if (Object.prototype.hasOwnProperty.call(context, normalized)) {
        return context[normalized];
      }
      return match;
    },
  );
}

function buildEmailSubject(docType, clientName, expiryDate) {
  var docLabel = docType ? docType : "Visa/Permit";
  return (
    "Reminder: " +
    docLabel +
    " Expiry on " +
    formatDate(expiryDate) +
    " \u2013 " +
    clientName
  );
}

function getStaticRedirectLinkUrl() {
  var targetUrl = String(CONFIG.STATIC_REDIRECT_URL || "").trim();
  if (!targetUrl) return "";

  var baseUrl = getOpenTrackingBaseUrl();
  if (!baseUrl) return targetUrl;

  var separator = baseUrl.indexOf("?") >= 0 ? "&" : "?";
  return baseUrl + separator + "mode=click&u=" + encodeURIComponent(targetUrl);
}

// ═══════════════════════════════════════════════════════════════════════════
// EmailDelivery — Gmail send, CC resolution, sender helpers
// ═══════════════════════════════════════════════════════════════════════════

// 33 Email Delivery

function getDefaultCcEmails() {
  var raw = getPropString(PROP_KEYS.DEFAULT_CC_EMAILS, "");
  if (!raw) return [];

  try {
    var parsed = JSON.parse(raw);
    return normalizeEmailList(parsed);
  } catch (e) {
    return normalizeEmailList(raw);
  }
}

function setDefaultCcEmails(emails) {
  var normalized = normalizeEmailList(emails);
  if (normalized.length === 0) {
    setPropString(PROP_KEYS.DEFAULT_CC_EMAILS, "");
    return;
  }

  setPropString(PROP_KEYS.DEFAULT_CC_EMAILS, JSON.stringify(normalized));
}

function resolveCcEmails(clientEmails, staffEmail) {
  var combined = mergeUniqueEmails(getDefaultCcEmails(), staffEmail);
  var clientList = Array.isArray(clientEmails)
    ? clientEmails
    : normalizeEmailList(clientEmails);
  var clientKeys = {};
  for (var k = 0; k < clientList.length; k++) {
    clientKeys[clientList[k].toLowerCase()] = true;
  }
  if (Object.keys(clientKeys).length === 0) return combined;

  var filtered = [];
  for (var i = 0; i < combined.length; i++) {
    if (!clientKeys[combined[i].toLowerCase()]) filtered.push(combined[i]);
  }
  return filtered;
}

function formatCcDisplay(ccEmails) {
  var ccList = normalizeEmailList(ccEmails);
  return ccList.length > 0 ? ccList.join(", ") : "(none)";
}

function setDefaultCcEmailsDialog() {
  var ui = SpreadsheetApp.getUi();
  var current = getDefaultCcEmails();
  var response = ui.prompt(
    "Set Default CC Emails",
    "Enter one or more default CC email addresses.\n" +
      "Use commas, semicolons, or new lines.\n" +
      "Leave blank to clear.\n\n" +
      "Current: " +
      formatCcDisplay(current),
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var emails = normalizeEmailList(response.getResponseText());
  var invalid = validateEmailList(emails);
  if (invalid.length > 0) {
    ui.alert(
      "Invalid Email(s)",
      "These email addresses are invalid:\n" + invalid.join("\n"),
      ui.ButtonSet.OK,
    );
    return;
  }

  setDefaultCcEmails(emails);
  ui.alert(
    "Default CC Emails Saved",
    emails.length > 0
      ? "Default CC list:\n" + emails.join("\n")
      : "Default CC emails cleared.",
    ui.ButtonSet.OK,
  );
}

function getSenderAccountEmail() {
  try {
    var effectiveEmail = Session.getEffectiveUser().getEmail();
    if (effectiveEmail) return effectiveEmail;
  } catch (e) {}

  try {
    var activeEmail = Session.getActiveUser().getEmail();
    if (activeEmail) return activeEmail;
  } catch (e) {}

  return "";
}

function getSenderDisplayName(email) {
  var raw = String(email || "").trim();
  var atIndex = raw.indexOf("@");
  if (atIndex < 0) return CONFIG.SENDER_NAME;

  var domain = raw.slice(atIndex + 1);
  var domainBase = domain.split(".")[0];
  if (!domainBase) return CONFIG.SENDER_NAME;

  return domainBase.charAt(0).toUpperCase() + domainBase.slice(1).toLowerCase();
}

function sendReminderEmail(
  clientEmail,
  ccEmails,
  subject,
  htmlBody,
  blobItems,
  senderName,
  fromEmail,
) {
  var options = {
    htmlBody: htmlBody,
    name: senderName || CONFIG.SENDER_NAME,
  };

  // From-address must be a verified Gmail alias of the script runner.
  // Caller is responsible for that check; we just pass it through.
  if (fromEmail) options.from = fromEmail;

  var normalizedCc = normalizeEmailList(ccEmails);
  if (normalizedCc.length > 0) {
    options.cc = normalizedCc.join(",");
  }

  if (blobItems && blobItems.length > 0) {
    var attachments = blobItems.map(function (item) {
      return item.blob.setName(item.name || "attachment");
    });
    options.attachments = attachments;
  }

  GmailApp.sendEmail(clientEmail, subject, "", options);

  // Best-effort metadata lookup from Sent mailbox.
  return lookupRecentSentMessageMeta(clientEmail, subject);
}

function sendReminderEmails(
  clientEmails,
  ccEmails,
  subject,
  htmlBody,
  blobItems,
  senderName,
  fromEmail,
) {
  var recipients = normalizeEmailList(clientEmails);
  var result = {
    success: [],
    failed: [],
    meta: { threadId: "", messageId: "" },
  };

  for (var i = 0; i < recipients.length; i++) {
    var recipient = recipients[i];

    if (!isValidEmail(recipient)) {
      result.failed.push({
        email: recipient,
        error: "Invalid email address format.",
      });
      continue;
    }

    try {
      var meta = sendReminderEmail(
        recipient,
        ccEmails,
        subject,
        htmlBody,
        blobItems,
        senderName,
        fromEmail,
      );
      if (result.success.length === 0) result.meta = meta;
      result.success.push(recipient);
    } catch (e) {
      result.failed.push({
        email: recipient,
        error: e && e.message ? e.message : String(e),
      });
    }
  }

  if (result.success.length === 0 && result.failed.length > 0) {
    throw new Error(
      result.failed
        .map(function (item) {
          return item.email + " (" + item.error + ")";
        })
        .join("; "),
    );
  }

  return result;
}

function lookupRecentSentMessageMeta(clientEmail, subject) {
  var meta = { threadId: "", messageId: "" };
  if (!clientEmail || !subject) return meta;

  try {
    var query =
      "in:sent to:(" +
      escapeGmailQueryValue(clientEmail) +
      ') subject:("' +
      escapeGmailQueryValue(subject) +
      '") newer_than:7d';
    var threads = GmailApp.search(query, 0, 5);
    if (!threads || threads.length === 0) return meta;

    var thread = threads[0];
    var messages = thread.getMessages();
    var latest = messages[messages.length - 1];
    meta.threadId = thread.getId() || "";
    meta.messageId = latest && latest.getId ? latest.getId() : "";
  } catch (e) {}

  return meta;
}

function escapeForQuotedPrintable(value) {
  return String(value || "").replace(/"/g, '\\"');
}
