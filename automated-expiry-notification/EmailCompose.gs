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

  htmlBody = [
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#333;max-width:600px;line-height:1.6;">',
    htmlBody,
    '<hr style="border:none;border-top:1px solid #eee;margin:24px 0;">',
    '<p style="font-size:11px;color:#999;">This is an automated reminder. Email opens may be tracked for delivery visibility.</p>',
    "</div>",
  ].join("\n");

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
