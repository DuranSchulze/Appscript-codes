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
) {
  var options = {
    htmlBody: htmlBody,
    name: senderName || CONFIG.SENDER_NAME,
  };

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
