/**
 * Main function for processing AI-Pending emails.
 * Triggered every 5 minutes or manually from menu.
 */
function processAutoDrafts() {
  assertMonitoredMailboxAccount_('AI auto-reply processing');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return { processed: 0, reason: 'Another AI job is active' };
  try {
    return processAutoDraftsUnlocked_();
  } finally {
    lock.releaseLock();
  }
}

function processAutoDraftsUnlocked_() {
  if (!isAiProcessingEnabled_()) return { processed: 0, reason: 'AI processing is paused' };
  const label = GmailApp.getUserLabelByName(AI_CONFIG.AI_PENDING_LABEL);
  if (!label) {
    console.warn(`Label "${AI_CONFIG.AI_PENDING_LABEL}" not found.`);
    return { processed: 0, reason: 'AI label not found' };
  }
  const threads = label.getThreads(0, AI_CONFIG.MAX_AI_PENDING_THREADS_PER_RUN);
  const apiKey = getGeminiApiKey();
  if (!apiKey) {
    console.error('Gemini API key not set. Use menu to configure.');
    return { processed: 0, reason: 'Gemini not configured' };
  }

  let processed = 0;
  let aiAttempts = 0;
  for (const thread of threads) {
    try {
      const messages = thread.getMessages();
      const lastMessage = messages[messages.length - 1];
      if (!lastMessage.isUnread()) continue;

      const messageId = lastMessage.getId();
      const senderEmail = extractEmailAddress(lastMessage.getFrom());
      const subject = lastMessage.getSubject();
      const body = lastMessage.getPlainBody();

      // Local FAQ answers do not consume Gemini quota.
      const faqMatch = checkFAQ(body, subject);
      if (faqMatch && faqMatch.confidence >= AI_CONFIG.FAQ_CONFIDENCE_THRESHOLD) {
        sendAutoReply(senderEmail, subject, faqMatch.answer, faqMatch.category);
        lastMessage.markRead();
        thread.removeLabel(label);
        try { logAutoReply(senderEmail, subject, true, faqMatch.category); } catch (logError) { console.error(logError); }
        processed++;
        continue;
      }

      let checkpoint = getAiDraftCheckpoint_(messageId);
      if (!checkpoint) {
        if (aiAttempts >= AI_CONFIG.MAX_AUTO_DRAFTS_PER_RUN) break;
        if (!canAttemptAiWork_('auto_reply', messageId)) continue;
        aiAttempts++;
        const generatedDraft = generateDraftWithGemini(body, subject, senderEmail);
        if (!generatedDraft) {
          recordAiWorkFailure_('auto_reply', messageId);
          console.error('Gemini draft generation failed for message ' + messageId + '. Retry was deferred.');
          continue;
        }

        clearAiWorkFailure_('auto_reply', messageId);
        const draft = GmailApp.createDraft(senderEmail, `Re: ${subject}`, generatedDraft.reply);
        checkpoint = {
          draftId: draft.getId(),
          draftUrl: `https://mail.google.com/mail/u/0/#drafts?compose=${draft.getId()}`,
          category: generatedDraft.category,
          createdAt: new Date().toISOString()
        };
        saveAiDraftCheckpoint_(messageId, checkpoint);
      }

      // If notification fails, the checkpoint prevents another AI call or draft.
      sendChatApprovalCard(senderEmail, subject, body, checkpoint.draftId, checkpoint.draftUrl, checkpoint.category);
      lastMessage.markRead();
      thread.removeLabel(label);
      clearAiDraftCheckpoint_(messageId);
      try { logAutoReply(senderEmail, subject, false, checkpoint.category); } catch (logError) { console.error(logError); }
      processed++;
    } catch (error) {
      console.error('AI-Pending thread deferred: ' + error.message);
    }
  }
  return { processed: processed, aiAttempts: aiAttempts };
}

/**
 * Generates a draft reply using Gemini.
 * Uses the primary model first and automatically retries with the configured
 * fallback model when the primary request or response fails.
 * @returns {object|null} { reply, category, model, usedFallback }
 */
function generateDraftWithGemini(emailBody, subject, senderEmail) {
  const apiKey = getGeminiApiKey();
  const primaryModel = getGeminiModel();
  const fallbackModel = getGeminiFallbackModel();

  const consultationLinksText = AI_CONFIG.CONSULTATION_LINKS
    .map(l => `${l.lawyer}: ${l.url}`)
    .join('\n');

  const cleanEmailBody = sanitizeEmailForDraft_(emailBody);
  const prompt = `You are a professional legal assistant for Duran Schulze Law.
Draft a polite, concise reply to the following client email.
If the client asks a complex legal question, suggest they schedule a consultation with one of our specialized lawyers.
Include these consultation links at the end:
${consultationLinksText}

Also, categorize the email into one of: Visa, Legal, Business Formation, Trademark, Accounting, General.
Return your response in JSON format:
{
  "reply": "the full email body",
  "category": "Visa"
}

Original email subject: ${subject}
Original email: ${cleanEmailBody}`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { responseMimeType: "application/json", temperature: 0.2, maxOutputTokens: 700 }
  };

  const route = [primaryModel];
  if (fallbackModel && fallbackModel !== primaryModel) {
    route.push(fallbackModel);
  }

  for (let i = 0; i < route.length; i++) {
    const model = route[i];
    try {
      const parsed = requestGeminiJson_(apiKey, model, payload, { purpose: 'auto_reply' });
      if (!parsed.reply || typeof parsed.reply !== 'string') {
        throw new Error('The response did not contain a valid reply.');
      }

      const usedFallback = i > 0;
      const properties = {
        GEMINI_LAST_MODEL_USED: model,
        GEMINI_LAST_SUCCESS_AT: new Date().toISOString()
      };
      if (usedFallback) {
        properties.GEMINI_LAST_FALLBACK_AT = new Date().toISOString();
      }
      PropertiesService.getScriptProperties().setProperties(properties);

      return {
        reply: parsed.reply.trim(),
        category: parsed.category || 'General',
        model: model,
        usedFallback: usedFallback
      };
    } catch (error) {
      console.error('Gemini model ' + model + ' failed: ' + error.message);
      if (shouldStopAiRoute_(error)) break;
    }
  }

  return null;
}

function sanitizeEmailForDraft_(body) {
  let text = String(body || '').replace(/\r/g, '');
  text = text.split(/\nOn .{0,300}wrote:\s*\n/i)[0];
  text = text.split(/\nFrom:\s.+\nSent:\s.+\nTo:\s.+/i)[0];
  text = text.split('\n').filter(function(line) { return !/^\s*>/.test(line); }).join('\n');
  return text.replace(/\n{3,}/g, '\n\n').trim().substring(0, AI_CONFIG.MAX_AI_DRAFT_BODY_CHARS);
}

function requestGeminiJson_(apiKey, model, payload, options) {
  options = options || {};
  reserveAiRequest_(options.purpose || 'general', model);
  const url = AI_CONFIG.GEMINI_API_BASE_URL + '/models/' + encodeURIComponent(model) + ':generateContent';
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { 'x-goog-api-key': apiKey },
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (statusCode < 200 || statusCode >= 300) {
    recordAiHttpFailure_(statusCode);
    let message = 'HTTP ' + statusCode;
    try {
      const errorData = JSON.parse(responseText);
      message = errorData && errorData.error && errorData.error.message
        ? errorData.error.message
        : message;
    } catch (ignore) {}
    const requestError = new Error(message);
    requestError.httpStatus = statusCode;
    requestError.aiStopRoute = statusCode === 429 || statusCode === 401 || statusCode === 403;
    requestError.aiBlocked = requestError.aiStopRoute;
    throw requestError;
  }

  const data = JSON.parse(responseText);
  if (!data.candidates || !data.candidates.length ||
      !data.candidates[0].content || !data.candidates[0].content.parts ||
      !data.candidates[0].content.parts.length) {
    throw new Error('Gemini returned no candidate content.');
  }

  let text = String(data.candidates[0].content.parts[0].text || '').trim();
  text = text.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');
  const parsed = JSON.parse(text);
  recordAiRequestSuccess_();
  return parsed;
}

/**
 * Checks the FAQ sheet for a matching question.
 * @returns {object|null} { question, answer, category, confidence }
 */
/**
 * Checks the Map Sheet (type 'FAQ') for a matching question.
 */
function checkFAQ(emailBody, subject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapSheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);
  if (!mapSheet || mapSheet.getLastRow() < 2) return null;

  const data = mapSheet.getDataRange().getValues();
  const faqEntries = data.slice(1).filter(row => row[0] === 'FAQ' && row[1] && row[2])
    .map(row => ({ category: row[1], question: row[1], answer: row[2] })); // category same as question for simplicity
  // Actually, we need a structure: we can use column B as question, column C as answer, and column A as type.
  // But the Map Sheet structure is Type, Key, Value. So we'll treat Key as question and Value as answer.
  // Let's filter properly:
  const questions = data.slice(1)
    .filter(row => row[0] === 'FAQ' && row[1] && row[2])
    .map(row => ({
      question: row[1],
      answer: row[2]
    }));

  const combined = (emailBody + ' ' + subject).toLowerCase();
  const match = questions.find(q => combined.includes(q.question.toLowerCase()));
  if (match) {
    return { ...match, confidence: 100 };
  }
  return null;
}

/**
 * Sends an auto-reply immediately.
 */
function sendAutoReply(to, subject, body, category) {
  assertMonitoredMailboxAccount_('automatic email reply');
  GmailApp.sendEmail(to, `Re: ${subject}`, body, {
    name: 'Duran Schulze Law',
    replyTo: getCategoryRoutingEmail(category) || 'info@duranschulze.com'
  });
}

/**
 * Logs the auto-response in the Engagement Information sheet (or a dedicated log).
 */
function logAutoReply(senderEmail, subject, isAuto, category) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);
  if (!sheet) return;

  // Find or create row for this sender
  // For simplicity, we'll append a new row with minimal info.
  const newRow = [
    new Date(),           // Contact Date
    extractSenderName(senderEmail), // Client Name (approximate)
    '',                   // Service
    '',                   // Contact Person
    senderEmail,          // Email
    '', '',               // Phone, Address
    subject,              // Remarks (subject)
    '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
    isAuto ? 'Auto-Replied' : 'Pending Approval',
    '', '', '', '', '', '',
    category || 'General'
  ];
  sheet.appendRow(newRow);
}

/**
 * Retrieves the category routing email from Map Sheet.
 */
function getCategoryRoutingEmail(category) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapSheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);
  if (!mapSheet) return null;
  const data = mapSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'CategoryRouting' && data[i][1] === category) {
      return data[i][2];
    }
  }
  return null;
}

/**
 * Menu function: generate draft for selected row.
 */
function menuGenerateDraftForRow() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);
  const activeRange = sheet.getActiveRange();
  const row = activeRange.getRow();
  if (row < 2) {
    ui.alert('Please select a data row.');
    return;
  }
  const email = sheet.getRange(row, 5).getValue(); // Column E: Email Address
  const subject = sheet.getRange(row, 8).getValue(); // Remarks (maybe contains original subject)
  const body = 'Client inquiry'; // We don't have full body, so we can't generate a meaningful draft without more data. For demo, we'll just show that it's not feasible.
  ui.alert('Draft Generation', 'To generate a draft, please use the AI-Pending label workflow or provide the full email content.', ui.ButtonSet.OK);
}
