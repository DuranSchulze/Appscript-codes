/**
 * Daily check for leads that need follow-up.
 * Saves draft reminders and sends Chat notifications.
 */
function processFollowUpReminders() {
  assertMonitoredMailboxAccount_('follow-up reminder processing');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return;

  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const delayDays = AI_CONFIG.FOLLOW_UP_DELAY_DAYS;
  const cutoff = new Date(now.getTime() - delayDays * 24 * 60 * 60 * 1000);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const contactDate = row[0]; // Column A
    const status = row[20]; // Engagement Status
    const lastFollowUp = row[30]; // Last Follow-Up Date (new column, maybe not populated yet)

    if (status === 'Engaged' || status === 'Client') continue;
    if (!contactDate || contactDate > cutoff) continue;

    const clientName = row[1] || row[4];
    const subject = `Follow-up: ${row[7] || 'Your inquiry'}`;
    const draftBody = `Dear ${clientName},\n\nWe wanted to follow up on our previous conversation. If you have any questions or would like to schedule a consultation, please let us know.\n\nBest regards,\nDuran Schulze Law`;

    const draft = GmailApp.createDraft(row[4], subject, draftBody);
    // Send Chat notification
    sendChatFollowUpCard(row[4], clientName, draft.getId());
    // Update Last Follow-Up Date column (add if not exists)
    // For now, set column AF (32) as placeholder
    if (sheet.getLastColumn() >= 32) {
      sheet.getRange(i + 1, 32).setValue(new Date());
    }
  }
}

function sendChatFollowUpCard(email, name, draftId) {
  const webhookUrl = getChatWebhookForCategory('FollowUp');
  if (!webhookUrl) return;
  const approveUrl = createSecureDraftActionUrl_('approve', draftId);
  const card = {
    cards: [{
      header: { title: '⏰ Follow-Up Reminder' },
      sections: [{
        widgets: [
          { keyValue: { topLabel: 'Client', content: `${name} (${email})` } },
          { textParagraph: { text: 'A follow-up draft is ready. Click to review and send.' } },
          {
            buttons: [
              {
                textButton: {
                  text: '✅ Review & Send',
                  onClick: { openLink: { url: approveUrl } }
                }
              }
            ]
          }
        ]
      }]
    }]
  };
  UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(card),
    muteHttpExceptions: true
  });
}
