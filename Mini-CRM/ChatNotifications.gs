/**
 * Sends a Google Chat card with Approve/Reject buttons for a draft.
 */
function sendChatApprovalCard(senderEmail, subject, snippet, draftId, draftUrl, category) {
  const webhookUrl = getChatWebhookForCategory(category || 'General');
  if (!webhookUrl) return;

  const approveUrl = createSecureDraftActionUrl_('approve', draftId);
  const rejectUrl = createSecureDraftActionUrl_('reject', draftId);

  const card = {
    cards: [{
      header: { title: '📧 AI Draft Ready for Review' },
      sections: [{
        widgets: [
          { keyValue: { topLabel: 'From', content: senderEmail } },
          { keyValue: { topLabel: 'Subject', content: subject } },
          { textParagraph: { text: snippet.substring(0, 300) } },
          {
            buttons: [
              {
                textButton: {
                  text: '✅ Approve & Send',
                  onClick: { openLink: { url: approveUrl } }
                }
              },
              {
                textButton: {
                  text: '❌ Discard',
                  onClick: { openLink: { url: rejectUrl } }
                }
              }
            ]
          }
        ]
      }]
    }]
  };

  const response = UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(card),
    muteHttpExceptions: true
  });
  const statusCode = response.getResponseCode();
  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('Google Chat notification failed with HTTP ' + statusCode + '.');
  }
}

function getChatWebhookForCategory(category) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapSheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);
  if (!mapSheet) return null;
  const data = mapSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'ChatWebhook' && data[i][1] === category) {
      return data[i][2];
    }
  }
  return null;
}
