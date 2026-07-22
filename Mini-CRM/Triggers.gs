function setupAllTriggers() {
  setupAutomationTriggers();
  SpreadsheetApp.getUi().alert(
    '✅ Automation Ready',
    'Configured AI processing, follow-up reminders, Gmail sync, Dashboard refresh, and Engagement edit automation. Unrelated project triggers were preserved.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
