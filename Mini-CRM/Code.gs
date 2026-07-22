/**
 * Enhanced CRM Gmail Tracker
 */

const CRM_VERSION = '5.7';

// ========================================
// CONFIGURATION
// ========================================

const CONFIG = {
  SHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(),
  ROOT_FOLDER_ID: '1TkUKqBF9kti-qCqUpBpOz9dpVbUGqOYP',
  GMAIL_SEARCH_QUERY: 'in:inbox newer_than:30d',
  MAX_THREADS: 200,
  BACKFILL_BATCH_SIZE: 100,
  MAX_BACKFILL_THREADS: 2000,
  WEB_APP_ACTION_TOKEN_TTL_HOURS: 24,
  MONITORED_MAILBOX_EMAIL: 'info@filepino.com',
  ENGAGEMENT_START_DATE: new Date('2025-10-01'),

  // INTERNAL/COMPANY DOMAINS - Status = "N/A"
  INTERNAL_DOMAINS: [
    '@duranschulze.com',
    '@filepino.com'
  ],

  // EXCLUDED DOMAINS - Status = "Trash"
  EXCLUDED_DOMAINS: [
    '@basecamp.com',
    '@blaze.com',
    'duranschulze.mail@gmail.com',
    '@securitybank.com',
    '@rebump.cc',
    '@stripe.com',
    '@marketing.securitybank.com'
  ],

  // EXCLUDED PATTERNS
  EXCLUDED_PATTERNS: [
    'no-reply',
    'noreply',
    'donotreply',
    'do-not-reply',
    'support-donotreply'
  ],

  // Email Classifications
  MEETING_KEYWORDS: [
    'meeting', 'schedule', 'appointment', 'zoom', 'google meet',
    'calendar', 'call', 'conference', 'demo', 'consultation'
  ],
  CONVERSION_KEYWORDS: [
    'engagement letter signed', 'engagement letter executed', 'service invoice paid',
    'retainer paid', 'engagement agreement signed', 'legal services agreement',
    'attorney-client agreement', 'invoice payment confirmed', 'retainer received',
    'legal fee paid', 'retained as counsel', 'engaged as attorney'
  ],
  PAYMENT_KEYWORDS: [
    'payment sent', 'payment made', 'payment received', 'invoice paid',
    'paypal', 'paid invoice', 'payment confirmation'
  ],
  UNSUBSCRIBE_KEYWORDS: [
    'unsubscribe', 'unsubscribed', 'opt out', 'opt-out', 'remove me',
    'stop sending', 'no longer interested', 'take me off', 'cancel subscription',
    'do not contact', 'stop emails', 'remove from list', 'stop communication'
  ],

  // Trash filters
  TRASH_EMAILS: [
    'jeremy.pajarillo@makatimed.net.ph',
    'lucky@yfzcip.com',
    'noreply@google.com',
    'info@pvpi.ph'
  ],
  TRASH_DOMAINS: [
    '@facebookmail.com'
  ],

  // Special email types
  ESIGN_EMAILS: [
    'noreply@mail.hellosign.com',
    'no-reply@hellosign.com'
  ],

  // Sheet names
  MAP_SHEET_NAME: 'Map Sheet',
  ENGAGE_SHEET_NAME: 'Engagement Information Sheet',
  CONVERSION_SHEET_NAME: 'Conversion Tracking',
  DASHBOARD_SHEET_NAME: 'Dashboard',
  POTENTIAL_CLIENTS_SHEET_NAME: 'Potential Clients',

  // System settings
  ARCHIVE_MONTHS_THRESHOLD: 12,
  RECENT_MONTHS_DASHBOARD: 6,
  AUTO_REFRESH_HOUR: 8,
  TIMEZONE: Session.getScriptTimeZone(),
  VAT_RATE: 0.12

  // Include AI configuration
  // (already present if you paste Config.gs content, but we'll keep them separate)

};

// ========================================
// AUTHORIZED MAILBOX GUARD
// Kept in Code.gs so a partial Apps Script upload cannot omit it.
// ========================================

/**
 * Returns the Google account whose OAuth authority is used for this execution.
 * For installable triggers, this is the account that created the trigger.
 */
function getEffectiveAutomationAccount_() {
  try {
    return String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  } catch (error) {
    return '';
  }
}

function getConfiguredMonitoredMailbox_() {
  return String(CONFIG.MONITORED_MAILBOX_EMAIL || '').trim().toLowerCase();
}

/**
 * Fails closed before Gmail is read or modified under the wrong Google account.
 */
function assertMonitoredMailboxAccount_(operation) {
  const expected = getConfiguredMonitoredMailbox_();
  const effective = getEffectiveAutomationAccount_();
  if (!expected) throw new Error('CONFIG.MONITORED_MAILBOX_EMAIL is not configured.');
  if (!effective) {
    throw new Error('Could not verify the Google account running ' + operation + '. Mailbox access was blocked.');
  }
  if (effective !== expected) {
    throw new Error(
      'Mailbox access blocked for ' + operation + '. This execution is authorized as ' + effective +
      ', but Mini-CRM is locked to ' + expected + '. Sign in as ' + expected +
      ' and run CRM Tracker > AI Auto-Reply > Setup AI Triggers.'
    );
  }
  return expected;
}

function menuVerifyAutomationAccount() {
  const ui = SpreadsheetApp.getUi();
  const expected = getConfiguredMonitoredMailbox_();
  const effective = getEffectiveAutomationAccount_();
  const allowed = Boolean(expected && effective && expected === effective);
  ui.alert(
    allowed ? '✅ Automation account verified' : '⛔ Wrong automation account',
    'Required mailbox: ' + (expected || '(not configured)') + '\n' +
    'Current effective account: ' + (effective || '(unavailable)') + '\n\n' +
    (allowed
      ? 'This account may create triggers and access the CRM mailbox.'
      : 'Gmail sync, drafts, sends, and mailbox automation are blocked. Sign in with the required mailbox account before setting up triggers.'),
    ui.ButtonSet.OK
  );
  return { allowed: allowed, expected: expected, effective: effective };
}

/**
 * Global cache for Map Sheet data
 * Stores parsed data to avoid repeated reads
 */
var MAP_CACHE = {
  departments: {},
  services: {},
  notificationEmails: {
    engagement: [],
    paidClient: []
  },
  dropdownOptions: {},
  lastRefresh: null
};


/**
 * ========================================
 * INITIAL SETUP WIZARD - For Fresh Installations
 * ========================================
 * This is the FIRST function to run on a clean system
 * It will:
 * 1. Create all required sheets
 * 2. Set up headers and formatting
 * 3. Populate Map Sheet with defaults
 * 4. Set up dropdowns
 * 5. Configure triggers
 * 6. Run initial email sync
 */
function initialSystemSetup() {
  const ui = SpreadsheetApp.getUi();

  try {
    assertMonitoredMailboxAccount_('initial system setup');
    // Welcome screen
    const welcomeResponse = ui.alert(
      '🎉 Welcome to Enhanced CRM Gmail Tracker v' + CRM_VERSION,
      '👋 Welcome! This wizard will set up your CRM system.\n\n' +
      '📋 What this setup will do:\n' +
      '✓ Create all required sheets\n' +
      '✓ Set up Map Sheet with default services\n' +
      '✓ Configure Conversion Tracking\n' +
      '✓ Create the Potential Clients review queue\n' +
      '✓ Set up Engagement Information Sheet\n' +
      '✓ Apply formulas and dropdowns\n' +
      '✓ Configure automatic triggers\n\n' +
      '⏱️ Estimated time: 2-3 minutes\n\n' +
      'Ready to begin?',
      ui.ButtonSet.YES_NO
    );

    if (welcomeResponse !== ui.Button.YES) {
      ui.alert('Setup Cancelled', 'You can run this setup anytime from the menu.', ui.ButtonSet.OK);
      return;
    }

    Logger.log('========================================');
    Logger.log('STARTING INITIAL SYSTEM SETUP');
    Logger.log('========================================');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('Creating and validating CRM tabs…', 'Initial setup', 5);
    const setupResult = initializeWorkbook_(ss);

    // STEP 6: Ask if user wants historical sync or just recent emails
    const historicalResponse = ui.alert(
      '📧 Email Sync Options',
      'How would you like to start?\n\n' +
      '✅ RECOMMENDED: Sync recent emails (last 30 days)\n' +
      '⏰ ADVANCED: Sync all emails from Jan 2024 (takes 10-20 min)\n\n' +
      'Choose "YES" for recent (30 days)\n' +
      'Choose "NO" to skip sync for now\n' +
      'Choose "CANCEL" for advanced (Jan 2024)',
      ui.ButtonSet.YES_NO_CANCEL
    );

    if (historicalResponse === ui.Button.YES) {
      // Sync recent emails (30 days)
      ss.toast('Syncing recent emails from Gmail…', 'Initial setup', 5);
      Logger.log('Step 6: Syncing recent emails...');
      syncNewEmails();
      Logger.log('✓ Recent emails synced');
    } else if (historicalResponse === ui.Button.CANCEL) {
      // User wants historical sync
      const confirmHistorical = ui.alert(
        '⚠️ Confirm Historical Sync',
        'This will sync ALL emails from January 2024 to present.\n\n' +
        '⏱️ This may take 10-20 minutes for large mailboxes.\n' +
        '⚠️ Do NOT close this window during sync.\n\n' +
        'Continue with historical sync?',
        ui.ButtonSet.YES_NO
      );

      if (confirmHistorical === ui.Button.YES) {
        ss.toast('Starting historical Gmail sync…', 'Initial setup', 5);
        Logger.log('Step 6: Running historical sync...');
        syncHistoricalEmailsFromJan2024();
        Logger.log('✓ Historical emails synced');
      } else {
        Logger.log('Step 6: Skipped email sync (user choice)');
      }
    } else {
      Logger.log('Step 6: Skipped email sync (user choice)');
    }

    // Final refresh and organization. Dashboard is created even when email
    // import was skipped.
    ss.toast('Refreshing Dashboard and finishing workbook setup…', 'Initial setup', 5);
    buildEnhancedDashboard();
    formatInitializedWorkbook_(ss);
    organizeSheets(ss);
    Logger.log('✓ Workbook formatted and organized');

    Logger.log('========================================');
    Logger.log('INITIAL SETUP COMPLETE!');
    Logger.log('========================================');

    // Show completion summary
    const completionMessage =
      '🎉 Setup Complete!\n\n' +
      '✅ Your CRM system is ready to use!\n\n' +
      '📊 What was created:\n' +
      '• Map Sheet (services & settings)\n' +
      '• Conversion Tracking (email aggregation)\n' +
      '• Potential Clients (AI qualification review queue)\n' +
      '• Engagement Information Sheet (prospects/leads)\n' +
      '• Dashboard (metrics and priority actions)\n' +
      '• AI-Pending Gmail label\n' +
      '• Sync, Dashboard, AI, follow-up, and edit triggers\n\n' +
      '🔒 Existing rows preserved: ' + (setupResult.preservedExistingData ? 'Yes' : 'New workbook') + '\n\n' +
      '📋 Next Steps:\n' +
      '1. Check Map Sheet - add your services\n' +
      '2. Review Qualified and Review rows in Potential Clients\n' +
      '3. Promote approved candidates into Engagement\n' +
      '4. Use menu options for ongoing management\n\n' +
      '💡 TIP: Use "Full Update" from menu for daily syncs!';

    ui.alert('🎉 Setup Complete!', completionMessage, ui.ButtonSet.OK);

    // Show quick start guide
    showQuickStartGuide();

  } catch (error) {
    Logger.log('========================================');
    Logger.log('ERROR in initial setup: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========================================');
    ui.alert('⚠️ Setup Error',
             'Setup encountered an error:\n\n' + error.message + '\n\n' +
             'Please check the execution log (Apps Script > View > Logs) for details.',
             ui.ButtonSet.OK);
  }
}

/**
 * Creates or upgrades every currently functional CRM component. This routine
 * is intentionally idempotent: rerunning setup preserves existing rows and
 * user configuration while restoring required headers, formats, and triggers.
 */
function initializeWorkbook_(ss) {
  const hadExistingData = ss.getSheets().some(function(sheet) {
    return sheet.getLastRow() > 1;
  });

  setupMapSheet(ss);
  setupFAQInMapSheet();
  ensureCategoryRouting();

  getOrCreateConversionTrackingSheet(ss);
  setupEngagementInformationSheet(ss);
  ensurePotentialClientsSheet_(ss);

  const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);
  setupFinancialFormulas(infoSheet);
  setupDropdownValidations(infoSheet);

  ensureAiPendingLabel_();
  setupAutomationTriggers();
  buildEnhancedDashboard();
  formatInitializedWorkbook_(ss);
  organizeSheets(ss);

  PropertiesService.getScriptProperties().setProperties({
    CRM_INITIALIZED_VERSION: CRM_VERSION,
    CRM_INITIALIZED_AT: new Date().toISOString()
  });

  return { preservedExistingData: hadExistingData };
}


/**
 * Sets up Map Sheet with default data including notification email provisions
 * Updated to include email notification configurations
 */
function setupMapSheet(ss) {
  let mapSheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);

  if (!mapSheet) {
    mapSheet = ss.insertSheet(CONFIG.MAP_SHEET_NAME);
  } else {
    if (mapSheet.getLastRow() > 1) {
      ensureMapSheetHeaders_(mapSheet);
      formatMapSheet_(mapSheet);
      Logger.log('  ℹ Existing Map Sheet preserved and reformatted');
      return;
    }
    // An empty/headers-only sheet is safe to seed.
    mapSheet.clear();
  }

  // Map Sheet data structure with notification emails
  const mapData = [
    ['Type', 'Key', 'Value'],

    // SERVICES - Add your services here with Google Docs template IDs
    ['Service', 'Immigration Visa', '1g8vauhd9r437PlAzCKJh5Qt11emZa75vm-DmWxGkaUQ'],
    ['Service', 'Litigation', '1NTfGpOGUSR08GOMImFM3EnT80tSn_zDSgBeo_XIVQNc'],
    ['Service', 'Business Registration', ''],
    ['Service', 'Legal Consultation', ''],
    ['Service', 'Corporate Services', ''],
    ['Service', 'Tax Advisory', ''],

    // BLANK ROW FOR VISUAL SEPARATION
    ['', '', ''],

    // DEPARTMENTS - Add all departments with email addresses
    ['Department', 'Legal', 'marywendy@duranschulze.com'],
    ['Department', 'Visa', 'marywendy@duranschulze.com'],
    ['Department', 'HR', 'marywendy@duranschulze.com'],
    ['Department', 'Accounting', 'marywendy@duranschulze.com'],
    ['Department', 'Operations', 'marywendy@duranschulze.com'],
    ['Department', 'Admin', 'marywendy@duranschulze.com'],

    // BLANK ROW FOR VISUAL SEPARATION
    ['', '', ''],

    // EMAIL NOTIFICATIONS FOR ENGAGEMENT STATUS (when status = "Engaged")
    // These emails will be notified when a client is marked as "Engaged"
    // Typically: Billing, Accounting, Management
    ['Email Notification for Engagement Only', 'Billing Department', 'marywendy@duranschulze.com'],
    ['Email Notification for Engagement Only', 'Accounting Head', 'marywendy@duranschulze.com'],
    ['Email Notification for Engagement Only', 'Finance Manager', 'marywendy@duranschulze.com'],
    ['Email Notification for Engagement Only', 'Management', 'marywendy@duranschulze.com'],

    // BLANK ROW FOR VISUAL SEPARATION
    ['', '', ''],

    // EMAIL NOTIFICATIONS FOR PAID CLIENT (when payment = "Paid" + department assigned)
    // These emails will be notified when a client has paid and is ready for service
    // Typically: Operations, Service Manager, Assigned Department
    ['Email Notification for Paid Client', 'Operations Manager', 'marywendy@duranschulze.com'],
    ['Email Notification for Paid Client', 'Service Coordinator', 'marywendy@duranschulze.com'],
    ['Email Notification for Paid Client', 'Project Manager', 'marywendy@duranschulze.com'],
    ['Email Notification for Paid Client', 'Quality Assurance', 'marywendy@duranschulze.com'],

    // BLANK ROW FOR VISUAL SEPARATION
    ['', '', ''],

    // DROPDOWN OPTIONS - Follow-Up Status
    ['FollowUp', 'Follow Up', '–'],
    ['FollowUp', 'Declined', '–'],
    ['FollowUp', 'On Hold', '–'],

    // BLANK ROW FOR VISUAL SEPARATION
    ['', '', ''],

    // DROPDOWN OPTIONS - Payment Status
    ['Payment', 'Paid', '–'],
    ['Payment', 'Unpaid', '–'],
    ['Payment', 'Partial', '–'],
    ['Payment', 'Pending', '–'],

    // BLANK ROW FOR VISUAL SEPARATION
    ['', '', ''],

    // DROPDOWN OPTIONS - Quote Actions
    ['QuoteAction', 'Generate Quote', '–'],
    ['QuoteAction', 'Hold for Now', '–'],
    ['QuoteAction', 'Revise Quote', '–'],

    // BLANK ROW FOR VISUAL SEPARATION
    ['', '', ''],

    // DROPDOWN OPTIONS - Renewal Status
    ['Renewal', 'Yes', '–'],
    ['Renewal', 'No', '–'],
    ['Renewal', 'Cancelled', '–'],
    ['Renewal', 'Pending', '–']
  ];

  // Write data to sheet
  mapSheet.getRange(1, 1, mapData.length, 3).setValues(mapData);

  // Format header row
  mapSheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground('#4285F4')
    .setFontColor('white')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Set column widths
  mapSheet.setColumnWidth(1, 300);  // Type column (wider for notification types)
  mapSheet.setColumnWidth(2, 200);  // Key column
  mapSheet.setColumnWidth(3, 350);  // Value column (wider for emails)

  // Freeze header row
  mapSheet.setFrozenRows(1);

  // Add color coding for different sections
  let currentRow = 2;

  // Services section (light green)
  for (let i = 0; i < 6; i++) {
    mapSheet.getRange(currentRow + i, 1, 1, 3).setBackground('#D9EAD3');
  }
  currentRow += 7; // 6 services + 1 blank

  // Departments section (light blue)
  for (let i = 0; i < 6; i++) {
    mapSheet.getRange(currentRow + i, 1, 1, 3).setBackground('#CFE2F3');
  }
  currentRow += 7; // 6 departments + 1 blank

  // Engagement notifications (light yellow)
  for (let i = 0; i < 4; i++) {
    mapSheet.getRange(currentRow + i, 1, 1, 3).setBackground('#FFF2CC');
  }
  currentRow += 5; // 4 emails + 1 blank

  // Paid client notifications (light orange)
  for (let i = 0; i < 4; i++) {
    mapSheet.getRange(currentRow + i, 1, 1, 3).setBackground('#FCE5CD');
  }
  currentRow += 5; // 4 emails + 1 blank

  // Dropdown options (light gray)
  const remainingRows = mapData.length - currentRow;
  if (remainingRows > 0) {
    mapSheet.getRange(currentRow, 1, remainingRows, 3).setBackground('#F3F3F3');
  }

  // Add rich text instructions in cell A1 note
  mapSheet.getRange('A1').setNote(
    '═══════════════════════════════════════════════════\n' +
    '                   MAP SHEET GUIDE\n' +
    '═══════════════════════════════════════════════════\n\n' +
    '🎯 PURPOSE:\n' +
    'This sheet controls all system configurations:\n' +
    '• Services and their document templates\n' +
    '• Departments and their email addresses\n' +
    '• Email notifications for status changes\n' +
    '• Dropdown options for Engagement Information\n\n' +
    '📋 SECTIONS:\n\n' +
    '1️⃣ SERVICES (Light Green)\n' +
    '   • Add services your firm offers\n' +
    '   • Value = Google Docs template ID (optional)\n' +
    '   • Leave blank if no template\n\n' +
    '2️⃣ DEPARTMENTS (Light Blue)\n' +
    '   • Add all departments in your firm\n' +
    '   • Value = Department email address (REQUIRED)\n' +
    '   • Used for auto-assignment and notifications\n\n' +
    '3️⃣ ENGAGEMENT NOTIFICATIONS (Light Yellow)\n' +
    '   • Type: "Email Notification for Engagement Only"\n' +
    '   • These emails get notified when status = "Engaged"\n' +
    '   • Typically: Billing, Accounting, Finance, Management\n' +
    '   • Value = Email address (REQUIRED)\n\n' +
    '4️⃣ PAID CLIENT NOTIFICATIONS (Light Orange)\n' +
    '   • Type: "Email Notification for Paid Client"\n' +
    '   • These emails get notified when payment = "Paid" + dept assigned\n' +
    '   • Typically: Operations, Service Manager, Project Manager\n' +
    '   • Value = Email address (REQUIRED)\n\n' +
    '5️⃣ DROPDOWN OPTIONS (Light Gray)\n' +
    '   • Controls dropdown lists in Engagement Information\n' +
    '   • Add/remove options as needed\n\n' +
    '⚠️ IMPORTANT RULES:\n' +
    '• Do NOT delete the header row (Row 1)\n' +
    '• Do NOT change column structure (Type, Key, Value)\n' +
    '• Email addresses must be valid format\n' +
    '• Type names must match EXACTLY (case-sensitive)\n' +
    '• After making changes, run: Menu > System Management > Refresh Map Sheet\n\n' +
    '💡 TIPS:\n' +
    '• Add as many notification emails as needed\n' +
    '• Different emails for different notification types\n' +
    '• Test with Menu > System Management > Refresh Map Sheet\n' +
    '• Check the report for errors/warnings\n\n' +
    '═══════════════════════════════════════════════════'
  );

  // Add instructions in a separate "Instructions" section
  const instructionRow = mapData.length + 2;

  mapSheet.getRange(instructionRow, 1).setValue('📖 INSTRUCTIONS');
  mapSheet.getRange(instructionRow, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground('#000000')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');

  const instructions = [
    ['', '', ''],
    ['HOW TO ADD NOTIFICATION EMAILS:', '', ''],
    ['For Engagement Status:', 'Add row with Type:', 'Email Notification for Engagement Only'],
    ['', 'Example:', 'Email Notification for Engagement Only | Billing Manager | billing@company.com'],
    ['', '', ''],
    ['For Paid Client:', 'Add row with Type:', 'Email Notification for Paid Client'],
    ['', 'Example:', 'Email Notification for Paid Client | Operations Manager | ops@company.com'],
    ['', '', ''],
    ['After making changes:', 'Run from menu:', 'System Management > Refresh Map Sheet Data'],
    ['', 'This will:', '• Validate all email addresses'],
    ['', '', '• Show detailed report of configuration'],
    ['', '', '• Update system immediately']
  ];

  mapSheet.getRange(instructionRow + 1, 1, instructions.length, 3).setValues(instructions);
  mapSheet.getRange(instructionRow + 1, 1, instructions.length, 3).setBackground('#FFFEF7');

  formatMapSheet_(mapSheet);

  Logger.log('  ✓ Map Sheet created with notification email provisions');
  Logger.log('  ✓ Total rows: ' + mapData.length);
  Logger.log('  ✓ Sections: Services, Departments, Notifications, Dropdowns');
}

function ensureMapSheetHeaders_(mapSheet) {
  const expected = ['Type', 'Key', 'Value'];
  const current = mapSheet.getRange(1, 1, 1, 3).getValues()[0];
  const blank = current.every(function(value) { return !String(value || '').trim(); });
  const valid = expected.every(function(value, index) {
    return String(current[index] || '').trim() === value;
  });

  if (blank) {
    mapSheet.getRange(1, 1, 1, 3).setValues([expected]);
  } else if (!valid) {
    throw new Error('Map Sheet must use the headers: Type, Key, Value. Existing data was not changed.');
  }
}

function formatMapSheet_(mapSheet) {
  ensureMapSheetHeaders_(mapSheet);
  const lastRow = Math.max(mapSheet.getLastRow(), 1);
  const header = mapSheet.getRange(1, 1, 1, 3);
  header
    .setFontWeight('bold')
    .setBackground('#1A73E8')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  mapSheet.setRowHeight(1, 34);
  mapSheet.setFrozenRows(1);
  mapSheet.setColumnWidth(1, 300);
  mapSheet.setColumnWidth(2, 220);
  mapSheet.setColumnWidth(3, 360);
  mapSheet.setTabColor('#5F6368');

  if (lastRow > 1) {
    const data = mapSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const fontColors = [];
    const colors = data.map(function(row) {
      const type = String(row[0] || '');
      let color = '#F8F9FA';
      if (type === 'Service' || type === 'FAQ') color = '#E6F4EA';
      else if (type === 'Department' || type === 'CategoryRouting' || type === 'TeamMember') color = '#E8F0FE';
      else if (type === 'Email Notification for Engagement Only') color = '#FEF7E0';
      else if (type === 'Email Notification for Paid Client' || type === 'ChatWebhook') color = '#FCE8E6';
      else if (type === '📖 INSTRUCTIONS') color = '#263238';
      const fontColor = type === '📖 INSTRUCTIONS' ? '#FFFFFF' : '#202124';
      fontColors.push([fontColor, fontColor, fontColor]);
      return [color, color, color];
    });
    mapSheet.getRange(2, 1, data.length, 3)
      .setBackgrounds(colors)
      .setFontColors(fontColors)
      .setVerticalAlignment('middle')
      .setWrap(true);
  }
}


/**
 * Helper: Sets up Engagement Information Sheet
 */
function setupEngagementInformationSheet(ss) {
  let infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);

  if (!infoSheet) {
    infoSheet = ss.insertSheet(CONFIG.ENGAGE_SHEET_NAME);
  } else {
    // Clear existing data but keep if it has data
    if (infoSheet.getLastRow() <= 1) {
      infoSheet.clear();
    } else {
      formatEngagementSheet_(infoSheet);
      formatEngagementSheet_(infoSheet);
      Logger.log('  ℹ Engagement Information Sheet already has data; schema verified without clearing rows');
      return;
    }
  }

  const headers = ENGAGEMENT_INFO_HEADERS;

  infoSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header
  infoSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#EA4335')
    .setFontColor('white')
    .setWrap(true)
    .setHorizontalAlignment('center');

  // Set column widths
  infoSheet.setColumnWidth(1, 100);  // Contact Date
  infoSheet.setColumnWidth(2, 180);  // Client Name
  infoSheet.setColumnWidth(3, 150);  // Project/Service
  infoSheet.setColumnWidth(4, 120);  // Contact Person
  infoSheet.setColumnWidth(5, 200);  // Email
  infoSheet.setColumnWidth(8, 250);  // Remarks
  infoSheet.setColumnWidth(14, 150); // Service Ref#
  infoSheet.setColumnWidth(23, 130); // Department
  infoSheet.setColumnWidth(ENGAGEMENT_COLUMNS.SOURCE_MONTH, 100);

  // Freeze header row
  infoSheet.setFrozenRows(1);
  formatEngagementSheet_(infoSheet);

  Logger.log('  ✓ Engagement Information Sheet created with ' + headers.length + ' columns');
}

/**
 * Helper: Sets up automatic triggers for periodic sync
 */
function setupAutomationTriggers() {
  try {
    assertMonitoredMailboxAccount_('automation trigger setup');
    const managedHandlers = [
      'processAutoDrafts',
      'processPendingAiQualifications',
      'processFollowUpReminders',
      'syncNewEmails',
      'buildEnhancedDashboard',
      'onEditInstallable'
    ];

    // Remove only CRM-managed triggers. Unrelated project triggers are preserved.
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      if (managedHandlers.indexOf(triggers[i].getHandlerFunction()) !== -1) {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }

    ScriptApp.newTrigger('processAutoDrafts')
      .timeBased()
      .everyMinutes(5)
      .create();

    ScriptApp.newTrigger('processPendingAiQualifications')
      .timeBased()
      .everyMinutes(10)
      .create();

    ScriptApp.newTrigger('processFollowUpReminders')
      .timeBased()
      .atHour(9)
      .everyDays(1)
      .create();

    ScriptApp.newTrigger('syncNewEmails')
      .timeBased()
      .everyHours(4)
      .create();

    ScriptApp.newTrigger('buildEnhancedDashboard')
      .timeBased()
      .atHour(CONFIG.AUTO_REFRESH_HOUR)
      .everyDays(1)
      .create();

    ScriptApp.newTrigger('onEditInstallable')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit()
      .create();

    Logger.log('  ✓ CRM triggers configured: AI drafts, AI qualification, follow-up, sync, Dashboard, and on-edit');

  } catch (error) {
    Logger.log('  ✗ Could not set up CRM triggers: ' + error.message);
    throw error;
  }
}

function ensureAiPendingLabel_() {
  let label = GmailApp.getUserLabelByName(AI_CONFIG.AI_PENDING_LABEL);
  if (!label) {
    label = GmailApp.createLabel(AI_CONFIG.AI_PENDING_LABEL);
    Logger.log('  ✓ Created Gmail label: ' + AI_CONFIG.AI_PENDING_LABEL);
  }
  return label;
}

/**
 * Helper: Organizes sheets in logical order
 */
function organizeSheets(ss) {
  try {
    const sheets = ss.getSheets();

    const coreOrder = [
      CONFIG.DASHBOARD_SHEET_NAME,
      CONFIG.POTENTIAL_CLIENTS_SHEET_NAME,
      CONFIG.ENGAGE_SHEET_NAME,
      CONFIG.CONVERSION_SHEET_NAME,
      CONFIG.MAP_SHEET_NAME
    ];
    const monthlyOrder = sheets
      .filter(function(sheet) { return /^[A-Za-z]{3}-\d{4}$/.test(sheet.getName()); })
      .sort(function(a, b) { return parseMonthSheetName(b.getName()) - parseMonthSheetName(a.getName()); })
      .map(function(sheet) { return sheet.getName(); });
    const desiredOrder = coreOrder.concat(monthlyOrder);

    // Move sheets to desired positions
    for (let i = 0; i < desiredOrder.length; i++) {
      const sheetName = desiredOrder[i];
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        ss.setActiveSheet(sheet);
        ss.moveActiveSheet(i + 1);
      }
    }

    // Open on Dashboard when available, otherwise Engagement Information.
    const landingSheet = ss.getSheetByName(CONFIG.DASHBOARD_SHEET_NAME) ||
      ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);
    if (landingSheet) {
      ss.setActiveSheet(landingSheet);
    }

    hideUnusedDefaultSheets_(ss, desiredOrder);

    Logger.log('  ✓ Sheets organized in logical order');

  } catch (error) {
    Logger.log('  ⚠️ Warning: Could not organize sheets: ' + error.message);
  }
}

function hideUnusedDefaultSheets_(ss, retainedNames) {
  const defaultNames = ['Sheet1', 'Sheet 1'];
  ss.getSheets().forEach(function(sheet) {
    if (defaultNames.indexOf(sheet.getName()) === -1 || retainedNames.indexOf(sheet.getName()) !== -1) return;
    const isBlank = sheet.getLastRow() <= 1 && sheet.getLastColumn() <= 1 && !sheet.getRange('A1').getValue();
    if (isBlank && !sheet.isSheetHidden()) {
      sheet.hideSheet();
      Logger.log('  ✓ Hid unused default sheet: ' + sheet.getName());
    }
  });
}

function formatInitializedWorkbook_(ss) {
  const mapSheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);
  const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);
  const conversionSheet = ss.getSheetByName(CONFIG.CONVERSION_SHEET_NAME);
  const potentialSheet = ss.getSheetByName(CONFIG.POTENTIAL_CLIENTS_SHEET_NAME);
  const dashboard = ss.getSheetByName(CONFIG.DASHBOARD_SHEET_NAME);

  if (mapSheet) formatMapSheet_(mapSheet);
  if (infoSheet) formatEngagementSheet_(infoSheet);
  if (conversionSheet) formatConversionTrackingSheet_(conversionSheet);
  if (potentialSheet) formatPotentialClientsSheet_(potentialSheet);
  if (dashboard) {
    dashboard.setTabColor('#1A73E8');
    dashboard.setHiddenGridlines(true);
  }

  ss.getSheets().forEach(function(sheet) {
    if (/^[A-Za-z]{3}-\d{4}$/.test(sheet.getName())) {
      try {
        formatMonthlySheet_(sheet);
      } catch (error) {
        Logger.log('  ⚠ Skipped monthly formatting for ' + sheet.getName() + ': ' + error.message);
      }
    }
  });
}

function formatEngagementSheet_(sheet) {
  if (!ensureEngagementSheetSchema_(sheet)) {
    throw new Error('Engagement Information Sheet has a customized or invalid core schema. Existing data was preserved.');
  }
  const columns = ENGAGEMENT_INFO_HEADERS.length;
  sheet.getRange(1, 1, 1, columns)
    .setFontWeight('bold')
    .setBackground('#B3261E')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(1, 42);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  sheet.setTabColor('#B3261E');

  const widths = {
    1: 110, 2: 190, 3: 160, 4: 150, 5: 210, 6: 125, 7: 210, 8: 260,
    9: 110, 10: 100, 11: 110, 12: 110, 13: 125, 14: 145, 15: 135,
    16: 120, 17: 130, 18: 135, 19: 115, 20: 120, 21: 125, 22: 110,
    23: 145, 24: 220, 25: 120, 26: 210, 27: 190, 28: 190, 29: 125,
    30: 105, 31: 145, 32: 165, 33: 90, 34: 135, 35: 150, 36: 125, 37: 125
  };
  Object.keys(widths).forEach(function(column) {
    sheet.setColumnWidth(Number(column), widths[column]);
  });

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const rows = lastRow - 1;
    [1, 16, 20, 25, 29, 31, 32, 37].forEach(function(column) {
      sheet.getRange(2, column, rows, 1).setNumberFormat('mmm d, yyyy');
    });
    sheet.getRange(2, 9, rows, 5).setNumberFormat('₱#,##0.00');
    sheet.getRange(2, 1, rows, columns).setVerticalAlignment('middle');
  }
  ensureSheetFilter_(sheet, columns);
  applyDefaultConditionalFormats_(sheet, 'engagement');
}

function formatConversionTrackingSheet_(sheet) {
  const columns = CONVERSION_TRACKING_HEADERS.length;
  sheet.getRange(1, 1, 1, columns)
    .setFontWeight('bold')
    .setBackground('#188038')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  sheet.setTabColor('#188038');
  [200, 180, 260, 125, 125, 90, 105, 320].forEach(function(width, index) {
    sheet.setColumnWidth(index + 1, width);
  });
  if (sheet.getLastRow() > 1) {
    const rows = sheet.getLastRow() - 1;
    sheet.getRange(2, 4, rows, 2).setNumberFormat('mmm d, yyyy');
    sheet.getRange(2, 6, rows, 1).setNumberFormat('0');
    sheet.getRange(2, 1, rows, columns).setVerticalAlignment('middle');
  }
  ensureSheetFilter_(sheet, columns);
  applyDefaultConditionalFormats_(sheet, 'conversion');
}

function formatMonthlySheet_(sheet) {
  const columns = MONTHLY_EMAIL_HEADERS.length;
  const currentHeaders = sheet.getRange(1, 1, 1, columns).getValues()[0];
  const headersBlank = currentHeaders.every(function(value) { return !String(value || '').trim(); });
  const headersValid = MONTHLY_EMAIL_HEADERS.every(function(header, index) {
    return String(currentHeaders[index] || '').trim() === header;
  });
  const legacyHeadersValid = MONTHLY_EMAIL_HEADERS.slice(0, 10).every(function(header, index) {
    return String(currentHeaders[index] || '').trim() === header;
  }) && currentHeaders.slice(10).every(function(value) { return !String(value || '').trim(); });
  const draftHeadersValid = MONTHLY_EMAIL_HEADERS.slice(0, 17).every(function(header, index) {
    return String(currentHeaders[index] || '').trim() === header;
  }) && String(currentHeaders[17] || '').trim() === 'Qualified At';
  if (headersBlank) {
    sheet.getRange(1, 1, 1, columns).setValues([MONTHLY_EMAIL_HEADERS]);
  } else if (legacyHeadersValid) {
    sheet.getRange(1, 11, 1, MONTHLY_EMAIL_HEADERS.length - 10)
      .setValues([MONTHLY_EMAIL_HEADERS.slice(10)]);
  } else if (draftHeadersValid) {
    sheet.getRange(1, 18).setValue('Classified At');
  } else if (!headersValid) {
    throw new Error('Monthly sheet "' + sheet.getName() + '" has unexpected headers. Existing data was preserved.');
  }

  sheet.getRange(1, 1, 1, columns)
    .setFontWeight('bold')
    .setBackground('#0F766E')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  sheet.setTabColor('#0F766E');
  [125, 125, 125, 190, 210, 280, 110, 120, 105, 360, 125, 100, 135, 300, 150, 180, 180, 135].forEach(function(width, index) {
    sheet.setColumnWidth(index + 1, width);
  });
  if (sheet.getLastRow() > 1) {
    const rows = sheet.getLastRow() - 1;
    sheet.getRange(2, 1, rows, 3).setNumberFormat('mmm d, yyyy h:mm AM/PM');
    sheet.getRange(2, 18, rows, 1).setNumberFormat('mmm d, yyyy h:mm AM/PM');
    sheet.getRange(2, 1, rows, columns).setVerticalAlignment('middle');
  }
  ensureSheetFilter_(sheet, columns);
  applyDefaultConditionalFormats_(sheet, 'monthly');
}

function ensureSheetFilter_(sheet, columns) {
  if (!sheet.getFilter()) {
    sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 2), columns).createFilter();
  }
}

function applyDefaultConditionalFormats_(sheet, sheetType) {
  // Preserve any rules the user has already configured and avoid duplicates
  // when initialization is run again.
  if (sheet.getConditionalFormatRules().length > 0) return;

  const maxDataRows = Math.max(sheet.getMaxRows() - 1, 1);
  const rules = [];

  if (sheetType === 'engagement') {
    const paymentRange = sheet.getRange(2, 19, maxDataRows, 1);
    const engagementRange = sheet.getRange(2, 21, maxDataRows, 1);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Paid')
        .setBackground('#D9EAD3')
        .setFontColor('#137333')
        .setRanges([paymentRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Pending')
        .setBackground('#FEF7E0')
        .setFontColor('#8A4B00')
        .setRanges([paymentRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Engaged')
        .setBackground('#D9EAD3')
        .setFontColor('#137333')
        .setRanges([engagementRange])
        .build()
    );
  } else if (sheetType === 'conversion') {
    const tableRange = sheet.getRange(2, 1, maxDataRows, CONVERSION_TRACKING_HEADERS.length);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$G2="Client"')
        .setBackground('#E6F4EA')
        .setRanges([tableRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$G2="N/A"')
        .setBackground('#F1F3F4')
        .setFontColor('#5F6368')
        .setRanges([tableRange])
        .build()
    );
  } else if (sheetType === 'monthly') {
    const tableRange = sheet.getRange(2, 1, maxDataRows, MONTHLY_EMAIL_HEADERS.length);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$I2="Client"')
        .setBackground('#E6F4EA')
        .setRanges([tableRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$I2="Prospect"')
        .setBackground('#FEF7E0')
        .setRanges([tableRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=OR($I2="Trash",$I2="N/A")')
        .setBackground('#F1F3F4')
        .setFontColor('#5F6368')
        .setRanges([tableRange])
        .build()
    );
  }

  if (rules.length > 0) sheet.setConditionalFormatRules(rules);
}

/**
 * Helper: Shows quick start guide after setup
 */
function showQuickStartGuide() {
  const ui = SpreadsheetApp.getUi();

  const guide =
    '📖 QUICK START GUIDE\n\n' +
    '🔧 CUSTOMIZATION:\n' +
    '1. Go to "Map Sheet"\n' +
    '2. Add your services (Type: Service)\n' +
    '3. Add departments (Type: Department)\n' +
    '4. Each service can have a Google Docs template ID\n\n' +
    '📧 DAILY USE:\n' +
    '• Menu > Full Update (sync + dashboard)\n' +
    '• Check Engagement Information for new prospects\n' +
    '• Select row > Generate Quote\n' +
    '• Assign departments as needed\n\n' +
    '🎯 KEY FEATURES:\n' +
    '• Auto-sync: Every 4 hours\n' +
    '• Internal emails (@duranschulze.com, @filepino.com): N/A status\n' +
    '• Excluded emails: Trash status\n' +
    '• Prospects/Leads only in Engagement Info\n\n' +
    '💡 TIPS:\n' +
    '• Remarks column auto-filled with email subject\n' +
    '• Financial formulas calculate automatically\n' +
    '• Use filters to view specific statuses\n\n' +
    'Need help? Check execution logs or menu > About';

  ui.alert('📖 Quick Start Guide', guide, ui.ButtonSet.OK);
}

/**
 * Checks if an email should be excluded from Engagement Information Sheet
 * Internal domains (@duranschulze.com, @filepino.com) are NOT excluded - they're marked as "N/A"
 * @param {string} email - The email address to check
 * @returns {object} - {isExcluded: boolean, reason: string, isInternalDomain: boolean}
 */
function isEmailExcluded(email) {
  if (!email) {
    return { isExcluded: true, reason: 'Empty email', isInternalDomain: false };
  }

  const emailLower = email.toLowerCase().trim();

  // Check if it's internal domain - NOT excluded, just marked as "N/A"
  for (const domain of CONFIG.INTERNAL_DOMAINS) {
    if (emailLower.includes(domain.toLowerCase())) {
      return { isExcluded: false, reason: 'Internal domain (N/A status): ' + domain, isInternalDomain: true };
    }
  }

  // Check excluded domains
  for (const domain of CONFIG.EXCLUDED_DOMAINS) {
    if (emailLower.includes(domain.toLowerCase())) {
      return { isExcluded: true, reason: 'Excluded domain: ' + domain, isInternalDomain: false };
    }
  }

  // Check excluded patterns (no-reply, noreply, etc.)
  for (const pattern of CONFIG.EXCLUDED_PATTERNS) {
    if (emailLower.includes(pattern.toLowerCase())) {
      return { isExcluded: true, reason: 'Excluded pattern: ' + pattern, isInternalDomain: false };
    }
  }

  return { isExcluded: false, reason: '', isInternalDomain: false };
}


// ========================================
// SHEET SCHEMAS
// ========================================

/**
 * Engagement Information Sheet Column Headers
 * Updated structure with separate columns for tracking
 */
const ENGAGEMENT_INFO_HEADERS = [
  'Contact Date',                                    // Column 1 (A)
  'Client Name / Company Name',                      // Column 2 (B)
  'Project / Service',                               // Column 3 (C)
  'Contact Person',                                  // Column 4 (D)
  'Email Address',                                   // Column 5 (E)
  'Telephone Number',                                // Column 6 (F)
  'Registered Address',                              // Column 7 (G)
  'Remarks',                                         // Column 8 (H)
  'Engagement Fee',                                  // Column 9 (I)
  'VAT',                                             // Column 10 (J)
  'Total Gross',                                     // Column 11 (K)
  'Miscellaneous',                                   // Column 12 (L)
  'Total Service Fee',                               // Column 13 (M)
  'Service Ref#',                                    // Column 14 (N)
  'Service Quote Created?',                          // Column 15 (O)
  'Service Quote Date Sent',                         // Column 16 (P)
  'Service Quote Status',                            // Column 17 (Q)
  'Service Quote Follow-Up',                         // Column 18 (R)
  'Payment Status',                                  // Column 19 (S)
  'Payment Status Date',                             // Column 20 (T)
  'Engagement Status',                               // Column 21 (U)
  'Needs Renewal',                                   // Column 22 (V)
  'Assigned Department',                             // Column 23 (W)
  'Assigned Department Email Address',               // Column 24 (X)
  'Assigned Date',                                   // Column 25 (Y)
  'PDF - Signed Agreement and Invoice',              // Column 26 (Z) - MANUAL UPLOAD
  'File Folder',                                     // Column 27 (AA) - Hyperlink
  'Generated Quote PDF',                             // Column 28 (AB) - Auto-generated
  'Date Due for Renewal',                            // Column 29 (AC)
  'Source Month',                                    // Column 30 (AD) - Auto from Contact Date
  'Engagement Notification Sent',                    // Column 31 (AE) - Status = Engaged
  'Paid Assignment Notification Sent',                // Column 32 (AF) - Payment = Paid + Dept Assigned
  'AI Score',                // Column 33 (AG)
'Intent Category',         // Column 34 (AH)
'Assigned Team Member',    // Column 35 (AI)
'Auto-Response Sent',      // Column 36 (AJ)
'Last Follow-Up Date'      // Column 37 (AK)
];

// Canonical one-based columns for Sheet.getRange(). Array access uses column - 1.
const ENGAGEMENT_COLUMNS = {
  GENERATED_QUOTE_PDF: 28,
  DATE_DUE_FOR_RENEWAL: 29,
  SOURCE_MONTH: 30,
  ENGAGEMENT_NOTIFICATION_SENT: 31,
  PAID_ASSIGNMENT_NOTIFICATION_SENT: 32,
  AI_SCORE: 33,
  INTENT_CATEGORY: 34,
  ASSIGNED_TEAM_MEMBER: 35,
  AUTO_RESPONSE_SENT: 36,
  LAST_FOLLOW_UP_DATE: 37
};

const CONVERSION_TRACKING_COLUMNS = {
  TOTAL_EMAILS: 6,
  STATUS: 7
};

const CONVERSION_TRACKING_HEADERS = [
  "Sender Email",
  "Sender's Name",
  "Original Subject",
  "First Contact Date",
  "Date Last Contacted",
  "Total Emails",
  "Status",
  "Latest Summary"
];

const MONTHLY_EMAIL_HEADERS = [
  'Date Received',
  'Date Responded',
  'Date Follow-up',
  'Sender',
  'Sender Email',
  'Subject',
  'Email Type',
  'Meeting Requested',
  'Status',
  'Summary',
  'Qualification Status',
  'AI Confidence',
  'Intent Category',
  'Qualification Reason',
  'Gemini Model',
  'Gmail Message ID',
  'Gmail Thread ID',
  'Classified At'
];

/**
 * Finds a one-based sheet column by header and uses the supplied canonical
 * fallback only when the header is absent.
 */
function getHeaderColumn_(headers, headerName, fallbackColumn) {
  const normalizedTarget = String(headerName || '').trim().toLowerCase();
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').trim().toLowerCase() === normalizedTarget) {
      return i + 1;
    }
  }
  return fallbackColumn;
}

/**
 * Safely upgrades the known legacy Engagement layout. Older installations
 * omitted "Generated Quote PDF", leaving Date Due and Source Month one column
 * too far left. Inserting the missing column preserves the existing data.
 */
function ensureEngagementSheetSchema_(sheet) {
  let lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return false;

  let headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const generatedQuoteIndex = headers.indexOf('Generated Quote PDF');
  const dateDueIndex = headers.indexOf('Date Due for Renewal');
  const sourceMonthIndex = headers.indexOf('Source Month');

  const hasKnownLegacyLayout = generatedQuoteIndex === -1 &&
    dateDueIndex === 27 && sourceMonthIndex === 28;

  if (hasKnownLegacyLayout) {
    sheet.insertColumnBefore(ENGAGEMENT_COLUMNS.GENERATED_QUOTE_PDF);
    Logger.log('  ✓ Inserted missing Generated Quote PDF column without overwriting existing data');
  }

  lastColumn = sheet.getLastColumn();
  if (sheet.getMaxColumns() < ENGAGEMENT_INFO_HEADERS.length) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      ENGAGEMENT_INFO_HEADERS.length - sheet.getMaxColumns()
    );
  }

  headers = sheet.getRange(1, 1, 1, Math.max(lastColumn, ENGAGEMENT_INFO_HEADERS.length)).getValues()[0];
  const firstCoreHeadersMatch = ENGAGEMENT_INFO_HEADERS.slice(0, 27).every(function(header, index) {
    return String(headers[index] || '').trim() === header;
  });

  if (!firstCoreHeadersMatch) {
    Logger.log('  ⚠ Engagement schema was not automatically rewritten because its core columns are customized');
    return false;
  }

  // At this point either the sheet was already canonical or the only known
  // missing column was inserted. Updating headers is safe and preserves rows.
  sheet.getRange(1, 1, 1, ENGAGEMENT_INFO_HEADERS.length).setValues([ENGAGEMENT_INFO_HEADERS]);
  repairMisplacedSourceMonths_(sheet);
  return true;
}

/**
 * Repairs rows written by the old push logic, which placed MMM-yyyy values in
 * Date Due for Renewal (AC) instead of Source Month (AD).
 */
function repairMisplacedSourceMonths_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const startColumn = ENGAGEMENT_COLUMNS.DATE_DUE_FOR_RENEWAL;
  const values = sheet.getRange(2, startColumn, lastRow - 1, 2).getValues();
  let repaired = 0;

  for (let i = 0; i < values.length; i++) {
    const dateDueValue = values[i][0];
    const sourceMonthValue = values[i][1];
    if (/^[A-Za-z]{3}-\d{4}$/.test(String(dateDueValue || '').trim()) && !sourceMonthValue) {
      values[i][1] = dateDueValue;
      values[i][0] = '';
      repaired++;
    }
  }

  if (repaired > 0) {
    sheet.getRange(2, startColumn, values.length, 2).setValues(values);
    Logger.log('  ✓ Repaired ' + repaired + ' misplaced Source Month value(s)');
  }
  return repaired;
}



const MAP_SEED_DATA = [
  ['Type', 'Key', 'Value'],
  ['Service', 'Immigration Visa', '1g8vauhd9r437PlAzCKJh5Qt11emZa75vm-DmWxGkaUQ'],
  ['Service', 'Litigation', '1NTfGpOGUSR08GOMImFM3EnT80tSn_zDSgBeo_XIVQNc'],
  ['Department', 'Legal', 'legal@company.com'],
  ['Department', 'Visa', 'visa@company.com'],
  ['Department', 'HR', 'hr@company.com'],
  ['Department', 'Accounting', 'accounting@company.com'],
  ['FollowUp', 'Follow Up', '–'],
  ['FollowUp', 'Declined', '–'],
  ['Payment', 'Paid', '–'],
  ['Payment', 'Unpaid', '–'],
  ['QuoteAction', 'Generate Quote', '–'],
  ['QuoteAction', 'Hold for Now', '–'],
  ['Renewal', 'Yes', '–'],
  ['Renewal', 'No', '–'],
  ['Renewal', 'Cancelled', '–']
];

// ========================================
// MENU SYSTEM
// ========================================

function onOpen() {

  createEnhancedCRMMenu();
}

function createEnhancedCRMMenu() {
  try {
    const ui = SpreadsheetApp.getUi();

    ui.createMenu('📧 CRM Tracker')
      .addItem('🚀 Setup / Upgrade Workbook', 'initialSystemSetup')
      .addSeparator()
      .addItem('🔄 Sync New Emails', 'menuSyncEmails')
      .addItem('📊 Refresh Dashboard', 'menuRefreshDashboard')
      .addItem('⚡ Full Update (Sync + Dashboard)', 'menuFullUpdate')
      .addSeparator()
      .addItem('⏰ ONE-TIME: Sync Historical Emails (Jan 2024)', 'syncHistoricalEmailsFromJan2024')
      .addItem('📅 Backfill a Custom Date Range', 'menuBackfillEmailsByDateRange')
      .addSeparator()
      .addSubMenu(ui.createMenu('🎯 Potential Clients')
        .addItem('🔄 Process Pending AI Qualifications', 'menuProcessPendingAiQualifications')
        .addItem('🧠 Reclassify Selected Candidate', 'menuReclassifySelectedPotentialClient')
        .addItem('👍 Approve Selected as Qualified', 'menuApproveSelectedPotentialClient')
        .addItem('✅ Promote Selected to Engagement', 'menuPromoteSelectedPotentialClient')
        .addItem('🚫 Reject Selected Candidate', 'menuRejectSelectedPotentialClient'))
      .addSeparator()
      .addSubMenu(ui.createMenu('🤖 AI Auto-Reply')
      .addItem('🔄 Process AI-Pending Emails', 'processAutoDrafts')
      .addItem('📊 Show Auto-Responded Emails', 'showAutoRespondedEmails')
      .addItem('🛡️ Show AI Usage & Limits', 'menuShowAiUsage')
      .addItem('⏯️ Enable / Disable AI Processing', 'menuToggleAiProcessing')
      .addItem('⚙️ Manage Category Assignments', 'manageCategoryAssignments')
      .addItem('💾 Save & Hide Old Sheets', 'performMaintenance')
      .addItem('🔑 Configure Gemini Model', 'setAPIKeyAndModel')
      .addItem('🛠️ Setup AI Triggers', 'setupAllTriggers'))
        .addSeparator()
      .addSubMenu(ui.createMenu('📋 Engagement Actions')
        .addItem('📝 Generate Quote for Selected Row', 'menuGenerateQuote')
        .addItem('👨 Assign Department to Row', 'menuAssignDepartment')
        .addItem('💰 Mark Payment Received', 'menuMarkPaymentReceived')
        .addItem('📁 Open File Folder', 'menuOpenFileFolder'))
      .addSeparator()
      .addSubMenu(ui.createMenu('🔧 System Management')
        .addItem('⚙️ Setup Edit Trigger (Enable Notifications)', 'setupEditTrigger') // NEW - TOP
        .addItem('🔍 Show Current Triggers', 'showCurrentTriggersFixed') // NEW
        .addItem('🔍 Verify Notification Columns', 'verifyNotificationColumns')
        .addItem('🧹 Remove All Triggers', 'removeAllTriggersFixed') // NEW
        .addSeparator()
        .addItem('🔄 Refresh Map Sheet Data', 'refreshMapSheet')
        .addSeparator()
        .addItem('📋 Refresh Dropdowns from Map', 'menuRefreshDropdowns')
        .addItem('🧮 Refresh Financial Formulas', 'refreshAllFinancialFormulas')
        .addItem('📅 Populate All Source Months', 'populateAllSourceMonths')
        .addItem('🔄 Recalculate All Renewal Dates', 'recalculateAllRenewalDates')
        .addItem('🧹 Clean All Excluded Emails', 'cleanupAllExcludedEmails')
        .addItem('🎨 Reformat Monthly Sheets Colors', 'reformatMonthlySheets')
        .addItem('🔄 Update Internal Domains to N/A', 'updateInternalDomainsToNA')
        .addItem('🔧 Fix Sender Names (Email→Name)', 'fixConversionTrackingSenderNames')
        .addItem('📁 Set Root Drive Folder', 'menuSetRootFolder')
        .addItem('🧹 Run Maintenance', 'menuRunMaintenance')
        .addItem('📮 Verify Automation Account', 'menuVerifyAutomationAccount')
        .addItem('🔧 Test Configuration', 'menuTestConfiguration'))
      .addSeparator()
      .addItem('📤 Export All Data', 'menuExportAllData')
      .addItem('📤 Export Current Month', 'menuExportCurrentMonth')
      .addSeparator()
      .addSubMenu(ui.createMenu('ℹ️ INFO')
        .addItem('📖 About Mini-CRM', 'menuShowAbout')
        .addItem('👥 Credits', 'menuShowCredits'))
      .addToUi();

    SpreadsheetApp.getActiveSpreadsheet().toast(
      '📧 CRM Tracker menu loaded!',
      'System Ready',
      3
    );

  } catch (error) {
    Logger.log('Error creating menu: ' + error.message);
  }
}

// ========================================
// CORE SHEET MANAGEMENT
// ========================================

function ensureCoreSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Map Sheet
    let mapSheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);
    if (!mapSheet) {
      mapSheet = ss.insertSheet(CONFIG.MAP_SHEET_NAME);
      mapSheet.getRange(1, 1, MAP_SEED_DATA.length, 3).setValues(MAP_SEED_DATA);
      mapSheet.getRange(1, 1, 1, 3)
        .setFontWeight('bold')
        .setBackground('#4285F4')
        .setFontColor('white');
      mapSheet.setFrozenRows(1);
    }

    // 2. Engagement Information Sheet (NOT Dashboard - Dashboard removed completely)
    let infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);
    if (!infoSheet) {
      infoSheet = ss.insertSheet(CONFIG.ENGAGE_SHEET_NAME);
      infoSheet.getRange(1, 1, 1, ENGAGEMENT_INFO_HEADERS.length).setValues([ENGAGEMENT_INFO_HEADERS]);
      infoSheet.getRange(1, 1, 1, ENGAGEMENT_INFO_HEADERS.length)
        .setFontWeight('bold')
        .setBackground('#EA4335')
        .setFontColor('white')
        .setWrap(true);
      infoSheet.setFrozenRows(1);

      // Set column widths for better visibility
      infoSheet.setColumnWidth(1, 100);  // Contact Date
      infoSheet.setColumnWidth(2, 180);  // Client Name
      infoSheet.setColumnWidth(3, 150);  // Project/Service
      infoSheet.setColumnWidth(4, 120);  // Contact Person
      infoSheet.setColumnWidth(5, 200);  // Email
      infoSheet.setColumnWidth(8, 200);  // Remarks
      infoSheet.setColumnWidth(14, 150); // Service Ref#
      infoSheet.setColumnWidth(23, 130); // Department
      infoSheet.setColumnWidth(ENGAGEMENT_COLUMNS.SOURCE_MONTH, 100);
    } else {
      ensureEngagementSheetSchema_(infoSheet);
    }

    // Apply formulas and dropdowns
    setupFinancialFormulas(infoSheet);
    setupDropdownValidations(infoSheet);

    Logger.log('Core sheets ensured successfully');

  } catch (error) {
    Logger.log('Error ensuring core sheets: ' + error.message);
    throw error;
  }
}

function setupFinancialFormulas(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    Logger.log('Setting up financial formulas for rows 2 to ' + lastRow);

    // Apply formulas to all data rows
    for (let row = 2; row <= lastRow; row++) {
      // Column J (10): VAT = Engagement Fee × 12%
      const vatFormula = `=IF(OR(I${row}="",I${row}=0),"",I${row}*${CONFIG.VAT_RATE})`;
      sheet.getRange(row, 10).setFormula(vatFormula);

      // Column K (11): Total Gross = Engagement Fee + VAT
      const grossFormula = `=IF(OR(I${row}="",I${row}=0),"",I${row}+J${row})`;
      sheet.getRange(row, 11).setFormula(grossFormula);

      // Column M (13): Total Service Fee = Total Gross + Miscellaneous
      const totalFormula = `=IF(OR(K${row}="",K${row}=0),"",K${row}+IF(OR(L${row}="",ISBLANK(L${row})),0,L${row}))`;
      sheet.getRange(row, 13).setFormula(totalFormula);
    }

    Logger.log('✓ Financial formulas applied successfully to ' + (lastRow - 1) + ' rows');

  } catch (error) {
    Logger.log('✗ Error setting up financial formulas: ' + error.message);
  }
}


function setupDropdownValidations(sheet) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mapSheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);
    if (!mapSheet) return;

    const lastRow = Math.max(sheet.getLastRow(), 2);
    const maxRows = Math.max(lastRow - 1, 500); // Apply to at least 500 rows for future entries

    // Get map data
    const mapData = mapSheet.getDataRange().getValues();

    // Build dropdown lists from Map Sheet
    const services = mapData.filter(r => r[0] === 'Service').map(r => r[1]);
    const departments = mapData.filter(r => r[0] === 'Department').map(r => r[1]);
    const followUps = mapData.filter(r => r[0] === 'FollowUp').map(r => r[1]);
    const payments = mapData.filter(r => r[0] === 'Payment').map(r => r[1]);
    const quoteActions = mapData.filter(r => r[0] === 'QuoteAction').map(r => r[1]);
    const renewals = mapData.filter(r => r[0] === 'Renewal').map(r => r[1]);

    Logger.log('Dropdown lists extracted:');
    Logger.log('Services: ' + services.join(', '));
    Logger.log('Departments: ' + departments.join(', '));
    Logger.log('Quote Actions: ' + quoteActions.join(', '));
    Logger.log('Follow Ups: ' + followUps.join(', '));
    Logger.log('Payments: ' + payments.join(', '));
    Logger.log('Renewals: ' + renewals.join(', '));

    // Column 3: Project / Service
    if (services.length > 0) {
      const serviceRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(services, true)
        .setAllowInvalid(false)
        .setHelpText('Select a service from the list')
        .build();
      sheet.getRange(2, 3, maxRows, 1).setDataValidation(serviceRule);
      Logger.log('Applied Service dropdown to column 3');
    }

    // Column 15: Service Quote Created?
    if (quoteActions.length > 0) {
      const quoteRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(quoteActions, true)
        .setAllowInvalid(false)
        .setHelpText('Generate Quote or Hold for Now')
        .build();
      sheet.getRange(2, 15, maxRows, 1).setDataValidation(quoteRule);
      Logger.log('Applied Quote Action dropdown to column 15');
    }

    // Column 18: Service Quote Follow-Up
    if (followUps.length > 0) {
      const followUpRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(followUps, true)
        .setAllowInvalid(false)
        .setHelpText('Follow Up or Declined')
        .build();
      sheet.getRange(2, 18, maxRows, 1).setDataValidation(followUpRule);
      Logger.log('Applied Follow-Up dropdown to column 18');
    }

    // Column 19: Payment Status
    if (payments.length > 0) {
      const paymentRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(payments, true)
        .setAllowInvalid(false)
        .setHelpText('Paid or Unpaid')
        .build();
      sheet.getRange(2, 19, maxRows, 1).setDataValidation(paymentRule);
      Logger.log('Applied Payment dropdown to column 19');
    }

    // Column 22: Needs Renewal
    if (renewals.length > 0) {
      const renewalRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(renewals, true)
        .setAllowInvalid(false)
        .setHelpText('Yes, No, or Cancelled')
        .build();
      sheet.getRange(2, 22, maxRows, 1).setDataValidation(renewalRule);
      Logger.log('Applied Renewal dropdown to column 22');
    }

    // Column 23: Assigned Department
    if (departments.length > 0) {
      const deptRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(departments, true)
        .setAllowInvalid(false)
        .setHelpText('Select department')
        .build();
      sheet.getRange(2, 23, maxRows, 1).setDataValidation(deptRule);
      Logger.log('Applied Department dropdown to column 23');
    }

    Logger.log('All dropdown validations applied successfully');

  } catch (error) {
    Logger.log('Error setting up dropdowns: ' + error.message);
  }
}

// ========================================
// GMAIL SYNC WITH MONTHLY SHEETS
// ========================================

function syncNewEmails() {
  assertMonitoredMailboxAccount_('Gmail sync');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    Logger.log('Email sync skipped because another sync is already running.');
    return { success: false, skipped: true, reason: 'Another email sync is active' };
  }
  try {
    return syncNewEmailsUnlocked_();
  } finally {
    lock.releaseLock();
  }
}

function syncNewEmailsUnlocked_() {
  try {
    Logger.log('Starting enhanced email sync v' + CRM_VERSION + ' with AI qualification...');

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Fetch Gmail threads
    const threads = GmailApp.search(CONFIG.GMAIL_SEARCH_QUERY, 0, CONFIG.MAX_THREADS);
    const identityIndex = getExistingEmailIdentityIndex_(ss);
    const mailboxContext = getMonitoredMailboxContext_();
    const userEmail = mailboxContext.primaryEmail;

    let newEmailsCount = 0;
    let outboundSkipped = 0;
    let duplicateSkipped = 0;
    const monthlySheetCache = {};
    const qualificationBudget = { remaining: AI_CONFIG.MAX_QUALIFICATIONS_PER_SYNC };

    // Process each thread
    for (let i = 0; i < threads.length; i++) {
      const thread = threads[i];
      const messages = thread.getMessages();

      for (let j = 0; j < messages.length; j++) {
        const msg = messages[j];
        const inbound = inspectInboundMessage_(msg, mailboxContext);
        if (!inbound.isInbound) {
          outboundSkipped++;
          continue;
        }
        const dateReceived = safeGetDate(msg.getDate());
        const senderEmail = extractEmailAddress(msg.getFrom());
        const subject = msg.getSubject() || "(No Subject)";

        const identity = inspectDuplicateMessage_(msg, senderEmail, dateReceived, subject, identityIndex);
        if (identity.isDuplicate) {
          duplicateSkipped++;
          continue;
        }

        // Process the email
        let emailData = processEmailMessage(msg, thread, userEmail, identityIndex.fallbackIds);
        emailData = qualifyEmailMessageForCrm_(msg, thread, emailData, qualificationBudget);

        // Route to the CORRECT month-year sheet based on email's received date
        const emailMonth = Utilities.formatDate(dateReceived, CONFIG.TIMEZONE, 'MMM-yyyy');

        // Get or create the month sheet for THIS email's date
        let monthSheet = monthlySheetCache[emailMonth];
        if (!monthSheet) {
          monthSheet = ss.getSheetByName(emailMonth);
          if (!monthSheet) {
            monthSheet = ss.insertSheet(emailMonth);
            monthSheet.getRange(1, 1, 1, MONTHLY_EMAIL_HEADERS.length)
              .setValues([MONTHLY_EMAIL_HEADERS]);
            Logger.log('Created new month sheet: ' + emailMonth);
          }
          formatMonthlySheet_(monthSheet);
          monthlySheetCache[emailMonth] = monthSheet;
        }

        // Append email to the CORRECT month's sheet
        monthSheet.appendRow(emailData);
        recordMessageIdentity_(identityIndex, identity);

        // Apply LIGHT GRAY formatting for Trash and N/A status
        const status = emailData[8]; // Column I (Status)
        if (status === 'Trash' || status === 'N/A') {
          const lastRow = monthSheet.getLastRow();
          monthSheet.getRange(lastRow, 1, 1, emailData.length)
            .setBackground('#D3D3D3')  // Light gray
            .setFontColor('#666666');  // Dark gray text
        }

        newEmailsCount++;

        // Update Conversion Tracking
        const conversionSheet = getOrCreateConversionTrackingSheet(ss);
        updateConversionTracking(conversionSheet, emailData);
      }
    }

    Logger.log('Emails routed to ' + Object.keys(monthlySheetCache).length + ' different month sheets');
    Logger.log('Month sheets used: ' + Object.keys(monthlySheetCache).join(', '));
    Logger.log('Skipped ' + outboundSkipped + ' non-inbound messages and ' + duplicateSkipped + ' duplicate messages.');

    // Refresh the AI-assisted review queue. Promotion to Engagement is manual.
    syncPotentialClientsFromMonthlySheets_(ss);

    // Process special emails (Dropbox Sign, Payments)
    processDropboxSignEmails(ss);
    processPaymentEmails(ss);

    Logger.log('Email sync complete: ' + newEmailsCount + ' new emails');

    return {
      success: true,
      newEmails: newEmailsCount,
      outboundSkipped: outboundSkipped,
      duplicateSkipped: duplicateSkipped,
      totalThreads: threads.length,
      monthsAffected: Object.keys(monthlySheetCache)
    };

  } catch (error) {
    Logger.log('Error in syncNewEmails: ' + error.message);
    throw error;
  }
}

/**
 * ========================================
 * INSTALLABLE ONEDIT TRIGGER v5.0
 * ========================================
 * This is an INSTALLABLE trigger that can send emails
 * Must be installed via setupEditTrigger() function
 */
function onEditInstallable(e) {
  try {
    assertMonitoredMailboxAccount_('installable edit automation');
    // Check if event object exists
    if (!e) {
      Logger.log('No event object - trigger not fired properly');
      return;
    }

    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();

    Logger.log('========================================');
    Logger.log('📧 INSTALLABLE onEdit TRIGGERED');
    Logger.log('Sheet: ' + sheetName);
    Logger.log('Target Sheet: ' + CONFIG.ENGAGE_SHEET_NAME);
    Logger.log('========================================');

    const row = e.range.getRow();
    const col = e.range.getColumn();
    const value = e.value || e.range.getValue();

    // Record direct reviewer decisions so later AI syncs do not overwrite them.
    if (sheetName === CONFIG.POTENTIAL_CLIENTS_SHEET_NAME) {
      if (row >= 2 && col === POTENTIAL_CLIENT_COLUMNS.STATUS) {
        sheet.getRange(row, POTENTIAL_CLIENT_COLUMNS.DECISION_SOURCE).setValue(getManualDecisionSource_());
        Logger.log('Potential-client status manually changed to: ' + value);
        buildEnhancedDashboard();
      }
      return;
    }

    // Only process other edits in Engagement Information Sheet
    if (sheetName !== CONFIG.ENGAGE_SHEET_NAME) {
      Logger.log('Not Engagement Information Sheet - exiting');
      return;
    }

    // Skip header row
    if (row < 2) {
      Logger.log('Header row - skipping');
      return;
    }

    Logger.log('Row: ' + row);
    Logger.log('Column: ' + col);
    Logger.log('Column Name: ' + getColumnName(col));
    Logger.log('New Value: ' + value);
    Logger.log('========================================');

    // Column 1 (A): Contact Date changed → Update Source Month
    if (col === 1) {
      Logger.log('🔔 Contact Date changed - updating Source Month');
      handleContactDateChange(sheet, row);
    }

    // Column 19 (S): Payment Status changed
    if (col === 19) {
      Logger.log('🔔 Payment Status changed to: ' + value);
      handlePaymentStatusChange(sheet, row);
    }

    // Column 21 (U): Engagement Status changed
    if (col === 21) {
      Logger.log('🔔 Engagement Status changed to: ' + value);
      handleEngagementStatusChange(sheet, row);
    }

    // Column 23 (W): Assigned Department changed
    if (col === 23) {
      Logger.log('🔔 Assigned Department changed to: ' + value);
      handleDepartmentAssignment(sheet, row);
    }

    Logger.log('========================================');
    Logger.log('✅ onEdit processing complete');
    Logger.log('========================================');

  } catch (error) {
    Logger.log('========================================');
    Logger.log('❌ ERROR in onEditInstallable');
    Logger.log('Error: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('Line: ' + error.lineNumber);
    Logger.log('========================================');

    // Show error to user
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Error in onEdit: ' + error.message + '\n\nCheck execution logs for details.',
        '❌ Error',
        10
      );
    } catch (e) {
      // Ignore toast errors
    }
  }
}

/**
 * ========================================
 * SETUP INSTALLABLE ONEDIT TRIGGER
 * ========================================
 * Installs the onEdit trigger that can send emails
 * Run this ONCE to set up
 */
function setupEditTrigger() {
  const ui = SpreadsheetApp.getUi();

  try {
    assertMonitoredMailboxAccount_('edit trigger setup');
    Logger.log('========================================');
    Logger.log('SETTING UP INSTALLABLE ONEDIT TRIGGER');
    Logger.log('========================================');

    // Remove any existing onEdit triggers first
    const triggers = ScriptApp.getProjectTriggers();
    let removedCount = 0;

    for (let i = 0; i < triggers.length; i++) {
      const trigger = triggers[i];
      const eventType = trigger.getEventType();

      // Remove any existing onEdit triggers
      if (eventType === ScriptApp.EventType.ON_EDIT) {
        ScriptApp.deleteTrigger(trigger);
        removedCount++;
        Logger.log('Removed old onEdit trigger: ' + trigger.getHandlerFunction());
      }
    }

    if (removedCount > 0) {
      Logger.log('Removed ' + removedCount + ' old trigger(s)');
    }

    // Create new installable onEdit trigger
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    ScriptApp.newTrigger('onEditInstallable')
      .forSpreadsheet(ss)
      .onEdit()
      .create();

    Logger.log('✅ Installable onEdit trigger created successfully');
    Logger.log('Handler function: onEditInstallable');
    Logger.log('========================================');

    ui.alert(
      '✅ Trigger Installed',
      'Installable onEdit trigger has been set up successfully!\n\n' +
      'Features enabled:\n' +
      '✓ Auto-populate department email\n' +
      '✓ Auto-update source month\n' +
      '✓ Send engagement notifications\n' +
      '✓ Send paid client notifications\n' +
      '✓ Calculate renewal dates\n\n' +
      'The trigger is now active and will work automatically.',
      ui.ButtonSet.OK
    );

  } catch (error) {
    Logger.log('========================================');
    Logger.log('❌ ERROR setting up trigger');
    Logger.log('Error: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========================================');

    ui.alert(
      '❌ Setup Failed',
      'Error setting up trigger:\n\n' + error.message + '\n\n' +
      'Make sure you have authorized the script with all permissions.',
      ui.ButtonSet.OK
    );
  }
}


/**
 * Helper function to get column name (A, B, C, etc.)
 */
function getColumnName(col) {
  const columns = ['', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF'];
  return columns[col] || 'Column ' + col;
}



/**
 * Handles Payment Status changes
 * - Alerts user if "Paid" but no department assigned
 * - Triggers notification if both Paid + Department assigned
 * @param {Sheet} sheet - The Engagement Information Sheet
 * @param {number} row - The row number that was edited
 */
function handlePaymentStatusChange(sheet, row) {
  try {
    Logger.log('Processing Payment Status change for row ' + row);

    const paymentStatus = sheet.getRange(row, 19).getValue(); // Column S

    // Only proceed if status is "Paid"
    if (paymentStatus !== 'Paid') {
      Logger.log('  Payment Status is not "Paid", skipping');
      return;
    }

    // Check if department is assigned
    const assignedDepartment = sheet.getRange(row, 23).getValue(); // Column W

    if (!assignedDepartment || assignedDepartment.toString().trim() === '') {
      // ALERT USER: Payment is Paid but no department assigned
      Logger.log('  ⚠️ Payment is Paid but no department assigned');

      // Highlight the Assigned Department cell
      sheet.getRange(row, 23)
        .setBackground('#FFD966') // Yellow highlight
        .setBorder(true, true, true, true, null, null, '#FF6B6B', SpreadsheetApp.BorderStyle.SOLID_THICK);

      // Show toast notification
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Please assign a department for this paid client (Column W)',
        '⚠️ Payment Received - Department Required',
        10
      );

      return;
    }

    // Both conditions met: Paid + Department Assigned
    // Check if notification already sent
    const notificationSentDate = sheet.getRange(row, 32).getValue(); // Column AF

    if (notificationSentDate) {
      Logger.log('  Paid assignment notification already sent on: ' + notificationSentDate);
      return;
    }

    // Send "Paid & Assigned" notification
    sendPaidAssignmentNotification(sheet, row);

  } catch (error) {
    Logger.log('ERROR in handlePaymentStatusChange: ' + error.message);
  }
}


/**
 * Auto-updates Source Month when Contact Date changes
 * @param {Sheet} sheet - The Engagement Information Sheet
 * @param {number} row - The row number that was edited
 */
function handleContactDateChange(sheet, row) {
  try {
    Logger.log('Processing Contact Date change for row ' + row);

    const contactDate = sheet.getRange(row, 1).getValue(); // Column A

    if (!contactDate || !(contactDate instanceof Date)) {
      // Clear Source Month if date is invalid
      sheet.getRange(row, ENGAGEMENT_COLUMNS.SOURCE_MONTH).setValue('');
      Logger.log('  Cleared Source Month (invalid date)');
      return;
    }

    // Format as MMM-yyyy
    const sourceMonth = Utilities.formatDate(contactDate, CONFIG.TIMEZONE, 'MMM-yyyy');

    // Update Source Month
    sheet.getRange(row, ENGAGEMENT_COLUMNS.SOURCE_MONTH).setValue(sourceMonth);

    Logger.log('  ✓ Source Month updated to: ' + sourceMonth);

    // Show toast notification
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Source Month: ' + sourceMonth,
      '✓ Updated',
      2
    );

  } catch (error) {
    Logger.log('ERROR in handleContactDateChange: ' + error.message);
  }
}


/**
 * Handles department assignment
 * - Auto-populates department email
 * - Removes highlight if present
 * - Triggers notification if Payment = Paid
 * @param {Sheet} sheet - The Engagement Information Sheet
 * @param {number} row - The row number that was edited
 */
function handleDepartmentAssignment(sheet, row) {
  try {
    Logger.log('Processing department assignment for row ' + row);

    const departmentName = sheet.getRange(row, 23).getValue(); // Column W

    if (!departmentName || departmentName.toString().trim() === '') {
      // Clear department email if department is cleared
      sheet.getRange(row, 24).setValue(''); // Column X
      Logger.log('  Cleared department email (department removed)');
      return;
    }

    // Look up department email from Map Sheet
    const departmentEmail = getDepartmentEmail(departmentName);

    if (departmentEmail) {
      // Populate department email
      sheet.getRange(row, 24).setValue(departmentEmail); // Column X

      // Set assigned date if not already set
      const assignedDate = sheet.getRange(row, 25).getValue(); // Column Y
      if (!assignedDate) {
        sheet.getRange(row, 25).setValue(new Date()); // Column Y
      }

      // Remove highlight and border if present
      sheet.getRange(row, 23)
        .setBackground(null)
        .setBorder(false, false, false, false, false, false);

      Logger.log('  ✓ Department email populated: ' + departmentEmail);

      // Show toast notification
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Department email: ' + departmentEmail,
        '✓ ' + departmentName + ' assigned',
        3
      );

      // Check if payment is already "Paid"
      const paymentStatus = sheet.getRange(row, 19).getValue(); // Column S

      if (paymentStatus === 'Paid') {
        // Check if notification already sent
        const notificationSentDate = sheet.getRange(row, 32).getValue(); // Column AF

        if (!notificationSentDate) {
          // Send "Paid & Assigned" notification
          Logger.log('  Payment is Paid and Department now assigned, sending notification');
          sendPaidAssignmentNotification(sheet, row);
        }
      }

    } else {
      Logger.log('  ⚠️ No email found for department: ' + departmentName);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Please add email for "' + departmentName + '" in Map Sheet',
        '⚠️ Department Email Not Found',
        5
      );
    }

  } catch (error) {
    Logger.log('ERROR in handleDepartmentAssignment: ' + error.message);
  }
}

/**
 * Looks up department email from Map Sheet
 * Uses cache if available, refreshes if needed
 * @param {string} departmentName - Name of the department
 * @returns {string} - Email address or empty string
 */
function getDepartmentEmail(departmentName) {
  try {
    // Check cache first
    if (MAP_CACHE.departments && MAP_CACHE.departments[departmentName]) {
      Logger.log('  ✓ Found department email in cache: ' + MAP_CACHE.departments[departmentName]);
      return MAP_CACHE.departments[departmentName];
    }

    // Cache miss - refresh from Map Sheet
    Logger.log('  Cache miss for department: ' + departmentName);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mapSheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);

    if (!mapSheet) return '';

    const data = mapSheet.getDataRange().getValues();

    // Search for department
    for (let i = 1; i < data.length; i++) {
      const type = data[i][0];
      const key = data[i][1];
      const value = data[i][2];

      if (type === 'Department' && key === departmentName) {
        // Update cache
        if (!MAP_CACHE.departments) {
          MAP_CACHE.departments = {};
        }
        MAP_CACHE.departments[departmentName] = value;

        return value;
      }
    }

    return '';

  } catch (error) {
    Logger.log('ERROR getting department email: ' + error.message);
    return '';
  }
}


/**
 * Handles engagement status changes
 * - Sends notifications when status = "Engaged" to BILLING recipients only
 * - Calculates Date Due for Renewal
 * @param {Sheet} sheet - The Engagement Information Sheet
 * @param {number} row - The row number that was edited
 */
function handleEngagementStatusChange(sheet, row) {
  try {
    Logger.log('Processing engagement status change for row ' + row);

    const engagementStatus = sheet.getRange(row, 21).getValue(); // Column U

    // Only proceed if status is "Engaged"
    if (engagementStatus !== 'Engaged') {
      Logger.log('  Status is not "Engaged", skipping notification');
      return;
    }

    // Calculate Date Due for Renewal
    calculateRenewalDate(sheet, row);

    // Check if notification already sent
    const notificationSentDate = sheet.getRange(row, 31).getValue(); // Column AE

    if (notificationSentDate) {
      Logger.log('  Engagement notification already sent on: ' + notificationSentDate);
      return;
    }

    // Get row data
    const rowData = sheet.getRange(row, 1, 1, 32).getValues()[0];

    const clientData = extractClientData(rowData);

    // Send "Engagement" notification (BILLING ONLY)
    sendEngagementNotification(sheet, row, clientData);

  } catch (error) {
    Logger.log('ERROR in handleEngagementStatusChange: ' + error.message);
  }
}


/**
 * Sends "Engagement" notification to BILLING recipients only
 * CORRECTED VERSION with proper column references
 */
function sendEngagementNotification(sheet, row, clientData) {
  assertMonitoredMailboxAccount_('engagement notification');
  try {
    Logger.log('========================================');
    Logger.log('📧 SENDING ENGAGEMENT NOTIFICATION');
    Logger.log('Row: ' + row);
    Logger.log('Client: ' + clientData.clientName);
    Logger.log('========================================');

    // IMPORTANT: Find the correct column for "Engagement Notification Sent"
    // We need to get the actual column number, not assume it
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let notificationColumn = -1;

    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i]).toLowerCase();
      if (header.includes('engagement notification sent')) {
        notificationColumn = i + 1;
        break;
      }
    }

    if (notificationColumn === -1) {
      Logger.log('⚠️ WARNING: "Engagement Notification Sent" column not found!');
      Logger.log('Looking for column containing "engagement notification sent"');
      Logger.log('Available columns: ' + headers.join(', '));

      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Column "Engagement Notification Sent" not found in sheet.\n\n' +
        'Please verify column headers.',
        '⚠️ Column Missing',
        10
      );
      return;
    }

    Logger.log('Found "Engagement Notification Sent" at column ' + notificationColumn + ' (' + getColumnLetter(notificationColumn) + ')');

    // Check if notification already sent
    const notificationSentDate = sheet.getRange(row, notificationColumn).getValue();

    if (notificationSentDate) {
      Logger.log('  Engagement notification already sent on: ' + notificationSentDate);
      Logger.log('  Skipping duplicate notification');

      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Engagement notification was already sent on ' +
        Utilities.formatDate(new Date(notificationSentDate), CONFIG.TIMEZONE, 'MMM dd, yyyy HH:mm'),
        'ℹ️ Already Notified',
        5
      );
      return;
    }

    // Get ONLY billing notification emails
    const recipients = getNotificationEmails('Email Notification for Engagement Only');

    Logger.log('Recipients found: ' + recipients.length);

    if (recipients.length === 0) {
      Logger.log('❌ ERROR: No billing notification emails found in Map Sheet');
      Logger.log('Please add entries with Type: "Email Notification for Engagement Only"');

      SpreadsheetApp.getActiveSpreadsheet().toast(
        'No "Email Notification for Engagement Only" emails found in Map Sheet.\n\n' +
        'Please add billing notification emails to Map Sheet.',
        '⚠️ No Recipients Found',
        10
      );
      return;
    }

    // List recipients
    for (let i = 0; i < recipients.length; i++) {
      Logger.log('  Recipient ' + (i + 1) + ': ' + recipients[i].email + ' (' + recipients[i].name + ')');
    }

    // Build email
    const emailSubject = '[ENGAGED] New Client Engagement: ' + clientData.clientName;
    const emailBody = buildEngagementEmailBody(clientData, row);

    Logger.log('Email Subject: ' + emailSubject);
    Logger.log('Email Body Length: ' + emailBody.length + ' characters');

    // Send to billing recipients
    let sentCount = 0;
    const errors = [];

    for (let i = 0; i < recipients.length; i++) {
      try {
        Logger.log('Attempting to send to: ' + recipients[i].email);

        GmailApp.sendEmail(
          recipients[i].email,
          emailSubject,
          emailBody,
          {
            htmlBody: emailBody.replace(/\n/g, '<br>'),
            name: 'Duran Schulze CRM System'
          }
        );

        sentCount++;
        Logger.log('  ✅ SUCCESS: Sent to ' + recipients[i].email);

      } catch (emailError) {
        const errorMsg = recipients[i].email + ': ' + emailError.message;
        errors.push(errorMsg);
        Logger.log('  ❌ FAILED: ' + errorMsg);
      }
    }

    Logger.log('========================================');
    Logger.log('Send Summary: ' + sentCount + ' / ' + recipients.length + ' succeeded');
    Logger.log('========================================');

    if (sentCount > 0) {
      // Mark notification as sent - CORRECTED LINE
      const now = new Date();
      sheet.getRange(row, notificationColumn).setValue(now);

      Logger.log('✅ Marked notification as sent in Column ' + notificationColumn + ' (' + getColumnLetter(notificationColumn) + ')');
      Logger.log('   Timestamp: ' + Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'));

      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Engagement notification sent to ' + sentCount + ' billing recipient(s)\n' +
        'Notification logged in column ' + getColumnLetter(notificationColumn),
        '✅ Notification Sent',
        5
      );

    } else {
      Logger.log('❌ All email sends failed');
      Logger.log('Errors: ' + errors.join('; '));

      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Failed to send notifications:\n' + errors.join('\n'),
        '❌ Send Failed',
        10
      );
    }

    if (errors.length > 0 && sentCount > 0) {
      Logger.log('⚠️ Partial success - some emails failed: ' + errors.join('; '));
    }

  } catch (error) {
    Logger.log('========================================');
    Logger.log('❌ CRITICAL ERROR in sendEngagementNotification');
    Logger.log('Error: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========================================');

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Critical error sending engagement notification: ' + error.message,
      '❌ Error',
      10
    );
  }
}



/**
 * Gets mandatory notification email addresses from Map Sheet
 * @returns {Array} - Array of {email: string, name: string} objects
 */
function getMandatoryNotificationEmails() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mapSheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);

    if (!mapSheet) return [];

    const data = mapSheet.getDataRange().getValues();
    const mandatoryEmails = [];

    // Search for NotifyOnEngaged entries
    for (let i = 1; i < data.length; i++) { // Skip header
      const type = data[i][0];
      const key = data[i][1];
      const value = data[i][2];

      if (type === 'NotifyOnEngaged' && value && value.trim() !== '') {
        mandatoryEmails.push({
          email: value.trim(),
          name: key || 'Notification Recipient'
        });
      }
    }

    Logger.log('Found ' + mandatoryEmails.length + ' mandatory notification emails');

    return mandatoryEmails;

  } catch (error) {
    Logger.log('Error getting mandatory emails: ' + error.message);
    return [];
  }
}

/**
 * Builds email body for Engagement notification (billing focus)
 */
function buildEngagementEmailBody(clientData, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetUrl = ss.getUrl() + '#gid=' + ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME).getSheetId() + '&range=A' + row;

  const formatCurrency = function(value) {
    if (!value || value === 0) return '0.00';
    return parseFloat(value).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  };

  const formatDate = function(date) {
    if (!date) return 'N/A';
    if (date instanceof Date) {
      return Utilities.formatDate(date, CONFIG.TIMEZONE, 'MMMM dd, yyyy');
    }
    return date;
  };

  return '═══════════════════════════════════════════════════════════\n' +
         '       NEW CLIENT ENGAGEMENT - BILLING NOTIFICATION\n' +
         '═══════════════════════════════════════════════════════════\n\n' +
         'A new client has been engaged and requires billing setup.\n\n' +
         '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
         'CLIENT DETAILS\n' +
         '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
         'Client Name:        ' + (clientData.clientName || 'N/A') + '\n' +
         'Contact Person:     ' + (clientData.contactPerson || 'N/A') + '\n' +
         'Email:              ' + (clientData.email || 'N/A') + '\n' +
         'Phone:              ' + (clientData.phone || 'N/A') + '\n' +
         'Service:            ' + (clientData.service || 'N/A') + '\n' +
         'Service Ref#:       ' + (clientData.serviceRef || 'Not yet generated') + '\n\n' +
         '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
         'FINANCIAL SUMMARY\n' +
         '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
         'Engagement Fee:     PHP ' + formatCurrency(clientData.engagementFee) + '\n' +
         'VAT (12%):          PHP ' + formatCurrency(clientData.vat) + '\n' +
         'Total Gross:        PHP ' + formatCurrency(clientData.totalGross) + '\n' +
         'Miscellaneous:      PHP ' + formatCurrency(clientData.miscellaneous) + '\n' +
         '                    ─────────────────────────\n' +
         'TOTAL SERVICE FEE:  PHP ' + formatCurrency(clientData.totalServiceFee) + '\n\n' +
         '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
         'ACTION REQUIRED\n' +
         '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
         '1. Create billing account\n' +
         '2. Set up payment tracking\n' +
         '3. Issue invoice when ready\n\n' +
         'View engagement details: ' + sheetUrl + '\n\n' +
         '═══════════════════════════════════════════════════════════\n' +
         'Duran Schulze CRM - Automated Notification\n' +
         '═══════════════════════════════════════════════════════════';
}

/**
 * Builds email body for Paid & Assigned notification (service delivery focus)
 */
function buildPaidAssignmentEmailBody(clientData, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetUrl = ss.getUrl() + '#gid=' + ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME).getSheetId() + '&range=A' + row;

  const formatCurrency = function(value) {
    if (!value || value === 0) return '0.00';
    return parseFloat(value).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  };

  const formatDate = function(date) {
    if (!date) return 'N/A';
    if (date instanceof Date) {
      return Utilities.formatDate(date, CONFIG.TIMEZONE, 'MMMM dd, yyyy');
    }
    return date;
  };

  return '═══════════════════════════════════════════════════════════\n' +
         '      CLIENT PAID & READY FOR SERVICE DELIVERY\n' +
         '═══════════════════════════════════════════════════════════\n\n' +
         'Hello ' + (clientData.assignedDepartment || 'Team') + ',\n\n' +
         '💰 PAYMENT RECEIVED - Client is now ready for service delivery.\n\n' +
         '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
         'CLIENT DETAILS\n' +
         '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
         'Client Name:        ' + (clientData.clientName || 'N/A') + '\n' +
         'Contact Person:     ' + (clientData.contactPerson || 'N/A') + '\n' +
         'Email:              ' + (clientData.email || 'N/A') + '\n' +
         'Phone:              ' + (clientData.phone || 'N/A') + '\n' +
         'Address:            ' + (clientData.address || 'N/A') + '\n\n' +
         '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
         'SERVICE INFORMATION\n' +
         '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
         'Service Type:       ' + (clientData.service || 'N/A') + '\n' +
         'Service Ref#:       ' + (clientData.serviceRef || 'Not yet generated') + '\n' +
         'Assigned Date:      ' + formatDate(clientData.assignedDate) + '\n\n' +
         '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
         'PAYMENT STATUS\n' +
         '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
         '✓ PAYMENT RECEIVED: PHP ' + formatCurrency(clientData.totalServiceFee) + '\n' +
         'Payment Date:       ' + formatDate(clientData.paymentStatusDate) + '\n\n' +
         '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
         'DOCUMENTS\n' +
         '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
         (clientData.generatedQuotePdf ? 'Quote PDF: ' + clientData.generatedQuotePdf + '\n' : '') +
         (clientData.folderLink ? 'Client Folder: ' + clientData.folderLink + '\n' : '') +
         'Engagement Details: ' + sheetUrl + '\n\n' +
         '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
         'IMMEDIATE ACTION REQUIRED\n' +
         '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
         '1. Contact client within 24 hours\n' +
         '2. Begin service delivery process\n' +
         '3. Schedule necessary appointments\n' +
         '4. Update client on timeline\n\n' +
         (clientData.remarks ? '📝 NOTES: ' + clientData.remarks + '\n\n' : '') +
         '═══════════════════════════════════════════════════════════\n' +
         'Duran Schulze CRM - Service Delivery Alert\n' +
         '═══════════════════════════════════════════════════════════';
}


// ========================================
// CONVERSION TRACKING TO ENGAGEMENT INFO
// ========================================

function pushProspectsLeadsToEngagement(ss) {
  try {
    Logger.log('Pushing Prospects/Leads to Engagement Information Sheet...');

    const conversionSheet = ss.getSheetByName(CONFIG.CONVERSION_SHEET_NAME);
    const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);

    if (!conversionSheet || !infoSheet) {
      Logger.log('Required sheets not found');
      return;
    }

    if (!ensureEngagementSheetSchema_(infoSheet)) {
      throw new Error('Engagement Information Sheet schema is invalid. Existing data was preserved.');
    }
    const engagementHeaders = infoSheet
      .getRange(1, 1, 1, infoSheet.getLastColumn())
      .getValues()[0];
    const sourceMonthColumn = getHeaderColumn_(
      engagementHeaders,
      'Source Month',
      ENGAGEMENT_COLUMNS.SOURCE_MONTH
    );
    const sourceMonthIndex = sourceMonthColumn - 1;

    const conversionData = conversionSheet.getDataRange().getValues();
    conversionData.shift();

    const infoData = infoSheet.getDataRange().getValues();
    const existingEntries = new Map();

    for (let i = 1; i < infoData.length; i++) {
      const email = String(infoData[i][4]).toLowerCase().trim();
      const sourceMonth = infoData[i][sourceMonthIndex];
      const key = email + '|' + sourceMonth;
      existingEntries.set(key, i + 1);
    }

    let newRecordsCount = 0;
    let updatedRecordsCount = 0;
    let skippedRecordsCount = 0;
    const exclusionReasons = {};

    for (const row of conversionData) {
      const senderEmail = String(row[0]).toLowerCase().trim();
      const senderName = row[1] || senderEmail;
      const originalSubject = row[2];  // Column C: Original Subject
      const firstContactDate = safeGetDate(row[3]);
      const dateLastContacted = safeGetDate(row[4]);
      const totalEmails = row[5];
      const status = row[6];
      const latestSummary = row[7];

      if (status !== 'Prospect' && status !== 'Lead') {
        if (status === 'N/A') {
          Logger.log('Skipping N/A status (internal domain): ' + senderEmail);
        }
        continue;
      }

      if (firstContactDate < CONFIG.ENGAGEMENT_START_DATE) {
        continue;
      }

      const exclusionCheck = isEmailExcluded(senderEmail);

      if (exclusionCheck.isInternalDomain) {  // FIXED: was isCompanyDomain
        Logger.log('Skipping internal domain email: ' + senderEmail + ' (' + exclusionCheck.reason + ')');
        skippedRecordsCount++;

        if (!exclusionReasons[exclusionCheck.reason]) {
          exclusionReasons[exclusionCheck.reason] = 0;
        }
        exclusionReasons[exclusionCheck.reason]++;

        continue;
      }

      if (exclusionCheck.isExcluded) {
        Logger.log('Excluding email from Engagement: ' + senderEmail + ' (' + exclusionCheck.reason + ')');
        skippedRecordsCount++;

        if (!exclusionReasons[exclusionCheck.reason]) {
          exclusionReasons[exclusionCheck.reason] = 0;
        }
        exclusionReasons[exclusionCheck.reason]++;

        continue;
      }

      const clientName = senderName;
      const sourceMonth = Utilities.formatDate(firstContactDate, CONFIG.TIMEZONE, 'MMM-yyyy');
      const entryKey = senderEmail + '|' + sourceMonth;

      if (existingEntries.has(entryKey)) {
        // UPDATE existing entry
        const existingRowIndex = existingEntries.get(entryKey);

        infoSheet.getRange(existingRowIndex, 1).setValue(firstContactDate);
        infoSheet.getRange(existingRowIndex, 2).setValue(clientName);
        infoSheet.getRange(existingRowIndex, 5).setValue(senderEmail);

        // OPTION B: Only populate Remarks if currently empty
        const currentRemarks = infoSheet.getRange(existingRowIndex, 8).getValue();
        if (!currentRemarks || currentRemarks.toString().trim() === '') {
          infoSheet.getRange(existingRowIndex, 8).setValue(originalSubject);
          Logger.log('  ✓ Updated Remarks with Original Subject: ' + originalSubject);
        } else {
          Logger.log('  ℹ Preserving existing Remarks (not overwriting)');
        }

        infoSheet.getRange(existingRowIndex, sourceMonthColumn).setValue(sourceMonth);

        // ENSURE FINANCIAL FORMULAS ARE APPLIED
        const vatFormula = `=IF(OR(I${existingRowIndex}="",I${existingRowIndex}=0),"",I${existingRowIndex}*${CONFIG.VAT_RATE})`;
        infoSheet.getRange(existingRowIndex, 10).setFormula(vatFormula);

        const grossFormula = `=IF(OR(I${existingRowIndex}="",I${existingRowIndex}=0),"",I${existingRowIndex}+J${existingRowIndex})`;
        infoSheet.getRange(existingRowIndex, 11).setFormula(grossFormula);

        const totalFormula = `=IF(OR(K${existingRowIndex}="",K${existingRowIndex}=0),"",K${existingRowIndex}+IF(OR(L${existingRowIndex}="",ISBLANK(L${existingRowIndex})),0,L${existingRowIndex}))`;
        infoSheet.getRange(existingRowIndex, 13).setFormula(totalFormula);

        Logger.log('Updated existing entry: ' + senderEmail + ' (' + clientName + ') for ' + sourceMonth + ' (Row ' + existingRowIndex + ')');
        updatedRecordsCount++;

      } else {
        // ADD NEW ROW
        const newRow = new Array(ENGAGEMENT_INFO_HEADERS.length).fill('');
        newRow[0] = firstContactDate;           // Contact Date
        newRow[1] = clientName;                 // Client Name
        newRow[2] = '';                         // Project / Service
        newRow[3] = '';                         // Contact Person
        newRow[4] = senderEmail;                // Email Address
        newRow[5] = '';                         // Telephone Number
        newRow[6] = '';                         // Registered Address
        newRow[7] = originalSubject;            // Remarks (populated with Original Subject)
        newRow[8] = '';                         // Engagement Fee
        newRow[9] = '';                         // VAT
        newRow[10] = '';                        // Total Gross
        newRow[11] = '';                        // Miscellaneous
        newRow[12] = '';                        // Total Service Fee
        newRow[13] = '';                        // Service Ref#
        newRow[14] = '';                        // Service Quote Created?
        newRow[15] = '';                        // Service Quote Date Sent
        newRow[16] = '';                        // Service Quote Status
        newRow[17] = '';                        // Service Quote Follow-Up
        newRow[18] = '';                        // Payment Status
        newRow[19] = '';                        // Payment Status Date
        newRow[20] = '';                        // Engagement Status
        newRow[21] = '';                        // Needs Renewal
        newRow[22] = '';                        // Assigned Department
        newRow[23] = '';                        // Assigned Department Email
        newRow[24] = '';                        // Assigned Date
        newRow[25] = '';                        // PDF - Signed Agreement
        newRow[26] = '';                        // File Folder
        newRow[27] = '';                        // Generated Quote PDF
        newRow[28] = '';                        // Date Due for Renewal
        newRow[sourceMonthIndex] = sourceMonth; // Source Month
        // AI enhancements: scoring, categorization, assignment
        newRow[32] = getGeminiScore(originalSubject || ''); // AI Score (column AG)
        newRow[33] = categorizeIntent(originalSubject || ''); // Intent Category (column AH)
        newRow[34] = getNextTeamMember() || ''; // Assigned Team Member (column AI)
// These columns must be added to ENGAGEMENT_INFO_HEADERS array (see below)
        infoSheet.appendRow(newRow);

        const newRowIndex = infoSheet.getLastRow();

        const vatFormula = `=IF(OR(I${newRowIndex}="",I${newRowIndex}=0),"",I${newRowIndex}*${CONFIG.VAT_RATE})`;
        infoSheet.getRange(newRowIndex, 10).setFormula(vatFormula);

        const grossFormula = `=IF(OR(I${newRowIndex}="",I${newRowIndex}=0),"",I${newRowIndex}+J${newRowIndex})`;
        infoSheet.getRange(newRowIndex, 11).setFormula(grossFormula);

        const totalFormula = `=IF(OR(K${newRowIndex}="",K${newRowIndex}=0),"",K${newRowIndex}+IF(OR(L${newRowIndex}="",ISBLANK(L${newRowIndex})),0,L${newRowIndex}))`;
        infoSheet.getRange(newRowIndex, 13).setFormula(totalFormula);

        Logger.log('Added new entry: ' + senderEmail + ' (' + clientName + ') for ' + sourceMonth + ' (Row ' + newRowIndex + ')');
        Logger.log('  ✓ Remarks populated with: ' + originalSubject);
        newRecordsCount++;
      }
    }

    Logger.log('========================================');
    Logger.log('Engagement Information Sheet sync complete:');
    Logger.log('  • New records added: ' + newRecordsCount);
    Logger.log('  • Existing records updated: ' + updatedRecordsCount);
    Logger.log('  • Records skipped (exclusions + N/A): ' + skippedRecordsCount);
    Logger.log('========================================');
    Logger.log('Exclusion breakdown:');
    for (const reason in exclusionReasons) {
      Logger.log('  • ' + reason + ': ' + exclusionReasons[reason]);
    }
    Logger.log('========================================');
    Logger.log('Internal domains (N/A status): ' + CONFIG.INTERNAL_DOMAINS.join(', '));
    Logger.log('Excluded domains: ' + CONFIG.EXCLUDED_DOMAINS.join(', '));
    Logger.log('Excluded patterns: ' + CONFIG.EXCLUDED_PATTERNS.join(', '));
    Logger.log('========================================');

    setupDropdownValidations(infoSheet);

  } catch (error) {
    Logger.log('Error pushing prospects/leads: ' + error.message);
  }
}



// ========================================
// DROPBOX SIGN & PAYMENT EMAIL PROCESSING
// ========================================

function processDropboxSignEmails(ss) {
  assertMonitoredMailboxAccount_('Dropbox Sign mailbox processing');
  try {
    const threads = GmailApp.search('from:noreply@mail.hellosign.com newer_than:7d', 0, 50);
    const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);

    if (!infoSheet || infoSheet.getLastRow() < 2) return;

    const infoData = infoSheet.getRange(2, 1, infoSheet.getLastRow() - 1, ENGAGEMENT_INFO_HEADERS.length).getValues();

    for (const thread of threads) {
      for (const msg of thread.getMessages()) {
        const body = msg.getPlainBody() || '';

        // Extract signer's email from body using regex
        const emailMatch = body.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/gi);

        if (emailMatch && emailMatch.length > 0) {
          const signerEmail = emailMatch[0].toLowerCase().trim();

          // Find matching row in Engagement Information
          for (let i = 0; i < infoData.length; i++) {
            const rowEmail = String(infoData[i][4]).toLowerCase().trim();

            if (rowEmail === signerEmail) {
              const rowIndex = i + 2;

              // Update Service Quote Status to "Engaged"
              infoSheet.getRange(rowIndex, 17).setValue('Engaged');

              // Send notification
              sendNotification('Dropbox Sign', signerEmail, infoData[i]);

              Logger.log('Updated Dropbox Sign status for: ' + signerEmail);
              break;
            }
          }
        }
      }
    }

  } catch (error) {
    Logger.log('Error processing Dropbox Sign emails: ' + error.message);
  }
}

function processPaymentEmails(ss) {
  assertMonitoredMailboxAccount_('payment mailbox processing');
  try {
    const searchQuery = CONFIG.PAYMENT_KEYWORDS.map(kw => `"${kw}"`).join(' OR ');
    const threads = GmailApp.search(`(${searchQuery}) newer_than:7d`, 0, 50);
    const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);

    if (!infoSheet || infoSheet.getLastRow() < 2) return;

    const infoData = infoSheet.getRange(2, 1, infoSheet.getLastRow() - 1, ENGAGEMENT_INFO_HEADERS.length).getValues();

    for (const thread of threads) {
      for (const msg of thread.getMessages()) {
        const senderEmail = extractEmailAddress(msg.getFrom());
        const body = msg.getPlainBody() || '';
        const subject = msg.getSubject() || '';

        // Check if payment keywords present
        if (!detectKeywords(body + ' ' + subject, CONFIG.PAYMENT_KEYWORDS)) continue;

        // Find matching row
        for (let i = 0; i < infoData.length; i++) {
          const rowEmail = String(infoData[i][4]).toLowerCase().trim();

          if (rowEmail === senderEmail.toLowerCase().trim()) {
            const rowIndex = i + 2;

            // Update Payment Status = "Paid"
            infoSheet.getRange(rowIndex, 19).setValue('Paid');

            // Update Payment Status Date
            infoSheet.getRange(rowIndex, 20).setValue(new Date());

            // Update Engagement Status = "Engaged"
            infoSheet.getRange(rowIndex, 21).setValue('Engaged');

            // Send notification
            sendNotification('Payment', senderEmail, infoData[i]);

            Logger.log('Updated payment status for: ' + senderEmail);
            break;
          }
        }
      }
    }

  } catch (error) {
    Logger.log('Error processing payment emails: ' + error.message);
  }
}

function sendNotification(type, email, rowData) {
  assertMonitoredMailboxAccount_('CRM email notification');
  try {
    const clientName = rowData[1] || email;
    const service = rowData[2] || 'N/A';
    const assignedDept = rowData[22] || 'Billing';
    const deptEmail = rowData[23] || 'billing@duranschulze.com';

    let subject = '';
    let body = '';

    if (type === 'Dropbox Sign') {
      subject = `✅ Document Signed: ${clientName}`;
      body = `Hello,\n\nA document has been signed by ${clientName} (${email}).\n\n` +
             `Service: ${service}\n` +
             `Status: Engaged\n\n` +
             `Please proceed with next steps.\n\n` +
             `Best regards,\nCRM System`;
    } else if (type === 'Payment') {
      subject = `💰 Payment Received: ${clientName}`;
      body = `Hello,\n\nPayment has been received from ${clientName} (${email}).\n\n` +
             `Service: ${service}\n` +
             `Status: Paid & Engaged\n\n` +
             `Please update your records accordingly.\n\n` +
             `Best regards,\nCRM System`;
    }

    if (deptEmail && deptEmail.includes('@')) {
      GmailApp.sendEmail(deptEmail, subject, body);
      Logger.log('Notification sent to: ' + deptEmail);
    }

  } catch (error) {
    Logger.log('Error sending notification: ' + error.message);
  }
}

// ========================================
// SERVICE REF# GENERATION
// ========================================

function generateServiceRefNumber(infoSheet) {
  try {
    const today = new Date();
    const dateStr = Utilities.formatDate(today, CONFIG.TIMEZONE, 'yyyyMMdd');
    const year = today.getFullYear();

    // Get all ref numbers for current year
    const lastRow = infoSheet.getLastRow();
    if (lastRow < 2) return `SRF-${dateStr}-001`;

    const refData = infoSheet.getRange(2, 14, lastRow - 1, 1).getValues();
    let maxCounter = 0;

    for (const row of refData) {
      const ref = String(row[0]);
      if (ref.startsWith(`SRF-${year}`)) {
        const parts = ref.split('-');
        if (parts.length === 3) {
          const counter = parseInt(parts[2]);
          if (!isNaN(counter) && counter > maxCounter) {
            maxCounter = counter;
          }
        }
      }
    }

    const newCounter = String(maxCounter + 1).padStart(3, '0');
    return `SRF-${dateStr}-${newCounter}`;

  } catch (error) {
    Logger.log('Error generating service ref: ' + error.message);
    return `SRF-${Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd')}-001`;
  }
}

// ========================================
// DOCUMENT GENERATION & DRIVE MANAGEMENT
// ========================================

function generateQuoteForSelectedRow() {
  try {
    Logger.log('========================================');
    Logger.log('Starting quote generation...');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);
    const activeRange = infoSheet.getActiveRange();

    if (!activeRange) {
      throw new Error('Please select a row in the Engagement Information Sheet');
    }

    const rowIndex = activeRange.getRow();
    if (rowIndex < 2) {
      throw new Error('Please select a data row (not the header)');
    }

    Logger.log('Selected row: ' + rowIndex);

    const rowData = infoSheet.getRange(rowIndex, 1, 1, ENGAGEMENT_INFO_HEADERS.length).getValues()[0];

    // Extract data from row
    const contactDate = rowData[0];
    const clientName = rowData[1] || 'Client Name';
    const service = rowData[2];
    const contactPerson = rowData[3] || '';
    const emailAddress = rowData[4] || '';
    const telephoneNumber = rowData[5] || '';
    const registeredAddress = rowData[6] || '';
    const remarks = rowData[7] || '';
    const engagementFee = rowData[8] || 0;
    const vat = rowData[9] || 0;
    const totalGross = rowData[10] || 0;
    const miscellaneous = rowData[11] || 0;
    const totalServiceFee = rowData[12] || 0;

    Logger.log('Client: ' + clientName);
    Logger.log('Service: ' + service);
    Logger.log('Email: ' + emailAddress);
    Logger.log('Engagement Fee: ' + engagementFee);

    // Validate required fields
    if (!service) {
      throw new Error('Please select a Project / Service first (Column C)');
    }

    if (!engagementFee || engagementFee === 0) {
      throw new Error('Please enter an Engagement Fee first (Column I)');
    }

    // Get template ID for service
    const templateId = getTemplateIdForService(service);
    if (!templateId) {
      throw new Error('No template found for service: ' + service + '\n\nPlease add the template ID in Map Sheet.');
    }

    Logger.log('Template ID: ' + templateId);

    // Generate or get Service Ref#
    let serviceRef = rowData[13];
    if (!serviceRef) {
      serviceRef = generateServiceRefNumber(infoSheet);
      infoSheet.getRange(rowIndex, 14).setValue(serviceRef);
      Logger.log('Generated Service Ref#: ' + serviceRef);
    } else {
      Logger.log('Using existing Service Ref#: ' + serviceRef);
    }

    // Create month folder
    const monthFolder = getOrCreateMonthFolder();
    Logger.log('Month folder: ' + monthFolder.getName());

    // Copy template document
    const templateDoc = DriveApp.getFileById(templateId);
    const docName = `${service} - ${clientName} - ${serviceRef}`;
    const newDoc = templateDoc.makeCopy(docName, monthFolder);
    const docId = newDoc.getId();

    Logger.log('Created document copy: ' + docName);
    Logger.log('Document ID: ' + docId);

    // Open document and replace placeholders
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // Current date for document
    const currentDate = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'MMMM dd, yyyy');

    // Format currency values
    const formatCurrency = function(value) {
      if (!value || value === 0) return '0.00';
      return parseFloat(value).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    };

    Logger.log('Replacing placeholders in document...');

    // Replace all placeholders - CASE INSENSITIVE
    const replacements = {
      '{{DATE}}': currentDate,
      '{{REFERENCE_NUMBER}}': serviceRef,
      '{{SERVICE_REF}}': serviceRef,
      '{{SERVICE_REFERENCE}}': serviceRef,
      '{{CLIENT_NAME}}': clientName,
      '{{CLIENT}}': clientName,
      '{{CONTACT_PERSON}}': contactPerson || clientName,
      '{{EMAIL}}': emailAddress,
      '{{EMAIL_ADDRESS}}': emailAddress,
      '{{PHONE}}': telephoneNumber,
      '{{TELEPHONE}}': telephoneNumber,
      '{{ADDRESS}}': registeredAddress,
      '{{REGISTERED_ADDRESS}}': registeredAddress,
      '{{SERVICE}}': service,
      '{{PROJECT}}': service,
      '{{ENGAGEMENT_FEE}}': formatCurrency(engagementFee),
      '{{PROFESSIONAL_FEE}}': formatCurrency(engagementFee),
      '{{VAT}}': formatCurrency(vat),
      '{{TOTAL_GROSS}}': formatCurrency(totalGross),
      '{{GOVERNMENT_FEES}}': formatCurrency(miscellaneous),
      '{{MISCELLANEOUS}}': formatCurrency(miscellaneous),
      '{{MISC}}': formatCurrency(miscellaneous),
      '{{TOTAL}}': formatCurrency(totalServiceFee),
      '{{TOTAL_SERVICE_FEE}}': formatCurrency(totalServiceFee),
      '{{TOTAL_FEE}}': formatCurrency(totalServiceFee)
    };

    // Perform replacements
    for (const placeholder in replacements) {
      const value = replacements[placeholder];
      body.replaceText(placeholder, value);
      // Also try case variations
      body.replaceText(placeholder.toLowerCase(), value);
      body.replaceText(placeholder.toUpperCase(), value);
    }

    Logger.log('✓ Placeholders replaced');

    // Save and close document
    doc.saveAndClose();
    Logger.log('✓ Document saved');

    // Wait a moment for Google to process
    Utilities.sleep(2000);

    // Export to PDF
    Logger.log('Exporting to PDF...');
    const pdfBlob = newDoc.getAs('application/pdf');
    const pdfName = `${service} - ${clientName} - ${serviceRef}.pdf`;
    const pdfFile = monthFolder.createFile(pdfBlob);
    pdfFile.setName(pdfName);

    Logger.log('✓ PDF created: ' + pdfName);

    // Update sheet with links and metadata
    infoSheet.getRange(rowIndex, 15).setValue('Generate Quote'); // Service Quote Created?
    infoSheet.getRange(rowIndex, 16).setValue(new Date()); // Service Quote Date Sent
    infoSheet.getRange(rowIndex, 27).setValue(monthFolder.getUrl()); // Folder link
    infoSheet.getRange(rowIndex, 28).setValue(pdfFile.getUrl()); // Column AB (Generated Quote PDF)

 // Format Column AA (File Folder) as hyperlink
if (monthFolder) {
  const folderUrl = monthFolder.getUrl();
  const formula = '=HYPERLINK("' + folderUrl + '", "View Proposed Agreement and Invoice")';
  infoSheet.getRange(rowIndex, 27).setFormula(formula); // Column AA
}

    Logger.log('✓ Sheet updated with links');
    Logger.log('========================================');
    Logger.log('Quote generation completed successfully!');
    Logger.log('========================================');

    return {
      success: true,
      pdfUrl: pdfFile.getUrl(),
      folderUrl: monthFolder.getUrl(),
      docUrl: newDoc.getUrl(),
      serviceRef: serviceRef
    };

  } catch (error) {
    Logger.log('========================================');
    Logger.log('✗ Error generating quote: ' + error.message);
    Logger.log('Stack trace: ' + error.stack);
    Logger.log('========================================');
    throw error;
  }
}


function getTemplateIdForService(serviceName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mapSheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);

    if (!mapSheet) return null;

    const mapData = mapSheet.getDataRange().getValues();

    for (const row of mapData) {
      if (row[0] === 'Service' && row[1] === serviceName) {
        return row[2];
      }
    }

    return null;

  } catch (error) {
    Logger.log('Error getting template ID: ' + error.message);
    return null;
  }
}

function getOrCreateMonthFolder() {
  try {
    const rootFolder = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
    const monthName = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'MMM yyyy');

    // Check if month folder exists
    const folders = rootFolder.getFoldersByName(monthName);

    if (folders.hasNext()) {
      return folders.next();
    } else {
      return rootFolder.createFolder(monthName);
    }

  } catch (error) {
    Logger.log('Error creating month folder: ' + error.message);
    throw error;
  }
}

function getDepartmentEmail(deptName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mapSheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);

    if (!mapSheet) return '';

    const mapData = mapSheet.getDataRange().getValues();

    for (const row of mapData) {
      if (row[0] === 'Department' && row[1] === deptName) {
        return row[2];
      }
    }

    return '';

  } catch (error) {
    Logger.log('Error getting department email: ' + error.message);
    return '';
  }
}

// ========================================
// MENU ACTIONS
// ========================================

function menuSyncEmails() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.alert('📧 Syncing Emails', 'Starting email synchronization...', ui.ButtonSet.OK);

    const result = syncNewEmails();
    if (result.skipped) {
      ui.alert('Sync Already Running', result.reason, ui.ButtonSet.OK);
      return;
    }

    const monthsList = result.monthsAffected && result.monthsAffected.length > 0
      ? result.monthsAffected.join(', ')
      : 'None';

    const message = `✅ Email sync completed successfully!\n\n` +
                    `📊 Results:\n` +
                    `• New emails processed: ${result.newEmails}\n` +
                    `• Total threads scanned: ${result.totalThreads}\n` +
                    `• Month sheets affected: ${monthsList}\n` +
                    `• Outbound/non-inbox messages skipped: ${result.outboundSkipped}\n` +
                    `• Duplicate messages prevented: ${result.duplicateSkipped}\n` +
                    `• Time: ${new Date().toLocaleString()}`;

    ui.alert('🎉 Sync Complete', message, ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('⚠️ Sync Error', `Error: ${error.message}`, ui.ButtonSet.OK);
    Logger.log('Menu sync error: ' + error.message);
  }
}


function menuRefreshDashboard() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.alert('📊 Refreshing Dashboard', 'Updating dashboard...', ui.ButtonSet.OK);

    buildEnhancedDashboard();

    ui.alert('🎉 Dashboard Updated', 'Dashboard refreshed successfully!', ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('⚠️ Error', `Error: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Menu function: Full system update
 */
function menuFullUpdate() {
  const ui = SpreadsheetApp.getUi();

  try {
    const response = ui.alert(
      '⚡ Full System Update',
      'This will:\n' +
      '1. Refresh Map Sheet data\n' +
      '2. Sync new emails from Gmail\n' +
      '3. Refresh the AI Potential Clients queue\n' +
      '4. Refresh Dashboard\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    ui.alert('⏳ Processing', 'Starting full system update...', ui.ButtonSet.OK);

    // Step 1: Refresh Map Sheet
    Logger.log('Step 1: Refreshing Map Sheet...');
    refreshMapSheet();

    // Step 2: Sync emails
    Logger.log('Step 2: Syncing emails...');
    const syncResult = syncNewEmails();

    // Step 3: Refresh the review queue
    Logger.log('Step 3: Refreshing Potential Clients...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    syncPotentialClientsFromMonthlySheets_(ss);

    // Step 4: Refresh dashboard
    Logger.log('Step 4: Refreshing Dashboard...');
    buildEnhancedDashboard();

    ui.alert(
      '✅ Update Complete',
      'Full system update completed successfully!\n\n' +
      'All data refreshed and synchronized.',
      ui.ButtonSet.OK
    );

  } catch (error) {
    Logger.log('ERROR in full update: ' + error.message);
    ui.alert('⚠️ Error', 'Update error: ' + error.message, ui.ButtonSet.OK);
  }
}


function menuPushToEngagement() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.alert('🔁 Pushing Data', 'Pushing Prospects/Leads to Engagement Information...', ui.ButtonSet.OK);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    pushProspectsLeadsToEngagement(ss);

    ui.alert('🎉 Push Complete', 'Engagement Information Sheet updated!', ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('⚠️ Error', error.message, ui.ButtonSet.OK);
  }
}

function menuGenerateQuote() {
  const ui = SpreadsheetApp.getUi();
  try {
    // Confirm before generating
    const response = ui.alert(
      '📝 Generate Quote',
      'Generate quote document for the selected row?\n\nMake sure you have:\n• Selected a row with data\n• Service selected\n• Engagement Fee entered',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    ui.alert('⏳ Processing', 'Generating quote... This may take a few seconds.', ui.ButtonSet.OK);

    const result = generateQuoteForSelectedRow();

    const message = `✅ Quote generated successfully!\n\n` +
                    `📄 Details:\n` +
                    `• Service Ref#: ${result.serviceRef}\n` +
                    `• PDF: ${result.pdfUrl}\n\n` +
                    `• Folder: ${result.folderUrl}\n\n` +
                    `The quote has been saved and links added to the sheet.`;

    ui.alert('🎉 Quote Generated', message, ui.ButtonSet.OK);

  } catch (error) {
    const errorMsg = `⚠️ Error generating quote:\n\n${error.message}\n\nPlease check:\n` +
                     `• Row is selected\n` +
                     `• Service is selected\n` +
                     `• Engagement Fee is entered\n` +
                     `• Template exists in Map Sheet`;
    ui.alert('⚠️ Generation Error', errorMsg, ui.ButtonSet.OK);
    Logger.log('Menu generate quote error: ' + error.message);
  }
}


function menuAssignDepartment() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);
    const activeRange = infoSheet.getActiveRange();

    if (!activeRange || activeRange.getRow() < 2) {
      throw new Error('Please select a data row');
    }

    const rowIndex = activeRange.getRow();
    const response = ui.prompt('Assign Department', 'Enter department name (Legal, Visa, HR, Accounting):', ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() === ui.Button.OK) {
      const deptName = response.getResponseText();
      const deptEmail = getDepartmentEmail(deptName);

      if (!deptEmail) {
        throw new Error('Department not found in Map Sheet');
      }

      infoSheet.getRange(rowIndex, 23).setValue(deptName);
      infoSheet.getRange(rowIndex, 24).setValue(deptEmail);
      infoSheet.getRange(rowIndex, 25).setValue(new Date());

      ui.alert('✅ Success', `Department assigned: ${deptName}`, ui.ButtonSet.OK);
    }

  } catch (error) {
    ui.alert('⚠️ Error', error.message, ui.ButtonSet.OK);
  }
}

function menuMarkPaymentReceived() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);
    const activeRange = infoSheet.getActiveRange();

    if (!activeRange || activeRange.getRow() < 2) {
      throw new Error('Please select a data row');
    }

    const rowIndex = activeRange.getRow();

    infoSheet.getRange(rowIndex, 19).setValue('Paid');
    infoSheet.getRange(rowIndex, 20).setValue(new Date());
    infoSheet.getRange(rowIndex, 21).setValue('Engaged');

    ui.alert('✅ Success', 'Payment marked as received!', ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('⚠️ Error', error.message, ui.ButtonSet.OK);
  }
}

function menuOpenFileFolder() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);
    const activeRange = infoSheet.getActiveRange();

    if (!activeRange || activeRange.getRow() < 2) {
      throw new Error('Please select a data row');
    }

    const rowIndex = activeRange.getRow();
    const folderUrl = infoSheet.getRange(rowIndex, 27).getValue();

    if (!folderUrl) {
      throw new Error('No folder URL found for this row');
    }

    ui.alert('📁 Folder URL', folderUrl, ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('⚠️ Error', error.message, ui.ButtonSet.OK);
  }
}

function menuRefreshDropdowns() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);
    setupDropdownValidations(infoSheet);
    ui.alert('✅ Success', 'Dropdowns refreshed from Map Sheet!', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('⚠️ Error', error.message, ui.ButtonSet.OK);
  }
}

function menuSetRootFolder() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Set Root Folder ID', `Current: ${CONFIG.ROOT_FOLDER_ID}\n\nEnter new folder ID:`, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK) {
    const newId = response.getResponseText();
    PropertiesService.getScriptProperties().setProperty('ROOT_FOLDER_ID', newId);
    ui.alert('✅ Success', 'Root folder ID updated! Please update CONFIG in script.', ui.ButtonSet.OK);
  }
}

function menuRunMaintenance() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('🧹 Run Maintenance',
                            'This will archive old sheets and clean up data.\n\nDo you want to continue?',
                            ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    try {
      performMaintenance();
      ui.alert('✅ Success', 'Maintenance completed!', ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('⚠️ Error', error.message, ui.ButtonSet.OK);
    }
  }
}

function menuTestConfiguration() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = testEnhancedConfiguration();
    if (result) {
      ui.alert('✅ Test Passed', 'All systems working correctly!', ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('⚠️ Test Failed', error.message, ui.ButtonSet.OK);
  }
}

function menuExportAllData() {
  const ui = SpreadsheetApp.getUi();
  try {
    const fileUrl = exportDataToCSV('All');
    if (fileUrl) {
      ui.alert('🎉 Export Complete', `CSV exported successfully!\n\n${fileUrl}`, ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('⚠️ Error', error.message, ui.ButtonSet.OK);
  }
}

function menuExportCurrentMonth() {
  const ui = SpreadsheetApp.getUi();
  try {
    const currentMonth = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'MMM-yyyy');
    const fileUrl = exportDataToCSV(currentMonth);
    if (fileUrl) {
      ui.alert('🎉 Export Complete', `CSV for ${currentMonth} exported!\n\n${fileUrl}`, ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('⚠️ Error', error.message, ui.ButtonSet.OK);
  }
}

function menuShowAbout() {
  const ui = SpreadsheetApp.getUi();
  const aboutMessage = `📧 Enhanced CRM Gmail Tracker\n\n` +
                       `Created by Atty. MWAD\n` +
                       `Later revised and enhanced by ZACK\n\n` +
                       `🚀 Key Features:\n` +
                       `• Monthly Gmail logging\n` +
                       `• Gemini subject/body lead qualification\n` +
                       `• Potential Clients review and promotion queue\n` +
                       `• Conversion Tracking (aggregated by email)\n` +
                       `• Engagement Information Sheet (Prospect/Lead)\n` +
                       `• Auto-filters: Oct 2025+\n` +
                       `• Internal domains: ${CONFIG.INTERNAL_DOMAINS.length}\n` +
                       `• Excluded domains: ${CONFIG.EXCLUDED_DOMAINS.length}\n` +
                       `• Excluded patterns: ${CONFIG.EXCLUDED_PATTERNS.length}\n` +
                       `• Document generation from service templates\n` +
                       `• Dropbox Sign & payment detection\n` +
                       `• Department assignment with email lookup\n` +
                       `• Drive folder management (MMM YYYY)\n` +
                       `• Service Ref# auto-generation (yearly reset)\n` +
                       `• Financial calculations (VAT 12%, totals)\n` +
                       `• Full dropdown data validations\n\n` +
                       `⚙️ System:\n` +
                       `• Internal version: ${CRM_VERSION}\n` +
                       `• Auto-sync: Every 4 hours\n` +
                       `• Deferred qualification: Every 10 minutes\n` +
                       `• Dashboard: Daily at ${CONFIG.AUTO_REFRESH_HOUR}:00\n\n` +
                       `🏢 Internal Domains (N/A status):\n` +
                       CONFIG.INTERNAL_DOMAINS.map(d => `  • ${d}`).join('\n') + '\n\n' +
                       `🚫 Excluded Domains (Trash status):\n` +
                       CONFIG.EXCLUDED_DOMAINS.map(d => `  • ${d}`).join('\n') + '\n\n' +
                       `🚫 Excluded Patterns:\n` +
                       CONFIG.EXCLUDED_PATTERNS.map(p => `  • ${p}`).join('\n') + '\n\n' +
                       `Data Flow:\n` +
                       `Gmail → Monthly Sheets → Potential Clients → reviewer promotion → Engagement Info\n\n` +
                       `✅ Current refinements:\n` +
                       `• @duranschulze.com AND @filepino.com = N/A status\n` +
                       `• Both internal domains excluded from Engagement Info\n` +
                       `• Light gray formatting for N/A and Trash\n` +
                       `• Sender's Name in Client Name column\n\n` +
                       `Use the menu to manage your CRM system!`;

  ui.alert('ℹ️ About Mini-CRM', aboutMessage, ui.ButtonSet.OK);
}

function menuShowCredits() {
  SpreadsheetApp.getUi().alert(
    '👥 Mini-CRM Credits',
    'Created by Atty. MWAD\n\nLater revised and enhanced by ZACK',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}



// ========================================
// ORIGINAL v3.5 FUNCTIONS (PRESERVED)
// ========================================

function logIncomingEmails() {
  return syncNewEmails();
}

function buildEnhancedDashboard() {
  assertMonitoredMailboxAccount_('Dashboard automation');
  try {
    Logger.log('Building dashboard...');

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let dashboard = ss.getSheetByName(CONFIG.DASHBOARD_SHEET_NAME);
    if (!dashboard) {
      dashboard = ss.insertSheet(CONFIG.DASHBOARD_SHEET_NAME);
    } else {
      dashboard.getDataRange().breakApart();
      dashboard.clear();
    }

    const now = new Date();
    setupDashboardHeader(dashboard, now);

    const recentSheets = getRecentMonthSheets(ss, CONFIG.RECENT_MONTHS_DASHBOARD);
    displayMetrics(dashboard, recentSheets);
    displayConversionSummary(dashboard, ss);
    displayPotentialCandidateSummary_(dashboard, ss);
    displayActionItems(dashboard, ss, recentSheets);

    formatDashboard(dashboard);
    Logger.log('Dashboard built successfully');

  } catch (error) {
    Logger.log('Error building dashboard: ' + error.message);
    throw error;
  }
}

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function safeGetDate(dateInput) {
  try {
    if (!dateInput) return new Date();
    if (dateInput instanceof Date && !isNaN(dateInput.getTime())) {
      return dateInput;
    }
    const parsed = new Date(dateInput);
    return isNaN(parsed.getTime()) ? new Date() : parsed;
  } catch (error) {
    return new Date();
  }
}

function createUniqueEmailId(senderEmail, dateReceived, subject) {
  const timestamp = dateReceived.getTime();
  const normalizedSubject = normalizeSubject(subject);
  return senderEmail + '_' + timestamp + '_' + normalizedSubject;
}

/**
 * Returns every address that can represent the mailbox running this script.
 * Gmail aliases are included so replies sent from an alias are never imported
 * as new inbound opportunities.
 */
function getMonitoredMailboxContext_() {
  const addresses = new Set();
  const primaryEmail = assertMonitoredMailboxAccount_('mailbox access');

  function addAddress(value) {
    const email = extractEmailAddress(String(value || '')).toLowerCase().trim();
    if (!email || email.indexOf('@') === -1) return;
    addresses.add(email);
  }

  addAddress(primaryEmail);
  try {
    const aliases = GmailApp.getAliases() || [];
    aliases.forEach(addAddress);
  } catch (error) {
    Logger.log('Could not read Gmail aliases; continuing with the active mailbox: ' + error.message);
  }

  return { primaryEmail: primaryEmail, addresses: addresses };
}

/**
 * Message-level inbound gate. Inbox searches return whole conversations, so a
 * matching thread can contain sent replies and drafts that must not be logged.
 */
function inspectInboundMessage_(msg, mailboxContext) {
  const senderEmail = extractEmailAddress(msg.getFrom()).toLowerCase().trim();
  if (!senderEmail) return { isInbound: false, reason: 'missing sender address' };
  if (mailboxContext.addresses.has(senderEmail)) {
    return { isInbound: false, reason: 'sent by monitored mailbox or alias' };
  }

  try {
    if (typeof msg.isDraft === 'function' && msg.isDraft()) {
      return { isInbound: false, reason: 'draft message' };
    }
    if (typeof msg.isInTrash === 'function' && msg.isInTrash()) {
      return { isInbound: false, reason: 'trash message' };
    }
    if (typeof msg.isInInbox === 'function' && !msg.isInInbox()) {
      return { isInbound: false, reason: 'message is not in Inbox' };
    }
  } catch (error) {
    Logger.log('Could not inspect Gmail message state; sender-direction rules still apply: ' + error.message);
  }

  return { isInbound: true, reason: 'external message in Inbox' };
}

/**
 * Builds a workbook-wide identity index. Gmail message ID is authoritative for
 * new rows; the sender/date/subject key is retained only to recognize legacy
 * rows created before message IDs were stored.
 */
function getExistingEmailIdentityIndex_(ss) {
  const index = { messageIds: new Set(), fallbackIds: new Set() };
  const messageIdColumn = MONTHLY_EMAIL_HEADERS.indexOf('Gmail Message ID') + 1;
  const monthlySheets = ss.getSheets().filter(function(sheet) {
    return /^[A-Za-z]{3}-\d{4}$/.test(sheet.getName());
  });

  monthlySheets.forEach(function(sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const readableColumns = Math.min(sheet.getMaxColumns(), MONTHLY_EMAIL_HEADERS.length);
    const rows = sheet.getRange(2, 1, lastRow - 1, readableColumns).getValues();
    rows.forEach(function(row) {
      const messageId = row.length >= messageIdColumn
        ? String(row[messageIdColumn - 1] || '').trim()
        : '';
      if (messageId) index.messageIds.add(messageId);
      if (!messageId && row[0] && row[4] && row[5]) {
        index.fallbackIds.add(createUniqueEmailId(
          String(row[4]).toLowerCase().trim(),
          safeGetDate(row[0]),
          row[5]
        ));
      }
    });
  });

  Logger.log('Identity index: ' + index.messageIds.size + ' Gmail IDs, ' +
    index.fallbackIds.size + ' legacy fallback IDs across ' + monthlySheets.length + ' monthly sheets.');
  return index;
}

function inspectDuplicateMessage_(msg, senderEmail, dateReceived, subject, identityIndex) {
  const messageId = getGmailMessageId_(msg);
  const fallbackId = createUniqueEmailId(senderEmail, dateReceived, subject);
  if (messageId && identityIndex.messageIds.has(messageId)) {
    return { isDuplicate: true, reason: 'Gmail message ID', messageId: messageId, fallbackId: fallbackId };
  }
  if (identityIndex.fallbackIds.has(fallbackId)) {
    return { isDuplicate: true, reason: 'legacy sender/date/subject key', messageId: messageId, fallbackId: fallbackId };
  }
  return { isDuplicate: false, reason: '', messageId: messageId, fallbackId: fallbackId };
}

function getGmailMessageId_(msg) {
  try {
    return msg && typeof msg.getId === 'function' ? String(msg.getId() || '').trim() : '';
  } catch (ignore) {
    return '';
  }
}

function recordMessageIdentity_(identityIndex, identity) {
  if (identity.messageId) {
    identityIndex.messageIds.add(identity.messageId);
  } else if (identity.fallbackId) {
    identityIndex.fallbackIds.add(identity.fallbackId);
  }
}

function normalizeSubject(subject) {
  if (!subject) return "";
  return subject
    .replace(/^(Re:|RE:|Fwd:|FWD:|Fw:|FW:)\s*/gi, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function extractEmailAddress(senderString) {
  try {
    if (!senderString) return "";
    const match = senderString.match(/<(.+?)>/);
    if (match) {
      return match[1].trim().toLowerCase();
    }
    if (senderString.includes('@')) {
      return senderString.trim().toLowerCase();
    }
    return senderString.trim();
  } catch (error) {
    return senderString || "";
  }
}

/**
 * Extracts sender name from the "Sender" field in monthly sheets
 * Format examples:
 *   "John Doe <john@email.com>" → "John Doe"
 *   "john@email.com" → "john" (username from email)
 *   "Company Name <info@company.com>" → "Company Name"
 * @param {string} senderString - The full sender string
 * @returns {string} - The extracted name
 */
function extractSenderName(senderString) {
  try {
    if (!senderString) return "Unknown Sender";

    const senderTrimmed = senderString.trim();

    // Try to extract name from "Name <email>" format
    const match = senderTrimmed.match(/^(.+?)\s*<(.+?)>/);
    if (match) {
      let name = match[1].trim();
      const email = match[2].trim();

      // Remove quotes if present
      name = name.replace(/^["']|["']$/g, '').trim();

      // If name is empty after cleaning, use email username
      if (!name || name.length === 0) {
        const emailUsername = email.split('@')[0];
        return emailUsername.charAt(0).toUpperCase() + emailUsername.slice(1);
      }

      return name;
    }

    // If no angle brackets, check if it's just an email
    if (senderTrimmed.includes('@')) {
      // Extract username from email and capitalize
      const emailParts = senderTrimmed.split('@');
      const username = emailParts[0];

      // Convert dots and underscores to spaces, then capitalize each word
      const cleanedName = username
        .replace(/[._-]/g, ' ')
        .split(' ')
        .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
        .join(' ');

      return cleanedName;
    }

    // Otherwise return as-is (it's already a name)
    return senderTrimmed;

  } catch (error) {
    Logger.log('Error extracting sender name: ' + error.message);

    // Fallback: try to extract something meaningful
    if (senderString && senderString.includes('@')) {
      const username = senderString.split('@')[0].replace(/[^a-zA-Z0-9]/g, '');
      return username.charAt(0).toUpperCase() + username.slice(1);
    }

    return "Unknown Sender";
  }
}

function detectKeywords(text, keywords) {
  try {
    if (!text || !keywords) return false;
    const lowerText = text.toLowerCase();
    for (let i = 0; i < keywords.length; i++) {
      if (lowerText.includes(keywords[i].toLowerCase())) {
        return true;
      }
    }
    return false;
  } catch (error) {
    return false;
  }
}

function processEmailMessage(msg, thread, userEmail, existingFallbackIds) {
  const dateReceived = safeGetDate(msg.getDate());
  const sender = msg.getFrom();
  const senderEmail = extractEmailAddress(sender);
  const subject = msg.getSubject() || "(No Subject)";
  const body = msg.getPlainBody() || "";

  const bodySnippet = body.substring(0, 200).replace(/\n/g, " ").trim();
  const summary = subject + ' — ' + bodySnippet;

  const replies = detectEnhancedReplies(thread, userEmail, senderEmail);
  const dateResponded = replies.dateResponded;
  const dateFollowUp = replies.dateFollowUp;

  const emailType = determineEmailType(msg, thread, senderEmail, subject, existingFallbackIds);
  const classification = classifyEmailByAddress(senderEmail);

  const isUnsubscribed = classification.isTrash || detectKeywords(body + " " + subject, CONFIG.UNSUBSCRIBE_KEYWORDS);
  const isESign = classification.isESign;
  const isInternalDomain = classification.isInternalDomain; // Changed from isCompanyDomain

  // Special handling for internal domains - no meeting detection
  const meetingRequested = (isUnsubscribed || isESign || isInternalDomain) ? "No" :
    (detectKeywords(body + " " + subject, CONFIG.MEETING_KEYWORDS) ? "Yes" : "No");

  let status = "Lead";

  // Priority 1: Check if it's internal domain (@duranschulze.com OR @filepino.com) - set to "N/A"
  if (isInternalDomain) {
    status = "N/A";
  }
  // Priority 2: Check if trash/unsubscribed
  else if (isUnsubscribed) {
    status = "Trash";
  }
  // Priority 3: Check if e-sign
  else if (isESign) {
    status = "E-sign";
  }
  // Priority 4: Normal classification
  else {
    if (meetingRequested === "Yes") {
      status = "Prospect";
    }
    if (detectKeywords(body + " " + subject, CONFIG.CONVERSION_KEYWORDS)) {
      status = "Client";
    }
  }

  return [
    dateReceived, dateResponded, dateFollowUp,
    sender, senderEmail, subject, emailType,
    meetingRequested, status, summary
  ];
}


function detectEnhancedReplies(thread, userEmail, originalSenderEmail) {
  let dateResponded = "";
  let dateFollowUp = "";

  try {
    const messages = thread.getMessages();
    const ourResponses = [];

    for (let i = 0; i < messages.length; i++) {
      const msg = messages[i];
      const fromEmail = extractEmailAddress(msg.getFrom());
      const msgDate = safeGetDate(msg.getDate());

      if (isInternalSender_(fromEmail, userEmail)) {
        ourResponses.push({
          date: msgDate,
          sender: fromEmail
        });
      }
    }

    ourResponses.sort(function(a, b) { return a.date - b.date; });

    if (ourResponses.length > 0) {
      dateResponded = ourResponses[0].date;
    }

    if (ourResponses.length > 1) {
      dateFollowUp = ourResponses[ourResponses.length - 1].date;
    }

  } catch (error) {
    Logger.log('Error in reply detection: ' + error.message);
  }

  return { dateResponded: dateResponded, dateFollowUp: dateFollowUp };
}

/**
 * Returns true only for the active mailbox or an address ending in one of the
 * configured internal domains. endsWith prevents lookalike external domains
 * such as "@duranschulze.com.example" from being treated as internal.
 */
function isInternalSender_(email, userEmail) {
  const normalizedEmail = String(email || '').trim().toLowerCase();
  const normalizedUserEmail = String(userEmail || '').trim().toLowerCase();
  if (!normalizedEmail) return false;
  if (normalizedUserEmail && normalizedEmail === normalizedUserEmail) return true;

  return CONFIG.INTERNAL_DOMAINS.some(function(domain) {
    const normalizedDomain = String(domain || '').trim().toLowerCase();
    return normalizedDomain && normalizedEmail.endsWith(normalizedDomain);
  });
}

function determineEmailType(currentMsg, thread, senderEmail, subject, existingFallbackIds) {
  try {
    const normalizedSubject = normalizeSubject(subject);
    const currentMsgDate = safeGetDate(currentMsg.getDate());
    const currentMsgId = currentMsg.getId();

    const threadMessages = thread.getMessages();
    let isFirstInThread = true;

    for (let i = 0; i < threadMessages.length; i++) {
      const msg = threadMessages[i];
      const msgSubject = msg.getSubject() || "";
      const msgNormalizedSubject = normalizeSubject(msgSubject);
      const msgDate = safeGetDate(msg.getDate());
      const msgId = msg.getId();

      if (msgId === currentMsgId) continue;

      if (msgNormalizedSubject === normalizedSubject && msgDate < currentMsgDate) {
        isFirstInThread = false;
        break;
      }
    }

    if (isFirstInThread && existingFallbackIds) {
      const existingEmailsArray = Array.from(existingFallbackIds);
      const senderPrefix = String(senderEmail || '').toLowerCase() + '_';
      for (let i = 0; i < existingEmailsArray.length; i++) {
        const existingId = String(existingEmailsArray[i] || '');
        if (!existingId.toLowerCase().startsWith(senderPrefix)) continue;

        const remainder = existingId.substring(senderPrefix.length);
        const timestampSeparator = remainder.indexOf('_');
        if (timestampSeparator === -1) continue;

        const existingSubject = remainder.substring(timestampSeparator + 1);
        if (existingSubject === normalizedSubject) {
          isFirstInThread = false;
          break;
        }
      }
    }

    return isFirstInThread ? "New" : "Continuation";

  } catch (error) {
    return "New";
  }
}

function classifyEmailByAddress(senderEmail) {
  const senderLower = senderEmail.toLowerCase();
  let isTrash = false;
  let isESign = false;
  let isInternalDomain = false; // NEW: Changed from isCompanyDomain

  // Check if it's internal domain (@duranschulze.com OR @filepino.com)
  for (let i = 0; i < CONFIG.INTERNAL_DOMAINS.length; i++) {
    if (senderLower.includes(CONFIG.INTERNAL_DOMAINS[i].toLowerCase())) {
      isInternalDomain = true;
      return { isTrash: false, isESign: false, isInternalDomain: true }; // Return early
    }
  }

  // Check trash emails
  for (let i = 0; i < CONFIG.TRASH_EMAILS.length; i++) {
    if (senderLower === CONFIG.TRASH_EMAILS[i].toLowerCase()) {
      isTrash = true;
      break;
    }
  }

  // Check trash domains
  if (!isTrash) {
    for (let i = 0; i < CONFIG.TRASH_DOMAINS.length; i++) {
      if (senderLower.includes(CONFIG.TRASH_DOMAINS[i].toLowerCase())) {
        isTrash = true;
        break;
      }
    }
  }

  // Check e-sign emails
  for (let i = 0; i < CONFIG.ESIGN_EMAILS.length; i++) {
    if (senderLower === CONFIG.ESIGN_EMAILS[i].toLowerCase()) {
      isESign = true;
      break;
    }
  }

  return { isTrash: isTrash, isESign: isESign, isInternalDomain: isInternalDomain };
}


function getRecentMonthSheets(ss, monthsCount) {
  const sheets = ss.getSheets();
  const monthlySheets = sheets.filter(function(sh) {
    return /^[A-Za-z]{3}-\d{4}$/.test(sh.getName());
  });

  monthlySheets.sort(function(a, b) {
    const dateA = parseMonthSheetName(a.getName());
    const dateB = parseMonthSheetName(b.getName());
    return dateB - dateA;
  });

  return monthlySheets.slice(0, monthsCount);
}

/**
 * Parses month-year sheet name to Date object
 * @param {string} sheetName - Format: "MMM-yyyy" (e.g., "Oct-2025")
 * @returns {Date} - Date object representing that month
 */
function parseMonthSheetName(sheetName) {
  try {
    const parts = sheetName.split('-');
    if (parts.length !== 2) return new Date(0);

    const month = parts[0];
    const year = parseInt(parts[1]);

    const monthIndex = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                       'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'].indexOf(month);

    if (monthIndex === -1) return new Date(0);

    return new Date(year, monthIndex, 1);
  } catch (error) {
    Logger.log('Error parsing month sheet name: ' + error.message);
    return new Date(0);
  }
}


function updateConversionTracking(conversionSheet, emailData) {
  try {
    const dateReceived = emailData[0];
    const dateResponded = emailData[1];
    const dateFollowUp = emailData[2];
    const sender = emailData[3];
    const senderEmail = emailData[4];
    const subject = emailData[5];
    const status = emailData[8];
    const summary = emailData[9];

    // Skip ONLY trash and e-sign emails
    // "N/A" status (company domain) IS allowed in Conversion Tracking
    if (status === "Trash" || status === "E-sign") {
      return;
    }

    // Extract sender name from full sender string
    const senderName = extractSenderName(sender);

    const normalizedSubject = normalizeSubject(subject);
    const lastRow = conversionSheet.getLastRow();
    let found = false;
    let rowIndex = -1;

    if (lastRow > 1) {
      const data = conversionSheet.getRange(2, 1, lastRow - 1, 8).getValues();

      for (let i = 0; i < data.length; i++) {
        const rowSenderEmail = String(data[i][0]).toLowerCase().trim();
        const rowSubject = normalizeSubject(data[i][2]);

        if (rowSenderEmail === senderEmail.toLowerCase().trim() && rowSubject === normalizedSubject) {
          rowIndex = i + 2;
          found = true;
          break;
        }
      }
    }

    if (found && rowIndex > 0) {
      // UPDATE existing entry
      const currentTotalEmails = conversionSheet.getRange(rowIndex, 6).getValue() || 0;

      let lastContactedDate = dateReceived;
      if (dateFollowUp && dateFollowUp instanceof Date) {
        lastContactedDate = dateFollowUp;
      } else if (dateResponded && dateResponded instanceof Date) {
        lastContactedDate = dateResponded;
      }

      const existingLastContacted = conversionSheet.getRange(rowIndex, 5).getValue();
      if (!existingLastContacted || lastContactedDate > existingLastContacted) {
        conversionSheet.getRange(rowIndex, 5).setValue(lastContactedDate);
      }

      conversionSheet.getRange(rowIndex, 2).setValue(senderName);
      conversionSheet.getRange(rowIndex, 6).setValue(currentTotalEmails + 1);
      conversionSheet.getRange(rowIndex, 7).setValue(status); // Can be "N/A", "Lead", "Prospect", "Client"
      conversionSheet.getRange(rowIndex, 8).setValue(summary);

      Logger.log('Updated Conversion Tracking: ' + senderEmail + ' (' + senderName + ') - Status: ' + status + ' - Row ' + rowIndex);

    } else {
      // ADD new entry
      let firstContactDate = dateReceived;
      let lastContactedDate = dateReceived;

      if (dateFollowUp && dateFollowUp instanceof Date) {
        lastContactedDate = dateFollowUp;
      } else if (dateResponded && dateResponded instanceof Date) {
        lastContactedDate = dateResponded;
      }

      const newRow = [
        senderEmail,
        senderName,
        subject,
        firstContactDate,
        lastContactedDate,
        1,
        status, // Can be "N/A", "Lead", "Prospect", "Client"
        summary
      ];

      conversionSheet.appendRow(newRow);
      Logger.log('Added new Conversion Tracking entry: ' + senderEmail + ' (' + senderName + ') - Status: ' + status);
    }

  } catch (error) {
    Logger.log('Error updating conversion tracking: ' + error.message);
  }
}


function setupDashboardHeader(dashboard, now) {
  dashboard.getRange("A1").setValue("📊 Enhanced CRM Dashboard v" + CRM_VERSION)
    .setFontWeight("bold").setFontSize(18).setFontColor('#1f4e79');

  dashboard.getRange("A2").setValue("Last Sync:")
    .setFontWeight("bold").setFontColor('#666666');

  dashboard.getRange("B2").setValue(now)
    .setNumberFormat("yyyy-MM-dd HH:mm:ss").setFontColor('#333333');

  dashboard.getRange("A3").setValue('Enhanced CRM v' + CRM_VERSION)
    .setFontSize(8).setFontColor('#999999');
}

function displayMetrics(dashboard, recentSheets) {
  dashboard.getRange("A5").setValue("📈 Recent Performance").setFontWeight("bold").setFontSize(14);

  let totalEmails = 0;
  let totalNew = 0;
  let totalContinuation = 0;
  let totalLeads = 0;
  let totalProspects = 0;
  let totalClients = 0;
  let totalTrash = 0;
  let totalESign = 0;
  let totalInternal = 0;
  let totalNonSales = 0;
  let totalReview = 0;
  let followUpsDue = 0;
  let meetingRequests = 0;

  const nowDate = new Date();

  for (let i = 0; i < recentSheets.length; i++) {
    const sheet = recentSheets[i];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;

    try {
      const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

      for (let j = 0; j < data.length; j++) {
        const row = data[j];
        if (!row[0]) continue;

        totalEmails++;

        const emailType = row[6] || "New";
        const status = row[8] || "Lead";

        if (status === "Trash") {
          totalTrash++;
          continue;
        } else if (status === "E-sign") {
          totalESign++;
          continue;
        } else if (status === "N/A") {
          totalInternal++;
          continue;
        } else if (status === "Non-Sales") {
          totalNonSales++;
          continue;
        } else if (status === "Review") {
          totalReview++;
          continue;
        }

        if (emailType === "New") totalNew++;
        if (emailType === "Continuation") totalContinuation++;

        if (status === "Lead") totalLeads++;
        if (status === "Prospect") totalProspects++;
        if (status === "Client") totalClients++;

        const followUpDate = row[2];
        const meetingRequested = row[7] || "No";

        if (followUpDate && followUpDate instanceof Date && followUpDate < nowDate) {
          followUpsDue++;
        }

        if (meetingRequested === "Yes") meetingRequests++;
      }

    } catch (error) {
      Logger.log('Error processing metrics: ' + error.message);
    }
  }

  const validEmails = totalEmails - totalTrash - totalESign - totalInternal - totalNonSales - totalReview;
  const conversionRate = validEmails > 0 ? ((totalClients / validEmails) * 100).toFixed(2) + '%' : '0%';

  const metricsData = [
    ["📧 Total Emails", totalEmails],
    ["🗑️ Trash/Unsubscribed", totalTrash],
    ["📝 E-sign Documents", totalESign],
    ["✅ Valid Business Emails", validEmails],
    ["🆕 New Conversations", totalNew],
    ["💬 Continuations", totalContinuation],
    ["🔵 Leads", totalLeads],
    ["🟡 Prospects", totalProspects],
    ["🟢 Clients", totalClients],
    ["📊 Conversion Rate", conversionRate],
    ["⏰ Follow-ups Due", followUpsDue],
    ["🤝 Meeting Requests", meetingRequests]
  ];

  dashboard.getRange(6, 1, metricsData.length, 2).setValues(metricsData);
  dashboard.getRange(6, 1, metricsData.length, 1).setFontWeight("bold");

  if (followUpsDue > 0) {
    dashboard.getRange(16, 2).setBackground("#ff4c4c").setFontColor("white");
  }

  if (totalTrash > 0) {
    dashboard.getRange(7, 2).setBackground("#ffeb3b").setFontColor("#333");
  }

  if (totalESign > 0) {
    dashboard.getRange(8, 2).setBackground("#e1f5fe").setFontColor("#0277bd");
  }
}

function displayConversionSummary(dashboard, ss) {
  dashboard.getRange("D5").setValue("💬 Conversation Tracking").setFontWeight("bold").setFontSize(14);

  const conversionSheet = ss.getSheetByName(CONFIG.CONVERSION_SHEET_NAME);
  if (!conversionSheet || conversionSheet.getLastRow() < 2) {
    dashboard.getRange("D6").setValue("No ongoing conversations tracked yet");
    return;
  }

  try {
    const data = conversionSheet
      .getRange(2, 1, conversionSheet.getLastRow() - 1, 8)
      .getValues();
    const summary = calculateConversionSummary_(data);

    const conversionSummary = [
      ["Active Conversations", summary.activeConversations],
      ["Avg Emails/Conversation", summary.averageEmails],
      ["Client Conversions", summary.clientConversions],
      ["Conversation→Client Rate", summary.conversionRate]
    ];

    dashboard.getRange(6, 4, conversionSummary.length, 2).setValues(conversionSummary);
    dashboard.getRange(6, 4, conversionSummary.length, 1).setFontWeight("bold");

  } catch (error) {
    Logger.log('Error displaying conversion summary: ' + error.message);
  }
}

/**
 * Calculates business-conversation metrics from Conversion Tracking rows.
 * Total Emails is column F (index 5) and Status is column G (index 6).
 * Internal N/A records are deliberately excluded from sales metrics.
 */
function calculateConversionSummary_(rows) {
  let activeConversations = 0;
  let totalEmailCount = 0;
  let clientConversions = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    const status = String(row[CONVERSION_TRACKING_COLUMNS.STATUS - 1] || 'Lead').trim();
    if (status === 'N/A' || status === 'Trash' || status === 'E-sign' || status === 'Non-Sales' || status === 'Review') continue;

    const parsedEmailCount = Number(row[CONVERSION_TRACKING_COLUMNS.TOTAL_EMAILS - 1]);
    const totalEmails = isNaN(parsedEmailCount) || parsedEmailCount < 1 ? 1 : parsedEmailCount;

    activeConversations++;
    totalEmailCount += totalEmails;
    if (status === 'Client') clientConversions++;
  }

  return {
    activeConversations: activeConversations,
    averageEmails: activeConversations > 0
      ? (totalEmailCount / activeConversations).toFixed(1)
      : '0.0',
    clientConversions: clientConversions,
    conversionRate: activeConversations > 0
      ? ((clientConversions / activeConversations) * 100).toFixed(1) + '%'
      : '0.0%'
  };
}

function displayActionItems(dashboard, ss, recentSheets) {
  dashboard.getRange("A20").setValue("⚡ Priority Actions").setFontWeight("bold").setFontSize(14);

  const actions = [];
  const now = new Date();

  for (let sheetIndex = 0; sheetIndex < recentSheets.length; sheetIndex++) {
    const sheet = recentSheets[sheetIndex];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;

    try {
      const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

      for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
        const row = data[rowIndex];
        const dateReceived = row[0];
        const dateResponded = row[1];
        const dateFollowUp = row[2];
        const senderEmail = row[4];
        const subject = row[5];
        const emailType = row[6];
        const meetingRequested = row[7];
        const status = row[8];

        if (status === "Trash" || status === "E-sign") continue;

        if (dateFollowUp instanceof Date && dateFollowUp < now) {
          const daysOverdue = Math.floor((now - dateFollowUp) / (1000 * 60 * 60 * 24));
          actions.push('🔴 OVERDUE: ' + senderEmail + ' (' + daysOverdue + ' days)');
        }

        if (meetingRequested === "Yes" && emailType === "New" && !dateResponded) {
          const hoursSinceReceived = Math.floor((now - dateReceived) / (1000 * 60 * 60));
          actions.push('📅 MEETING REQUEST: ' + senderEmail + ' (' + hoursSinceReceived + 'h ago)');
        }

        if (emailType === "Continuation" && !dateResponded) {
          const hoursSinceEmail = (now - dateReceived) / (1000 * 60 * 60);
          if (hoursSinceEmail > 24) {
            const hoursRounded = Math.floor(hoursSinceEmail);
            actions.push('💬 PENDING: ' + senderEmail + ' (' + hoursRounded + 'h ago)');
          }
        }
      }

    } catch (error) {
      Logger.log('Error processing actions: ' + error.message);
    }
  }

  if (actions.length > 0) {
    dashboard.getRange("A21").setValue(actions.length + ' items need attention:');
    const maxActions = Math.min(actions.length, 8);
    for (let i = 0; i < maxActions; i++) {
      dashboard.getRange(22 + i, 1, 1, 4).merge().setValue(actions[i]);
    }

    if (actions.length > 8) {
      dashboard.getRange(30, 1).setValue('... and ' + (actions.length - 8) + ' more');
    }
  } else {
    dashboard.getRange("A21").setValue("✅ All conversations up to date!");
  }
}

function formatDashboard(dashboard) {
  try {
    dashboard.autoResizeColumns(1, dashboard.getLastColumn());

    dashboard.setColumnWidth(1, 180);
    dashboard.setColumnWidth(2, 120);
    dashboard.setColumnWidth(4, 180);
    dashboard.setColumnWidth(5, 120);
    dashboard.setColumnWidth(7, 180);
    dashboard.setColumnWidth(8, 120);

    const metricsRange = dashboard.getRange(5, 1, 12, 2);
    metricsRange.setBorder(true, true, true, true, true, true);

    const conversionRange = dashboard.getRange(5, 4, 8, 2);
    conversionRange.setBorder(true, true, true, true, true, true);

    const potentialRange = dashboard.getRange(5, 7, 7, 2);
    potentialRange.setBorder(true, true, true, true, true, true);

    dashboard.setFrozenRows(4);
  } catch (error) {
    Logger.log('Error formatting dashboard: ' + error.message);
  }
}

function setupEnhancedCRM() {
  try {
    Logger.log('Setting up Enhanced CRM System v' + CRM_VERSION + '...');

    const ss = getSpreadsheet();
    if (!ss) {
      throw new Error('Unable to access spreadsheet');
    }

    Logger.log('Spreadsheet access confirmed');

    initializeWorkbook_(ss);
    Logger.log('Workbook, formatting, labels, and automation initialized');

    const result = syncNewEmails();
    Logger.log('Initial email import complete: ' + result.newEmails + ' emails processed');

    buildEnhancedDashboard();
    Logger.log('Enhanced dashboard created');

    Logger.log('Enhanced CRM v' + CRM_VERSION + ' Setup Complete!');
    Logger.log('Visit your spreadsheet: https://docs.google.com/spreadsheets/d/' + ss.getId());

  } catch (error) {
    Logger.log('Setup failed: ' + error.message);
    throw error;
  }
}

function createTriggers() {
  setupAutomationTriggers();
}

function testEnhancedConfiguration() {
  try {
    assertMonitoredMailboxAccount_('configuration test');
    Logger.log('Testing Enhanced CRM v4.2 Configuration...');

    const ss = getSpreadsheet();
    if (!ss) {
      throw new Error('Spreadsheet access failed');
    }
    Logger.log('✓ Spreadsheet access OK: ' + ss.getName());

    const userEmail = Session.getActiveUser().getEmail();
    const testThreads = GmailApp.search('in:inbox', 0, 3);

    Logger.log('✓ Gmail access OK: ' + userEmail);
    Logger.log('✓ Found ' + testThreads.length + ' test threads');

    // Test Drive access
    try {
      const rootFolder = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
      Logger.log('✓ Drive access OK: ' + rootFolder.getName());
    } catch (e) {
      Logger.log('⚠ Drive folder access warning: ' + e.message);
    }

    // Test sheet operations
    const testSheet = ss.insertSheet('Test_' + new Date().getTime());
    testSheet.getRange('A1').setValue('Test successful');
    ss.deleteSheet(testSheet);
    Logger.log('✓ Sheet creation OK');

    const existingTriggers = ScriptApp.getProjectTriggers();
    Logger.log('✓ Trigger access OK: ' + existingTriggers.length + ' existing triggers');

    // Test Map Sheet
    const mapSheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);
    if (mapSheet) {
      Logger.log('✓ Map Sheet found');
    } else {
      Logger.log('⚠ Map Sheet not found');
    }

    // Test Engagement Info Sheet
    const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);
    if (infoSheet) {
      Logger.log('✓ Engagement Information Sheet found');
    } else {
      Logger.log('⚠ Engagement Information Sheet not found');
    }

    Logger.log('✅ All tests passed! Configuration is valid.');
    return true;

  } catch (error) {
    Logger.log('❌ Configuration test failed: ' + error.message);
    return false;
  }
}

function performMaintenance() {
  try {
    Logger.log('Starting maintenance...');

    const ss = getSpreadsheet();
    if (!ss) throw new Error('Unable to access spreadsheet');

    const sheets = ss.getSheets();
    const monthlySheets = sheets.filter(function(sh) {
      return /^[A-Za-z]{3}-\d{4}$/.test(sh.getName());
    });

    const cutoffDate = new Date();
    cutoffDate.setMonth(cutoffDate.getMonth() - CONFIG.ARCHIVE_MONTHS_THRESHOLD);

    let archivedCount = 0;

    for (let i = 0; i < monthlySheets.length; i++) {
      const sheet = monthlySheets[i];
      const sheetDate = parseMonthSheetName(sheet.getName());

      if (sheetDate < cutoffDate) {
        try {
          sheet.hideSheet();
          archivedCount++;
        } catch (error) {
          Logger.log('Could not archive sheet: ' + error.message);
        }
      }
    }

    if (archivedCount > 0) {
      Logger.log('Archived ' + archivedCount + ' old sheets');
    }

    buildEnhancedDashboard();
    Logger.log('Maintenance completed');

  } catch (error) {
    Logger.log('Maintenance failed: ' + error.message);
  }
}

function exportDataToCSV(monthFilter) {
  if (!monthFilter) monthFilter = 'All';

  try {
    Logger.log('Exporting data to CSV...');

    const ss = getSpreadsheet();
    if (!ss) throw new Error('Unable to access spreadsheet');

    let dataToExport = [];
    const headers = [
      "Date Received", "Date Responded", "Date Follow-up",
      "Sender", "Sender Email", "Subject", "Email Type",
      "Meeting Requested", "Status", "Summary"
    ];

    dataToExport.push(headers);

    if (monthFilter === 'All') {
      const sheets = ss.getSheets();
      const monthlySheets = sheets.filter(function(sh) {
        return /^[A-Za-z]{3}-\d{4}$/.test(sh.getName());
      });

      for (let i = 0; i < monthlySheets.length; i++) {
        const sheet = monthlySheets[i];
        if (sheet.getLastRow() > 1) {
          const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
          dataToExport = dataToExport.concat(data);
        }
      }
    } else {
      const monthSheet = ss.getSheetByName(monthFilter);
      if (monthSheet && monthSheet.getLastRow() > 1) {
        const data = monthSheet.getRange(2, 1, monthSheet.getLastRow() - 1, 10).getValues();
        dataToExport = dataToExport.concat(data);
      }
    }

    const csvRows = [];
    for (let i = 0; i < dataToExport.length; i++) {
      const row = dataToExport[i];
      const csvRow = [];

      for (let j = 0; j < row.length; j++) {
        const cell = row[j];
        if (cell instanceof Date) {
          csvRow.push(Utilities.formatDate(cell, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'));
        } else {
          csvRow.push('"' + String(cell).replace(/"/g, '""') + '"');
        }
      }
      csvRows.push(csvRow.join(','));
    }

    const csvContent = csvRows.join('\n');
    const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd_HH-mm');
    const blob = Utilities.newBlob(csvContent, 'text/csv', 'CRM_Export_' + monthFilter + '_' + timestamp + '.csv');
    const file = DriveApp.createFile(blob);

    Logger.log('CSV export completed: ' + file.getName());
    return file.getUrl();

  } catch (error) {
    Logger.log('Export failed: ' + error.message);
    return null;
  }
}

function cleanupExcludedDomainsFromEngagement() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);

    if (!infoSheet || infoSheet.getLastRow() < 2) {
      Logger.log('No data to clean');
      SpreadsheetApp.getUi().alert('No Data', 'Engagement Information Sheet is empty.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const data = infoSheet.getDataRange().getValues();
    const header = data[0];
    const rows = data.slice(1);

    let deletedCount = 0;
    const deletionReasons = {};

    // Process from bottom to top to avoid index shifting
    for (let i = rows.length - 1; i >= 0; i--) {
      const email = String(rows[i][4]).toLowerCase().trim(); // Column E (Email Address)

      const exclusionCheck = isEmailExcluded(email);

      if (exclusionCheck.isExcluded) {
        Logger.log('Deleting row ' + (i + 2) + ': ' + email + ' (' + exclusionCheck.reason + ')');
        infoSheet.deleteRow(i + 2); // +2 because: +1 for header, +1 for 1-based indexing
        deletedCount++;

        // Track deletion reasons
        if (!deletionReasons[exclusionCheck.reason]) {
          deletionReasons[exclusionCheck.reason] = 0;
        }
        deletionReasons[exclusionCheck.reason]++;
      }
    }

    let reasonsText = '';
    for (const reason in deletionReasons) {
      reasonsText += '\n• ' + reason + ': ' + deletionReasons[reason];
    }

    Logger.log('========================================');
    Logger.log('Cleanup complete. Deleted ' + deletedCount + ' rows with excluded emails.');
    Logger.log('Breakdown:' + reasonsText);
    Logger.log('========================================');

    SpreadsheetApp.getUi().alert('Cleanup Complete',
      'Removed ' + deletedCount + ' rows with excluded emails from Engagement Information Sheet.\n\nBreakdown:' + reasonsText,
      SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    Logger.log('Error in cleanup: ' + error.message);
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Master cleanup function - removes excluded emails from Conversion Tracking
 * and Engagement Information, and tags them as "Trash" in monthly sheets
 */
function cleanupAllExcludedEmails() {
  const ui = SpreadsheetApp.getUi();

  try {
    const response = ui.alert(
      '🧹 Clean All Excluded Emails',
      'This will:\n\n• TAG excluded emails as "Trash" or "N/A" in Monthly Sheets\n• DELETE from Conversion Tracking\n• DELETE from Engagement Information Sheet\n\nContinue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Step 1: Tag in monthly sheets
    const monthlyResults = tagExcludedEmailsAsTrash(ss);

    // Step 2: Delete from Conversion Tracking
    const conversionResults = cleanupConversionTracking(ss);

    // Step 3: Delete from Engagement Information
    const engagementResults = cleanupEngagementInformation(ss);

    const summary = '🧹 Cleanup Complete!\n\n' +
                    '📊 Results:\n' +
                    '• Monthly Sheets: ' + monthlyResults.totalTagged + ' rows tagged\n' +
                    '• Conversion Tracking: ' + conversionResults.deleted + ' rows deleted\n' +
                    '• Engagement Information: ' + engagementResults.deleted + ' rows deleted';

    ui.alert('✅ Cleanup Complete', summary, ui.ButtonSet.OK);

  } catch (error) {
    Logger.log('Error in cleanup: ' + error.message);
    ui.alert('⚠️ Error', error.message, ui.ButtonSet.OK);
  }
}


/**
 * Tags excluded emails as "Trash" or "N/A" in monthly sheets
 * Internal domains (@duranschulze.com, @filepino.com) get "N/A" status
 * Other excluded domains/patterns get "Trash" status
 * Both use light gray background
 */
function tagExcludedEmailsAsTrash(ss) {

  try {
    Logger.log('Tagging excluded emails in monthly sheets...');

    const sheets = ss.getSheets();
    const monthlySheets = sheets.filter(function(sh) {
      return /^[A-Za-z]{3}-\d{4}$/.test(sh.getName());
    });

    let totalTagged = 0;
    let companyDomainTagged = 0;
    let sheetsProcessed = 0;

    for (let s = 0; s < monthlySheets.length; s++) {
      const sheet = monthlySheets[s];
      const sheetName = sheet.getName();
      const lastRow = sheet.getLastRow();

      if (lastRow < 2) continue;

      Logger.log('Processing sheet: ' + sheetName);

      const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
      let taggedInSheet = 0;

      for (let i = 0; i < data.length; i++) {
        const rowIndex = i + 2;
        const senderEmail = String(data[i][4]).toLowerCase().trim();
        const currentStatus = data[i][8];

        // Skip if already marked as Trash or N/A
        if (currentStatus === 'Trash' || currentStatus === 'N/A') continue;

        const exclusionCheck = isEmailExcluded(senderEmail);

        // Company domain gets "N/A" status
        if (exclusionCheck.isCompanyDomain) {
          sheet.getRange(rowIndex, 9).setValue('N/A'); // Column I (Status)

          // Apply LIGHT GRAY background for company domain
          sheet.getRange(rowIndex, 1, 1, 10)
            .setBackground('#D3D3D3')  // Light gray background
            .setFontColor('#666666');  // Dark gray text for readability

          taggedInSheet++;
          companyDomainTagged++;
          totalTagged++;

          Logger.log('  Tagged row ' + rowIndex + ' as N/A (company domain): ' + senderEmail);
        }
        // Other excluded emails get "Trash" status
        else if (exclusionCheck.isExcluded) {
          sheet.getRange(rowIndex, 9).setValue('Trash'); // Column I (Status)

          // Apply LIGHT GRAY background for trash (same as N/A)
          sheet.getRange(rowIndex, 1, 1, 10)
            .setBackground('#D3D3D3')  // Light gray background
            .setFontColor('#666666');  // Dark gray text for readability

          taggedInSheet++;
          totalTagged++;

          Logger.log('  Tagged row ' + rowIndex + ' as Trash: ' + senderEmail + ' (' + exclusionCheck.reason + ')');
        }
      }

      if (taggedInSheet > 0) {
        Logger.log('✓ ' + sheetName + ': Tagged ' + taggedInSheet + ' rows');
        sheetsProcessed++;
      }
    }

    Logger.log('Monthly sheets tagging complete:');
    Logger.log('  • Total tagged: ' + totalTagged);
    Logger.log('  • Company domain (N/A): ' + companyDomainTagged);
    Logger.log('  • Trash: ' + (totalTagged - companyDomainTagged));
    Logger.log('  • Sheets processed: ' + sheetsProcessed);

    return {
      totalTagged: totalTagged,
      sheetsProcessed: sheetsProcessed,
      companyDomainTagged: companyDomainTagged
    };

  } catch (error) {
    Logger.log('Error tagging monthly sheets: ' + error.message);
    return { totalTagged: 0, sheetsProcessed: 0, companyDomainTagged: 0 };
  }
}



/**
 * DELETES excluded emails from Conversion Tracking sheet
 */
function cleanupConversionTracking(ss) {
  try {
    const conversionSheet = ss.getSheetByName(CONFIG.CONVERSION_SHEET_NAME);

    if (!conversionSheet || conversionSheet.getLastRow() < 2) {
      return { deleted: 0 };
    }

    const data = conversionSheet.getDataRange().getValues();
    const rows = data.slice(1);

    let deletedCount = 0;

    for (let i = rows.length - 1; i >= 0; i--) {
      const senderEmail = String(rows[i][0]).toLowerCase().trim();
      const exclusionCheck = isEmailExcluded(senderEmail);

      if (exclusionCheck.isExcluded) {
        conversionSheet.deleteRow(i + 2);
        deletedCount++;
      }
    }

    return { deleted: deletedCount };

  } catch (error) {
    Logger.log('Error cleaning Conversion Tracking: ' + error.message);
    return { deleted: 0 };
  }
}


/**
 * DELETES excluded emails from Engagement Information Sheet
 */
function cleanupEngagementInformation(ss) {
  try {
    const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);

    if (!infoSheet || infoSheet.getLastRow() < 2) {
      return { deleted: 0 };
    }

    const data = infoSheet.getDataRange().getValues();
    const rows = data.slice(1);

    let deletedCount = 0;

    for (let i = rows.length - 1; i >= 0; i--) {
      const email = String(rows[i][4]).toLowerCase().trim();
      const exclusionCheck = isEmailExcluded(email);

      if (exclusionCheck.isExcluded) {
        infoSheet.deleteRow(i + 2);
        deletedCount++;
      }
    }

    return { deleted: deletedCount };

  } catch (error) {
    Logger.log('Error cleaning Engagement Information: ' + error.message);
    return { deleted: 0 };
  }
}



function refreshAllFinancialFormulas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);

    if (!infoSheet) {
      throw new Error('Engagement Information Sheet not found');
    }

    setupFinancialFormulas(infoSheet);

    SpreadsheetApp.getUi().alert(
      '✅ Success',
      'All financial formulas have been refreshed!',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    Logger.log('Financial formulas refreshed successfully');

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('Error refreshing formulas: ' + error.message);
  }
}

/**
 * Reformats all existing monthly sheets to use light gray for Trash and N/A
 * Run this ONCE to update old data with new color scheme
 */
function reformatMonthlySheets() {
  const ui = SpreadsheetApp.getUi();

  try {
    const response = ui.alert(
      '🎨 Reformat Monthly Sheets',
      'This will update the background color of all Trash and N/A rows to light gray.\n\nContinue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const monthlySheets = sheets.filter(function(sh) {
      return /^[A-Za-z]{3}-\d{4}$/.test(sh.getName());
    });

    let totalReformatted = 0;

    for (let s = 0; s < monthlySheets.length; s++) {
      const sheet = monthlySheets[s];
      const lastRow = sheet.getLastRow();

      if (lastRow < 2) continue;

      const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

      for (let i = 0; i < data.length; i++) {
        const rowIndex = i + 2;
        const status = data[i][8]; // Column I (Status)

        if (status === 'Trash' || status === 'N/A') {
          sheet.getRange(rowIndex, 1, 1, 10)
            .setBackground('#D3D3D3')
            .setFontColor('#666666');

          totalReformatted++;
        }
      }
    }

    ui.alert('✅ Complete', 'Reformatted ' + totalReformatted + ' rows', ui.ButtonSet.OK);

  } catch (error) {
    Logger.log('Error reformatting: ' + error.message);
    ui.alert('⚠️ Error', error.message, ui.ButtonSet.OK);
  }
}

/**
 * Diagnostic function - checks Conversion Tracking sender names
 * Run this from Apps Script to see what names are currently stored
 */
function diagnoseConversionTrackingNames() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const conversionSheet = ss.getSheetByName(CONFIG.CONVERSION_SHEET_NAME);

    if (!conversionSheet || conversionSheet.getLastRow() < 2) {
      Logger.log('No data in Conversion Tracking');
      return;
    }

    Logger.log('========================================');
    Logger.log('CONVERSION TRACKING - Current Names:');
    Logger.log('========================================');

    const data = conversionSheet.getRange(2, 1, conversionSheet.getLastRow() - 1, 2).getValues();

    for (let i = 0; i < data.length; i++) {
      const email = data[i][0];
      const name = data[i][1];
      const isEmail = name && name.includes('@');

      Logger.log((i + 1) + '. Email: ' + email);
      Logger.log('   Name: "' + name + '"' + (isEmail ? ' ⚠️ (IS EMAIL - NEEDS FIX)' : ' ✓'));
      Logger.log('');
    }

    Logger.log('========================================');

  } catch (error) {
    Logger.log('Error in diagnostic: ' + error.message);
  }
}

/**
 * Fixes existing Conversion Tracking data by re-extracting sender names from monthly sheets
 * Run this ONCE to fix all existing rows that have email addresses instead of names
 */
function fixConversionTrackingSenderNames() {
  const ui = SpreadsheetApp.getUi();

  try {
    const response = ui.alert(
      '🔧 Fix Sender Names',
      'This will update all Sender Names in Conversion Tracking by re-extracting from monthly sheets.\n\n' +
      'This will replace email addresses with actual names.\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      ui.alert('Cancelled', 'Fix cancelled by user.', ui.ButtonSet.OK);
      return;
    }

    Logger.log('========================================');
    Logger.log('Fixing Conversion Tracking sender names...');
    Logger.log('========================================');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const conversionSheet = ss.getSheetByName(CONFIG.CONVERSION_SHEET_NAME);

    if (!conversionSheet || conversionSheet.getLastRow() < 2) {
      ui.alert('No Data', 'Conversion Tracking is empty.', ui.ButtonSet.OK);
      return;
    }

    // Get all monthly sheets
    const sheets = ss.getSheets();
    const monthlySheets = sheets.filter(function(sh) {
      return /^[A-Za-z]{3}-\d{4}$/.test(sh.getName());
    });

    // Build a lookup map: email → sender name (from monthly sheets)
    const emailToNameMap = {};

    Logger.log('Building email-to-name map from monthly sheets...');

    for (let s = 0; s < monthlySheets.length; s++) {
      const sheet = monthlySheets[s];
      const lastRow = sheet.getLastRow();

      if (lastRow < 2) continue;

      const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

      for (let i = 0; i < data.length; i++) {
        const sender = data[i][3];      // Column D: Sender
        const senderEmail = String(data[i][4]).toLowerCase().trim(); // Column E: Sender Email

        if (senderEmail && sender) {
          const extractedName = extractSenderName(sender);

          // Only store if it's a better name (not an email)
          if (!extractedName.includes('@') && extractedName !== 'Unknown Sender') {
            emailToNameMap[senderEmail] = extractedName;
          }
        }
      }
    }

    Logger.log('✓ Built map with ' + Object.keys(emailToNameMap).length + ' email-to-name mappings');

    // Update Conversion Tracking
    const conversionData = conversionSheet.getRange(2, 1, conversionSheet.getLastRow() - 1, 2).getValues();
    let updatedCount = 0;

    Logger.log('Updating Conversion Tracking names...');

    for (let i = 0; i < conversionData.length; i++) {
      const rowIndex = i + 2;
      const email = String(conversionData[i][0]).toLowerCase().trim();
      const currentName = conversionData[i][1];

      // Check if we have a better name
      if (emailToNameMap[email]) {
        const newName = emailToNameMap[email];

        // Only update if current name is an email or different
        if (currentName !== newName) {
          conversionSheet.getRange(rowIndex, 2).setValue(newName);
          updatedCount++;
          Logger.log('  Row ' + rowIndex + ': "' + currentName + '" → "' + newName + '"');
        }
      }
    }

    Logger.log('========================================');
    Logger.log('✓ Fix complete! Updated ' + updatedCount + ' rows');
    Logger.log('========================================');

    // Now push to Engagement Information
    Logger.log('Pushing updated names to Engagement Information...');
    pushProspectsLeadsToEngagement(ss);

    const summary = `🔧 Fix Complete!\n\n` +
                    `📊 Results:\n` +
                    `• Conversion Tracking rows updated: ${updatedCount}\n` +
                    `• Total email-name mappings found: ${Object.keys(emailToNameMap).length}\n` +
                    `• Engagement Information also updated\n\n` +
                    `Check Engagement Information Sheet - Client Name column should now show actual names!`;

    ui.alert('✅ Fix Complete', summary, ui.ButtonSet.OK);

  } catch (error) {
    Logger.log('Error fixing names: ' + error.message);
    ui.alert('⚠️ Error', 'Fix error: ' + error.message, ui.ButtonSet.OK);
  }
}

function testBasicSetup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('✓ Spreadsheet access OK: ' + ss.getName());

    const sheets = ss.getSheets();
    Logger.log('✓ Total sheets: ' + sheets.length);

    for (let i = 0; i < sheets.length; i++) {
      Logger.log('  - ' + sheets[i].getName());
    }

    return true;
  } catch (error) {
    Logger.log('✗ Error: ' + error.message);
    Logger.log('✗ Stack: ' + error.stack);
    return false;
  }
}

function testCoreSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const mapSheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);
    Logger.log('Map Sheet: ' + (mapSheet ? '✓ Found' : '✗ Missing'));

    const conversionSheet = ss.getSheetByName(CONFIG.CONVERSION_SHEET_NAME);
    Logger.log('Conversion Tracking: ' + (conversionSheet ? '✓ Found' : '✗ Missing'));

    const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);
    Logger.log('Engagement Info: ' + (infoSheet ? '✓ Found' : '✗ Missing'));

    return true;
  } catch (error) {
    Logger.log('✗ Error: ' + error.message);
    Logger.log('✗ Stack: ' + error.stack);
    return false;
  }
}

function testConfig() {
  try {
    Logger.log('Testing CONFIG...');
    Logger.log('ROOT_FOLDER_ID: ' + CONFIG.ROOT_FOLDER_ID);
    Logger.log('INTERNAL_DOMAINS: ' + CONFIG.INTERNAL_DOMAINS.join(', '));
    Logger.log('EXCLUDED_DOMAINS: ' + CONFIG.EXCLUDED_DOMAINS.join(', '));
    Logger.log('EXCLUDED_PATTERNS: ' + CONFIG.EXCLUDED_PATTERNS.join(', '));

    return true;
  } catch (error) {
    Logger.log('✗ Error: ' + error.message);
    Logger.log('✗ Stack: ' + error.stack);
    return false;
  }
}

function emergencyTest() {
  try {
    Logger.log('========================================');
    Logger.log('EMERGENCY TEST - Step by step');
    Logger.log('========================================');

    // Step 1: Access spreadsheet
    Logger.log('Step 1: Accessing spreadsheet...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('✓ Success: ' + ss.getName());

    // Step 2: List all sheets
    Logger.log('Step 2: Listing sheets...');
    const sheets = ss.getSheets();
    Logger.log('✓ Found ' + sheets.length + ' sheets');
    sheets.forEach(function(sheet) {
      Logger.log('  - ' + sheet.getName() + ' (' + sheet.getLastRow() + ' rows)');
    });

    // Step 3: Check for monthly sheets
    Logger.log('Step 3: Checking for monthly sheets...');
    const monthlySheets = sheets.filter(function(sh) {
      return /^[A-Za-z]{3}-\d{4}$/.test(sh.getName());
    });
    Logger.log('✓ Found ' + monthlySheets.length + ' monthly sheets');

    // Step 4: Test extractSenderName function
    Logger.log('Step 4: Testing extractSenderName...');
    const testName1 = extractSenderName('John Doe <john@email.com>');
    Logger.log('  Test 1: "John Doe <john@email.com>" → "' + testName1 + '"');

    const testName2 = extractSenderName('jane.smith@company.com');
    Logger.log('  Test 2: "jane.smith@company.com" → "' + testName2 + '"');

    Logger.log('========================================');
    Logger.log('✓ ALL TESTS PASSED!');
    Logger.log('========================================');

    SpreadsheetApp.getUi().alert('✅ Test Passed', 'Emergency test completed successfully!', SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    Logger.log('========================================');
    Logger.log('✗ ERROR FOUND:');
    Logger.log('Message: ' + error.message);
    Logger.log('Line: ' + error.lineNumber);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========================================');

    SpreadsheetApp.getUi().alert('⚠️ Test Failed', 'Error: ' + error.message + '\n\nCheck execution logs for details.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function checkForDuplicates() {
  Logger.log('========================================');
  Logger.log('CHECKING FOR DUPLICATE FUNCTIONS');
  Logger.log('========================================');

  // This will fail if duplicates exist
  try {
    updateConversionTracking(null, []);
    Logger.log('❌ updateConversionTracking is defined (should fail with null)');
  } catch (e) {
    Logger.log('✓ updateConversionTracking check passed');
  }

  try {
    const name = extractSenderName('John Doe <john@test.com>');
    Logger.log('✓ extractSenderName returned: ' + name);

    if (name === 'john@test.com' || name.includes('@')) {
      Logger.log('❌ ERROR: extractSenderName is returning email, not name!');
      Logger.log('   You have the WRONG version of the function');
    } else {
      Logger.log('✓ extractSenderName is working correctly');
    }
  } catch (e) {
    Logger.log('❌ extractSenderName error: ' + e.message);
  }

  Logger.log('========================================');
}

// ========================================
// MISSING FUNCTION #1: getOrCreateConversionTrackingSheet
// ========================================

function getOrCreateConversionTrackingSheet(ss) {
  let conversionSheet = ss.getSheetByName(CONFIG.CONVERSION_SHEET_NAME);

  if (!conversionSheet) {
    conversionSheet = ss.insertSheet(CONFIG.CONVERSION_SHEET_NAME);
  }

  const currentHeaders = conversionSheet
    .getRange(1, 1, 1, CONVERSION_TRACKING_HEADERS.length)
    .getValues()[0];
  const blankHeaders = currentHeaders.every(function(value) { return !String(value || '').trim(); });
  const validHeaders = CONVERSION_TRACKING_HEADERS.every(function(header, index) {
    return String(currentHeaders[index] || '').trim() === header;
  });
  if (blankHeaders) {
    conversionSheet
      .getRange(1, 1, 1, CONVERSION_TRACKING_HEADERS.length)
      .setValues([CONVERSION_TRACKING_HEADERS]);
  } else if (!validHeaders) {
    throw new Error('Conversion Tracking headers do not match the required schema. Existing data was not changed.');
  }

  formatConversionTrackingSheet_(conversionSheet);
  Logger.log('Conversion Tracking sheet verified and formatted');

  return conversionSheet;
}

/**
 * Updates existing @duranschulze.com and @filepino.com emails from "Trash" to "N/A"
 * Applies to: Monthly Sheets, Conversion Tracking, Engagement Information
 */
function updateInternalDomainsToNA() {
  const ui = SpreadsheetApp.getUi();

  try {
    const response = ui.alert(
      '🔄 Update Internal Domains to N/A',
      'This will update all @duranschulze.com and @filepino.com emails:\n\n' +
      '• Change status from "Trash" → "N/A" in Monthly Sheets\n' +
      '• Update status to "N/A" in Conversion Tracking\n' +
      '• Change background to light gray\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      ui.alert('Cancelled', 'Update cancelled by user.', ui.ButtonSet.OK);
      return;
    }

    Logger.log('========================================');
    Logger.log('Updating internal domains to N/A status...');
    Logger.log('========================================');

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Step 1: Update Monthly Sheets
    const monthlyResults = updateMonthlySheetInternalDomains(ss);

    // Step 2: Update Conversion Tracking
    const conversionResults = updateConversionTrackingInternalDomains(ss);

    // Step 3: Remove from Engagement Information (internal domains shouldn't be there)
    const engagementResults = removeInternalDomainsFromEngagement(ss);

    Logger.log('========================================');
    Logger.log('Update complete!');
    Logger.log('========================================');

    const summary = `🔄 Update Complete!\n\n` +
                    `📊 Results:\n` +
                    `• Monthly Sheets: ${monthlyResults.updated} rows updated to N/A\n` +
                    `• Conversion Tracking: ${conversionResults.updated} rows updated to N/A\n` +
                    `• Engagement Information: ${engagementResults.removed} internal domain rows removed\n\n` +
                    `All @duranschulze.com and @filepino.com emails now have N/A status!`;

    ui.alert('✅ Update Complete', summary, ui.ButtonSet.OK);

  } catch (error) {
    Logger.log('Error updating internal domains: ' + error.message);
    ui.alert('⚠️ Error', 'Update error: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Helper: Updates monthly sheets internal domains to N/A
 */
function updateMonthlySheetInternalDomains(ss) {
  try {
    Logger.log('Updating monthly sheets...');

    const sheets = ss.getSheets();
    const monthlySheets = sheets.filter(function(sh) {
      return /^[A-Za-z]{3}-\d{4}$/.test(sh.getName());
    });

    let totalUpdated = 0;

    for (let s = 0; s < monthlySheets.length; s++) {
      const sheet = monthlySheets[s];
      const sheetName = sheet.getName();
      const lastRow = sheet.getLastRow();

      if (lastRow < 2) continue;

      Logger.log('Processing sheet: ' + sheetName);

      const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

      for (let i = 0; i < data.length; i++) {
        const rowIndex = i + 2;
        const senderEmail = String(data[i][4]).toLowerCase().trim(); // Column E
        const currentStatus = data[i][8]; // Column I

        // Check if email is from internal domains
        let isInternalDomain = false;
        for (const domain of CONFIG.INTERNAL_DOMAINS) {
          if (senderEmail.includes(domain.toLowerCase())) {
            isInternalDomain = true;
            break;
          }
        }

        // If it's internal domain and NOT already N/A, update it
        if (isInternalDomain && currentStatus !== 'N/A') {
          // Update status to N/A
          sheet.getRange(rowIndex, 9).setValue('N/A');

          // Apply light gray background
          sheet.getRange(rowIndex, 1, 1, 10)
            .setBackground('#D3D3D3')
            .setFontColor('#666666');

          totalUpdated++;
          Logger.log('  ✓ Updated row ' + rowIndex + ': ' + senderEmail + ' → N/A');
        }
      }
    }

    Logger.log('Monthly sheets: Updated ' + totalUpdated + ' rows to N/A status');

    return { updated: totalUpdated };

  } catch (error) {
    Logger.log('Error updating monthly sheets: ' + error.message);
    return { updated: 0 };
  }
}

/**
 * Helper: Updates Conversion Tracking internal domains to N/A
 */
function updateConversionTrackingInternalDomains(ss) {
  try {
    Logger.log('Updating Conversion Tracking...');

    const conversionSheet = ss.getSheetByName(CONFIG.CONVERSION_SHEET_NAME);

    if (!conversionSheet || conversionSheet.getLastRow() < 2) {
      Logger.log('No data in Conversion Tracking');
      return { updated: 0 };
    }

    const data = conversionSheet.getRange(2, 1, conversionSheet.getLastRow() - 1, 8).getValues();
    let updatedCount = 0;

    for (let i = 0; i < data.length; i++) {
      const rowIndex = i + 2;
      const senderEmail = String(data[i][0]).toLowerCase().trim(); // Column A
      const currentStatus = data[i][6]; // Column G

      // Check if email is from internal domains
      let isInternalDomain = false;
      for (const domain of CONFIG.INTERNAL_DOMAINS) {
        if (senderEmail.includes(domain.toLowerCase())) {
          isInternalDomain = true;
          break;
        }
      }

      // If it's internal domain and NOT already N/A, update it
      if (isInternalDomain && currentStatus !== 'N/A') {
        conversionSheet.getRange(rowIndex, 7).setValue('N/A'); // Column G: Status
        updatedCount++;
        Logger.log('  ✓ Updated row ' + rowIndex + ': ' + senderEmail + ' → N/A');
      }
    }

    Logger.log('Conversion Tracking: Updated ' + updatedCount + ' rows to N/A status');

    return { updated: updatedCount };

  } catch (error) {
    Logger.log('Error updating Conversion Tracking: ' + error.message);
    return { updated: 0 };
  }
}

/**
 * Helper: Removes internal domain emails from Engagement Information
 */
function removeInternalDomainsFromEngagement(ss) {
  try {
    Logger.log('Removing internal domains from Engagement Information...');

    const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);

    if (!infoSheet || infoSheet.getLastRow() < 2) {
      Logger.log('No data in Engagement Information');
      return { removed: 0 };
    }

    const data = infoSheet.getDataRange().getValues();
    const rows = data.slice(1);

    let removedCount = 0;

    // Process from bottom to top to avoid index shifting
    for (let i = rows.length - 1; i >= 0; i--) {
      const email = String(rows[i][4]).toLowerCase().trim(); // Column E

      // Check if email is from internal domains
      let isInternalDomain = false;
      for (const domain of CONFIG.INTERNAL_DOMAINS) {
        if (email.includes(domain.toLowerCase())) {
          isInternalDomain = true;
          break;
        }
      }

      // If it's internal domain, DELETE it from Engagement Info
      if (isInternalDomain) {
        infoSheet.deleteRow(i + 2);
        removedCount++;
        Logger.log('  ✓ Removed row ' + (i + 2) + ': ' + email);
      }
    }

    Logger.log('Engagement Information: Removed ' + removedCount + ' internal domain rows');

    return { removed: removedCount };

  } catch (error) {
    Logger.log('Error removing from Engagement Information: ' + error.message);
    return { removed: 0 };
  }
}

/**
 * ONE-TIME HISTORICAL SYNC - Syncs emails from January 2024 to present
 * WARNING: This processes 10,000+ emails and may take 10-20 minutes
 * Only run this ONCE, then use regular syncNewEmails() for ongoing sync
 */
function syncHistoricalEmailsFromJan2024() {
  const ui = SpreadsheetApp.getUi();

  try {
    assertMonitoredMailboxAccount_('historical Gmail sync');
    const response = ui.alert(
      '⏰ Historical Email Sync',
      '🚨 WARNING: This will sync ALL emails from January 2024 to present.\n\n' +
      '⏱️ Estimated time: 10-20 minutes for 10,000+ emails\n' +
      '📧 This creates separate monthly sheets for each month\n' +
      '💾 This only needs to run ONCE\n\n' +
      '⚠️ Do NOT close this window while processing!\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      ui.alert('Cancelled', 'Historical sync cancelled.', ui.ButtonSet.OK);
      return;
    }

    Logger.log('========================================');
    Logger.log('STARTING HISTORICAL EMAIL SYNC');
    Logger.log('From: January 1, 2024');
    Logger.log('To: ' + new Date().toISOString());
    Logger.log('========================================');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const startDate = new Date('2024-01-01');
    const endDate = new Date();

    // Calculate date range in Gmail search format
    const startDateStr = Utilities.formatDate(startDate, CONFIG.TIMEZONE, 'yyyy/MM/dd');

    // Build Gmail search query for historical range
    const historicalQuery = 'in:inbox after:' + startDateStr;

    Logger.log('Gmail Search Query: ' + historicalQuery);
    ui.alert('⏳ Processing', 'Starting to fetch emails... This may take several minutes.', ui.ButtonSet.OK);

    // Fetch ALL threads from Jan 2024 onwards (in batches)
    let allThreads = [];
    let startIndex = 0;
    const batchSize = 500;
    let hasMore = true;

    Logger.log('Fetching email threads in batches...');

    while (hasMore) {
      try {
        const threads = GmailApp.search(historicalQuery, startIndex, batchSize);

        if (threads.length === 0) {
          hasMore = false;
        } else {
          allThreads = allThreads.concat(threads);
          startIndex += batchSize;
          Logger.log('  Fetched batch: ' + allThreads.length + ' threads so far...');

          // Show progress every 1000 threads
          if (allThreads.length % 1000 === 0) {
            ui.alert('⏳ Progress', 'Fetched ' + allThreads.length + ' email threads...', ui.ButtonSet.OK);
          }

          if (threads.length < batchSize) {
            hasMore = false;
          }
        }
      } catch (e) {
        Logger.log('Error fetching batch at index ' + startIndex + ': ' + e.message);
        hasMore = false;
      }
    }

    Logger.log('========================================');
    Logger.log('Total threads fetched: ' + allThreads.length);
    Logger.log('========================================');

    const identityIndex = getExistingEmailIdentityIndex_(ss);
    const mailboxContext = getMonitoredMailboxContext_();
    const userEmail = mailboxContext.primaryEmail;
    const monthlySheetCache = {};
    const qualificationBudget = { remaining: AI_CONFIG.MAX_QUALIFICATIONS_PER_SYNC };

    let newEmailsCount = 0;
    let processedCount = 0;
    let outboundSkipped = 0;
    let duplicateSkipped = 0;

    Logger.log('Processing emails and routing to monthly sheets...');

    // Process each thread
    for (let i = 0; i < allThreads.length; i++) {
      const thread = allThreads[i];
      const messages = thread.getMessages();

      for (let j = 0; j < messages.length; j++) {
        const msg = messages[j];
        const inbound = inspectInboundMessage_(msg, mailboxContext);
        if (!inbound.isInbound) {
          outboundSkipped++;
          continue;
        }
        const dateReceived = safeGetDate(msg.getDate());
        const senderEmail = extractEmailAddress(msg.getFrom());
        const subject = msg.getSubject() || "(No Subject)";

        // Skip if before Jan 2024
        if (dateReceived < startDate) continue;

        const identity = inspectDuplicateMessage_(msg, senderEmail, dateReceived, subject, identityIndex);
        if (identity.isDuplicate) {
          duplicateSkipped++;
          continue;
        }

        // Process the email
        let emailData = processEmailMessage(msg, thread, userEmail, identityIndex.fallbackIds);
        emailData = qualifyEmailMessageForCrm_(msg, thread, emailData, qualificationBudget);

        // Route to correct month-year sheet
        const emailMonth = Utilities.formatDate(dateReceived, CONFIG.TIMEZONE, 'MMM-yyyy');

        let monthSheet = monthlySheetCache[emailMonth];
        if (!monthSheet) {
          monthSheet = ss.getSheetByName(emailMonth);
          if (!monthSheet) {
            monthSheet = ss.insertSheet(emailMonth);
            monthSheet.getRange(1, 1, 1, MONTHLY_EMAIL_HEADERS.length)
              .setValues([MONTHLY_EMAIL_HEADERS]);
            Logger.log('  ✓ Created new month sheet: ' + emailMonth);
          }
          formatMonthlySheet_(monthSheet);
          monthlySheetCache[emailMonth] = monthSheet;
        }

        monthSheet.appendRow(emailData);
        recordMessageIdentity_(identityIndex, identity);

        // Apply light gray formatting for Trash and N/A
        const status = emailData[8];
        if (status === 'Trash' || status === 'N/A') {
          const lastRow = monthSheet.getLastRow();
          monthSheet.getRange(lastRow, 1, 1, emailData.length)
            .setBackground('#D3D3D3')
            .setFontColor('#666666');
        }

        newEmailsCount++;

        // Update Conversion Tracking
        const conversionSheet = getOrCreateConversionTrackingSheet(ss);
        updateConversionTracking(conversionSheet, emailData);

        processedCount++;

        // Show progress every 500 emails
        if (processedCount % 500 === 0) {
          Logger.log('  Progress: Processed ' + processedCount + ' emails...');
        }
      }
    }

    Logger.log('========================================');
    Logger.log('Email processing complete!');
    Logger.log('  • Total emails processed: ' + newEmailsCount);
    Logger.log('  • Non-inbound messages skipped: ' + outboundSkipped);
    Logger.log('  • Duplicate messages skipped: ' + duplicateSkipped);
    Logger.log('  • Monthly sheets created/updated: ' + Object.keys(monthlySheetCache).length);
    Logger.log('  • Months: ' + Object.keys(monthlySheetCache).sort().join(', '));
    Logger.log('========================================');

    // Refresh the review queue; promotion to Engagement remains explicit.
    Logger.log('Refreshing Potential Clients review queue...');
    syncPotentialClientsFromMonthlySheets_(ss);

    Logger.log('========================================');
    Logger.log('HISTORICAL SYNC COMPLETE!');
    Logger.log('========================================');

    const summary = '✅ Historical Sync Complete!\n\n' +
                    '📊 Results:\n' +
                    '• Total emails synced: ' + newEmailsCount + '\n' +
                    '• Non-inbound messages skipped: ' + outboundSkipped + '\n' +
                    '• Duplicate messages skipped: ' + duplicateSkipped + '\n' +
                    '• Monthly sheets: ' + Object.keys(monthlySheetCache).length + '\n' +
                    '• Date range: Jan 2024 - ' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'MMM yyyy') + '\n\n' +
                    '📋 Sheets created:\n' +
                    Object.keys(monthlySheetCache).sort().join(', ') + '\n\n' +
                    '✅ You can now use regular "Sync New Emails" for ongoing updates!';

    ui.alert('🎉 Historical Sync Complete', summary, ui.ButtonSet.OK);

  } catch (error) {
    Logger.log('========================================');
    Logger.log('ERROR in historical sync: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========================================');
    ui.alert('⚠️ Error', 'Historical sync error: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Extracts client data from row array
 * @param {Array} rowData - Array of row values
 * @returns {Object} - Client data object
 */
function extractClientData(rowData) {
  return {
    contactDate: rowData[0],
    clientName: rowData[1],
    service: rowData[2],
    contactPerson: rowData[3],
    email: rowData[4],
    phone: rowData[5],
    address: rowData[6],
    remarks: rowData[7],
    engagementFee: rowData[8],
    vat: rowData[9],
    totalGross: rowData[10],
    miscellaneous: rowData[11],
    totalServiceFee: rowData[12],
    serviceRef: rowData[13],
    quoteCreated: rowData[14],
    quoteDateSent: rowData[15],
    quoteStatus: rowData[16],
    quoteFollowUp: rowData[17],
    paymentStatus: rowData[18],
    paymentStatusDate: rowData[19],
    engagementStatus: rowData[20],
    needsRenewal: rowData[21],
    assignedDepartment: rowData[22],
    assignedDepartmentEmail: rowData[23],
    assignedDate: rowData[24],
    signedAgreementPdf: rowData[25],
    folderLink: rowData[26],
    generatedQuotePdf: rowData[27],
    dateDueForRenewal: rowData[ENGAGEMENT_COLUMNS.DATE_DUE_FOR_RENEWAL - 1],
    sourceMonth: rowData[ENGAGEMENT_COLUMNS.SOURCE_MONTH - 1],
    engagementNotificationSent: rowData[ENGAGEMENT_COLUMNS.ENGAGEMENT_NOTIFICATION_SENT - 1],
    paidAssignmentNotificationSent: rowData[ENGAGEMENT_COLUMNS.PAID_ASSIGNMENT_NOTIFICATION_SENT - 1]
  };
}

/**
 * Sends "Paid & Assigned" notification
 * CORRECTED VERSION with proper column references
 */
function sendPaidAssignmentNotification(sheet, row) {
  assertMonitoredMailboxAccount_('paid assignment notification');
  try {
    Logger.log('========================================');
    Logger.log('📧 SENDING PAID & ASSIGNED NOTIFICATION');
    Logger.log('Row: ' + row);
    Logger.log('========================================');

    // Find the correct column for "Paid Assignment Notification Sent"
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let notificationColumn = -1;

    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i]).toLowerCase();
      if (header.includes('paid assignment notification sent') ||
          header.includes('paid client notification')) {
        notificationColumn = i + 1;
        break;
      }
    }

    if (notificationColumn === -1) {
      Logger.log('⚠️ WARNING: "Paid Assignment Notification Sent" column not found!');
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Column "Paid Assignment Notification Sent" not found in sheet.',
        '⚠️ Column Missing',
        10
      );
      return;
    }

    Logger.log('Found "Paid Assignment Notification Sent" at column ' + notificationColumn + ' (' + getColumnLetter(notificationColumn) + ')');

    // Check if notification already sent
    const notificationSentDate = sheet.getRange(row, notificationColumn).getValue();

    if (notificationSentDate) {
      Logger.log('  Paid assignment notification already sent on: ' + notificationSentDate);
      return;
    }

    const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const clientData = extractClientData(rowData);

    Logger.log('Client: ' + clientData.clientName);
    Logger.log('Assigned Dept: ' + clientData.assignedDepartment);
    Logger.log('Assigned Dept Email: ' + clientData.assignedDepartmentEmail);
    Logger.log('Payment Status: ' + clientData.paymentStatus);

    // Collect recipients
    const recipients = [];

    // 1. Add assigned department email
    if (clientData.assignedDepartmentEmail && clientData.assignedDepartmentEmail.trim() !== '') {
      recipients.push({
        email: clientData.assignedDepartmentEmail,
        name: clientData.assignedDepartment || 'Assigned Department'
      });
      Logger.log('Added: ' + clientData.assignedDepartmentEmail + ' (Assigned Dept)');
    }

    // 2. Add mandatory recipients
    const mandatoryRecipients = getNotificationEmails('Email Notification for Paid Client');
    Logger.log('Mandatory recipients found: ' + mandatoryRecipients.length);

    for (let i = 0; i < mandatoryRecipients.length; i++) {
      recipients.push(mandatoryRecipients[i]);
      Logger.log('Added: ' + mandatoryRecipients[i].email + ' (Mandatory)');
    }

    if (recipients.length === 0) {
      Logger.log('❌ ERROR: No recipients found');
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'No recipients found for paid client notification.',
        '⚠️ No Recipients',
        10
      );
      return;
    }

    Logger.log('Total recipients: ' + recipients.length);

    // Build email
    const emailSubject = '[PAID & READY] Client Ready for Service: ' + clientData.clientName;
    const emailBody = buildPaidAssignmentEmailBody(clientData, row);

    Logger.log('Email Subject: ' + emailSubject);

    // Send emails
    let sentCount = 0;
    const errors = [];

    for (let i = 0; i < recipients.length; i++) {
      try {
        Logger.log('Attempting to send to: ' + recipients[i].email);

        GmailApp.sendEmail(
          recipients[i].email,
          emailSubject,
          emailBody,
          {
            htmlBody: emailBody.replace(/\n/g, '<br>'),
            name: 'Duran Schulze CRM System'
          }
        );

        sentCount++;
        Logger.log('  ✅ SUCCESS: Sent to ' + recipients[i].email);

      } catch (emailError) {
        const errorMsg = recipients[i].email + ': ' + emailError.message;
        errors.push(errorMsg);
        Logger.log('  ❌ FAILED: ' + errorMsg);
      }
    }

    Logger.log('========================================');
    Logger.log('Send Summary: ' + sentCount + ' / ' + recipients.length + ' succeeded');
    Logger.log('========================================');

    if (sentCount > 0) {
      // Mark notification as sent - CORRECTED LINE
      const now = new Date();
      sheet.getRange(row, notificationColumn).setValue(now);

      Logger.log('✅ Marked notification as sent in Column ' + notificationColumn + ' (' + getColumnLetter(notificationColumn) + ')');
      Logger.log('   Timestamp: ' + Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'));

      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Paid client notification sent to ' + sentCount + ' recipient(s)\n' +
        'Notification logged in column ' + getColumnLetter(notificationColumn),
        '✅ Service Ready Notification Sent',
        5
      );

    } else {
      Logger.log('❌ All email sends failed');
      Logger.log('Errors: ' + errors.join('; '));

      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Failed to send notifications:\n' + errors.join('\n'),
        '❌ Send Failed',
        10
      );
    }

  } catch (error) {
    Logger.log('========================================');
    Logger.log('❌ CRITICAL ERROR in sendPaidAssignmentNotification');
    Logger.log('Error: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========================================');

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Critical error sending paid notification: ' + error.message,
      '❌ Error',
      10
    );
  }
}



/**
 * Gets notification emails from Map Sheet by Type
 * Uses cache if available, refreshes if needed
 */
function getNotificationEmails(notificationType) {
  try {
    Logger.log('========================================');
    Logger.log('🔍 SEARCHING FOR NOTIFICATION EMAILS');
    Logger.log('Type: "' + notificationType + '"');
    Logger.log('========================================');

    // Check cache first
    if (MAP_CACHE.notificationEmails) {
      if (notificationType === 'Email Notification for Engagement Only' &&
          MAP_CACHE.notificationEmails.engagement &&
          MAP_CACHE.notificationEmails.engagement.length > 0) {
        Logger.log('✓ Found ' + MAP_CACHE.notificationEmails.engagement.length + ' emails in cache');
        return MAP_CACHE.notificationEmails.engagement;
      }

      if (notificationType === 'Email Notification for Paid Client' &&
          MAP_CACHE.notificationEmails.paidClient &&
          MAP_CACHE.notificationEmails.paidClient.length > 0) {
        Logger.log('✓ Found ' + MAP_CACHE.notificationEmails.paidClient.length + ' emails in cache');
        return MAP_CACHE.notificationEmails.paidClient;
      }
    }

    // Cache miss - read from Map Sheet
    Logger.log('Cache miss - reading from Map Sheet');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mapSheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);

    if (!mapSheet) {
      Logger.log('❌ ERROR: Map Sheet not found');
      return [];
    }

    const data = mapSheet.getDataRange().getValues();
    const emails = [];

    // Search for matching notification type
    for (let i = 1; i < data.length; i++) {
      const type = String(data[i][0]).trim();
      const key = String(data[i][1]).trim();
      const value = String(data[i][2]).trim();

      if (type === notificationType && value && value !== '') {
        emails.push({
          email: value,
          name: key || 'Recipient'
        });
        Logger.log('  ✅ MATCH: ' + value + ' (' + key + ')');
      }
    }

    Logger.log('========================================');
    Logger.log('Total matches: ' + emails.length);
    Logger.log('========================================');

    return emails;

  } catch (error) {
    Logger.log('❌ ERROR in getNotificationEmails: ' + error.message);
    return [];
  }
}



/**
 * ONE-TIME FUNCTION: Updates Engagement Information Sheet structure
 * Adds new columns and reorganizes headers
 */
function updateEngagementSheetStructure() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);

    if (!infoSheet) {
      throw new Error('Engagement Information Sheet not found');
    }

    SpreadsheetApp.getUi().alert(
      'Verify Engagement Structure',
      'This safely upgrades the known legacy Engagement layout and preserves existing row data.\n\n' +
      'A backup is still recommended before any structural migration.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    if (!ensureEngagementSheetSchema_(infoSheet)) {
      throw new Error('The Engagement sheet has customized core columns and could not be migrated automatically.');
    }

    const newHeaders = ENGAGEMENT_INFO_HEADERS;
    const currentHeaders = infoSheet.getRange(1, 1, 1, newHeaders.length).getValues()[0];
    const schemaMatches = newHeaders.every(function(header, index) {
      return currentHeaders[index] === header;
    });
    if (!schemaMatches) {
      throw new Error('The Engagement sheet has customized core columns and could not be migrated automatically.');
    }

    // Reapply canonical header formatting after verification/migration.
    infoSheet.getRange(1, 1, 1, newHeaders.length)
      .setFontWeight('bold')
      .setBackground('#EA4335')
      .setFontColor('white')
      .setWrap(true)
      .setHorizontalAlignment('center');

    // Set column widths for new columns
    infoSheet.setColumnWidth(ENGAGEMENT_COLUMNS.GENERATED_QUOTE_PDF, 200);
    infoSheet.setColumnWidth(ENGAGEMENT_COLUMNS.SOURCE_MONTH, 100);
    infoSheet.setColumnWidth(ENGAGEMENT_COLUMNS.ENGAGEMENT_NOTIFICATION_SENT, 150);
    infoSheet.setColumnWidth(ENGAGEMENT_COLUMNS.PAID_ASSIGNMENT_NOTIFICATION_SENT, 180);

    Logger.log('✓ Sheet structure updated to ' + newHeaders.length + ' columns');

    SpreadsheetApp.getUi().alert(
      '✅ Update Complete',
      'Engagement Information Sheet now has ' + newHeaders.length + ' columns.\n\n' +
      'New columns:\n' +
      '• Generated Quote PDF (AB)\n' +
      '• Source Month (AD)\n' +
      '• Engagement Notification Sent (AE)\n' +
      '• Paid Assignment Notification Sent (AF)',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    Logger.log('ERROR: ' + error.message);
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * ONE-TIME FUNCTION: Populates Source Month for ALL existing rows
 * Based on Contact Date (Column A)
 */
function populateAllSourceMonths() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);

    if (!infoSheet) {
      throw new Error('Engagement Information Sheet not found');
    }

    const lastRow = infoSheet.getLastRow();

    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert('No Data', 'No rows to process', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    Logger.log('========================================');
    Logger.log('POPULATING SOURCE MONTHS');
    Logger.log('Processing rows 2 to ' + lastRow);
    Logger.log('========================================');

    let updatedCount = 0;
    let skippedCount = 0;

    // Process each data row
    for (let row = 2; row <= lastRow; row++) {
      const contactDate = infoSheet.getRange(row, 1).getValue(); // Column A

      if (!contactDate || !(contactDate instanceof Date)) {
        Logger.log('Row ' + row + ': No valid contact date, skipping');
        skippedCount++;
        continue;
      }

      // Format as MMM-yyyy
      const sourceMonth = Utilities.formatDate(contactDate, CONFIG.TIMEZONE, 'MMM-yyyy');

      // Update Source Month column (Column AD = 30)
      infoSheet.getRange(row, ENGAGEMENT_COLUMNS.SOURCE_MONTH).setValue(sourceMonth);

      updatedCount++;

      if (updatedCount % 10 === 0) {
        Logger.log('  Processed ' + updatedCount + ' rows...');
      }
    }

    Logger.log('========================================');
    Logger.log('COMPLETE: Updated ' + updatedCount + ' rows');
    Logger.log('Skipped ' + skippedCount + ' rows (no valid date)');
    Logger.log('========================================');

    SpreadsheetApp.getUi().alert(
      '✅ Source Months Updated',
      'Updated ' + updatedCount + ' rows\n' +
      'Skipped ' + skippedCount + ' rows (no valid date)\n\n' +
      'Source Month now matches Contact Date for all rows.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    Logger.log('ERROR: ' + error.message);
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Calculates and sets Date Due for Renewal
 * Based on Contact Date + duration (6 months or 1 year)
 * Duration determined by "Needs Renewal" column value
 *
 * @param {Sheet} sheet - The Engagement Information Sheet
 * @param {number} row - The row number
 */
function calculateRenewalDate(sheet, row) {
  try {
    Logger.log('Calculating renewal date for row ' + row);

    // Get Contact Date (Column A)
    const contactDate = sheet.getRange(row, 1).getValue();

    if (!contactDate || !(contactDate instanceof Date)) {
      Logger.log('  No valid contact date, cannot calculate renewal date');
      return;
    }

    // Get Needs Renewal value (Column V = 22)
    const needsRenewal = sheet.getRange(row, 22).getValue();

    // Determine renewal duration
    let monthsToAdd = 12; // Default: 1 year

    // If "Needs Renewal" = "Yes", use 6 months
    if (needsRenewal === 'Yes') {
      monthsToAdd = 6;
      Logger.log('  Needs Renewal = Yes, using 6 months');
    } else {
      Logger.log('  Using default 1 year renewal period');
    }

    // Calculate renewal date
    const renewalDate = new Date(contactDate);
    renewalDate.setMonth(renewalDate.getMonth() + monthsToAdd);

    // Set Date Due for Renewal (Column AC = 29)
    sheet.getRange(row, ENGAGEMENT_COLUMNS.DATE_DUE_FOR_RENEWAL).setValue(renewalDate);

    const formattedDate = Utilities.formatDate(renewalDate, CONFIG.TIMEZONE, 'MMM dd, yyyy');
    Logger.log('  ✓ Renewal date set to: ' + formattedDate + ' (' + monthsToAdd + ' months from contact date)');

    // Show toast notification
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Renewal due: ' + formattedDate,
      '✓ Renewal Date Calculated',
      3
    );

  } catch (error) {
    Logger.log('ERROR in calculateRenewalDate: ' + error.message);
  }
}

/**
 * ONE-TIME FUNCTION: Recalculates Date Due for Renewal for ALL engaged clients
 */
function recalculateAllRenewalDates() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);

    if (!infoSheet) {
      throw new Error('Engagement Information Sheet not found');
    }

    const lastRow = infoSheet.getLastRow();

    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert('No Data', 'No rows to process', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    Logger.log('========================================');
    Logger.log('RECALCULATING RENEWAL DATES');
    Logger.log('Processing rows 2 to ' + lastRow);
    Logger.log('========================================');

    let updatedCount = 0;
    let skippedCount = 0;

    // Process each data row
    for (let row = 2; row <= lastRow; row++) {
      const engagementStatus = infoSheet.getRange(row, 21).getValue(); // Column U

      // Only process "Engaged" clients
      if (engagementStatus !== 'Engaged') {
        skippedCount++;
        continue;
      }

      // Calculate renewal date
      calculateRenewalDate(infoSheet, row);
      updatedCount++;

      if (updatedCount % 10 === 0) {
        Logger.log('  Processed ' + updatedCount + ' rows...');
      }
    }

    Logger.log('========================================');
    Logger.log('COMPLETE: Updated ' + updatedCount + ' engaged clients');
    Logger.log('Skipped ' + skippedCount + ' rows (not engaged)');
    Logger.log('========================================');

    SpreadsheetApp.getUi().alert(
      '✅ Renewal Dates Updated',
      'Updated ' + updatedCount + ' engaged clients\n' +
      'Skipped ' + skippedCount + ' rows (not engaged)\n\n' +
      'Renewal dates calculated based on:\n' +
      '• Contact Date + 6 months (if Needs Renewal = Yes)\n' +
      '• Contact Date + 1 year (default)',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    Logger.log('ERROR: ' + error.message);
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * TEST FUNCTION: Check if notification emails can be retrieved
 */
function testNotificationEmails() {
  Logger.log('========================================');
  Logger.log('TESTING NOTIFICATION EMAIL RETRIEVAL');
  Logger.log('========================================');

  // Test engagement emails
  const engagementEmails = getNotificationEmails('Email Notification for Engagement Only');
  Logger.log('Engagement notification emails: ' + engagementEmails.length);

  // Test paid emails
  const paidEmails = getNotificationEmails('Email Notification for Paid Client');
  Logger.log('Paid notification emails: ' + paidEmails.length);

  Logger.log('========================================');
  Logger.log('TEST COMPLETE');
  Logger.log('If counts are 0, check Map Sheet entries');
  Logger.log('========================================');

  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Test Results',
    'Engagement notification emails: ' + engagementEmails.length + '\n' +
    'Paid notification emails: ' + paidEmails.length + '\n\n' +
    'Check execution logs for details.',
    ui.ButtonSet.OK
  );
}

/**
 * TEST FUNCTION: Send test email
 */
function testSendEmail() {
  try {
    assertMonitoredMailboxAccount_('test email');
    const testEmail = 'marywendy@duranschulze.com'; // REPLACE WITH YOUR EMAIL

    Logger.log('Attempting to send test email to: ' + testEmail);

    GmailApp.sendEmail(
      testEmail,
      '[TEST] CRM Email Notification Test',
      'This is a test email from your CRM system.\n\nIf you receive this, email sending is working!',
      {
        htmlBody: 'This is a <b>test email</b> from your CRM system.<br><br>If you receive this, email sending is working!',
        name: 'Duran Schulze CRM System'
      }
    );

    Logger.log('✅ Test email sent successfully');

    SpreadsheetApp.getUi().alert(
      '✅ Test Email Sent',
      'Check ' + testEmail + ' for the test email.\n\n' +
      'If you receive it, email sending is working.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    Logger.log('❌ ERROR sending test email: ' + error.message);

    SpreadsheetApp.getUi().alert(
      '❌ Test Failed',
      'Error: ' + error.message + '\n\n' +
      'You may need to authorize Gmail permissions.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * TEST FUNCTION: Manual trigger of engagement notification for specific row
 */
function testEngagementNotificationManual() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);

    const testRow = 2; // CHANGE THIS to your test row number

    Logger.log('Testing engagement notification for row ' + testRow);

    const rowData = sheet.getRange(testRow, 1, 1, 32).getValues()[0];
    const clientData = extractClientData(rowData);

    Logger.log('Client: ' + clientData.clientName);

    sendEngagementNotification(sheet, testRow, clientData);

    SpreadsheetApp.getUi().alert(
      'Test Complete',
      'Check execution logs for details.\n\n' +
      'If successful, check recipient email inboxes.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    Logger.log('ERROR: ' + error.message);
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * ========================================
 * REFRESH MAP SHEET DATA
 * ========================================
 * Reads and validates all Map Sheet data
 * Caches for system use
 * Shows detailed report
 */
function refreshMapSheet() {
  try {
    const ui = SpreadsheetApp.getUi();

    Logger.log('========================================');
    Logger.log('REFRESHING MAP SHEET DATA');
    Logger.log('========================================');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mapSheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);

    if (!mapSheet) {
      throw new Error('Map Sheet not found. Please create it first.');
    }

    // Clear cache
    MAP_CACHE = {
      departments: {},
      services: {},
      notificationEmails: {
        engagement: [],
        paidClient: []
      },
      dropdownOptions: {},
      lastRefresh: new Date()
    };

    const data = mapSheet.getDataRange().getValues();

    if (data.length < 2) {
      throw new Error('Map Sheet is empty. Please add data first.');
    }

    Logger.log('Total rows in Map Sheet: ' + data.length);

    // Statistics
    let stats = {
      departments: 0,
      services: 0,
      engagementEmails: 0,
      paidClientEmails: 0,
      followUpOptions: 0,
      paymentOptions: 0,
      quoteActionOptions: 0,
      renewalOptions: 0,
      errors: [],
      warnings: []
    };

    // Process each row (skip header)
    for (let i = 1; i < data.length; i++) {
      const rowNum = i + 1;
      const type = String(data[i][0]).trim();
      const key = String(data[i][1]).trim();
      const value = String(data[i][2]).trim();

      // Skip empty rows
      if (!type && !key && !value) continue;

      // Validate row has all required fields
      if (!type) {
        stats.warnings.push('Row ' + rowNum + ': Missing Type');
        continue;
      }

      if (!key) {
        stats.warnings.push('Row ' + rowNum + ': Missing Key');
        continue;
      }

      // Process based on type
      switch (type) {
        case 'Department':
          if (!value) {
            stats.errors.push('Row ' + rowNum + ': Department "' + key + '" has no email address');
          } else if (!isValidEmail(value)) {
            stats.errors.push('Row ' + rowNum + ': Invalid email for department "' + key + '": ' + value);
          } else {
            MAP_CACHE.departments[key] = value;
            stats.departments++;
            Logger.log('  ✓ Department: ' + key + ' → ' + value);
          }
          break;

        case 'Service':
          MAP_CACHE.services[key] = value || '';
          stats.services++;
          Logger.log('  ✓ Service: ' + key + (value ? ' → Template ID: ' + value : ' (no template)'));
          break;

        case 'Email Notification for Engagement Only':
          if (!value) {
            stats.errors.push('Row ' + rowNum + ': Engagement notification "' + key + '" has no email');
          } else if (!isValidEmail(value)) {
            stats.errors.push('Row ' + rowNum + ': Invalid email for "' + key + '": ' + value);
          } else {
            MAP_CACHE.notificationEmails.engagement.push({
              name: key,
              email: value
            });
            stats.engagementEmails++;
            Logger.log('  ✓ Engagement Email: ' + key + ' → ' + value);
          }
          break;

        case 'Email Notification for Paid Client':
          if (!value) {
            stats.errors.push('Row ' + rowNum + ': Paid client notification "' + key + '" has no email');
          } else if (!isValidEmail(value)) {
            stats.errors.push('Row ' + rowNum + ': Invalid email for "' + key + '": ' + value);
          } else {
            MAP_CACHE.notificationEmails.paidClient.push({
              name: key,
              email: value
            });
            stats.paidClientEmails++;
            Logger.log('  ✓ Paid Client Email: ' + key + ' → ' + value);
          }
          break;

        case 'FollowUp':
          if (!MAP_CACHE.dropdownOptions.followUp) {
            MAP_CACHE.dropdownOptions.followUp = [];
          }
          MAP_CACHE.dropdownOptions.followUp.push(key);
          stats.followUpOptions++;
          break;

        case 'Payment':
          if (!MAP_CACHE.dropdownOptions.payment) {
            MAP_CACHE.dropdownOptions.payment = [];
          }
          MAP_CACHE.dropdownOptions.payment.push(key);
          stats.paymentOptions++;
          break;

        case 'QuoteAction':
          if (!MAP_CACHE.dropdownOptions.quoteAction) {
            MAP_CACHE.dropdownOptions.quoteAction = [];
          }
          MAP_CACHE.dropdownOptions.quoteAction.push(key);
          stats.quoteActionOptions++;
          break;

        case 'Renewal':
          if (!MAP_CACHE.dropdownOptions.renewal) {
            MAP_CACHE.dropdownOptions.renewal = [];
          }
          MAP_CACHE.dropdownOptions.renewal.push(key);
          stats.renewalOptions++;
          break;

        default:
          stats.warnings.push('Row ' + rowNum + ': Unknown type "' + type + '"');
      }
    }

    Logger.log('========================================');
    Logger.log('MAP SHEET REFRESH COMPLETE');
    Logger.log('========================================');
    Logger.log('Statistics:');
    Logger.log('  Departments: ' + stats.departments);
    Logger.log('  Services: ' + stats.services);
    Logger.log('  Engagement Emails: ' + stats.engagementEmails);
    Logger.log('  Paid Client Emails: ' + stats.paidClientEmails);
    Logger.log('  Follow-Up Options: ' + stats.followUpOptions);
    Logger.log('  Payment Options: ' + stats.paymentOptions);
    Logger.log('  Quote Action Options: ' + stats.quoteActionOptions);
    Logger.log('  Renewal Options: ' + stats.renewalOptions);
    Logger.log('  Errors: ' + stats.errors.length);
    Logger.log('  Warnings: ' + stats.warnings.length);
    Logger.log('========================================');

    // Build report message
    let report = '✅ Map Sheet Refreshed Successfully!\n\n';
    report += '📊 DATA SUMMARY:\n';
    report += '• Departments: ' + stats.departments + '\n';
    report += '• Services: ' + stats.services + '\n';
    report += '• Engagement Notification Emails: ' + stats.engagementEmails + '\n';
    report += '• Paid Client Notification Emails: ' + stats.paidClientEmails + '\n';
    report += '• Follow-Up Options: ' + stats.followUpOptions + '\n';
    report += '• Payment Options: ' + stats.paymentOptions + '\n';
    report += '• Quote Actions: ' + stats.quoteActionOptions + '\n';
    report += '• Renewal Options: ' + stats.renewalOptions + '\n\n';

    if (stats.errors.length > 0) {
      report += '❌ ERRORS (' + stats.errors.length + '):\n';
      for (let i = 0; i < Math.min(stats.errors.length, 5); i++) {
        report += '• ' + stats.errors[i] + '\n';
      }
      if (stats.errors.length > 5) {
        report += '• ... and ' + (stats.errors.length - 5) + ' more errors\n';
      }
      report += '\n';
    }

    if (stats.warnings.length > 0) {
      report += '⚠️ WARNINGS (' + stats.warnings.length + '):\n';
      for (let i = 0; i < Math.min(stats.warnings.length, 3); i++) {
        report += '• ' + stats.warnings[i] + '\n';
      }
      if (stats.warnings.length > 3) {
        report += '• ... and ' + (stats.warnings.length - 3) + ' more warnings\n';
      }
      report += '\n';
    }

    report += 'Last refreshed: ' + Utilities.formatDate(MAP_CACHE.lastRefresh, CONFIG.TIMEZONE, 'MMM dd, yyyy HH:mm:ss');

    // Show report
    ui.alert('📋 Map Sheet Refresh Report', report, ui.ButtonSet.OK);

    // Also show detailed report if errors exist
    if (stats.errors.length > 0) {
      const detailsResponse = ui.alert(
        '❌ Errors Found',
        'There are ' + stats.errors.length + ' errors in Map Sheet.\n\n' +
        'Would you like to see the full error list?',
        ui.ButtonSet.YES_NO
      );

      if (detailsResponse === ui.Button.YES) {
        let fullErrors = '❌ FULL ERROR LIST:\n\n';
        for (let i = 0; i < stats.errors.length; i++) {
          fullErrors += (i + 1) + '. ' + stats.errors[i] + '\n';
        }
        ui.alert('Error Details', fullErrors, ui.ButtonSet.OK);
      }
    }

    // Refresh dropdowns in Engagement Information Sheet
    setupDropdownValidations(ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME));

    return true;

  } catch (error) {
    Logger.log('========================================');
    Logger.log('❌ ERROR REFRESHING MAP SHEET');
    Logger.log('Error: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========================================');

    SpreadsheetApp.getUi().alert(
      '❌ Refresh Failed',
      'Error refreshing Map Sheet:\n\n' + error.message + '\n\n' +
      'Check execution logs for details.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return false;
  }
}

/**
 * Helper function to validate email format
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Removes all triggers
 * Fixed version with proper error handling
 */
function removeAllTriggersFixed() {
  const ui = SpreadsheetApp.getUi();

  try {
    const response = ui.alert(
      '🔧 Remove All Triggers',
      'This will remove ALL existing triggers.\n\n' +
      'You can reinstall them later if needed.\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    Logger.log('========================================');
    Logger.log('REMOVING ALL TRIGGERS');
    Logger.log('========================================');

    const triggers = ScriptApp.getProjectTriggers();

    Logger.log('Found ' + triggers.length + ' trigger(s)');

    let removedCount = 0;

    for (let i = 0; i < triggers.length; i++) {
      const trigger = triggers[i];

      try {
        const handlerFunction = trigger.getHandlerFunction();
        const eventType = trigger.getEventType();

        Logger.log('Removing: ' + handlerFunction + ' (' + eventType + ')');

        ScriptApp.deleteTrigger(trigger);
        removedCount++;

      } catch (e) {
        Logger.log('  ⚠️ Failed to remove trigger: ' + e.message);
      }
    }

    Logger.log('========================================');
    Logger.log('Removed ' + removedCount + ' trigger(s)');
    Logger.log('========================================');

    ui.alert(
      '✅ Triggers Removed',
      'Removed ' + removedCount + ' trigger(s).\n\n' +
      'To reinstall onEdit trigger, run:\n' +
      'Menu > System Management > Setup Edit Trigger',
      ui.ButtonSet.OK
    );

  } catch (error) {
    Logger.log('ERROR: ' + error.message);
    ui.alert(
      '❌ Error',
      'Error removing triggers:\n\n' + error.message + '\n\n' +
      'You may need to grant additional permissions.',
      ui.ButtonSet.OK
    );
  }
}


/**
 * DIAGNOSTIC: Lists all functions that might conflict with onEdit
 */
function diagnosticCheckFunctions() {
  const ui = SpreadsheetApp.getUi();

  Logger.log('========================================');
  Logger.log('DIAGNOSTIC: Checking for function conflicts');
  Logger.log('========================================');

  // List of functions that should exist
  const requiredFunctions = [
    'onEdit',
    'handleContactDateChange',
    'handlePaymentStatusChange',
    'handleEngagementStatusChange',
    'handleDepartmentAssignment',
    'getDepartmentEmail',
    'sendEngagementNotification',
    'sendPaidAssignmentNotification',
    'getNotificationEmails',
    'extractClientData',
    'buildEngagementEmailBody',
    'buildPaidAssignmentEmailBody'
  ];

  let report = '📋 FUNCTION CHECK REPORT\n\n';
  let allGood = true;

  for (let i = 0; i < requiredFunctions.length; i++) {
    const funcName = requiredFunctions[i];

    try {
      // Try to access the function
      const func = eval(funcName);

      if (typeof func === 'function') {
        report += '✅ ' + funcName + ' - EXISTS\n';
        Logger.log('✅ ' + funcName + ' found');
      } else {
        report += '❌ ' + funcName + ' - NOT A FUNCTION\n';
        Logger.log('❌ ' + funcName + ' exists but is not a function');
        allGood = false;
      }
    } catch (e) {
      report += '❌ ' + funcName + ' - MISSING\n';
      Logger.log('❌ ' + funcName + ' not found: ' + e.message);
      allGood = false;
    }
  }

  Logger.log('========================================');

  if (allGood) {
    report += '\n✅ All required functions are present!';
  } else {
    report += '\n⚠️ Some functions are missing or invalid.';
  }

  ui.alert('Function Diagnostic', report, ui.ButtonSet.OK);
}

/**
 * DIAGNOSTIC: Shows all current triggers
 */
function diagnosticShowTriggers() {
  const ui = SpreadsheetApp.getUi();

  Logger.log('========================================');
  Logger.log('CURRENT TRIGGERS');
  Logger.log('========================================');

  const triggers = ScriptApp.getProjectTriggers();

  let report = '📋 INSTALLED TRIGGERS\n\n';
  report += 'Total: ' + triggers.length + '\n\n';

  if (triggers.length === 0) {
    report += 'No installable triggers found.\n\n';
    report += 'Note: onEdit is a "simple trigger" and won\'t show here.\n';
    report += 'Simple triggers work automatically without installation.';
  } else {
    for (let i = 0; i < triggers.length; i++) {
      const trigger = triggers[i];
      const handlerFunction = trigger.getHandlerFunction();
      const eventType = trigger.getEventType();
      const source = trigger.getTriggerSource();

      report += (i + 1) + '. Handler: ' + handlerFunction + '\n';
      report += '   Event: ' + eventType + '\n';
      report += '   Source: ' + source + '\n\n';

      Logger.log('Trigger ' + (i + 1) + ': ' + handlerFunction + ' (' + eventType + ')');

      // Check if handler function exists
      try {
        eval(handlerFunction);
        report += '   ✅ Function exists\n';
      } catch (e) {
        report += '   ❌ Function NOT FOUND (broken trigger)\n';
      }

      report += '\n';
    }
  }

  Logger.log('========================================');

  ui.alert('Trigger Diagnostic', report, ui.ButtonSet.OK);
}


/**
 * Shows current triggers
 * Fixed version with proper error handling
 */
function showCurrentTriggersFixed() {
  const ui = SpreadsheetApp.getUi();

  try {
    Logger.log('========================================');
    Logger.log('CHECKING CURRENT TRIGGERS');
    Logger.log('========================================');

    const triggers = ScriptApp.getProjectTriggers();

    let report = '📋 INSTALLED TRIGGERS\n\n';
    report += 'Total: ' + triggers.length + '\n\n';

    if (triggers.length === 0) {
      report += 'No triggers currently installed.\n\n';
      report += 'To install onEdit trigger:\n';
      report += 'Menu > System Management > Setup Edit Trigger';
    } else {
      for (let i = 0; i < triggers.length; i++) {
        const trigger = triggers[i];

        try {
          const handlerFunction = trigger.getHandlerFunction();
          const eventType = trigger.getEventType();
          const source = trigger.getTriggerSource();

          report += (i + 1) + '. ' + handlerFunction + '\n';
          report += '   Event: ' + eventType + '\n';
          report += '   Source: ' + source + '\n';

          // Check if function exists
          try {
            const func = eval(handlerFunction);
            if (typeof func === 'function') {
              report += '   ✅ Function exists\n';
            } else {
              report += '   ❌ Not a function\n';
            }
          } catch (e) {
            report += '   ❌ Function not found\n';
          }

          report += '\n';

        } catch (e) {
          report += (i + 1) + '. (Error reading trigger)\n\n';
        }
      }
    }

    Logger.log(report);
    ui.alert('Current Triggers', report, ui.ButtonSet.OK);

  } catch (error) {
    Logger.log('ERROR: ' + error.message);
    ui.alert(
      '❌ Error',
      'Error checking triggers:\n\n' + error.message + '\n\n' +
      'You may need to grant additional permissions.',
      ui.ButtonSet.OK
    );
  }
}

/**
 * TEST: Manually trigger engagement notification for a specific row
 * IMPORTANT: Edit the row number below before running
 */
function testEngagementNotificationManual() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);

    // ⚠️ CHANGE THIS to your test row number
    const testRow = 2;

    Logger.log('========================================');
    Logger.log('MANUAL TEST: Engagement Notification');
    Logger.log('Row: ' + testRow);
    Logger.log('========================================');

    // Get the status to verify
    const status = sheet.getRange(testRow, 21).getValue();
    Logger.log('Current Engagement Status: ' + status);

    if (status !== 'Engaged') {
      Logger.log('⚠️ WARNING: Status is not "Engaged", setting it now...');
      sheet.getRange(testRow, 21).setValue('Engaged');
    }

    // Get row data
    const rowData = sheet.getRange(testRow, 1, 1, 32).getValues()[0];
    const clientData = extractClientData(rowData);

    Logger.log('Client: ' + clientData.clientName);
    Logger.log('Email: ' + clientData.email);
    Logger.log('Service: ' + clientData.service);

    // Call the notification function
    Logger.log('Calling sendEngagementNotification...');
    sendEngagementNotification(sheet, testRow, clientData);

    Logger.log('========================================');
    Logger.log('TEST COMPLETE');
    Logger.log('Check execution logs and email inbox');
    Logger.log('========================================');

    SpreadsheetApp.getUi().alert(
      '✅ Test Complete',
      'Manual test executed.\n\n' +
      'Check:\n' +
      '1. Execution logs (View > Logs)\n' +
      '2. Email inbox (billing emails from Map Sheet)\n' +
      '3. Column AE (Engagement Notification Sent)',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    Logger.log('========================================');
    Logger.log('❌ TEST ERROR');
    Logger.log('Error: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========================================');

    SpreadsheetApp.getUi().alert(
      '❌ Test Failed',
      'Error: ' + error.message + '\n\nCheck execution logs for details.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * DIAGNOSTIC: Show exact column numbers for notification tracking
 */
function verifyNotificationColumns() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);

    if (!infoSheet) {
      throw new Error('Engagement Information Sheet not found');
    }

    const lastCol = infoSheet.getLastColumn();
    const headers = infoSheet.getRange(1, 1, 1, lastCol).getValues()[0];

    let report = '📋 NOTIFICATION COLUMN VERIFICATION\n\n';

    // Find notification columns
    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i]).toLowerCase();

      if (header.includes('engagement notification') ||
          header.includes('paid assignment notification') ||
          header.includes('source month') ||
          header.includes('date due for renewal')) {

        report += 'Column ' + (i + 1) + ' (' + getColumnLetter(i + 1) + '): "' + headers[i] + '"\n';
      }
    }

    report += '\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n';
    report += 'EXPECTED STRUCTURE:\n';
    report += 'Column 29 (AC): Date Due for Renewal\n';
    report += 'Column 30 (AD): Source Month\n';
    report += 'Column 31 (AE): Engagement Notification Sent\n';
    report += 'Column 32 (AF): Paid Assignment Notification Sent\n';

    Logger.log(report);
    SpreadsheetApp.getUi().alert('Notification Columns', report, SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    Logger.log('ERROR: ' + error.message);
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function getColumnLetter(col) {
  let letter = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - mod) / 26);
  }
  return letter;
}

function showAutoRespondedEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);
  if (!sheet) return;
  // Filter rows where column Engagement Status = "Auto-Replied" or "Pending Approval"
  // For simplicity, we'll just create a filter view
  const lastRow = sheet.getLastRow();
  const filter = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).createFilter();
  // Not programmatically easy; we'll just alert the user to manually filter.
  SpreadsheetApp.getUi().alert('Filter manually on the Engagement Information sheet for "Auto-Replied" or "Pending Approval" in the Engagement Status column.');
}

function manageCategoryAssignments() {
  const html = HtmlService.createHtmlOutput(
    '<p>Edit the <b>CategoryRouting</b> rows in the Map Sheet to assign email addresses to categories.</p>'
  );
  SpreadsheetApp.getUi().showSidebar(html);
}

//============HELPER FUNCTIONS===============//
function getGeminiScore(emailSnippet) {
  // Simple mock: 5-10 based on length; you can replace with actual Gemini call
  return Math.min(10, 5 + Math.floor(emailSnippet.length / 100));
}

function categorizeIntent(emailSnippet) {
  // Simple keyword-based categorization; you can enhance with Gemini
  const lower = emailSnippet.toLowerCase();
  if (lower.includes('visa')) return 'Visa';
  if (lower.includes('trademark')) return 'Trademark';
  if (lower.includes('business')) return 'Business Formation';
  if (lower.includes('accounting')) return 'Accounting';
  if (lower.includes('legal')) return 'Legal';
  return 'General';
}

function restoreEngagementSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);
  if (!infoSheet) {
    // Create sheet with headers
    infoSheet = ss.insertSheet(CONFIG.ENGAGE_SHEET_NAME);
    infoSheet.getRange(1, 1, 1, ENGAGEMENT_INFO_HEADERS.length).setValues([ENGAGEMENT_INFO_HEADERS]);
    infoSheet.getRange(1, 1, 1, ENGAGEMENT_INFO_HEADERS.length)
      .setFontWeight('bold').setBackground('#EA4335').setFontColor('white');
    infoSheet.setFrozenRows(1);
    // Apply dropdowns & formulas
    setupDropdownValidations(infoSheet);
    setupFinancialFormulas(infoSheet);
  }
  // Repopulate from Conversion Tracking using the configured engagement cutoff.
  pushProspectsLeadsToEngagement(ss);
  SpreadsheetApp.getUi().alert('Engagement Information Sheet restored with currently eligible leads.');
}
