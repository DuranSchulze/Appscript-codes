const SHEET_NAMES = {
  PAGES: 'Pages',
  CONFIG: 'Config',
  LOGS: 'Logs',
  DASHBOARD: 'Dashboard',
  AI_META_DESCRIPTIONS: 'AI Meta Descriptions',
  AI_SEO_AUDIT: 'AI SEO Audit',
  PAGE_KEYWORDS: 'Page Keywords',
};

const GOOGLE_INSPECTION_COLUMNS = [
  'google_inspection_status',
  'google_index_verdict',
  'google_coverage_state',
  'google_last_crawl_time',
  'google_referring_url',
  'google_user_canonical',
  'google_google_canonical',
  'google_robots_allowed',
  'google_inspected_at',
];

const HEADERS = {
  CONFIG: [
    'website_url',
    'host',
    'search_console_property',
    'ga4_property_id',
    'indexnow_key',
    'indexnow_key_url',
    'default_date_range_days',
    'last_refresh_at',
    'last_refresh_status',
  ],
  PAGES: [
    'website_url',
    'host',
    'page_url',
    'page_path',
    'source_method',
    'status_code',
    'title',
    'meta_description',
    'h1',
    'canonical_url',
    'robots_meta',
    'is_indexable',
    'clicks',
    'impressions',
    'ctr',
    'avg_position',
    'sessions',
    'users',
    'conversions',
    'range_start',
    'range_end',
    'last_refreshed_at',
  ].concat(GOOGLE_INSPECTION_COLUMNS),
  LOGS: [
    'timestamp',
    'website_url',
    'action',
    'status',
    'pages_discovered',
    'pages_written',
    'message',
  ],
  DASHBOARD: [
    'website_url',
    'section',
    'metric',
    'value',
    'details',
    'updated_at',
  ],
  AI_META_DESCRIPTIONS: [
    'website_url',
    'host',
    'page_url',
    'page_path',
    'existing_meta_description',
    'existing_char_count',
    'length_status',
    'ai_recommended_meta_description',
    'ai_char_count',
    'recommendation_status',
    'ai_error_message',
    'generated_at',
  ],
  AI_SEO_AUDIT: [
    'website_url',
    'host',
    'page_url',
    'page_path',
    'priority_score',
    'priority_label',
    'issue_category',
    'issue_code',
    'issue_summary',
    'severity',
    'traffic_signal',
    'search_signal',
    'supporting_data',
    'ai_recommendation',
    'recommendation_status',
    'ai_error_message',
    'generated_at',
  ],
  PAGE_KEYWORDS: [
    'page_url',
    'keyword',
  ],
};

const DEFAULT_DATE_RANGE_DAYS = 28;
const ALLOWED_DATE_RANGES = [7, 28, 90];
const TIMEZONE = Session.getScriptTimeZone() || 'Asia/Manila';
const SITEMAP_CANDIDATE_PATHS = ['/sitemap.xml', '/sitemap_index.xml', '/sitemap-index.xml'];
const CRAWL_MAX_PAGES = 100;
const CRAWL_MAX_DEPTH = 2;
const AUDIT_BATCH_SIZE = 20;
const SEARCH_CONSOLE_ROW_LIMIT = 25000;
const GA4_ROW_LIMIT = 100000;
const INDEXNOW_ENDPOINT = 'https://api.indexnow.org/IndexNow';
const URL_INSPECTION_ENDPOINT = 'https://searchconsole.googleapis.com/v1/urlInspection/index:inspect';
const AI_PROVIDER_GEMINI = 'gemini';
const AI_PROVIDER_OPENAI = 'openai';
const GEMINI_API_KEY_PROPERTY = 'GEMINI_API_KEY';
const GEMINI_MODEL_PROPERTY = 'GEMINI_MODEL';
const DEFAULT_GEMINI_MODEL = 'gemini-2.5-flash';
const GEMINI_API_BASE_URL = 'https://generativelanguage.googleapis.com/v1beta/models/';
const OPENAI_API_KEY_PROPERTY = 'OPENAI_API_KEY';
const OPENAI_MODEL_PROPERTY = 'OPENAI_MODEL';
const ACTIVE_AI_PROVIDER_PROPERTY = 'ACTIVE_AI_PROVIDER';
const OPENAI_API_BASE_URL = 'https://api.openai.com/v1';
const META_DESCRIPTION_MIN_LENGTH = 120;
const META_DESCRIPTION_MAX_LENGTH = 150;
const GEMINI_MODEL_CACHE_TTL_SECONDS = 21600;
const GEMINI_RESPONSE_CACHE_TTL_SECONDS = 21600;
const GEMINI_MAX_REQUESTS_PER_RUN = 10;
const GEMINI_BATCH_SIZE = 5;
const GEMINI_REQUEST_DELAY_MS = 600;
const GEMINI_BATCH_DELAY_MS = 1200;
const SHEET_WRITE_BATCH_SIZE = 200;
const TITLE_MIN_LENGTH = 20;
const TITLE_MAX_LENGTH = 60;
const LOW_CTR_THRESHOLD = 0.02;
const HIGH_IMPRESSIONS_THRESHOLD = 100;
const HIGH_SESSIONS_THRESHOLD = 50;
const HIGH_USERS_THRESHOLD = 30;
const PAGE_KEYWORD_ROWS_PER_PAGE = 6;
const AI_SHEETS_FORMULA_FUNCTION = 'GEMINI';
const SOURCE_METHOD_PRIORITY = {
  sitemap: 2,
  crawl: 1,
};

function onOpen() {
  ensureRequiredSheets_();
  const ui = SpreadsheetApp.getUi();
  const websiteSetupMenu = ui
    .createMenu('Website Setup')
    .addItem('Setup Website', 'setupWebsite')
    .addItem('Refresh Website', 'refreshWebsite')
    .addItem('Refresh Website With Date Range', 'refreshWebsiteWithDateRange');
  const googleDataMenu = ui
    .createMenu('Google Data')
    .addItem('Check Google Connections', 'checkGoogleConnections')
    .addItem('Check Google Index Status', 'checkGoogleIndexStatus')
    .addItem('Submit to IndexNow', 'submitToIndexNow');
  const dashboardMenu = ui
    .createMenu('Dashboard')
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addItem('View Dashboard', 'viewDashboard');
  const aiToolsMenu = ui
    .createMenu('AI Tools')
    .addItem('Setup Gemini AI', 'setupGeminiAi')
    .addItem('Select Gemini AI Model', 'selectGeminiAiModel')
    .addItem('Setup OpenAI', 'setupOpenAi')
    .addItem('Select OpenAI AI Model', 'selectOpenAiModel')
    .addItem('Select Active AI Provider', 'selectActiveAiProvider')
    .addItem('Check Active AI Connection', 'checkActiveAiConnection')
    .addItem('Generate AI Meta Descriptions', 'generateAiMetaDescriptions')
    .addItem('Run AI SEO Audit', 'runAiSeoAudit')
    .addItem('Build Page Keywords Sheet', 'buildPageKeywordsSheet')
    .addItem('View AI Meta Descriptions', 'viewAiMetaDescriptions')
    .addItem('View AI SEO Audit', 'viewAiSeoAudit')
    .addItem('View Page Keywords Sheet', 'viewPageKeywords');
  const reportsMenu = ui
    .createMenu('Reports')
    .addItem('View Dashboard', 'viewDashboard')
    .addItem('View Logs', 'viewLogs')
    .addItem('View AI SEO Audit', 'viewAiSeoAudit')
    .addItem('View Page Keywords Sheet', 'viewPageKeywords');
  const maintenanceMenu = ui
    .createMenu('Maintenance')
    .addItem('Clear All Sheet Data', 'clearAllSheetData');

  ui.createMenu('SEO Workspace')
    .addSubMenu(websiteSetupMenu)
    .addSubMenu(googleDataMenu)
    .addSubMenu(dashboardMenu)
    .addSubMenu(aiToolsMenu)
    .addSubMenu(reportsMenu)
    .addSubMenu(maintenanceMenu)
    .addToUi();
}

function onInstall() {
  onOpen();
}

function doGet(e) {
  return HtmlService
    .createHtmlOutput('<html><body><p>SEO Workspace web output is not currently in use.</p></body></html>')
    .setTitle('SEO Workspace');
}

function setupWebsite() {
  ensureRequiredSheets_();
  const ui = SpreadsheetApp.getUi();

  const websiteUrl = promptForWebsiteUrl_();
  if (!websiteUrl) {
    return;
  }

  const normalizedWebsiteUrl = normalizeWebsiteUrl_(websiteUrl);
  const host = getHost_(normalizedWebsiteUrl);
  const searchConsoleProperty = promptForOptionalText_(
    'Search Console Property',
    'Optional: enter the Search Console property.\nExamples:\nhttps://www.example.com/\nsc-domain:example.com\n\nLeave blank to skip for now.'
  );
  const ga4PropertyId = promptForOptionalText_(
    'GA4 Property ID',
    'Optional: enter the numeric GA4 property ID.\nExample: 123456789\n\nLeave blank to skip for now.'
  );

  const configRecord = {
    website_url: normalizedWebsiteUrl,
    host: host,
    search_console_property: normalizeSearchConsoleProperty_(searchConsoleProperty),
    ga4_property_id: String(ga4PropertyId || '').trim(),
    indexnow_key: '',
    indexnow_key_url: '',
    default_date_range_days: String(DEFAULT_DATE_RANGE_DAYS),
    last_refresh_at: '',
    last_refresh_status: '',
  };

  upsertConfigRecord_(configRecord);

  ui.alert(
    'Website Saved',
    'The website was saved. You can now run "Refresh Website" for audit data. Search Console and GA4 can still be updated later in the Config tab.',
    ui.ButtonSet.OK
  );
}

function refreshWebsite() {
  runRefreshWorkflow_({
    promptForDateRange: false,
    actionName: 'refresh_website',
  });
}

function refreshWebsiteWithDateRange() {
  runRefreshWorkflow_({
    promptForDateRange: true,
    actionName: 'refresh_website_with_date_range',
  });
}

function checkGoogleConnections() {
  ensureRequiredSheets_();

  const ui = SpreadsheetApp.getUi();
  const configs = getConfigRecords_();
  if (!configs.length) {
    ui.alert('No websites are configured yet. Run "Setup Website" first.');
    return;
  }

  const selectedConfig = selectWebsiteConfig_(configs);
  if (!selectedConfig) {
    return;
  }

  const dateRange = getDateRange_(7);
  const refreshAt = formatDateTime_(new Date());
  const searchConsole = fetchOptionalSearchConsoleMetrics_(
    selectedConfig,
    dateRange.startDate,
    dateRange.endDate
  );
  const ga4 = fetchOptionalGa4Metrics_(
    selectedConfig.ga4_property_id,
    dateRange.startDate,
    dateRange.endDate
  );

  const status = searchConsole.included || ga4.included ? 'SUCCESS' : 'WARNING';
  const message =
    'Search Console: ' +
    searchConsole.statusMessage +
    '\nGA4: ' +
    ga4.statusMessage;

  appendLogRow_({
    timestamp: refreshAt,
    website_url: selectedConfig.website_url,
    action: 'check_google_connections',
    status: status,
    pages_discovered: 0,
    pages_written: 0,
    message: truncateText_(message.replace(/\n/g, ' '), 500),
  });

  ui.alert(
    'Google Connection Check',
    'Website: ' +
      selectedConfig.website_url +
      '\n\nSearch Console: ' +
      searchConsole.statusMessage +
      '\n\nGA4: ' +
      ga4.statusMessage,
    ui.ButtonSet.OK
  );
}

function checkGoogleIndexStatus() {
  ensureRequiredSheets_();

  const ui = SpreadsheetApp.getUi();
  const configs = getConfigRecords_();
  if (!configs.length) {
    ui.alert('No websites are configured yet. Run "Setup Website" first.');
    return;
  }

  const selectedConfig = selectWebsiteConfig_(configs);
  if (!selectedConfig) {
    return;
  }

  const refreshAt = formatDateTime_(new Date());

  try {
    const pageRows = getPageRowsForWebsite_(selectedConfig.website_url);
    if (!pageRows.length) {
      throw new Error('No pages were found in the Pages tab for ' + selectedConfig.website_url + '. Run "Refresh Website" first.');
    }

    const inspectionRun = inspectWebsiteGoogleIndexStatus_(
      selectedConfig,
      pageRows,
      refreshAt
    );

    writeGoogleInspectionResults_(selectedConfig.website_url, inspectionRun.resultsByUrl);
    refreshDashboardForWebsite_(selectedConfig.website_url, refreshAt);
    appendLogRow_({
      timestamp: refreshAt,
      website_url: selectedConfig.website_url,
      action: 'check_google_index_status',
      status: inspectionRun.status,
      pages_discovered: pageRows.length,
      pages_written: inspectionRun.updatedCount,
      message: inspectionRun.logMessage,
    });

    ui.alert(
      'Google Index Status Check',
      inspectionRun.popupMessage,
      ui.ButtonSet.OK
    );
  } catch (error) {
    const failureMessage = truncateText_(error && error.message ? error.message : String(error), 500);
    appendLogRow_({
      timestamp: refreshAt,
      website_url: selectedConfig.website_url,
      action: 'check_google_index_status',
      status: 'FAILED',
      pages_discovered: 0,
      pages_written: 0,
      message: failureMessage,
    });

    ui.alert(
      'Google Index Status Check Failed',
      failureMessage,
      ui.ButtonSet.OK
    );
  }
}

function submitToIndexNow() {
  ensureRequiredSheets_();

  const ui = SpreadsheetApp.getUi();
  const configs = getConfigRecords_();
  if (!configs.length) {
    ui.alert('No websites are configured yet. Run "Setup Website" first.');
    return;
  }

  const selectedConfig = selectWebsiteConfig_(configs);
  if (!selectedConfig) {
    return;
  }

  const refreshAt = formatDateTime_(new Date());

  try {
    const pageUrls = getPageUrlsForWebsite_(selectedConfig.website_url);
    if (!pageUrls.length) {
      throw new Error('No pages were found in the Pages tab for ' + selectedConfig.website_url + '. Run "Refresh Website" first.');
    }

    const verification = verifyIndexNowConfiguration_(selectedConfig);
    const responseCode = submitUrlsToIndexNow_(
      selectedConfig.website_url,
      verification.indexnowKey,
      verification.indexnowKeyUrl,
      pageUrls
    );
    const message =
      'IndexNow submission accepted for ' +
      pageUrls.length +
      ' URLs using the shared endpoint. This notifies IndexNow-supported engines and does not guarantee Google rankings.';

    appendLogRow_({
      timestamp: refreshAt,
      website_url: selectedConfig.website_url,
      action: 'submit_to_indexnow',
      status: 'SUCCESS',
      pages_discovered: pageUrls.length,
      pages_written: pageUrls.length,
      message: message + ' HTTP ' + responseCode + '.',
    });

    ui.alert(
      'IndexNow Submitted',
      message,
      ui.ButtonSet.OK
    );
  } catch (error) {
    const failureMessage = truncateText_(error && error.message ? error.message : String(error), 500);
    appendLogRow_({
      timestamp: refreshAt,
      website_url: selectedConfig.website_url,
      action: 'submit_to_indexnow',
      status: 'FAILED',
      pages_discovered: 0,
      pages_written: 0,
      message: failureMessage,
    });

    ui.alert(
      'IndexNow Submission Failed',
      failureMessage,
      ui.ButtonSet.OK
    );
  }
}

function setupGeminiAi() {
  ensureRequiredSheets_();

  const ui = SpreadsheetApp.getUi();
  const enteredApiKey = promptForText_(
    'Gemini API Key',
    'Enter your Gemini API key.\nIt will be saved once for this Apps Script project and reused for AI meta description generation.'
  );

  if (!enteredApiKey) {
    return;
  }

  const apiKey = String(enteredApiKey || '').trim();
  const refreshAt = formatDateTime_(new Date());

  try {
    validateAiProvider_(AI_PROVIDER_GEMINI, apiKey);
    PropertiesService.getScriptProperties().setProperty(GEMINI_API_KEY_PROPERTY, apiKey);
    const selectedModel = promptForProviderModelSelection_(
      AI_PROVIDER_GEMINI,
      apiKey,
      getProviderModel_(AI_PROVIDER_GEMINI)
    ) || DEFAULT_GEMINI_MODEL;
    saveProviderModel_(AI_PROVIDER_GEMINI, selectedModel);
    ensureActiveAiProviderSet_(AI_PROVIDER_GEMINI);

    appendLogRow_({
      timestamp: refreshAt,
      website_url: '',
      action: 'setup_gemini_ai',
      status: 'SUCCESS',
      pages_discovered: 0,
      pages_written: 0,
      message: 'Gemini API key was saved and validated successfully. Model: ' + selectedModel + '.',
    });

    ui.alert(
      'Gemini AI Connected',
      'The Gemini API key was saved successfully.\nSelected model: ' +
        selectedModel +
        '\n\nYou can now run "Generate AI Meta Descriptions".',
      ui.ButtonSet.OK
    );
  } catch (error) {
    appendLogRow_({
      timestamp: refreshAt,
      website_url: '',
      action: 'setup_gemini_ai',
      status: 'FAILED',
      pages_discovered: 0,
      pages_written: 0,
      message: truncateText_(error && error.message ? error.message : String(error), 500),
    });

    ui.alert(
      'Gemini AI Setup Failed',
      truncateText_(error && error.message ? error.message : String(error), 500),
      ui.ButtonSet.OK
    );
  }
}

function setupOpenAi() {
  ensureRequiredSheets_();

  const ui = SpreadsheetApp.getUi();
  const enteredApiKey = promptForText_(
    'OpenAI API Key',
    'Enter your OpenAI API key.\nIt will be saved once for this Apps Script project and can be used for AI meta descriptions and SEO audit.'
  );

  if (!enteredApiKey) {
    return;
  }

  const apiKey = String(enteredApiKey || '').trim();
  const refreshAt = formatDateTime_(new Date());

  try {
    validateAiProvider_(AI_PROVIDER_OPENAI, apiKey);
    const selectedModel = promptForProviderModelSelection_(
      AI_PROVIDER_OPENAI,
      apiKey,
      getProviderModel_(AI_PROVIDER_OPENAI)
    );

    if (!selectedModel) {
      return;
    }

    PropertiesService.getScriptProperties().setProperty(OPENAI_API_KEY_PROPERTY, apiKey);
    saveProviderModel_(AI_PROVIDER_OPENAI, selectedModel);
    ensureActiveAiProviderSet_(AI_PROVIDER_OPENAI);

    appendLogRow_({
      timestamp: refreshAt,
      website_url: '',
      action: 'setup_openai_ai',
      status: 'SUCCESS',
      pages_discovered: 0,
      pages_written: 0,
      message: 'OpenAI API key was saved and validated successfully. Model: ' + selectedModel + '.',
    });

    ui.alert(
      'OpenAI Connected',
      'The OpenAI API key was saved successfully.\nSelected model: ' +
        selectedModel +
        '\n\nYou can now use OpenAI as the active AI provider.',
      ui.ButtonSet.OK
    );
  } catch (error) {
    appendLogRow_({
      timestamp: refreshAt,
      website_url: '',
      action: 'setup_openai_ai',
      status: 'FAILED',
      pages_discovered: 0,
      pages_written: 0,
      message: truncateText_(error && error.message ? error.message : String(error), 500),
    });

    ui.alert(
      'OpenAI Setup Failed',
      truncateText_(error && error.message ? error.message : String(error), 500),
      ui.ButtonSet.OK
    );
  }
}

function selectGeminiAiModel() {
  ensureRequiredSheets_();

  const ui = SpreadsheetApp.getUi();
  const apiKey = getProviderApiKey_(AI_PROVIDER_GEMINI);
  if (!apiKey) {
    ui.alert('Gemini AI is not configured yet. Run "Setup Gemini AI" first.');
    return;
  }

  const refreshAt = formatDateTime_(new Date());

  try {
    const selectedModel = promptForProviderModelSelection_(
      AI_PROVIDER_GEMINI,
      apiKey,
      getProviderModel_(AI_PROVIDER_GEMINI)
    );
    if (!selectedModel) {
      return;
    }

    saveProviderModel_(AI_PROVIDER_GEMINI, selectedModel);
    appendLogRow_({
      timestamp: refreshAt,
      website_url: '',
      action: 'select_gemini_ai_model',
      status: 'SUCCESS',
      pages_discovered: 0,
      pages_written: 0,
      message: 'Gemini model updated to ' + selectedModel + '.',
    });

    ui.alert(
      'Gemini Model Updated',
      'The saved Gemini model is now: ' + selectedModel,
      ui.ButtonSet.OK
    );
  } catch (error) {
    appendLogRow_({
      timestamp: refreshAt,
      website_url: '',
      action: 'select_gemini_ai_model',
      status: 'FAILED',
      pages_discovered: 0,
      pages_written: 0,
      message: truncateText_(error && error.message ? error.message : String(error), 500),
    });

    ui.alert(
      'Gemini Model Selection Failed',
      truncateText_(error && error.message ? error.message : String(error), 500),
      ui.ButtonSet.OK
    );
  }
}

function selectOpenAiModel() {
  ensureRequiredSheets_();

  const ui = SpreadsheetApp.getUi();
  const apiKey = getProviderApiKey_(AI_PROVIDER_OPENAI);
  if (!apiKey) {
    ui.alert('OpenAI is not configured yet. Run "Setup OpenAI" first.');
    return;
  }

  const refreshAt = formatDateTime_(new Date());

  try {
    const selectedModel = promptForProviderModelSelection_(
      AI_PROVIDER_OPENAI,
      apiKey,
      getProviderModel_(AI_PROVIDER_OPENAI)
    );
    if (!selectedModel) {
      return;
    }

    saveProviderModel_(AI_PROVIDER_OPENAI, selectedModel);
    appendLogRow_({
      timestamp: refreshAt,
      website_url: '',
      action: 'select_openai_model',
      status: 'SUCCESS',
      pages_discovered: 0,
      pages_written: 0,
      message: 'OpenAI model updated to ' + selectedModel + '.',
    });

    ui.alert(
      'OpenAI Model Updated',
      'The saved OpenAI model is now: ' + selectedModel,
      ui.ButtonSet.OK
    );
  } catch (error) {
    appendLogRow_({
      timestamp: refreshAt,
      website_url: '',
      action: 'select_openai_model',
      status: 'FAILED',
      pages_discovered: 0,
      pages_written: 0,
      message: truncateText_(error && error.message ? error.message : String(error), 500),
    });

    ui.alert(
      'OpenAI Model Selection Failed',
      truncateText_(error && error.message ? error.message : String(error), 500),
      ui.ButtonSet.OK
    );
  }
}

function selectActiveAiProvider() {
  ensureRequiredSheets_();

  const ui = SpreadsheetApp.getUi();
  const configuredProviders = getConfiguredAiProviders_();
  if (!configuredProviders.length) {
    ui.alert('No AI providers are configured yet. Set up Gemini or OpenAI first.');
    return;
  }

  const currentProvider = getStoredActiveAiProvider_();
  const response = ui.prompt(
    'Select Active AI Provider',
    'Enter the exact provider to use for all AI actions.\nCurrent: ' +
      valueOrEmpty_(currentProvider || 'not set') +
      '\nAvailable: ' +
      configuredProviders.join(', '),
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const selectedProvider = String(response.getResponseText() || '').trim().toLowerCase() || configuredProviders[0];
  if (configuredProviders.indexOf(selectedProvider) === -1) {
    ui.alert('The selected provider "' + selectedProvider + '" is not configured.');
    return;
  }

  setActiveAiProvider_(selectedProvider);
  ui.alert(
    'Active AI Provider Updated',
    'The active AI provider is now: ' + selectedProvider +
      '\nModel: ' + getProviderModel_(selectedProvider),
    ui.ButtonSet.OK
  );
}

function checkActiveAiConnection() {
  ensureRequiredSheets_();

  const ui = SpreadsheetApp.getUi();
  const refreshAt = formatDateTime_(new Date());
  let connection;

  try {
    connection = ensureActiveAiProviderReady_();
  } catch (error) {
    appendLogRow_({
      timestamp: refreshAt,
      website_url: '',
      action: 'check_active_ai_connection',
      status: 'WARNING',
      pages_discovered: 0,
      pages_written: 0,
      message: truncateText_(error && error.message ? error.message : String(error), 500),
    });

    ui.alert(
      'AI Provider Not Ready',
      truncateText_(error && error.message ? error.message : String(error), 500),
      ui.ButtonSet.OK
    );
    return;
  }

  try {
    validateAiProvider_(connection.provider, connection.apiKey);
    appendLogRow_({
      timestamp: refreshAt,
      website_url: '',
      action: 'check_active_ai_connection',
      status: 'SUCCESS',
      pages_discovered: 0,
      pages_written: 0,
      message:
        'AI connection test succeeded. Provider: ' +
        connection.provider +
        ', model: ' +
        connection.model +
        '.',
    });

    ui.alert(
      'Active AI Connection Check',
      'The active AI provider is configured correctly.\nProvider: ' +
        connection.provider +
        '\nModel: ' +
        connection.model,
      ui.ButtonSet.OK
    );
  } catch (error) {
    appendLogRow_({
      timestamp: refreshAt,
      website_url: '',
      action: 'check_active_ai_connection',
      status: 'FAILED',
      pages_discovered: 0,
      pages_written: 0,
      message: truncateText_(error && error.message ? error.message : String(error), 500),
    });

    ui.alert(
      'Active AI Connection Failed',
      truncateText_(error && error.message ? error.message : String(error), 500),
      ui.ButtonSet.OK
    );
  }
}

function checkGeminiAiConnection() {
  checkActiveAiConnection();
}

function generateAiMetaDescriptions() {
  ensureRequiredSheets_();

  const ui = SpreadsheetApp.getUi();
  const configs = getConfigRecords_();
  if (!configs.length) {
    ui.alert('No websites are configured yet. Run "Setup Website" first.');
    return;
  }

  const selectedConfig = selectWebsiteConfig_(configs);
  if (!selectedConfig) {
    return;
  }

  const refreshAt = formatDateTime_(new Date());

  try {
    const pageRows = getDetailedPageRowsForWebsite_(selectedConfig.website_url);
    if (!pageRows.length) {
      throw new Error('No pages were found in the Pages tab for ' + selectedConfig.website_url + '. Run "Refresh Website" first.');
    }

    const generationRun = buildAiMetaDescriptionRows_(
      selectedConfig,
      pageRows,
      refreshAt
    );

    writeAiMetaDescriptionRows_(selectedConfig.website_url, generationRun.rows);
    refreshDashboardForWebsite_(selectedConfig.website_url, refreshAt);
    appendLogRow_({
      timestamp: refreshAt,
      website_url: selectedConfig.website_url,
      action: 'generate_ai_meta_descriptions',
      status: generationRun.status,
      pages_discovered: pageRows.length,
      pages_written: generationRun.rows.length,
      message: generationRun.logMessage,
    });

    ui.alert(
      'AI Meta Descriptions Generated',
      generationRun.popupMessage,
      ui.ButtonSet.OK
    );
  } catch (error) {
    appendLogRow_({
      timestamp: refreshAt,
      website_url: selectedConfig.website_url,
      action: 'generate_ai_meta_descriptions',
      status: 'FAILED',
      pages_discovered: 0,
      pages_written: 0,
      message: truncateText_(error && error.message ? error.message : String(error), 500),
    });

    ui.alert(
      'AI Meta Description Generation Failed',
      truncateText_(error && error.message ? error.message : String(error), 500),
      ui.ButtonSet.OK
    );
  }
}

function runAiSeoAudit() {
  ensureRequiredSheets_();

  const ui = SpreadsheetApp.getUi();
  const configs = getConfigRecords_();
  if (!configs.length) {
    ui.alert('No websites are configured yet. Run "Setup Website" first.');
    return;
  }

  const selectedConfig = selectWebsiteConfig_(configs);
  if (!selectedConfig) {
    return;
  }

  const refreshAt = formatDateTime_(new Date());

  try {
    const pageRows = getDetailedPageRowsForWebsite_(selectedConfig.website_url);
    if (!pageRows.length) {
      throw new Error(
        'No pages were found in the Pages tab for ' + selectedConfig.website_url + '. Run "Refresh Website" first.'
      );
    }

    const auditRun = buildAiSeoAuditRows_(
      selectedConfig,
      pageRows,
      refreshAt
    );

    writeAiSeoAuditRows_(selectedConfig.website_url, auditRun.rows);
    refreshDashboardForWebsite_(selectedConfig.website_url, refreshAt);
    appendLogRow_({
      timestamp: refreshAt,
      website_url: selectedConfig.website_url,
      action: 'run_ai_seo_audit',
      status: auditRun.status,
      pages_discovered: pageRows.length,
      pages_written: auditRun.rows.length,
      message: auditRun.logMessage,
    });

    ui.alert(
      'AI SEO Audit Complete',
      auditRun.popupMessage,
      ui.ButtonSet.OK
    );
  } catch (error) {
    appendLogRow_({
      timestamp: refreshAt,
      website_url: selectedConfig.website_url,
      action: 'run_ai_seo_audit',
      status: 'FAILED',
      pages_discovered: 0,
      pages_written: 0,
      message: truncateText_(error && error.message ? error.message : String(error), 500),
    });

    ui.alert(
      'AI SEO Audit Failed',
      truncateText_(error && error.message ? error.message : String(error), 500),
      ui.ButtonSet.OK
    );
  }
}

function viewAiMetaDescriptions() {
  ensureRequiredSheets_();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.AI_META_DESCRIPTIONS);
  spreadsheet.setActiveSheet(sheet);
}

function viewDashboard() {
  ensureRequiredSheets_();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.DASHBOARD);
  spreadsheet.setActiveSheet(sheet);
}

function viewAiSeoAudit() {
  ensureRequiredSheets_();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.AI_SEO_AUDIT);
  spreadsheet.setActiveSheet(sheet);
}

function buildPageKeywordsSheet() {
  ensureRequiredSheets_();

  const ui = SpreadsheetApp.getUi();
  const configs = getConfigRecords_();
  if (!configs.length) {
    ui.alert('No websites are configured yet. Run "Setup Website" first.');
    return;
  }

  const selectedConfig = selectWebsiteConfig_(configs);
  if (!selectedConfig) {
    return;
  }

  const refreshAt = formatDateTime_(new Date());

  try {
    const pageRows = getDetailedPageRowsForWebsite_(selectedConfig.website_url);
    if (!pageRows.length) {
      throw new Error(
        'No pages were found in the Pages tab for ' + selectedConfig.website_url + '. Run "Refresh Website" first.'
      );
    }

    const keywordSheetData = buildPageKeywordSheetRows_(pageRows);
    writePageKeywordSheet_(keywordSheetData);

    appendLogRow_({
      timestamp: refreshAt,
      website_url: selectedConfig.website_url,
      action: 'build_page_keywords_sheet',
      status: 'SUCCESS',
      pages_discovered: pageRows.length,
      pages_written: keywordSheetData.rows.length,
      message:
        'Page Keywords sheet rebuilt with ' +
        pageRows.length +
        ' pages and ' +
        PAGE_KEYWORD_ROWS_PER_PAGE +
        ' keyword slots per page using =' +
        AI_SHEETS_FORMULA_FUNCTION +
        ' formulas.',
    });

    ui.alert(
      'Page Keywords Sheet Built',
      'Website: ' +
        selectedConfig.website_url +
        '\n\nPages included: ' +
        pageRows.length +
        '\nKeyword rows written: ' +
        keywordSheetData.rows.length +
        '\nKeywords per page: ' +
        PAGE_KEYWORD_ROWS_PER_PAGE +
        '\nFormula used: =' +
        AI_SHEETS_FORMULA_FUNCTION +
        '(...)',
      ui.ButtonSet.OK
    );
  } catch (error) {
    appendLogRow_({
      timestamp: refreshAt,
      website_url: selectedConfig.website_url,
      action: 'build_page_keywords_sheet',
      status: 'FAILED',
      pages_discovered: 0,
      pages_written: 0,
      message: truncateText_(error && error.message ? error.message : String(error), 500),
    });

    ui.alert(
      'Page Keywords Sheet Failed',
      truncateText_(error && error.message ? error.message : String(error), 500),
      ui.ButtonSet.OK
    );
  }
}

function viewPageKeywords() {
  ensureRequiredSheets_();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.PAGE_KEYWORDS);
  spreadsheet.setActiveSheet(sheet);
}

function clearAllSheetData() {
  ensureRequiredSheets_();

  const ui = SpreadsheetApp.getUi();
  const confirmation = ui.alert(
    'Clear All Sheet Data',
    'This will clear all cell contents across every sheet in this spreadsheet. Required headers will be restored afterward.\n\nDo you want to continue?',
    ui.ButtonSet.YES_NO
  );

  if (confirmation !== ui.Button.YES) {
    return;
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const refreshAt = formatDateTime_(new Date());

  spreadsheet.getSheets().forEach(function(sheet) {
    sheet.clearContents();
  });

  ensureRequiredSheets_();

  appendLogRow_({
    timestamp: refreshAt,
    website_url: '',
    action: 'clear_all_sheet_data',
    status: 'SUCCESS',
    pages_discovered: 0,
    pages_written: 0,
    message: 'All sheet cell contents were cleared. Required headers were restored.',
  });

  ui.alert(
    'Sheet Data Cleared',
    'All cell contents were cleared across the spreadsheet. The required automation headers have been restored.',
    ui.ButtonSet.OK
  );
}

function viewLogs() {
  ensureRequiredSheets_();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.LOGS);
  spreadsheet.setActiveSheet(sheet);
}

function refreshDashboard() {
  ensureRequiredSheets_();

  const ui = SpreadsheetApp.getUi();
  const configs = getConfigRecords_();
  if (!configs.length) {
    ui.alert('No websites are configured yet. Run "Setup Website" first.');
    return;
  }

  const selectedConfig = selectWebsiteConfig_(configs);
  if (!selectedConfig) {
    return;
  }

  const refreshedAt = formatDateTime_(new Date());

  try {
    const summary = refreshDashboardForWebsite_(selectedConfig.website_url, refreshedAt);

    appendLogRow_({
      timestamp: refreshedAt,
      website_url: selectedConfig.website_url,
      action: 'refresh_dashboard',
      status: 'SUCCESS',
      pages_discovered: summary.pageCount,
      pages_written: summary.rowCount,
      message: 'Dashboard refreshed with ' + summary.rowCount + ' summary rows.',
    });

    ui.alert(
      'Dashboard Refreshed',
      'Website: ' +
        selectedConfig.website_url +
        '\n\nPages summarized: ' +
        summary.pageCount +
        '\nDashboard rows written: ' +
        summary.rowCount +
        '\nHigh-priority quick wins: ' +
        summary.highPriorityCount,
      ui.ButtonSet.OK
    );
  } catch (error) {
    appendLogRow_({
      timestamp: refreshedAt,
      website_url: selectedConfig.website_url,
      action: 'refresh_dashboard',
      status: 'FAILED',
      pages_discovered: 0,
      pages_written: 0,
      message: truncateText_(error && error.message ? error.message : String(error), 500),
    });

    ui.alert(
      'Dashboard Refresh Failed',
      truncateText_(error && error.message ? error.message : String(error), 500),
      ui.ButtonSet.OK
    );
  }
}

function runRefreshWorkflow_(options) {
  ensureRequiredSheets_();

  const ui = SpreadsheetApp.getUi();
  const configs = getConfigRecords_();
  if (!configs.length) {
    ui.alert('No websites are configured yet. Run "Setup Website" first.');
    return;
  }

  const selectedConfig = selectWebsiteConfig_(configs);
  if (!selectedConfig) {
    return;
  }

  const dateRangeDays = options.promptForDateRange
    ? promptForDateRangeDays_(selectedConfig.default_date_range_days || DEFAULT_DATE_RANGE_DAYS)
    : parseDateRangeDays_(selectedConfig.default_date_range_days || DEFAULT_DATE_RANGE_DAYS);

  if (!dateRangeDays) {
    return;
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const snapshot = buildWebsiteSnapshot_(selectedConfig, dateRangeDays);
    writePagesSnapshot_(selectedConfig.website_url, snapshot.rows);
    const refreshAt = formatDateTime_(new Date());
    updateConfigRefreshState_(
      selectedConfig.website_url,
      refreshAt,
      'SUCCESS'
    );
    refreshDashboardForWebsite_(selectedConfig.website_url, refreshAt);
    appendLogRow_({
      timestamp: refreshAt,
      website_url: selectedConfig.website_url,
      action: options.actionName,
      status: 'SUCCESS',
      pages_discovered: snapshot.discoveredCount,
      pages_written: snapshot.rows.length,
      message: snapshot.runMessage,
    });

    ui.alert(
      'Refresh Complete',
      snapshot.completionMessage +
        '\n\nFetched ' +
        snapshot.rows.length +
        ' pages for ' +
        selectedConfig.website_url +
        ' using a ' +
        dateRangeDays +
        '-day range.',
      ui.ButtonSet.OK
    );
  } catch (error) {
    updateConfigRefreshState_(
      selectedConfig.website_url,
      formatDateTime_(new Date()),
      'FAILED'
    );
    appendLogRow_({
      timestamp: formatDateTime_(new Date()),
      website_url: selectedConfig.website_url,
      action: options.actionName,
      status: 'FAILED',
      pages_discovered: 0,
      pages_written: 0,
      message: truncateText_(error && error.message ? error.message : String(error), 500),
    });

    ui.alert(
      'Refresh Failed',
      truncateText_(error && error.message ? error.message : String(error), 500),
      ui.ButtonSet.OK
    );
  } finally {
    lock.releaseLock();
  }
}

function buildWebsiteSnapshot_(config, dateRangeDays) {
  const dateRange = getDateRange_(dateRangeDays);
  const discovery = discoverPages_(config.website_url);
  if (!discovery.pages.length) {
    throw new Error('No pages were discovered for ' + config.website_url + '.');
  }

  const audits = auditPages_(discovery.pages);
  const canonicalizedPages = buildCanonicalizedPageEntries_(
    discovery.pages,
    audits,
    config.host
  );
  const preservedGoogleStatusMap = getExistingGoogleStatusMap_(config.website_url);
  const enrichment = buildGoogleEnrichment_(
    config,
    dateRange.startDate,
    dateRange.endDate
  );
  const refreshedAt = formatDateTime_(new Date());

  const rows = canonicalizedPages
    .map(function(pageEntry) {
      const pageUrl = normalizePageUrl_(pageEntry.page_url);
      const pagePath = extractPathFromUrl_(pageUrl);
      const audit =
        audits[pageEntry.audit_key] ||
        audits[pageUrl] ||
        {};
      const comparisonUrl = audit.canonical_url
        ? normalizeComparableUrl_(audit.canonical_url)
        : normalizeComparableUrl_(pageUrl);
      const searchMetric =
        enrichment.searchConsole.metrics[comparisonUrl] ||
        enrichment.searchConsole.metrics[normalizeComparableUrl_(pageUrl)] ||
        createEmptySearchConsoleMetrics_();
      const gaMetric =
        enrichment.ga4.metrics[normalizePathKey_(pagePath)] || createEmptyGa4Metrics_();
      const preservedGoogleStatus =
        preservedGoogleStatusMap[normalizeComparableUrl_(pageUrl)] ||
        preservedGoogleStatusMap[normalizeComparableUrl_(pageEntry.audit_key)] ||
        createEmptyGoogleInspectionFields_();

      return [
        config.website_url,
        config.host,
        pageUrl,
        pagePath,
        pageEntry.source_method,
        valueOrEmpty_(audit.status_code),
        valueOrEmpty_(audit.title),
        valueOrEmpty_(audit.meta_description),
        valueOrEmpty_(audit.h1),
        valueOrEmpty_(audit.canonical_url),
        valueOrEmpty_(audit.robots_meta),
        String(audit.is_indexable === true),
        searchMetric.clicks,
        searchMetric.impressions,
        searchMetric.ctr,
        searchMetric.avg_position,
        gaMetric.sessions,
        gaMetric.users,
        gaMetric.conversions,
        dateRange.startDate,
        dateRange.endDate,
        refreshedAt,
      ].concat(googleInspectionFieldsToRow_(preservedGoogleStatus));
    })
    .sort(function(left, right) {
      return left[2].localeCompare(right[2]);
    });

  return {
    rows: rows,
    discoveredCount: canonicalizedPages.length,
    runMessage: buildRunMessage_(discovery.message, enrichment.statusMessages),
    completionMessage: buildCompletionMessage_(enrichment.statusMessages),
  };
}

function buildGoogleEnrichment_(config, startDate, endDate) {
  const searchConsole = fetchOptionalSearchConsoleMetrics_(
    config,
    startDate,
    endDate
  );
  const ga4 = fetchOptionalGa4Metrics_(
    config.ga4_property_id,
    startDate,
    endDate
  );

  return {
    searchConsole: searchConsole,
    ga4: ga4,
    statusMessages: [searchConsole.statusMessage, ga4.statusMessage],
  };
}

function fetchOptionalSearchConsoleMetrics_(config, startDate, endDate) {
  const propertyCandidates = getSearchConsolePropertyCandidates_(config);
  if (!propertyCandidates.length) {
    return {
      metrics: {},
      statusMessage: 'Search Console skipped: no property candidates were available.',
      included: false,
    };
  }

  const errors = [];
  for (let i = 0; i < propertyCandidates.length; i += 1) {
    const propertyIdentifier = propertyCandidates[i];

    try {
      const metrics = fetchSearchConsoleMetrics_(propertyIdentifier, startDate, endDate);
      const metricCount = Object.keys(metrics).length;
      return {
        metrics: metrics,
        statusMessage:
          'Search Console metrics included using ' +
          propertyIdentifier +
          (metricCount ? '.' : ', but no page metrics were returned for the selected range.'),
        included: true,
      };
    } catch (error) {
      errors.push(
        propertyIdentifier +
          ': ' +
          truncateText_(error && error.message ? error.message : String(error), 160)
      );
    }
  }

  return {
    metrics: {},
    statusMessage: 'Search Console skipped: ' + errors.join(' | '),
    included: false,
  };
}

function getSearchConsolePropertyCandidates_(config) {
  const candidates = [];
  const seen = {};

  function addCandidate(value) {
    const trimmedValue = String(value || '').trim();
    if (!trimmedValue || seen[trimmedValue]) {
      return;
    }
    seen[trimmedValue] = true;
    candidates.push(trimmedValue);
  }

  const configuredProperty = normalizeSearchConsoleProperty_(config.search_console_property);
  if (configuredProperty) {
    addCandidate(configuredProperty);
  }

  if (config.website_url) {
    addCandidate(normalizeWebsiteUrl_(config.website_url));
  }

  if (config.host) {
    addCandidate('sc-domain:' + config.host);
  }

  return candidates;
}

function normalizeSearchConsoleProperty_(value) {
  const trimmedValue = String(value || '').trim();
  if (!trimmedValue) {
    return '';
  }

  if (/^sc-domain:/i.test(trimmedValue)) {
    return 'sc-domain:' + trimmedValue.replace(/^sc-domain:/i, '').trim().toLowerCase();
  }

  if (/^https?:\/\//i.test(trimmedValue)) {
    return normalizeWebsiteUrl_(trimmedValue);
  }

  return trimmedValue;
}

function fetchOptionalGa4Metrics_(propertyId, startDate, endDate) {
  if (!String(propertyId || '').trim()) {
    return {
      metrics: {},
      statusMessage: 'GA4 skipped: not configured.',
      included: false,
    };
  }

  try {
    return {
      metrics: fetchGa4Metrics_(propertyId, startDate, endDate),
      statusMessage: 'GA4 metrics included.',
      included: true,
    };
  } catch (error) {
    return {
      metrics: {},
      statusMessage:
        'GA4 skipped: ' +
        truncateText_(error && error.message ? error.message : String(error), 220),
      included: false,
    };
  }
}

function buildRunMessage_(discoveryMessage, enrichmentMessages) {
  return [discoveryMessage].concat(enrichmentMessages).filter(Boolean).join(' ');
}

function buildCompletionMessage_(enrichmentMessages) {
  const combined = enrichmentMessages.join(' ');
  if (/skipped/i.test(combined)) {
    return 'Audit completed. Some Google metrics were skipped.';
  }

  return 'Audit completed with Google metrics included.';
}

function buildCanonicalizedPageEntries_(pageEntries, audits, allowedHost) {
  const entriesByComparableUrl = {};

  pageEntries.forEach(function(pageEntry) {
    const originalUrl = normalizePageUrl_(pageEntry.page_url);
    const audit = audits[originalUrl] || {};
    let preferredUrl = originalUrl;

    if (audit.canonical_url) {
      const canonicalUrl = tryNormalizePageUrl_(audit.canonical_url);
      if (canonicalUrl && getHost_(canonicalUrl) === allowedHost) {
        preferredUrl = canonicalUrl;
      }
    }

    if (isLikelyBinaryAsset_(preferredUrl)) {
      return;
    }

    const comparableUrl = normalizeComparableUrl_(preferredUrl);
    const existingEntry = entriesByComparableUrl[comparableUrl];

    if (
      !existingEntry ||
      getSourceMethodPriority_(pageEntry.source_method) > getSourceMethodPriority_(existingEntry.source_method)
    ) {
      entriesByComparableUrl[comparableUrl] = {
        page_url: preferredUrl,
        source_method: pageEntry.source_method,
        audit_key: originalUrl,
      };
    }
  });

  return Object.keys(entriesByComparableUrl)
    .sort()
    .map(function(key) {
      return entriesByComparableUrl[key];
    });
}

function inspectWebsiteGoogleIndexStatus_(config, pageRows, inspectedAt) {
  const propertyResolution = resolveWorkingInspectionProperty_(config);
  if (!propertyResolution.propertyIdentifier) {
    const blockedStatus = propertyResolution.status || 'BLOCKED';
    const blockedResults = createUniformInspectionResultMap_(
      pageRows,
      blockedStatus,
      inspectedAt
    );

    return {
      resultsByUrl: blockedResults,
      updatedCount: pageRows.length,
      status: 'WARNING',
      logMessage: propertyResolution.message,
      popupMessage:
        'Website: ' +
        config.website_url +
        '\n\nGoogle inspection could not run.\n' +
        propertyResolution.message,
    };
  }

  const resultsByUrl = {};
  const summary = {
    inspected: 0,
    noResult: 0,
    errors: 0,
  };

  pageRows.forEach(function(pageRow) {
    const pageUrl = normalizePageUrl_(pageRow.page_url);

    try {
      const apiResponse = fetchUrlInspectionResult_(
        pageUrl,
        propertyResolution.propertyIdentifier
      );
      const inspectionFields = mapUrlInspectionResultToFields_(apiResponse, inspectedAt);
      resultsByUrl[normalizeComparableUrl_(pageUrl)] = inspectionFields;

      if (inspectionFields.google_inspection_status === 'INSPECTED') {
        summary.inspected += 1;
      } else {
        summary.noResult += 1;
      }
    } catch (error) {
      resultsByUrl[normalizeComparableUrl_(pageUrl)] = buildInspectionFailureFields_(
        'ERROR',
        inspectedAt
      );
      summary.errors += 1;
    }
  });

  const status = summary.errors ? 'WARNING' : 'SUCCESS';
  const logMessage =
    'Used property ' +
    propertyResolution.propertyIdentifier +
    '. Inspected: ' +
    summary.inspected +
    ', no result: ' +
    summary.noResult +
    ', errors: ' +
    summary.errors +
    '.';

  const popupMessage =
    'Website: ' +
    config.website_url +
    '\n\nGoogle property used: ' +
    propertyResolution.propertyIdentifier +
    '\nInspected rows: ' +
    summary.inspected +
    '\nNo result rows: ' +
    summary.noResult +
    '\nError rows: ' +
    summary.errors +
    '\n\nThis checks Google index status only and does not force reindexing.';

  return {
    resultsByUrl: resultsByUrl,
    updatedCount: pageRows.length,
    status: status,
    logMessage: logMessage,
    popupMessage: popupMessage,
  };
}

function resolveWorkingInspectionProperty_(config) {
  const propertyCandidates = getSearchConsolePropertyCandidates_(config);
  if (!propertyCandidates.length) {
    return {
      propertyIdentifier: '',
      status: 'NOT_CONFIGURED',
      message: 'No Search Console property candidates are available for this website.',
    };
  }

  const errors = [];
  for (let i = 0; i < propertyCandidates.length; i += 1) {
    const propertyIdentifier = propertyCandidates[i];

    try {
      fetchUrlInspectionResult_(config.website_url, propertyIdentifier);
      return {
        propertyIdentifier: propertyIdentifier,
        status: 'READY',
        message: 'Google inspection is available using ' + propertyIdentifier + '.',
      };
    } catch (error) {
      errors.push(
        propertyIdentifier +
          ': ' +
          truncateText_(error && error.message ? error.message : String(error), 180)
      );
    }
  }

  return {
    propertyIdentifier: '',
    status: 'BLOCKED',
    message: 'Google inspection is blocked. ' + errors.join(' | '),
  };
}

function fetchUrlInspectionResult_(inspectionUrl, siteUrl) {
  return callGoogleApi_(URL_INSPECTION_ENDPOINT, 'post', {
    inspectionUrl: inspectionUrl,
    siteUrl: siteUrl,
    languageCode: 'en-US',
  });
}

function mapUrlInspectionResultToFields_(apiResponse, inspectedAt) {
  const inspectionResult = apiResponse.inspectionResult || {};
  const indexStatusResult = inspectionResult.indexStatusResult || {};

  if (!Object.keys(indexStatusResult).length) {
    return buildInspectionFailureFields_('NO_RESULT', inspectedAt);
  }

  const robotsTxtState = String(indexStatusResult.robotsTxtState || '').toUpperCase();
  let robotsAllowed = '';
  if (robotsTxtState) {
    robotsAllowed = String(robotsTxtState === 'ALLOWED');
  }

  return {
    google_inspection_status: 'INSPECTED',
    google_index_verdict: valueOrEmpty_(indexStatusResult.verdict),
    google_coverage_state: valueOrEmpty_(indexStatusResult.coverageState || indexStatusResult.indexingState),
    google_last_crawl_time: valueOrEmpty_(indexStatusResult.lastCrawlTime),
    google_referring_url:
      indexStatusResult.referringUrls && indexStatusResult.referringUrls.length
        ? indexStatusResult.referringUrls[0]
        : '',
    google_user_canonical: valueOrEmpty_(indexStatusResult.userCanonical),
    google_google_canonical: valueOrEmpty_(indexStatusResult.googleCanonical),
    google_robots_allowed: robotsAllowed,
    google_inspected_at: inspectedAt,
  };
}

function buildInspectionFailureFields_(status, inspectedAt) {
  return {
    google_inspection_status: status,
    google_index_verdict: '',
    google_coverage_state: '',
    google_last_crawl_time: '',
    google_referring_url: '',
    google_user_canonical: '',
    google_google_canonical: '',
    google_robots_allowed: '',
    google_inspected_at: inspectedAt,
  };
}

function createUniformInspectionResultMap_(pageRows, status, inspectedAt) {
  const results = {};

  pageRows.forEach(function(pageRow) {
    results[normalizeComparableUrl_(pageRow.page_url)] = buildInspectionFailureFields_(
      status,
      inspectedAt
    );
  });

  return results;
}

function googleInspectionFieldsToRow_(fields) {
  return GOOGLE_INSPECTION_COLUMNS.map(function(columnName) {
    return valueOrEmpty_(fields[columnName]);
  });
}

function createEmptyGoogleInspectionFields_() {
  return buildInspectionFailureFields_('', '');
}

function getSourceMethodPriority_(sourceMethod) {
  return SOURCE_METHOD_PRIORITY[sourceMethod] || 0;
}

function discoverPages_(websiteUrl) {
  const host = getHost_(websiteUrl);
  const sitemapCandidates = getSitemapCandidates_(websiteUrl);
  const sitemapPages = fetchPagesFromSitemaps_(sitemapCandidates, host);

  if (sitemapPages.length) {
    return {
      pages: sitemapPages.map(function(pageUrl) {
        return {
          page_url: pageUrl,
          source_method: 'sitemap',
        };
      }),
      message: 'Discovered pages from sitemap.',
    };
  }

  const crawledPages = crawlWebsitePages_(websiteUrl, host);
  return {
    pages: crawledPages.map(function(pageUrl) {
      return {
        page_url: pageUrl,
        source_method: 'crawl',
      };
    }),
    message: 'Sitemap unavailable; used homepage crawl fallback.',
  };
}

function fetchPagesFromSitemaps_(candidateUrls, allowedHost) {
  const discoveredPages = {};
  const visitedSitemaps = {};

  candidateUrls.forEach(function(candidateUrl) {
    collectSitemapPages_(candidateUrl, allowedHost, visitedSitemaps, discoveredPages);
  });

  return Object.keys(discoveredPages).sort();
}

function collectSitemapPages_(sitemapUrl, allowedHost, visitedSitemaps, discoveredPages) {
  const normalizedSitemapUrl = normalizeUrlWithProtocol_(sitemapUrl);
  if (visitedSitemaps[normalizedSitemapUrl]) {
    return;
  }
  visitedSitemaps[normalizedSitemapUrl] = true;

  let response;
  try {
    response = fetchWithRetry_(normalizedSitemapUrl, {
      method: 'get',
      muteHttpExceptions: true,
      headers: {
        Accept: 'application/xml,text/xml,application/xhtml+xml,text/html',
      },
    });
  } catch (error) {
    return;
  }

  if (!isSuccessfulResponse_(response.getResponseCode())) {
    return;
  }

  const xml = response.getContentText();
  if (!xml) {
    return;
  }

  const locValues = extractXmlTagValues_(xml, 'loc');
  if (isSitemapIndexXml_(xml)) {
    locValues.forEach(function(childSitemapUrl) {
      const trimmedUrl = String(childSitemapUrl || '').trim();
      if (trimmedUrl) {
        collectSitemapPages_(trimmedUrl, allowedHost, visitedSitemaps, discoveredPages);
      }
    });
    return;
  }

  if (!isUrlSetXml_(xml)) {
    locValues.forEach(function(locValue) {
      const trimmedUrl = String(locValue || '').trim();
      if (/\.xml(?:$|[?#])/i.test(trimmedUrl)) {
        collectSitemapPages_(trimmedUrl, allowedHost, visitedSitemaps, discoveredPages);
      }
    });
    return;
  }

  locValues.forEach(function(pageUrl) {
    const normalizedPageUrl = tryNormalizePageUrl_(pageUrl);
    if (
      normalizedPageUrl &&
      getHost_(normalizedPageUrl) === allowedHost &&
      !isLikelyBinaryAsset_(normalizedPageUrl)
    ) {
      discoveredPages[normalizedPageUrl] = true;
    }
  });
}

function crawlWebsitePages_(websiteUrl, allowedHost) {
  const queue = [];
  const queued = {};
  const visited = {};
  const discovered = {};

  enqueueUrl_(websiteUrl, 0, allowedHost, queue, queued, visited);

  while (queue.length && Object.keys(discovered).length < CRAWL_MAX_PAGES) {
    const current = queue.shift();
    const comparableUrl = normalizeComparableUrl_(current.url);
    delete queued[comparableUrl];
    if (visited[comparableUrl]) {
      continue;
    }
    visited[comparableUrl] = true;

    let response;
    try {
      response = fetchWithRetry_(current.url, {
        method: 'get',
        muteHttpExceptions: true,
        followRedirects: true,
        headers: {
          Accept: 'text/html,application/xhtml+xml',
        },
      });
    } catch (error) {
      continue;
    }

    if (!isSuccessfulResponse_(response.getResponseCode())) {
      continue;
    }

    const pageUrl = tryNormalizePageUrl_(current.url);
    if (!pageUrl || getHost_(pageUrl) !== allowedHost) {
      continue;
    }

    discovered[pageUrl] = true;

    if (current.depth >= CRAWL_MAX_DEPTH) {
      continue;
    }

    extractInternalLinks_(response.getContentText(), pageUrl, allowedHost).forEach(function(linkUrl) {
      enqueueUrl_(linkUrl, current.depth + 1, allowedHost, queue, queued, visited);
    });
  }

  return Object.keys(discovered).sort();
}

function enqueueUrl_(url, depth, allowedHost, queue, queued, visited) {
  const normalizedUrl = tryNormalizePageUrl_(url);
  if (!normalizedUrl || getHost_(normalizedUrl) !== allowedHost || isLikelyBinaryAsset_(normalizedUrl)) {
    return;
  }

  const comparableUrl = normalizeComparableUrl_(normalizedUrl);
  if (visited[comparableUrl] || queued[comparableUrl]) {
    return;
  }

  queued[comparableUrl] = true;
  queue.push({
    url: normalizedUrl,
    depth: depth,
  });
}

function fetchSearchConsoleMetrics_(propertyIdentifier, startDate, endDate) {
  const metrics = {};
  let startRow = 0;

  while (true) {
    const payload = {
      startDate: startDate,
      endDate: endDate,
      dimensions: ['page'],
      rowLimit: SEARCH_CONSOLE_ROW_LIMIT,
      startRow: startRow,
    };
    const apiUrl =
      'https://searchconsole.googleapis.com/webmasters/v3/sites/' +
      encodeURIComponent(propertyIdentifier) +
      '/searchAnalytics/query';
    const response = callGoogleApi_(apiUrl, 'post', payload);
    const rows = response.rows || [];

    rows.forEach(function(row) {
      const pageUrl = row.keys && row.keys.length ? row.keys[0] : '';
      if (!pageUrl) {
        return;
      }

      const comparisonKey = normalizeComparableUrl_(pageUrl);
      metrics[comparisonKey] = {
        clicks: row.clicks || 0,
        impressions: row.impressions || 0,
        ctr: row.ctr || 0,
        avg_position: row.position || 0,
      };
    });

    if (rows.length < SEARCH_CONSOLE_ROW_LIMIT) {
      break;
    }
    startRow += rows.length;
  }

  return metrics;
}

function fetchGa4Metrics_(propertyId, startDate, endDate) {
  const metrics = {};
  let offset = 0;

  while (true) {
    const payload = {
      dateRanges: [{ startDate: startDate, endDate: endDate }],
      dimensions: [{ name: 'pagePathPlusQueryString' }],
      metrics: [
        { name: 'sessions' },
        { name: 'totalUsers' },
        { name: 'conversions' },
      ],
      keepEmptyRows: false,
      limit: String(GA4_ROW_LIMIT),
      offset: String(offset),
    };
    const apiUrl =
      'https://analyticsdata.googleapis.com/v1beta/properties/' +
      encodeURIComponent(propertyId) +
      ':runReport';
    const response = callGoogleApi_(apiUrl, 'post', payload);
    const rows = response.rows || [];

    rows.forEach(function(row) {
      const pathValue =
        row.dimensionValues && row.dimensionValues.length
          ? row.dimensionValues[0].value
          : '';
      if (!pathValue) {
        return;
      }

      metrics[normalizePathKey_(pathValue)] = {
        sessions: parseMetricValue_(row.metricValues, 0),
        users: parseMetricValue_(row.metricValues, 1),
        conversions: parseMetricValue_(row.metricValues, 2),
      };
    });

    if (rows.length < GA4_ROW_LIMIT) {
      break;
    }
    offset += rows.length;
  }

  return metrics;
}

function auditPages_(pageEntries) {
  const results = {};
  const urls = pageEntries.map(function(pageEntry) {
    return pageEntry.page_url;
  });

  for (let i = 0; i < urls.length; i += AUDIT_BATCH_SIZE) {
    const batch = urls.slice(i, i + AUDIT_BATCH_SIZE);
    const requests = batch.map(function(url) {
      return {
        url: url,
        method: 'get',
        muteHttpExceptions: true,
        followRedirects: true,
        headers: {
          Accept: 'text/html,application/xhtml+xml',
        },
      };
    });
    const responses = UrlFetchApp.fetchAll(requests);

    batch.forEach(function(requestedUrl, index) {
      results[normalizePageUrl_(requestedUrl)] = parseAuditResponse_(
        requestedUrl,
        responses[index]
      );
    });
  }

  return results;
}

function parseAuditResponse_(requestedUrl, response) {
  const statusCode = response.getResponseCode();
  const finalUrl = tryNormalizePageUrl_(requestedUrl) || normalizePageUrl_(requestedUrl);
  const html = response.getContentText() || '';
  const headerRobots = getHeaderValue_(response.getAllHeaders(), 'x-robots-tag');
  const metaRobots = extractMetaContent_(html, 'robots');
  const combinedRobots = [metaRobots, headerRobots].filter(Boolean).join(' | ');
  const canonicalUrl = resolveRelativeUrl_(extractCanonicalHref_(html), finalUrl);
  const isIndexable = statusCode >= 200 &&
    statusCode < 400 &&
    !/noindex/i.test(combinedRobots || '');

  return {
    status_code: statusCode,
    title: extractTagText_(html, 'title'),
    meta_description: extractMetaContent_(html, 'description'),
    h1: extractTagText_(html, 'h1'),
    canonical_url: canonicalUrl,
    robots_meta: combinedRobots,
    is_indexable: isIndexable,
  };
}

function writePagesSnapshot_(websiteUrl, rows) {
  const sheet = getOrCreateSheet_(SHEET_NAMES.PAGES, HEADERS.PAGES);
  const existingData = getSheetValues_(sheet);
  const retainedRows = existingData.rows.filter(function(row) {
    return row[0] !== websiteUrl;
  }).map(function(row) {
    return normalizeRowLength_(row, HEADERS.PAGES.length);
  });
  const allRows = retainedRows.concat(rows);

  sheet.clearContents();
  sheet.getRange(1, 1, 1, HEADERS.PAGES.length).setValues([HEADERS.PAGES]);
  if (allRows.length) {
    sheet
      .getRange(2, 1, allRows.length, HEADERS.PAGES.length)
      .setValues(allRows);
  }
  applySheetFormatting_(sheet, HEADERS.PAGES.length);
}

function writeDashboardRows_(websiteUrl, rows) {
  const sheet = getOrCreateSheet_(SHEET_NAMES.DASHBOARD, HEADERS.DASHBOARD);
  const existingData = getSheetValues_(sheet);
  const retainedRows = existingData.rows.filter(function(row) {
    return row[0] !== websiteUrl;
  }).map(function(row) {
    return normalizeRowLength_(row, HEADERS.DASHBOARD.length);
  });
  const allRows = retainedRows.concat(rows);

  sheet.clearContents();
  sheet
    .getRange(1, 1, 1, HEADERS.DASHBOARD.length)
    .setValues([HEADERS.DASHBOARD]);
  if (allRows.length) {
    sheet
      .getRange(2, 1, allRows.length, HEADERS.DASHBOARD.length)
      .setValues(allRows);
  }
  applyDashboardFormatting_(sheet);
}

function writeAiMetaDescriptionRows_(websiteUrl, rows) {
  const sheet = getOrCreateSheet_(SHEET_NAMES.AI_META_DESCRIPTIONS, HEADERS.AI_META_DESCRIPTIONS);
  const existingData = getSheetValues_(sheet);
  const retainedRows = existingData.rows.filter(function(row) {
    return row[0] !== websiteUrl;
  }).map(function(row) {
    return normalizeRowLength_(row, HEADERS.AI_META_DESCRIPTIONS.length);
  });
  const allRows = retainedRows.concat(rows);

  sheet.clearContents();
  sheet
    .getRange(1, 1, 1, HEADERS.AI_META_DESCRIPTIONS.length)
    .setValues([HEADERS.AI_META_DESCRIPTIONS]);
  if (allRows.length) {
    for (let startIndex = 0; startIndex < allRows.length; startIndex += SHEET_WRITE_BATCH_SIZE) {
      const batchRows = allRows.slice(startIndex, startIndex + SHEET_WRITE_BATCH_SIZE);
      sheet
        .getRange(2 + startIndex, 1, batchRows.length, HEADERS.AI_META_DESCRIPTIONS.length)
        .setValues(batchRows);
    }
  }
  applySheetFormatting_(sheet, HEADERS.AI_META_DESCRIPTIONS.length);
}

function writeAiSeoAuditRows_(websiteUrl, rows) {
  const sheet = getOrCreateSheet_(SHEET_NAMES.AI_SEO_AUDIT, HEADERS.AI_SEO_AUDIT);
  const existingData = getSheetValues_(sheet);
  const retainedRows = existingData.rows.filter(function(row) {
    return row[0] !== websiteUrl;
  }).map(function(row) {
    return normalizeRowLength_(row, HEADERS.AI_SEO_AUDIT.length);
  });
  const sortedRows = rows.slice().sort(function(left, right) {
    return Number(right[4] || 0) - Number(left[4] || 0);
  });
  const allRows = retainedRows.concat(sortedRows);

  sheet.clearContents();
  sheet
    .getRange(1, 1, 1, HEADERS.AI_SEO_AUDIT.length)
    .setValues([HEADERS.AI_SEO_AUDIT]);
  if (allRows.length) {
    for (let startIndex = 0; startIndex < allRows.length; startIndex += SHEET_WRITE_BATCH_SIZE) {
      const batchRows = allRows.slice(startIndex, startIndex + SHEET_WRITE_BATCH_SIZE);
      sheet
        .getRange(2 + startIndex, 1, batchRows.length, HEADERS.AI_SEO_AUDIT.length)
        .setValues(batchRows);
    }
  }
  applySheetFormatting_(sheet, HEADERS.AI_SEO_AUDIT.length);
}

function buildPageKeywordSheetRows_(pageRows) {
  const rows = [];
  const mergeRanges = [];
  let currentRowNumber = 2;

  pageRows.forEach(function(pageRow) {
    const pageUrl = normalizePageUrl_(pageRow.page_url);
    const groupStartRow = currentRowNumber;

    for (let slotIndex = 0; slotIndex < PAGE_KEYWORD_ROWS_PER_PAGE; slotIndex += 1) {
      const rowNumber = currentRowNumber + slotIndex;
      rows.push([
        pageUrl,
        buildPageKeywordFormulaForRow_(pageRow, slotIndex + 1, rowNumber, groupStartRow),
      ]);
    }

    mergeRanges.push({
      startRow: groupStartRow,
      rowCount: PAGE_KEYWORD_ROWS_PER_PAGE,
    });
    currentRowNumber += PAGE_KEYWORD_ROWS_PER_PAGE;
  });

  return {
    rows: rows,
    mergeRanges: mergeRanges,
  };
}

function writePageKeywordSheet_(keywordSheetData) {
  const sheet = getOrCreateSheet_(SHEET_NAMES.PAGE_KEYWORDS, HEADERS.PAGE_KEYWORDS);
  const rows = keywordSheetData.rows || [];
  const mergeRanges = keywordSheetData.mergeRanges || [];

  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastColumn = Math.max(sheet.getLastColumn(), HEADERS.PAGE_KEYWORDS.length);
  sheet.getRange(1, 1, lastRow, lastColumn).breakApart();
  sheet.clearContents();
  sheet
    .getRange(1, 1, 1, HEADERS.PAGE_KEYWORDS.length)
    .setValues([HEADERS.PAGE_KEYWORDS]);

  if (rows.length) {
    sheet
      .getRange(2, 1, rows.length, HEADERS.PAGE_KEYWORDS.length)
      .setValues(rows);
  }

  mergeRanges.forEach(function(rangeInfo) {
    if (rangeInfo.rowCount > 1) {
      sheet.getRange(rangeInfo.startRow, 1, rangeInfo.rowCount, 1).merge();
    }
  });

  applyPageKeywordSheetFormatting_(sheet, rows.length);
}

function appendLogRow_(logRecord) {
  const sheet = getOrCreateSheet_(SHEET_NAMES.LOGS, HEADERS.LOGS);
  const row = [
    logRecord.timestamp,
    logRecord.website_url,
    logRecord.action,
    logRecord.status,
    logRecord.pages_discovered,
    logRecord.pages_written,
    logRecord.message,
  ];
  sheet.appendRow(row);
  applySheetFormatting_(sheet, HEADERS.LOGS.length);
}

function getPageRowsForWebsite_(websiteUrl) {
  const sheet = getOrCreateSheet_(SHEET_NAMES.PAGES, HEADERS.PAGES);
  const data = getSheetValues_(sheet);
  const pageUrlIndex = HEADERS.PAGES.indexOf('page_url');

  return data.rows
    .filter(function(row) {
      return row[0] === websiteUrl && String(row[pageUrlIndex] || '').trim();
    })
    .map(function(row) {
      return {
        website_url: row[0],
        page_url: normalizePageUrl_(row[pageUrlIndex]),
      };
    });
}

function getDetailedPageRowsForWebsite_(websiteUrl) {
  const sheet = getOrCreateSheet_(SHEET_NAMES.PAGES, HEADERS.PAGES);
  const data = getSheetValues_(sheet);
  const pageUrlIndex = HEADERS.PAGES.indexOf('page_url');

  return data.rows
    .map(function(row, index) {
      const record = rowToObject_(HEADERS.PAGES, normalizeRowLength_(row, HEADERS.PAGES.length));
      record._sheet_row = index + 2;
      return record;
    })
    .filter(function(record) {
      return record.website_url === websiteUrl && String(record.page_url || '').trim();
    })
    .sort(function(left, right) {
      return String(left[HEADERS.PAGES[pageUrlIndex]] || left.page_url || '').localeCompare(
        String(right[HEADERS.PAGES[pageUrlIndex]] || right.page_url || '')
      );
    });
}

function getRowsByWebsite_(sheetName, headers, websiteUrl) {
  const sheet = getOrCreateSheet_(sheetName, headers);
  const data = getSheetValues_(sheet);

  return data.rows
    .map(function(row) {
      return rowToObject_(headers, normalizeRowLength_(row, headers.length));
    })
    .filter(function(record) {
      return record.website_url === websiteUrl;
    });
}

function getPageUrlsForWebsite_(websiteUrl) {
  return getPageRowsForWebsite_(websiteUrl).map(function(pageRow) {
    return pageRow.page_url;
  });
}

function getExistingGoogleStatusMap_(websiteUrl) {
  const sheet = getOrCreateSheet_(SHEET_NAMES.PAGES, HEADERS.PAGES);
  const data = getSheetValues_(sheet);
  const pageUrlIndex = HEADERS.PAGES.indexOf('page_url');

  return data.rows.reduce(function(map, row) {
    if (row[0] !== websiteUrl || !String(row[pageUrlIndex] || '').trim()) {
      return map;
    }

    const rowObject = rowToObject_(HEADERS.PAGES, normalizeRowLength_(row, HEADERS.PAGES.length));
    const fields = {};
    GOOGLE_INSPECTION_COLUMNS.forEach(function(columnName) {
      fields[columnName] = valueOrEmpty_(rowObject[columnName]);
    });
    map[normalizeComparableUrl_(rowObject.page_url)] = fields;
    return map;
  }, {});
}

function writeGoogleInspectionResults_(websiteUrl, resultsByUrl) {
  const sheet = getOrCreateSheet_(SHEET_NAMES.PAGES, HEADERS.PAGES);
  const data = getSheetValues_(sheet);

  if (!data.rows.length) {
    return;
  }

  const pageUrlIndex = HEADERS.PAGES.indexOf('page_url');
  const updatedRows = data.rows.map(function(row) {
    const normalizedRow = normalizeRowLength_(row, HEADERS.PAGES.length);
    if (normalizedRow[0] !== websiteUrl || !String(normalizedRow[pageUrlIndex] || '').trim()) {
      return normalizedRow;
    }

    const result =
      resultsByUrl[normalizeComparableUrl_(normalizedRow[pageUrlIndex])] ||
      createEmptyGoogleInspectionFields_();
    const rowObject = rowToObject_(HEADERS.PAGES, normalizedRow);

    GOOGLE_INSPECTION_COLUMNS.forEach(function(columnName) {
      rowObject[columnName] = valueOrEmpty_(result[columnName]);
    });

    return objectToRow_(HEADERS.PAGES, rowObject);
  });

  sheet
    .getRange(2, 1, updatedRows.length, HEADERS.PAGES.length)
    .setValues(updatedRows);
  applySheetFormatting_(sheet, HEADERS.PAGES.length);
}

function refreshDashboardForWebsite_(websiteUrl, refreshedAt) {
  const dashboardData = buildDashboardRows_(websiteUrl, refreshedAt || formatDateTime_(new Date()));
  writeDashboardRows_(websiteUrl, dashboardData.rows);
  return {
    pageCount: dashboardData.pageCount,
    rowCount: dashboardData.rows.length,
    highPriorityCount: dashboardData.highPriorityCount,
  };
}

function buildDashboardRows_(websiteUrl, refreshedAt) {
  const pageRows = getDetailedPageRowsForWebsite_(websiteUrl);
  const config = getConfigRecords_().find(function(record) {
    return record.website_url === websiteUrl;
  }) || {};

  if (!pageRows.length) {
    throw new Error('No pages were found for ' + websiteUrl + '. Run "Refresh Website" first.');
  }

  const metaRows = getRowsByWebsite_(SHEET_NAMES.AI_META_DESCRIPTIONS, HEADERS.AI_META_DESCRIPTIONS, websiteUrl);
  const auditRows = getRowsByWebsite_(SHEET_NAMES.AI_SEO_AUDIT, HEADERS.AI_SEO_AUDIT, websiteUrl);

  const overview = summarizeDashboardOverview_(pageRows);
  const issueSummary = summarizeDashboardIssues_(pageRows);
  const workflowSummary = summarizeDashboardWorkflow_(metaRows, auditRows);
  const quickWins = buildDashboardQuickWins_(pageRows);
  const rows = [];

  function pushRow(section, metric, value, details) {
    rows.push([
      websiteUrl,
      section,
      metric,
      value,
      details,
      refreshedAt,
    ]);
  }

  pushRow('Overview', 'Tracked pages', overview.pageCount, 'Pages currently stored in the Pages tab.');
  pushRow('Overview', 'Indexable pages', overview.indexableCount, 'Pages currently marked indexable from crawl and robots signals.');
  pushRow('Overview', 'Non-indexable pages', overview.nonIndexableCount, 'Pages that may need indexability review.');
  pushRow('Overview', 'Search visibility', overview.totalImpressions, 'Total impressions across tracked pages in the selected date range.');
  pushRow('Overview', 'Search clicks', overview.totalClicks, 'Total clicks across tracked pages.');
  pushRow('Overview', 'Average CTR', overview.averageCtrDisplay, 'Weighted CTR based on clicks and impressions.');
  pushRow('Overview', 'Sessions', overview.totalSessions, 'Total GA4 sessions across tracked pages.');
  pushRow('Overview', 'Users', overview.totalUsers, 'Total GA4 users across tracked pages.');
  pushRow('Overview', 'Conversions', overview.totalConversions, 'Total recorded GA4 conversions across tracked pages.');
  pushRow('Overview', 'Last refresh', valueOrEmpty_(config.last_refresh_at), 'Latest website refresh timestamp from Config.');
  pushRow('Overview', 'Refresh status', valueOrEmpty_(config.last_refresh_status), 'Latest refresh status from Config.');

  pushRow('Issues', 'Missing meta descriptions', issueSummary.missingMeta, 'Pages missing search snippet descriptions.');
  pushRow('Issues', 'Missing title tags', issueSummary.missingTitle, 'Pages with empty titles.');
  pushRow('Issues', 'Missing H1 tags', issueSummary.missingH1, 'Pages missing a main heading.');
  pushRow('Issues', 'Meta descriptions too long', issueSummary.longMeta, 'Descriptions above the configured length limit.');
  pushRow('Issues', 'Title tags too short', issueSummary.shortTitle, 'Titles shorter than the recommended minimum.');
  pushRow('Issues', 'Title tags too long', issueSummary.longTitle, 'Titles longer than the recommended maximum.');
  pushRow('Issues', 'Canonical mismatches', issueSummary.canonicalMismatch, 'Pages whose canonical points to a different URL.');
  pushRow('Issues', 'Google inspection issues', issueSummary.googleInspectionIssues, 'Pages not returning an INSPECTED status.');
  pushRow('Issues', 'Low CTR with impressions', issueSummary.lowCtrHighImpressions, 'Pages getting impressions but weak CTR.');

  pushRow('Workflow', 'AI meta generated', workflowSummary.meta.generated, 'Rows with generated AI meta descriptions.');
  pushRow('Workflow', 'AI meta failed', workflowSummary.meta.failed, 'Rows that need AI meta retry or review.');
  pushRow('Workflow', 'AI SEO audit generated', workflowSummary.audit.generated, 'Rows with generated AI SEO recommendations.');
  pushRow('Workflow', 'AI SEO audit failed', workflowSummary.audit.failed, 'Rows that need AI audit retry or review.');
  pushRow('Workflow', 'High-priority audit rows', workflowSummary.audit.highPriority, 'Flagged audit rows marked HIGH priority.');

  if (quickWins.length) {
    quickWins.forEach(function(item, index) {
      pushRow(
        'Quick Wins',
        'Priority ' + (index + 1),
        item.page_path,
        'Score ' + item.priorityScore + ' | ' + item.issueSummary + ' | ' + item.searchSignal + ' | ' + item.trafficSignal
      );
    });
  } else {
    pushRow('Quick Wins', 'Priority 1', 'No urgent quick wins', 'No pages currently met the quick-win thresholds.');
  }

  return {
    rows: rows,
    pageCount: pageRows.length,
    highPriorityCount: quickWins.filter(function(item) {
      return item.priorityLabel === 'HIGH';
    }).length,
  };
}

function summarizeDashboardOverview_(pageRows) {
  const totals = pageRows.reduce(function(summary, pageRow) {
    const isIndexable = String(pageRow.is_indexable || '').toLowerCase() === 'true';
    summary.pageCount += 1;
    summary.indexableCount += isIndexable ? 1 : 0;
    summary.nonIndexableCount += isIndexable ? 0 : 1;
    summary.totalClicks += toNumber_(pageRow.clicks);
    summary.totalImpressions += toNumber_(pageRow.impressions);
    summary.totalSessions += toNumber_(pageRow.sessions);
    summary.totalUsers += toNumber_(pageRow.users);
    summary.totalConversions += toNumber_(pageRow.conversions);
    return summary;
  }, {
    pageCount: 0,
    indexableCount: 0,
    nonIndexableCount: 0,
    totalClicks: 0,
    totalImpressions: 0,
    totalSessions: 0,
    totalUsers: 0,
    totalConversions: 0,
  });

  totals.averageCtrDisplay = totals.totalImpressions
    ? (Math.round((totals.totalClicks / totals.totalImpressions) * 10000) / 100) + '%'
    : '0%';

  return totals;
}

function summarizeDashboardIssues_(pageRows) {
  return pageRows.reduce(function(summary, pageRow) {
    const title = sanitizeSingleLineText_(pageRow.title || '');
    const metaDescription = sanitizeSingleLineText_(pageRow.meta_description || '');
    const h1 = sanitizeSingleLineText_(pageRow.h1 || '');
    const canonicalUrl = sanitizeSingleLineText_(pageRow.canonical_url || '');
    const pageUrl = sanitizeSingleLineText_(pageRow.page_url || '');
    const googleInspectionStatus = String(pageRow.google_inspection_status || '').toUpperCase();
    const impressions = toNumber_(pageRow.impressions);
    const ctr = toNumber_(pageRow.ctr);

    if (!metaDescription) {
      summary.missingMeta += 1;
    } else if (metaDescription.length > META_DESCRIPTION_MAX_LENGTH) {
      summary.longMeta += 1;
    }

    if (!title) {
      summary.missingTitle += 1;
    } else {
      if (title.length < TITLE_MIN_LENGTH) {
        summary.shortTitle += 1;
      }
      if (title.length > TITLE_MAX_LENGTH) {
        summary.longTitle += 1;
      }
    }

    if (!h1) {
      summary.missingH1 += 1;
    }

    if (canonicalUrl) {
      try {
        if (normalizeComparableUrl_(canonicalUrl) !== normalizeComparableUrl_(pageUrl)) {
          summary.canonicalMismatch += 1;
        }
      } catch (error) {
        summary.canonicalMismatch += 1;
      }
    }

    if (googleInspectionStatus && googleInspectionStatus !== 'INSPECTED') {
      summary.googleInspectionIssues += 1;
    }

    if (impressions >= HIGH_IMPRESSIONS_THRESHOLD && ctr > 0 && ctr < LOW_CTR_THRESHOLD) {
      summary.lowCtrHighImpressions += 1;
    }

    return summary;
  }, {
    missingMeta: 0,
    missingTitle: 0,
    missingH1: 0,
    longMeta: 0,
    shortTitle: 0,
    longTitle: 0,
    canonicalMismatch: 0,
    googleInspectionIssues: 0,
    lowCtrHighImpressions: 0,
  });
}

function summarizeDashboardWorkflow_(metaRows, auditRows) {
  function summarizeStatuses(rows) {
    return rows.reduce(function(summary, row) {
      const status = String(row.recommendation_status || '').toUpperCase();
      if (status === 'GENERATED') {
        summary.generated += 1;
      } else if (status === 'FAILED') {
        summary.failed += 1;
      } else if (status.indexOf('SKIPPED') === 0) {
        summary.skipped += 1;
      } else if (status === 'NOT_NEEDED') {
        summary.notNeeded += 1;
      }
      return summary;
    }, {
      generated: 0,
      failed: 0,
      skipped: 0,
      notNeeded: 0,
    });
  }

  const auditSummary = summarizeStatuses(auditRows);
  auditSummary.highPriority = auditRows.filter(function(row) {
    return String(row.priority_label || '').toUpperCase() === 'HIGH';
  }).length;

  return {
    meta: summarizeStatuses(metaRows),
    audit: auditSummary,
  };
}

function buildDashboardQuickWins_(pageRows) {
  return pageRows
    .map(function(pageRow) {
      return buildSeoAuditRecord_({ website_url: pageRow.website_url, host: pageRow.host }, pageRow);
    })
    .filter(function(auditRecord) {
      return auditRecord && auditRecord.priorityScore >= 40;
    })
    .sort(function(left, right) {
      return right.priorityScore - left.priorityScore;
    })
    .slice(0, 10);
}

function verifyIndexNowConfiguration_(config) {
  const indexnowKey = String(config.indexnow_key || '').trim();
  const indexnowKeyUrl = String(config.indexnow_key_url || '').trim();

  if (!indexnowKey) {
    throw new Error('IndexNow key is missing. Add it in the Config tab for ' + config.website_url + '.');
  }

  if (!indexnowKeyUrl) {
    throw new Error('IndexNow key URL is missing. Add it in the Config tab for ' + config.website_url + '.');
  }

  let response;
  try {
    response = fetchWithRetry_(indexnowKeyUrl, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
    });
  } catch (error) {
    throw new Error('IndexNow key file could not be reached at ' + indexnowKeyUrl + '.');
  }

  if (!isSuccessfulResponse_(response.getResponseCode())) {
    throw new Error(
      'IndexNow key file returned HTTP ' +
        response.getResponseCode() +
        ' at ' +
        indexnowKeyUrl +
        '.'
    );
  }

  const remoteKey = String(response.getContentText() || '').trim();
  if (remoteKey !== indexnowKey) {
    throw new Error(
      'IndexNow key mismatch. The hosted key file content does not match the key saved in Config.'
    );
  }

  return {
    indexnowKey: indexnowKey,
    indexnowKeyUrl: indexnowKeyUrl,
  };
}

function submitUrlsToIndexNow_(hostUrl, key, keyUrl, urlList) {
  const payload = {
    host: getHost_(hostUrl),
    key: key,
    keyLocation: keyUrl,
    urlList: urlList,
  };
  const response = fetchWithRetry_(INDEXNOW_ENDPOINT, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify(payload),
  });

  if (!isSuccessfulResponse_(response.getResponseCode())) {
    throw buildApiError_(INDEXNOW_ENDPOINT, response);
  }

  return response.getResponseCode();
}

function getConfigRecords_() {
  const sheet = getOrCreateSheet_(SHEET_NAMES.CONFIG, HEADERS.CONFIG);
  const data = getSheetValues_(sheet);

  return data.rows
    .map(function(row) {
      return rowToObject_(HEADERS.CONFIG, row);
    })
    .filter(function(record) {
      return record.website_url;
    });
}

function upsertConfigRecord_(record) {
  const sheet = getOrCreateSheet_(SHEET_NAMES.CONFIG, HEADERS.CONFIG);
  const data = getSheetValues_(sheet);
  const rows = data.rows.slice();
  const index = rows.findIndex(function(row) {
    return row[0] === record.website_url;
  });
  const newRow = objectToRow_(HEADERS.CONFIG, record);

  if (index === -1) {
    rows.push(newRow);
  } else {
    rows[index] = newRow;
  }

  sheet.clearContents();
  sheet.getRange(1, 1, 1, HEADERS.CONFIG.length).setValues([HEADERS.CONFIG]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, HEADERS.CONFIG.length).setValues(rows);
  }
  applySheetFormatting_(sheet, HEADERS.CONFIG.length);
}

function updateConfigRefreshState_(websiteUrl, refreshAt, status) {
  const configs = getConfigRecords_();
  const matchingConfig = configs.find(function(record) {
    return record.website_url === websiteUrl;
  });

  if (!matchingConfig) {
    return;
  }

  matchingConfig.last_refresh_at = refreshAt;
  matchingConfig.last_refresh_status = status;
  upsertConfigRecord_(matchingConfig);
}

function selectWebsiteConfig_(configs) {
  if (configs.length === 1) {
    return configs[0];
  }

  const ui = SpreadsheetApp.getUi();
  const examples = configs
    .map(function(config) {
      return config.host;
    })
    .slice(0, 8)
    .join(', ');

  const response = ui.prompt(
    'Select Website',
    'Enter the exact website URL or host to refresh.\nAvailable: ' + examples,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return null;
  }

  const input = String(response.getResponseText() || '').trim();
  if (!input) {
    ui.alert('Website selection was empty.');
    return null;
  }

  const normalizedInput = tryNormalizeWebsiteUrl_(input);
  const hostInput = normalizedInput ? getHost_(normalizedInput) : sanitizeHost_(input);

  const matchingConfig = configs.find(function(config) {
    return (
      config.website_url === normalizedInput ||
      config.host === hostInput ||
      config.website_url === input
    );
  });

  if (!matchingConfig) {
    ui.alert('No configured website matched "' + input + '".');
    return null;
  }

  return matchingConfig;
}

function ensureRequiredSheets_() {
  getOrCreateSheet_(SHEET_NAMES.PAGES, HEADERS.PAGES);
  getOrCreateSheet_(SHEET_NAMES.CONFIG, HEADERS.CONFIG);
  getOrCreateSheet_(SHEET_NAMES.LOGS, HEADERS.LOGS);
  getOrCreateSheet_(SHEET_NAMES.DASHBOARD, HEADERS.DASHBOARD);
  getOrCreateSheet_(SHEET_NAMES.AI_META_DESCRIPTIONS, HEADERS.AI_META_DESCRIPTIONS);
  getOrCreateSheet_(SHEET_NAMES.AI_SEO_AUDIT, HEADERS.AI_SEO_AUDIT);
  getOrCreateSheet_(SHEET_NAMES.PAGE_KEYWORDS, HEADERS.PAGE_KEYWORDS);
}

function getGeminiApiKey_() {
  return getProviderApiKey_(AI_PROVIDER_GEMINI);
}

function getGeminiModel_() {
  return getProviderModel_(AI_PROVIDER_GEMINI);
}

function getOpenAiApiKey_() {
  return getProviderApiKey_(AI_PROVIDER_OPENAI);
}

function getOpenAiModel_() {
  return getProviderModel_(AI_PROVIDER_OPENAI);
}

function getProviderApiKey_(provider) {
  const propertyName = provider === AI_PROVIDER_OPENAI
    ? OPENAI_API_KEY_PROPERTY
    : GEMINI_API_KEY_PROPERTY;
  return String(PropertiesService.getScriptProperties().getProperty(propertyName) || '').trim();
}

function getProviderModel_(provider) {
  const propertyName = provider === AI_PROVIDER_OPENAI
    ? OPENAI_MODEL_PROPERTY
    : GEMINI_MODEL_PROPERTY;
  const fallbackModel = provider === AI_PROVIDER_OPENAI ? '' : DEFAULT_GEMINI_MODEL;
  return String(PropertiesService.getScriptProperties().getProperty(propertyName) || fallbackModel).trim();
}

function saveProviderModel_(provider, modelId) {
  const propertyName = provider === AI_PROVIDER_OPENAI
    ? OPENAI_MODEL_PROPERTY
    : GEMINI_MODEL_PROPERTY;
  const fallbackModel = provider === AI_PROVIDER_OPENAI ? '' : DEFAULT_GEMINI_MODEL;
  PropertiesService.getScriptProperties().setProperty(
    propertyName,
    String(modelId || '').trim() || fallbackModel
  );
}

function getStoredActiveAiProvider_() {
  return String(
    PropertiesService.getScriptProperties().getProperty(ACTIVE_AI_PROVIDER_PROPERTY) || ''
  ).trim().toLowerCase();
}

function getConfiguredAiProviders_() {
  return [AI_PROVIDER_GEMINI, AI_PROVIDER_OPENAI].filter(function(provider) {
    return Boolean(getProviderApiKey_(provider));
  });
}

function getActiveAiProvider_() {
  const configuredProviders = getConfiguredAiProviders_();
  if (!configuredProviders.length) {
    return '';
  }

  const storedProvider = getStoredActiveAiProvider_();
  if (configuredProviders.indexOf(storedProvider) !== -1) {
    return storedProvider;
  }

  return configuredProviders[0];
}

function setActiveAiProvider_(provider) {
  PropertiesService.getScriptProperties().setProperty(
    ACTIVE_AI_PROVIDER_PROPERTY,
    String(provider || '').trim().toLowerCase()
  );
}

function ensureActiveAiProviderSet_(fallbackProvider) {
  const activeProvider = getActiveAiProvider_();
  if (activeProvider) {
    return activeProvider;
  }

  const configuredProviders = getConfiguredAiProviders_();
  const providerToSet =
    configuredProviders.indexOf(fallbackProvider) !== -1
      ? fallbackProvider
      : configuredProviders[0];
  if (providerToSet) {
    setActiveAiProvider_(providerToSet);
  }

  return providerToSet || '';
}

function getActiveAiModel_() {
  const provider = getActiveAiProvider_();
  if (!provider) {
    return '';
  }

  return getProviderModel_(provider);
}

function ensureActiveAiProviderReady_() {
  const provider = getActiveAiProvider_();
  if (!provider) {
    throw new Error('No AI provider is configured. Set up Gemini or OpenAI first.');
  }

  const apiKey = getProviderApiKey_(provider);
  if (!apiKey) {
    throw new Error(
      'The active AI provider "' + provider + '" does not have an API key configured. Set it up first.'
    );
  }

  const model = getProviderModel_(provider);
  if (!model) {
    throw new Error(
      'The active AI provider "' + provider + '" does not have a saved model. Select a model first.'
    );
  }

  const availableModels = fetchProviderModelOptions_(provider, apiKey);
  const matchingModel = availableModels.find(function(option) {
    return option.modelId === model;
  });
  if (!matchingModel) {
    throw new Error(
      'The saved model "' + model + '" is no longer available for the active provider "' +
        provider +
        '". Reselect the model first.'
    );
  }

  return {
    provider: provider,
    apiKey: apiKey,
    model: model,
  };
}

function validateAiProvider_(provider, apiKey) {
  if (provider === AI_PROVIDER_OPENAI) {
    validateOpenAiApiKey_(apiKey);
    return;
  }

  validateGeminiApiKey_(apiKey);
}

function fetchProviderModelOptions_(provider, apiKey) {
  if (provider === AI_PROVIDER_OPENAI) {
    return fetchOpenAiModelOptions_(apiKey);
  }

  return fetchGeminiModelOptions_(apiKey);
}

function validateGeminiApiKey_(apiKey) {
  if (!/^AIza/i.test(String(apiKey || '').trim())) {
    throw new Error(
      'The value entered does not look like a Gemini API key from Google AI Studio. Gemini API keys usually start with "AIza".'
    );
  }

  fetchGeminiModelOptions_(apiKey);
}

function validateOpenAiApiKey_(apiKey) {
  if (!/^sk-/i.test(String(apiKey || '').trim())) {
    throw new Error(
      'The value entered does not look like an OpenAI API key. OpenAI API keys usually start with "sk-".'
    );
  }

  fetchOpenAiModelOptions_(apiKey);
}

function buildAiMetaDescriptionRows_(config, pageRows, generatedAt) {
  const rows = [];
  const summary = {
    ok: 0,
    formulaRows: 0,
  };
  pageRows.forEach(function(pageRow) {
    const existingMetaDescription = sanitizeSingleLineText_(pageRow.meta_description || '');
    const existingCharCount = existingMetaDescription.length;
    const lengthStatus = getMetaDescriptionLengthStatus_(existingMetaDescription);
    const rowNumber = rows.length + 2;

    if (lengthStatus === 'OK') {
      summary.ok += 1;
    }

    rows.push([
      config.website_url,
      config.host,
      normalizePageUrl_(pageRow.page_url),
      pageRow.page_path || extractPathFromUrl_(pageRow.page_url),
      existingMetaDescription,
      existingCharCount,
      lengthStatus,
      buildAiMetaDescriptionFormulaForRow_(pageRow, rowNumber),
      '=IF(LEN(H' + rowNumber + ')=0,"",LEN(H' + rowNumber + '))',
      '=IF(LEN(H' + rowNumber + ')=0,"PENDING_FORMULA",IF(AND(I' + rowNumber + '>=' + META_DESCRIPTION_MIN_LENGTH + ',I' + rowNumber + '<=' + META_DESCRIPTION_MAX_LENGTH + '),"READY","CHECK_LENGTH"))',
      '',
      generatedAt,
    ]);
    summary.formulaRows += 1;
  });

  const finalStatus = 'SUCCESS';
  const logMessage =
    'Prepared ' +
    pageRows.length +
    ' rows. Existing meta already OK: ' +
    summary.ok +
    ', formula rows written: ' +
    summary.formulaRows +
    '.';
  const popupMessage =
    'Website: ' +
    config.website_url +
    '\n\nRows checked: ' +
    pageRows.length +
    '\nAlready OK: ' +
    summary.ok +
    '\nFormula rows written: ' +
    summary.formulaRows +
    '\nFormula used: =' +
    AI_SHEETS_FORMULA_FUNCTION +
    '(prompt, range)' +
    '\n\nThe ai_recommended_meta_description column now uses sheet formulas. Review ai_char_count and recommendation_status to spot rows that need manual tightening.';

  return {
    rows: rows,
    status: finalStatus,
    logMessage: logMessage,
    popupMessage: popupMessage,
  };
}

function buildAiSeoAuditRows_(config, pageRows, generatedAt) {
  const rows = [];
  const flaggedPages = pageRows
    .map(function(pageRow) {
      return buildSeoAuditRecord_(config, pageRow);
    })
    .filter(function(auditRecord) {
      return auditRecord !== null;
    })
    .sort(function(left, right) {
      return right.priorityScore - left.priorityScore;
    });

  const summary = {
    flagged: flaggedPages.length,
    formulaRows: 0,
  };
  flaggedPages.forEach(function(auditRecord) {
    const rowNumber = rows.length + 2;

    rows.push([
      auditRecord.website_url,
      auditRecord.host,
      auditRecord.page_url,
      auditRecord.page_path,
      auditRecord.priorityScore,
      auditRecord.priorityLabel,
      auditRecord.issueCategory,
      auditRecord.issueCode,
      auditRecord.issueSummary,
      auditRecord.severity,
      auditRecord.trafficSignal,
      auditRecord.searchSignal,
      auditRecord.supportingData,
      buildAiSeoAuditFormulaForRow_(auditRecord, rowNumber),
      '=IF(LEN(N' + rowNumber + ')=0,"PENDING_FORMULA","READY")',
      '',
      generatedAt,
    ]);
    summary.formulaRows += 1;
  });

  const status = 'SUCCESS';
  const logMessage =
    'Pages checked: ' +
    pageRows.length +
    ', flagged: ' +
    summary.flagged +
    ', formula rows written: ' +
    summary.formulaRows +
    '.';
  const popupMessage =
    'Website: ' +
    config.website_url +
    '\n\nPages checked: ' +
    pageRows.length +
    '\nFlagged pages: ' +
    summary.flagged +
    '\nFormula rows written: ' +
    summary.formulaRows +
    '\nFormula used: =' +
    AI_SHEETS_FORMULA_FUNCTION +
    '(prompt, range)' +
    '\n\nThe ai_recommendation column now uses sheet formulas. Review recommendation_status after the sheet finishes calculating.';

  return {
    rows: rows,
    status: status,
    logMessage: logMessage,
    popupMessage: popupMessage,
  };
}

function buildSeoAuditRecord_(config, pageRow) {
  const issues = detectSeoAuditIssues_(pageRow);
  if (!issues.length) {
    return null;
  }

  const primaryIssue = issues.sort(function(left, right) {
    return right.weight - left.weight;
  })[0];
  const priorityScore = calculateSeoAuditPriorityScore_(pageRow, primaryIssue);

  return {
    website_url: config.website_url,
    host: config.host,
    page_url: normalizePageUrl_(pageRow.page_url),
    page_path: valueOrEmpty_(pageRow.page_path || extractPathFromUrl_(pageRow.page_url)),
    priorityScore: priorityScore,
    priorityLabel: getSeoAuditPriorityLabel_(priorityScore),
    issueCategory: primaryIssue.category,
    issueCode: primaryIssue.code,
    issueSummary: primaryIssue.summary,
    severity: primaryIssue.severity,
    trafficSignal: buildTrafficSignal_(pageRow),
    searchSignal: buildSearchSignal_(pageRow),
    supportingData: buildSeoAuditSupportingData_(pageRow, issues),
    pageData: pageRow,
  };
}

function detectSeoAuditIssues_(pageRow) {
  const issues = [];
  const title = sanitizeSingleLineText_(pageRow.title || '');
  const metaDescription = sanitizeSingleLineText_(pageRow.meta_description || '');
  const h1 = sanitizeSingleLineText_(pageRow.h1 || '');
  const canonicalUrl = sanitizeSingleLineText_(pageRow.canonical_url || '');
  const robotsMeta = sanitizeSingleLineText_(pageRow.robots_meta || '');
  const isIndexable = String(pageRow.is_indexable || '').toLowerCase() === 'true';
  const googleInspectionStatus = String(pageRow.google_inspection_status || '').toUpperCase();
  const impressions = toNumber_(pageRow.impressions);
  const ctr = toNumber_(pageRow.ctr);
  const sessions = toNumber_(pageRow.sessions);
  const users = toNumber_(pageRow.users);

  if (!metaDescription) {
    issues.push(createSeoAuditIssue_('metadata', 'meta_missing', 'Missing meta description reduces snippet quality and click-through opportunity.', 'HIGH', 60));
  } else if (metaDescription.length > META_DESCRIPTION_MAX_LENGTH) {
    issues.push(createSeoAuditIssue_('metadata', 'meta_too_long', 'Meta description is too long and may be truncated in search results.', 'MEDIUM', 38));
  }

  if (!title) {
    issues.push(createSeoAuditIssue_('metadata', 'title_missing', 'Missing title tag weakens relevance and search result presentation.', 'HIGH', 58));
  } else {
    if (title.length < TITLE_MIN_LENGTH) {
      issues.push(createSeoAuditIssue_('metadata', 'title_too_short', 'Title tag is very short and may not communicate enough context to search users.', 'MEDIUM', 28));
    }
    if (title.length > TITLE_MAX_LENGTH) {
      issues.push(createSeoAuditIssue_('metadata', 'title_too_long', 'Title tag is long and may be truncated in search results.', 'MEDIUM', 32));
    }
  }

  if (!h1) {
    issues.push(createSeoAuditIssue_('content', 'h1_missing', 'Missing H1 weakens page clarity and topical alignment.', 'MEDIUM', 25));
  }

  if (!isIndexable || /noindex/i.test(robotsMeta)) {
    issues.push(createSeoAuditIssue_('indexability', 'not_indexable', 'Page is marked non-indexable or blocked from indexing signals.', 'CRITICAL', 72));
  }

  if (googleInspectionStatus && googleInspectionStatus !== 'INSPECTED') {
    issues.push(createSeoAuditIssue_('indexability', 'google_inspection_issue', 'Google inspection reported a non-ready status for this page.', 'HIGH', 55));
  }

  if (impressions >= HIGH_IMPRESSIONS_THRESHOLD && ctr > 0 && ctr < LOW_CTR_THRESHOLD) {
    issues.push(createSeoAuditIssue_('performance', 'low_ctr_high_impressions', 'Page has meaningful impressions but a weak CTR, suggesting snippet underperformance.', 'HIGH', 50));
  }

  if ((sessions >= HIGH_SESSIONS_THRESHOLD || users >= HIGH_USERS_THRESHOLD) && (!metaDescription || !title)) {
    issues.push(createSeoAuditIssue_('performance', 'important_page_weak_metadata', 'A traffic-relevant page has weak or missing search metadata.', 'HIGH', 52));
  }

  if (canonicalUrl) {
    const normalizedPageUrl = normalizeComparableUrl_(pageRow.page_url);
    let normalizedCanonical = '';
    try {
      normalizedCanonical = normalizeComparableUrl_(canonicalUrl);
    } catch (error) {
      normalizedCanonical = '';
    }
    if (normalizedCanonical && normalizedCanonical !== normalizedPageUrl) {
      issues.push(createSeoAuditIssue_('canonical', 'canonical_mismatch', 'Canonical URL points to a different page and may conflict with expected indexing.', 'HIGH', 48));
    }
  }

  return issues;
}

function createSeoAuditIssue_(category, code, summary, severity, weight) {
  return {
    category: category,
    code: code,
    summary: summary,
    severity: severity,
    weight: weight,
  };
}

function calculateSeoAuditPriorityScore_(pageRow, primaryIssue) {
  let score = Number(primaryIssue.weight || 0);
  const impressions = toNumber_(pageRow.impressions);
  const clicks = toNumber_(pageRow.clicks);
  const ctr = toNumber_(pageRow.ctr);
  const avgPosition = toNumber_(pageRow.avg_position);
  const sessions = toNumber_(pageRow.sessions);
  const users = toNumber_(pageRow.users);
  const conversions = toNumber_(pageRow.conversions);

  if (impressions >= 1000) {
    score += 18;
  } else if (impressions >= 300) {
    score += 12;
  } else if (impressions >= HIGH_IMPRESSIONS_THRESHOLD) {
    score += 6;
  }

  if (clicks >= 50) {
    score += 8;
  } else if (clicks >= 10) {
    score += 4;
  }

  if (ctr > 0 && ctr < LOW_CTR_THRESHOLD) {
    score += 12;
  }

  if (avgPosition >= 1 && avgPosition <= 20) {
    score += 6;
  }

  if (sessions >= 300) {
    score += 14;
  } else if (sessions >= HIGH_SESSIONS_THRESHOLD) {
    score += 8;
  }

  if (users >= 150) {
    score += 10;
  } else if (users >= HIGH_USERS_THRESHOLD) {
    score += 5;
  }

  if (conversions >= 5) {
    score += 10;
  } else if (conversions > 0) {
    score += 4;
  }

  return Math.min(100, Math.round(score));
}

function getSeoAuditPriorityLabel_(priorityScore) {
  if (priorityScore >= 70) {
    return 'HIGH';
  }
  if (priorityScore >= 40) {
    return 'MEDIUM';
  }
  return 'LOW';
}

function buildTrafficSignal_(pageRow) {
  return 'sessions=' + toNumber_(pageRow.sessions) +
    ', users=' + toNumber_(pageRow.users) +
    ', conversions=' + toNumber_(pageRow.conversions);
}

function buildSearchSignal_(pageRow) {
  return 'impressions=' + toNumber_(pageRow.impressions) +
    ', clicks=' + toNumber_(pageRow.clicks) +
    ', ctr=' + toNumber_(pageRow.ctr) +
    ', avg_position=' + toNumber_(pageRow.avg_position);
}

function buildSeoAuditSupportingData_(pageRow, issues) {
  const issueCodes = issues.map(function(issue) {
    return issue.code;
  }).join(', ');
  const googleStatus = String(pageRow.google_inspection_status || '').trim() || 'none';
  const robotsMeta = sanitizeSingleLineText_(pageRow.robots_meta || '') || 'none';
  const canonicalUrl = sanitizeSingleLineText_(pageRow.canonical_url || '') || 'none';

  return truncateText_(
    'issues=' +
      issueCodes +
      '; title_length=' +
      sanitizeSingleLineText_(pageRow.title || '').length +
      '; meta_length=' +
      sanitizeSingleLineText_(pageRow.meta_description || '').length +
      '; h1_present=' +
      String(Boolean(sanitizeSingleLineText_(pageRow.h1 || ''))) +
      '; indexable=' +
      String(String(pageRow.is_indexable || '').toLowerCase() === 'true') +
      '; google_status=' +
      googleStatus +
      '; robots=' +
      robotsMeta +
      '; canonical=' +
      canonicalUrl,
    500
  );
}

function getMetaDescriptionLengthStatus_(metaDescription) {
  if (!metaDescription) {
    return 'EMPTY';
  }

  if (metaDescription.length > META_DESCRIPTION_MAX_LENGTH) {
    return 'TOO_LONG';
  }

  return 'OK';
}

function generateRecommendedMetaDescription_(pageRow, aiConnection) {
  const prompt = buildGeminiMetaDescriptionPrompt_(pageRow);
  const cacheKey = buildAiTextCacheKey_(
    'meta',
    aiConnection.provider,
    aiConnection.model,
    prompt
  );
  const cache = CacheService.getScriptCache();
  const cachedResponse = cache.get(cacheKey);
  if (cachedResponse) {
    return {
      text: sanitizeGeminiMetaDescription_(cachedResponse),
      usedApiCall: false,
    };
  }

  const responseText = callAiGenerateText_(aiConnection.provider, aiConnection.apiKey, prompt);
  const sanitizedText = sanitizeGeminiMetaDescription_(responseText);

  if (!sanitizedText) {
    throw new Error('The AI provider returned an empty recommendation.');
  }

  cache.put(cacheKey, sanitizedText, GEMINI_RESPONSE_CACHE_TTL_SECONDS);
  return {
    text: sanitizedText,
    usedApiCall: true,
  };
}

function generateAiSeoAuditRecommendation_(auditRecord, aiConnection) {
  const prompt = buildAiSeoAuditPrompt_(auditRecord);
  const cacheKey = buildAiTextCacheKey_(
    'audit',
    aiConnection.provider,
    aiConnection.model,
    prompt
  );
  const cache = CacheService.getScriptCache();
  const cachedResponse = cache.get(cacheKey);
  if (cachedResponse) {
    return {
      text: sanitizeSingleLineText_(cachedResponse),
      usedApiCall: false,
    };
  }

  const responseText = callAiGenerateText_(aiConnection.provider, aiConnection.apiKey, prompt);
  const sanitizedText = sanitizeSingleLineText_(responseText);

  if (!sanitizedText) {
    throw new Error('The AI provider returned an empty SEO audit recommendation.');
  }

  cache.put(cacheKey, sanitizedText, GEMINI_RESPONSE_CACHE_TTL_SECONDS);
  return {
    text: sanitizedText,
    usedApiCall: true,
  };
}

function buildGeminiMetaDescriptionPrompt_(pageRow) {
  return [
    'You are rewriting a website meta description for SEO.',
    'Create exactly one meta description for the page using only the information provided below.',
    'Do not invent facts, offers, pricing, guarantees, locations, features, or claims that are not supported by the page URL, title, H1, or current meta description.',
    'Preserve the page intent and topic. If the source details are limited, stay generic and conservative.',
    'Return plain text only.',
    'Do not use quotation marks, bullets, labels, markdown, emojis, or multiple options.',
    'Write natural, clickworthy copy suitable for a search result snippet.',
    'Keep the final meta description at or below ' + META_DESCRIPTION_MAX_LENGTH + ' characters.',
    '',
    'Page URL: ' + valueOrEmpty_(pageRow.page_url),
    'Page path: ' + valueOrEmpty_(pageRow.page_path),
    'Title: ' + valueOrEmpty_(pageRow.title),
    'H1: ' + valueOrEmpty_(pageRow.h1),
    'Current meta description: ' + valueOrEmpty_(pageRow.meta_description),
  ].join('\n');
}

function buildAiSeoAuditPrompt_(auditRecord) {
  const pageRow = auditRecord.pageData;

  return [
    'You are an SEO consultant writing one concise audit recommendation for a single web page.',
    'Use only the provided page data and issue summary.',
    'Do not invent facts, services, offers, locations, pricing, or page content that is not supported by the supplied fields.',
    'Return plain text only as 2 short sentences maximum.',
    'Sentence 1 should explain the issue briefly.',
    'Sentence 2 should give the most practical next action.',
    'Keep the wording concise and spreadsheet-friendly.',
    '',
    'Primary issue category: ' + auditRecord.issueCategory,
    'Primary issue code: ' + auditRecord.issueCode,
    'Primary issue summary: ' + auditRecord.issueSummary,
    'Priority label: ' + auditRecord.priorityLabel,
    'Priority score: ' + auditRecord.priorityScore,
    'Page URL: ' + valueOrEmpty_(auditRecord.page_url),
    'Page path: ' + valueOrEmpty_(auditRecord.page_path),
    'Title: ' + valueOrEmpty_(pageRow.title),
    'Meta description: ' + valueOrEmpty_(pageRow.meta_description),
    'H1: ' + valueOrEmpty_(pageRow.h1),
    'Canonical URL: ' + valueOrEmpty_(pageRow.canonical_url),
    'Robots meta: ' + valueOrEmpty_(pageRow.robots_meta),
    'Is indexable: ' + valueOrEmpty_(pageRow.is_indexable),
    'Search metrics: ' + auditRecord.searchSignal,
    'Traffic metrics: ' + auditRecord.trafficSignal,
    'Google inspection status: ' + valueOrEmpty_(pageRow.google_inspection_status),
    'Supporting data: ' + valueOrEmpty_(auditRecord.supportingData),
  ].join('\n');
}


function callAiGenerateText_(provider, apiKey, prompt) {
  if (provider === AI_PROVIDER_OPENAI) {
    return callOpenAiGenerateText_(apiKey, prompt);
  }

  return callGeminiGenerateText_(apiKey, prompt);
}

function callGeminiGenerateText_(apiKey, prompt) {
  const modelId = getProviderModel_(AI_PROVIDER_GEMINI);
  const url =
    GEMINI_API_BASE_URL +
    encodeURIComponent(modelId) +
    ':generateContent?key=' +
    encodeURIComponent(apiKey);
  const response = fetchWithRetry_(url, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify({
      contents: [
        {
          parts: [
            {
              text: prompt,
            },
          ],
        },
      ],
      generationConfig: {
        temperature: 0.3,
        maxOutputTokens: 80,
      },
    }),
  });

  if (!isSuccessfulResponse_(response.getResponseCode())) {
    throw buildApiError_(url, response);
  }

  return extractGeminiTextFromResponse_(JSON.parse(response.getContentText() || '{}'));
}

function callOpenAiGenerateText_(apiKey, prompt) {
  const modelId = getProviderModel_(AI_PROVIDER_OPENAI);
  const url = OPENAI_API_BASE_URL + '/responses';
  const response = fetchWithRetry_(url, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + apiKey,
    },
    payload: JSON.stringify({
      model: modelId,
      input: [
        {
          role: 'user',
          content: [
            {
              type: 'input_text',
              text: prompt,
            },
          ],
        },
      ],
      max_output_tokens: 120,
    }),
  });

  if (!isSuccessfulResponse_(response.getResponseCode())) {
    throw buildApiError_(url, response);
  }

  return extractOpenAiTextFromResponse_(JSON.parse(response.getContentText() || '{}'));
}

function promptForProviderModelSelection_(provider, apiKey, currentModel) {
  if (provider === AI_PROVIDER_OPENAI) {
    return promptForOpenAiModelSelection_(apiKey, currentModel);
  }

  return promptForGeminiModelSelection_(apiKey, currentModel);
}

function promptForGeminiModelSelection_(apiKey, currentModel) {
  const ui = SpreadsheetApp.getUi();
  const modelOptions = fetchGeminiModelOptions_(apiKey);
  if (!modelOptions.length) {
    throw new Error('No Gemini models with generateContent support were returned for this API key.');
  }

  const exampleList = modelOptions
    .slice(0, 20)
    .map(function(option) {
      return option.modelId;
    })
    .join('\n');
  const recommendedModel = findRecommendedGeminiModel_(modelOptions, currentModel);
  const response = ui.prompt(
    'Select Gemini Model',
    'Enter the exact Gemini model ID to save.\nCurrent: ' +
      valueOrEmpty_(currentModel) +
      '\nRecommended: ' +
      recommendedModel +
      '\n\nAvailable model IDs:\n' +
      exampleList +
      (modelOptions.length > 20 ? '\n...' : ''),
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return '';
  }

  const requestedModel = String(response.getResponseText() || '').trim() || recommendedModel;
  const matchingModel = modelOptions.find(function(option) {
    return option.modelId === requestedModel;
  });

  if (!matchingModel) {
    throw new Error(
      'The selected model "' + requestedModel + '" was not found in the Gemini models list for this API key.'
    );
  }

  return matchingModel.modelId;
}

function promptForOpenAiModelSelection_(apiKey, currentModel) {
  const ui = SpreadsheetApp.getUi();
  const modelOptions = fetchOpenAiModelOptions_(apiKey);
  if (!modelOptions.length) {
    throw new Error('No supported OpenAI text-generation models were returned for this API key.');
  }

  const exampleList = modelOptions
    .slice(0, 20)
    .map(function(option) {
      return option.modelId;
    })
    .join('\n');
  const recommendedModel = findRecommendedOpenAiModel_(modelOptions, currentModel);
  const response = ui.prompt(
    'Select OpenAI Model',
    'Enter the exact OpenAI model ID to save.\nCurrent: ' +
      valueOrEmpty_(currentModel) +
      '\nRecommended: ' +
      recommendedModel +
      '\n\nAvailable model IDs:\n' +
      exampleList +
      (modelOptions.length > 20 ? '\n...' : ''),
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return '';
  }

  const requestedModel = String(response.getResponseText() || '').trim() || recommendedModel;
  const matchingModel = modelOptions.find(function(option) {
    return option.modelId === requestedModel;
  });

  if (!matchingModel) {
    throw new Error(
      'The selected model "' + requestedModel + '" was not found in the supported OpenAI model list for this API key.'
    );
  }

  return matchingModel.modelId;
}

function fetchGeminiModelOptions_(apiKey) {
  const cacheKey = buildProviderModelCacheKey_(AI_PROVIDER_GEMINI, apiKey);
  const cache = CacheService.getScriptCache();
  const cachedValue = cache.get(cacheKey);
  if (cachedValue) {
    return JSON.parse(cachedValue);
  }

  const modelsById = {};
  let pageToken = '';

  do {
    const url =
      'https://generativelanguage.googleapis.com/v1beta/models?key=' +
      encodeURIComponent(apiKey) +
      (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
    const response = fetchWithRetry_(url, {
      method: 'get',
      muteHttpExceptions: true,
    });

    if (!isSuccessfulResponse_(response.getResponseCode())) {
      throw buildApiError_(url, response);
    }

    const responseBody = JSON.parse(response.getContentText() || '{}');
    const models = responseBody.models || [];
    models.forEach(function(model) {
      const supportedMethods = model.supportedGenerationMethods || [];
      const modelName = String(model.name || '').trim();
      const modelId = extractGeminiModelId_(modelName);
      if (!modelId || supportedMethods.indexOf('generateContent') === -1) {
        return;
      }

      if (!modelsById[modelId]) {
        modelsById[modelId] = {
          modelId: modelId,
          name: modelName,
          displayName: String(model.displayName || modelId),
          description: String(model.description || '').trim(),
        };
      }
    });

    pageToken = String(responseBody.nextPageToken || '').trim();
  } while (pageToken);

  const modelOptions = Object.keys(modelsById)
    .sort()
    .map(function(modelId) {
      return modelsById[modelId];
    });

  cache.put(cacheKey, JSON.stringify(modelOptions), GEMINI_MODEL_CACHE_TTL_SECONDS);
  return modelOptions;
}

function fetchOpenAiModelOptions_(apiKey) {
  const cacheKey = buildProviderModelCacheKey_(AI_PROVIDER_OPENAI, apiKey);
  const cache = CacheService.getScriptCache();
  const cachedValue = cache.get(cacheKey);
  if (cachedValue) {
    return JSON.parse(cachedValue);
  }

  const url = OPENAI_API_BASE_URL + '/models';
  const response = fetchWithRetry_(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + apiKey,
    },
  });

  if (!isSuccessfulResponse_(response.getResponseCode())) {
    throw buildApiError_(url, response);
  }

  const responseBody = JSON.parse(response.getContentText() || '{}');
  const modelOptions = (responseBody.data || [])
    .map(function(model) {
      return {
        modelId: String(model.id || '').trim(),
        ownedBy: String(model.owned_by || '').trim(),
      };
    })
    .filter(function(option) {
      return isSupportedOpenAiModelId_(option.modelId);
    })
    .sort(function(left, right) {
      return left.modelId.localeCompare(right.modelId);
    });

  cache.put(cacheKey, JSON.stringify(modelOptions), GEMINI_MODEL_CACHE_TTL_SECONDS);
  return modelOptions;
}

function findRecommendedGeminiModel_(modelOptions, currentModel) {
  const preferredModels = [
    currentModel,
    DEFAULT_GEMINI_MODEL,
    'gemini-2.0-flash',
  ].filter(Boolean);

  for (let i = 0; i < preferredModels.length; i += 1) {
    const candidate = preferredModels[i];
    const match = modelOptions.find(function(option) {
      return option.modelId === candidate;
    });
    if (match) {
      return match.modelId;
    }
  }

  return modelOptions[0].modelId;
}

function findRecommendedOpenAiModel_(modelOptions, currentModel) {
  const preferredModels = [
    currentModel,
    'gpt-5.1-mini',
    'gpt-5.1',
    'gpt-4.1-mini',
    'gpt-4.1',
    'o4-mini',
    'o3-mini',
  ].filter(Boolean);

  for (let i = 0; i < preferredModels.length; i += 1) {
    const candidate = preferredModels[i];
    const match = modelOptions.find(function(option) {
      return option.modelId === candidate;
    });
    if (match) {
      return match.modelId;
    }
  }

  return modelOptions[0].modelId;
}

function isSupportedOpenAiModelId_(modelId) {
  const normalizedModelId = String(modelId || '').trim().toLowerCase();
  if (!normalizedModelId) {
    return false;
  }

  if (
    /embedding|whisper|transcribe|tts|speech|image|dall-e|moderation|realtime|search|computer-use|vision-preview|audio/.test(normalizedModelId)
  ) {
    return false;
  }

  return /^(gpt-|o[1-9]|o\d)/.test(normalizedModelId);
}

function extractGeminiModelId_(modelName) {
  const normalizedName = String(modelName || '').trim();
  if (!normalizedName) {
    return '';
  }

  return normalizedName.replace(/^models\//i, '').trim();
}

function chunkArray_(items, batchSize) {
  const batches = [];
  const safeBatchSize = Math.max(1, Number(batchSize) || 1);

  for (let index = 0; index < items.length; index += safeBatchSize) {
    batches.push(items.slice(index, index + safeBatchSize));
  }

  return batches;
}

function buildProviderModelCacheKey_(provider, apiKey) {
  return (
    'ai_models_' +
    hashTextForCache_(String(provider || '') + '|' + String(apiKey || '').trim())
  );
}

function buildAiTextCacheKey_(feature, provider, modelId, prompt) {
  return (
    'ai_text_' +
    hashTextForCache_(
      String(feature || '') +
        '|' +
        String(provider || '') +
        '|' +
        String(modelId || '') +
        '|' +
        String(prompt || '')
    )
  );
}

function hashTextForCache_(value) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value || ''),
    Utilities.Charset.UTF_8
  );
  return Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, '');
}

function extractGeminiTextFromResponse_(responseBody) {
  const candidates = responseBody.candidates || [];
  if (!candidates.length) {
    throw new Error('Gemini returned no candidates.');
  }

  const firstCandidate = candidates[0] || {};
  const content = firstCandidate.content || {};
  const parts = content.parts || [];
  const text = parts.map(function(part) {
    return String(part && part.text ? part.text : '');
  }).join(' ').trim();

  if (!text) {
    throw new Error('Gemini returned no text content.');
  }

  return text;
}

function extractOpenAiTextFromResponse_(responseBody) {
  if (typeof responseBody.output_text === 'string' && responseBody.output_text.trim()) {
    return responseBody.output_text.trim();
  }

  const textSegments = [];
  const output = responseBody.output || [];

  output.forEach(function(item) {
    const content = item && item.content ? item.content : [];
    content.forEach(function(part) {
      if (part && (part.type === 'output_text' || part.type === 'text') && part.text) {
        textSegments.push(String(part.text));
      }
    });
  });

  const text = textSegments.join(' ').trim();
  if (!text) {
    throw new Error('OpenAI returned no text content.');
  }

  return text;
}

function sanitizeGeminiMetaDescription_(value) {
  let sanitizedValue = sanitizeSingleLineText_(value)
    .replace(/^["'“”‘’]+|["'“”‘’]+$/g, '')
    .replace(/^(meta description|description)\s*:\s*/i, '')
    .trim();

  if (sanitizedValue.length > META_DESCRIPTION_MAX_LENGTH) {
    sanitizedValue = sanitizedValue.slice(0, META_DESCRIPTION_MAX_LENGTH).trim();
  }

  return sanitizedValue;
}

function sanitizeSingleLineText_(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .trim();
}

function buildPageKeywordFormulaForRow_(pageRow, keywordIndex, rowNumber, groupStartRow) {
  const pageUrl = normalizePageUrl_(pageRow.page_url);
  const pageSourceRange = buildPageKeywordSourceRange_(pageRow);
  const promptSegments = [
    'You are helping with SEO keyword research for a single page.',
    'Return exactly one keyword phrase only.',
    'Do not use numbering, bullets, labels, punctuation-only responses, or explanations.',
    'Keep the phrase realistic, search-friendly, and closely aligned to the page.',
    'If the page context is limited, stay conservative and relevant to the visible metadata.',
    'Generate keyword idea #' + keywordIndex + ' of ' + PAGE_KEYWORD_ROWS_PER_PAGE + ' for this page.',
    'Page URL: ' + pageUrl,
    'Page path: ' + valueOrEmpty_(pageRow.page_path),
    'Title: ' + valueOrEmpty_(pageRow.title),
    'Meta description: ' + valueOrEmpty_(pageRow.meta_description),
    'H1: ' + valueOrEmpty_(pageRow.h1),
  ];

  const escapedPrompt = escapeFormulaString_(promptSegments.join(' '));
  const duplicateClause = buildPageKeywordDuplicateClause_(rowNumber, groupStartRow);

  return (
    '=IF(LEN($A' +
    rowNumber +
    ')=0,"",' +
    AI_SHEETS_FORMULA_FUNCTION +
    '("' +
    escapedPrompt +
    '"&' +
    duplicateClause +
    ',' +
    pageSourceRange +
    '))'
  );
}

function buildAiMetaDescriptionFormulaForRow_(pageRow, rowNumber) {
  const pageUrl = normalizePageUrl_(pageRow.page_url);
  const pageSourceRange = buildPageKeywordSourceRange_(pageRow);
  const existingMetaDescription = sanitizeSingleLineText_(pageRow.meta_description || '');
  const promptSegments = [
    'Rewrite the meta description for this page as a stronger SEO recommendation.',
    'Use the existing meta description as a reference, but improve clarity and CTA strength.',
    'Return plain text only.',
    'Do not use quotes, labels, bullets, or multiple options.',
    'Do not invent facts, pricing, offers, locations, or claims not supported by the page context.',
    'Keep the result between ' + META_DESCRIPTION_MIN_LENGTH + ' and ' + META_DESCRIPTION_MAX_LENGTH + ' characters.',
    'Aim for a concise, clickworthy search snippet with a natural CTA.',
    'Page URL: ' + pageUrl,
    'Existing meta description: ' + existingMetaDescription,
    'Title: ' + valueOrEmpty_(pageRow.title),
    'H1: ' + valueOrEmpty_(pageRow.h1),
  ];

  return (
    '=IF(LEN($C' +
    rowNumber +
    ')=0,"",' +
    AI_SHEETS_FORMULA_FUNCTION +
    '("' +
    escapeFormulaString_(promptSegments.join(' ')) +
    '",' +
    pageSourceRange +
    '))'
  );
}

function buildAiSeoAuditFormulaForRow_(auditRecord, rowNumber) {
  const pageRow = auditRecord.pageData || {};
  const pageSourceRange = buildPageKeywordSourceRange_(pageRow);
  const promptSegments = [
    'You are an SEO consultant writing one concise recommendation for a single page.',
    'Use the issue summary and the provided page data only.',
    'Return plain text only in 2 short sentences maximum.',
    'Sentence 1 should explain the SEO issue clearly.',
    'Sentence 2 should suggest the most practical next action.',
    'Do not invent claims, content, offers, pricing, or unsupported details.',
    'Priority label: ' + valueOrEmpty_(auditRecord.priorityLabel),
    'Priority score: ' + valueOrEmpty_(auditRecord.priorityScore),
    'Issue category: ' + valueOrEmpty_(auditRecord.issueCategory),
    'Issue code: ' + valueOrEmpty_(auditRecord.issueCode),
    'Issue summary: ' + valueOrEmpty_(auditRecord.issueSummary),
    'Severity: ' + valueOrEmpty_(auditRecord.severity),
    'Traffic metrics: ' + valueOrEmpty_(auditRecord.trafficSignal),
    'Search metrics: ' + valueOrEmpty_(auditRecord.searchSignal),
    'Supporting data: ' + valueOrEmpty_(auditRecord.supportingData),
  ];

  return (
    '=IF(LEN($C' +
    rowNumber +
    ')=0,"",' +
    AI_SHEETS_FORMULA_FUNCTION +
    '("' +
    escapeFormulaString_(promptSegments.join(' ')) +
    '",' +
    pageSourceRange +
    '))'
  );
}

function buildPageKeywordDuplicateClause_(rowNumber, groupStartRow) {
  if (rowNumber <= groupStartRow) {
    return '" Avoid duplicates with earlier keywords for this page: none."';
  }

  return (
    '" Avoid duplicates with earlier keywords for this page: "&IFERROR(TEXTJOIN(", ", TRUE, FILTER($B$' +
    groupStartRow +
    ':B' +
    (rowNumber - 1) +
    ', $B$' +
    groupStartRow +
    ':B' +
    (rowNumber - 1) +
    '<>"")), "none")&"."'
  );
}

function escapeFormulaString_(value) {
  return String(value || '').replace(/"/g, '""');
}

function buildPageKeywordSourceRange_(pageRow) {
  const sheetRow = Number(pageRow._sheet_row || 0);
  if (!sheetRow) {
    return '$A$1';
  }

  return "'" + SHEET_NAMES.PAGES.replace(/'/g, "''") + "'!C" + sheetRow + ':I' + sheetRow;
}

function getOrCreateSheet_(sheetName, headers) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  ensureHeaders_(sheet, headers);
  applySheetFormatting_(sheet, headers.length);
  return sheet;
}

function ensureHeaders_(sheet, headers) {
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  const existingHeaders = headerRange.getValues()[0];

  const headersMatch = headers.every(function(header, index) {
    return existingHeaders[index] === header;
  });

  if (!headersMatch) {
    headerRange.setValues([headers]);
  }
}

function applySheetFormatting_(sheet, columnCount) {
  sheet.setFrozenRows(1);
  if (sheet.getMaxColumns() < columnCount) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), columnCount - sheet.getMaxColumns());
  }
  sheet.autoResizeColumns(1, columnCount);
}

function applyDashboardFormatting_(sheet) {
  applySheetFormatting_(sheet, HEADERS.DASHBOARD.length);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (spreadsheet.getSheetByName(SHEET_NAMES.DASHBOARD) !== sheet) {
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  sheet.getRange(2, 1, lastRow - 1, HEADERS.DASHBOARD.length).sort([
    { column: 2, ascending: true },
    { column: 3, ascending: true },
  ]);
}

function applyPageKeywordSheetFormatting_(sheet, rowCount) {
  applySheetFormatting_(sheet, HEADERS.PAGE_KEYWORDS.length);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 340);
  sheet.setColumnWidth(2, 280);

  if (rowCount <= 0) {
    return;
  }

  const pageRange = sheet.getRange(2, 1, rowCount, 1);
  pageRange.setWrap(true);
  pageRange.setVerticalAlignment('middle');

  const keywordRange = sheet.getRange(2, 2, rowCount, 1);
  keywordRange.setWrap(true);
  keywordRange.setVerticalAlignment('middle');
}

function getSheetValues_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn === 0) {
    return { rows: [] };
  }

  return {
    rows: sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues(),
  };
}

function getSitemapCandidates_(websiteUrl) {
  const origin = getOrigin_(websiteUrl);
  const candidates = [];
  const seen = {};

  function addCandidate(url) {
    const normalizedUrl = String(url || '').trim();
    if (!normalizedUrl || seen[normalizedUrl]) {
      return;
    }

    seen[normalizedUrl] = true;
    candidates.push(normalizedUrl);
  }

  try {
    const robotsResponse = fetchWithRetry_(origin + '/robots.txt', {
      method: 'get',
      muteHttpExceptions: true,
    });
    if (isSuccessfulResponse_(robotsResponse.getResponseCode())) {
      extractRobotsSitemapUrls_(robotsResponse.getContentText()).forEach(function(sitemapUrl) {
        addCandidate(sitemapUrl);
      });
    }
  } catch (error) {
    // Ignore robots.txt issues and continue with default sitemap candidates.
  }

  SITEMAP_CANDIDATE_PATHS.forEach(function(path) {
    addCandidate(origin + path);
  });

  return candidates;
}

function extractRobotsSitemapUrls_(robotsText) {
  const matches = [];
  const pattern = /^\s*Sitemap:\s*(\S+)\s*$/gim;
  let match = pattern.exec(robotsText);

  while (match) {
    matches.push(match[1]);
    match = pattern.exec(robotsText);
  }

  return matches;
}

function extractInternalLinks_(html, baseUrl, allowedHost) {
  const links = {};
  const pattern = /<a\b[^>]*href\s*=\s*["']([^"'#]+)["']/gi;
  let match = pattern.exec(html);

  while (match) {
    const resolvedUrl = resolveRelativeUrl_(match[1], baseUrl);
    const normalizedUrl = tryNormalizePageUrl_(resolvedUrl);
    if (
      normalizedUrl &&
      getHost_(normalizedUrl) === allowedHost &&
      !isLikelyBinaryAsset_(normalizedUrl)
    ) {
      links[normalizedUrl] = true;
    }
    match = pattern.exec(html);
  }

  return Object.keys(links);
}

function isLikelyBinaryAsset_(url) {
  return /\.(?:jpg|jpeg|png|gif|webp|svg|pdf|zip|mp4|mp3|avi|mov|woff|woff2|ttf|ico)$/i.test(url);
}

function callGoogleApi_(url, method, payload) {
  const response = fetchWithRetry_(url, {
    method: method,
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
    },
    payload: payload ? JSON.stringify(payload) : null,
  });
  const body = response.getContentText() || '{}';

  if (!isSuccessfulResponse_(response.getResponseCode())) {
    throw buildApiError_(url, response);
  }

  return JSON.parse(body);
}

function fetchWithRetry_(url, options) {
  const maxAttempts = 3;
  let lastError;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      if (responseCode === 429 || responseCode >= 500) {
        lastError = buildApiError_(url, response);
      } else {
        return response;
      }
    } catch (error) {
      lastError = error;
    }

    Utilities.sleep(500 * attempt);
  }

  throw lastError || new Error('Request failed for ' + url);
}

function buildApiError_(url, response) {
  const body = truncateText_(response.getContentText() || '', 300);
  return new Error(
    'Request failed for ' +
      url +
      ' with status ' +
      response.getResponseCode() +
      '. Response: ' +
      body
  );
}

function promptForWebsiteUrl_() {
  const response = SpreadsheetApp.getUi().prompt(
    'Website URL',
    'Enter the website URL to track.\nExample: https://www.example.com',
    SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) {
    return null;
  }

  const value = String(response.getResponseText() || '').trim();
  if (!value) {
    SpreadsheetApp.getUi().alert('Website URL is required.');
    return null;
  }

  try {
    return normalizeWebsiteUrl_(value);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Invalid website URL: ' + error.message);
    return null;
  }
}

function promptForText_(title, message) {
  const response = SpreadsheetApp.getUi().prompt(
    title,
    message,
    SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) {
    return null;
  }

  const value = String(response.getResponseText() || '').trim();
  if (!value) {
    SpreadsheetApp.getUi().alert(title + ' is required.');
    return null;
  }

  return value;
}

function promptForOptionalText_(title, message) {
  const response = SpreadsheetApp.getUi().prompt(
    title,
    message,
    SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) {
    return '';
  }

  return String(response.getResponseText() || '').trim();
}

function promptForDateRangeDays_(defaultValue) {
  const response = SpreadsheetApp.getUi().prompt(
    'Date Range',
    'Enter one of these day ranges: 7, 28, 90.\nDefault: ' + defaultValue,
    SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) {
    return null;
  }

  const value = parseDateRangeDays_(response.getResponseText() || defaultValue);
  if (!value) {
    SpreadsheetApp.getUi().alert('Date range must be 7, 28, or 90.');
    return null;
  }

  return value;
}

function parseDateRangeDays_(value) {
  const days = Number(value);
  if (ALLOWED_DATE_RANGES.indexOf(days) === -1) {
    return null;
  }
  return days;
}

function getDateRange_(days) {
  const endDate = new Date();
  const startDate = new Date(endDate.getTime());
  startDate.setDate(startDate.getDate() - (Number(days) - 1));

  return {
    startDate: formatDate_(startDate),
    endDate: formatDate_(endDate),
  };
}

function formatDate_(date) {
  return Utilities.formatDate(date, TIMEZONE, 'yyyy-MM-dd');
}

function formatDateTime_(date) {
  return Utilities.formatDate(date, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}

function normalizeWebsiteUrl_(input) {
  const url = parseAbsoluteUrl_(input);
  return url.protocol + '//' + url.host + '/';
}

function tryNormalizeWebsiteUrl_(input) {
  try {
    return normalizeWebsiteUrl_(input);
  } catch (error) {
    return null;
  }
}

function normalizePageUrl_(input) {
  const url = parseAbsoluteUrl_(input);
  return url.protocol + '//' + url.host + normalizePathname_(url.pathname);
}

function tryNormalizePageUrl_(input) {
  try {
    return normalizePageUrl_(input);
  } catch (error) {
    return null;
  }
}

function normalizeComparableUrl_(input) {
  const url = parseAbsoluteUrl_(input);
  return url.protocol.toLowerCase() + '//' + url.host.toLowerCase() + normalizePathname_(url.pathname);
}

function normalizePathname_(pathname) {
  if (!pathname) {
    return '/';
  }

  let normalized = pathname.replace(/\/{2,}/g, '/');
  if (normalized.length > 1 && normalized.endsWith('/')) {
    normalized = normalized.slice(0, -1);
  }
  return normalized || '/';
}

function normalizePathKey_(pathValue) {
  const value = String(pathValue || '').trim();
  if (!value) {
    return '/';
  }

  if (/^https?:\/\//i.test(value)) {
    return normalizePathname_(parseAbsoluteUrl_(value).pathname);
  }

  const withoutHash = value.split('#')[0];
  const withoutQuery = withoutHash.split('?')[0];
  const pathname = withoutQuery.startsWith('/') ? withoutQuery : '/' + withoutQuery;
  return normalizePathname_(pathname);
}

function normalizeUrlWithProtocol_(input) {
  const value = String(input || '').trim();
  if (!value) {
    throw new Error('URL is empty.');
  }

  if (/^https?:\/\//i.test(value)) {
    return value;
  }

  return 'https://' + value;
}

function getOrigin_(url) {
  const parsed = parseAbsoluteUrl_(url);
  return parsed.protocol + '//' + parsed.host;
}

function getHost_(url) {
  return parseAbsoluteUrl_(url).host.toLowerCase();
}

function sanitizeHost_(input) {
  return String(input || '')
    .trim()
    .replace(/^https?:\/\//i, '')
    .replace(/\/.*$/, '')
    .toLowerCase();
}

function extractPathFromUrl_(url) {
  return normalizePathname_(parseAbsoluteUrl_(url).pathname);
}

function resolveRelativeUrl_(value, baseUrl) {
  const trimmedValue = String(value || '').trim();
  if (!trimmedValue) {
    return '';
  }

  try {
    if (/^https?:\/\//i.test(trimmedValue)) {
      return normalizePageUrl_(trimmedValue);
    }

    if (/^\/\//.test(trimmedValue)) {
      const baseParts = parseAbsoluteUrl_(baseUrl);
      return normalizePageUrl_(baseParts.protocol + trimmedValue);
    }

    if (trimmedValue.startsWith('/')) {
      return normalizePageUrl_(getOrigin_(baseUrl) + trimmedValue);
    }

    const baseParts = parseAbsoluteUrl_(baseUrl);
    const baseDirectory = getDirectoryPath_(baseParts.pathname);
    const resolvedPath = joinAndNormalizePath_(baseDirectory, trimmedValue);
    return normalizePageUrl_(baseParts.protocol + '//' + baseParts.host + resolvedPath);
  } catch (error) {
    return '';
  }
}

function parseAbsoluteUrl_(input) {
  const value = normalizeUrlWithProtocol_(input);
  const match = /^(https?):\/\/([^\/?#]+)([^?#]*)?(?:\?[^#]*)?(?:#.*)?$/i.exec(value);

  if (!match) {
    throw new Error('Invalid URL format.');
  }

  return {
    protocol: match[1].toLowerCase() + ':',
    host: match[2],
    pathname: match[3] || '/',
  };
}

function getDirectoryPath_(pathname) {
  const normalizedPath = normalizePathname_(pathname);
  if (normalizedPath === '/') {
    return '/';
  }

  const lastSlashIndex = normalizedPath.lastIndexOf('/');
  if (lastSlashIndex <= 0) {
    return '/';
  }

  return normalizedPath.slice(0, lastSlashIndex + 1);
}

function joinAndNormalizePath_(baseDirectory, relativePath) {
  const cleanRelativePath = String(relativePath || '')
    .trim()
    .split('#')[0]
    .split('?')[0];

  const joinedPath = cleanRelativePath.startsWith('/')
    ? cleanRelativePath
    : baseDirectory + cleanRelativePath;

  const segments = joinedPath.split('/');
  const resolvedSegments = [];

  segments.forEach(function(segment) {
    if (!segment || segment === '.') {
      return;
    }

    if (segment === '..') {
      if (resolvedSegments.length) {
        resolvedSegments.pop();
      }
      return;
    }

    resolvedSegments.push(segment);
  });

  return '/' + resolvedSegments.join('/');
}

function isSuccessfulResponse_(statusCode) {
  return statusCode >= 200 && statusCode < 300;
}

function extractXmlTagValues_(xml, tagName) {
  const values = [];
  const pattern = new RegExp(
    '<(?:[a-z0-9_-]+:)?' +
      tagName +
      '\\b[^>]*>([\\s\\S]*?)<\\/(?:[a-z0-9_-]+:)?' +
      tagName +
      '>',
    'gi'
  );
  let match = pattern.exec(xml);

  while (match) {
    values.push(stripCdata_(match[1]).trim());
    match = pattern.exec(xml);
  }

  return values;
}

function stripCdata_(value) {
  return String(value || '')
    .replace(/^<!\[CDATA\[/i, '')
    .replace(/\]\]>$/i, '');
}

function isSitemapIndexXml_(xml) {
  return /<(?:[a-z0-9_-]+:)?sitemapindex\b/i.test(xml);
}

function isUrlSetXml_(xml) {
  return /<(?:[a-z0-9_-]+:)?urlset\b/i.test(xml);
}

function extractTagText_(html, tagName) {
  const pattern = new RegExp('<' + tagName + '\\b[^>]*>([\\s\\S]*?)<\\/' + tagName + '>', 'i');
  const match = pattern.exec(html);
  if (!match) {
    return '';
  }

  return cleanHtmlText_(match[1]);
}

function extractMetaContent_(html, metaName) {
  const patterns = [
    new RegExp(
      '<meta\\b[^>]*name=["\']' +
        metaName +
        '["\'][^>]*content=["\']([\\s\\S]*?)["\'][^>]*>',
      'i'
    ),
    new RegExp(
      '<meta\\b[^>]*content=["\']([\\s\\S]*?)["\'][^>]*name=["\']' +
        metaName +
        '["\'][^>]*>',
      'i'
    ),
  ];

  for (let i = 0; i < patterns.length; i += 1) {
    const match = patterns[i].exec(html);
    if (match) {
      return cleanHtmlText_(match[1]);
    }
  }

  return '';
}

function extractCanonicalHref_(html) {
  const patterns = [
    /<link\b[^>]*rel=["']canonical["'][^>]*href=["']([\s\S]*?)["'][^>]*>/i,
    /<link\b[^>]*href=["']([\s\S]*?)["'][^>]*rel=["']canonical["'][^>]*>/i,
  ];

  for (let i = 0; i < patterns.length; i += 1) {
    const match = patterns[i].exec(html);
    if (match) {
      return String(match[1] || '').trim();
    }
  }

  return '';
}

function cleanHtmlText_(value) {
  return String(value || '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function getHeaderValue_(headers, headerName) {
  const targetName = String(headerName || '').toLowerCase();
  const keys = Object.keys(headers || {});

  for (let i = 0; i < keys.length; i += 1) {
    if (keys[i].toLowerCase() === targetName) {
      return String(headers[keys[i]] || '').trim();
    }
  }

  return '';
}

function parseMetricValue_(metricValues, index) {
  if (!metricValues || !metricValues[index]) {
    return 0;
  }

  const rawValue = metricValues[index].value;
  const parsedValue = Number(rawValue);
  return isNaN(parsedValue) ? 0 : parsedValue;
}

function createEmptySearchConsoleMetrics_() {
  return {
    clicks: 0,
    impressions: 0,
    ctr: 0,
    avg_position: 0,
  };
}

function createEmptyGa4Metrics_() {
  return {
    sessions: 0,
    users: 0,
    conversions: 0,
  };
}

function rowToObject_(headers, row) {
  const object = {};
  headers.forEach(function(header, index) {
    object[header] = row[index];
  });
  return object;
}

function objectToRow_(headers, object) {
  return headers.map(function(header) {
    return valueOrEmpty_(object[header]);
  });
}

function normalizeRowLength_(row, targetLength) {
  const normalizedRow = row.slice(0, targetLength);

  while (normalizedRow.length < targetLength) {
    normalizedRow.push('');
  }

  return normalizedRow;
}

function valueOrEmpty_(value) {
  return value === null || value === undefined ? '' : value;
}

function toNumber_(value) {
  const numericValue = Number(value);
  return isNaN(numericValue) ? 0 : numericValue;
}

function truncateText_(value, maxLength) {
  const text = String(value || '');
  if (text.length <= maxLength) {
    return text;
  }
  return text.slice(0, maxLength - 3) + '...';
}
