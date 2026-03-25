const SHEET_NAMES = {
  PAGES: 'Pages',
  CONFIG: 'Config',
  LOGS: 'Logs',
};

const HEADERS = {
  CONFIG: [
    'website_url',
    'host',
    'search_console_property',
    'ga4_property_id',
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
  ],
  LOGS: [
    'timestamp',
    'website_url',
    'action',
    'status',
    'pages_discovered',
    'pages_written',
    'message',
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

function onOpen() {
  ensureRequiredSheets_();
  SpreadsheetApp.getUi()
    .createMenu('SEO Workspace')
    .addItem('Setup Website', 'setupWebsite')
    .addItem('Refresh Website', 'refreshWebsite')
    .addItem('Refresh Website With Date Range', 'refreshWebsiteWithDateRange')
    .addSeparator()
    .addItem('View Logs', 'viewLogs')
    .addToUi();
}

function onInstall() {
  onOpen();
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

  const configRecord = {
    website_url: normalizedWebsiteUrl,
    host: host,
    search_console_property: '',
    ga4_property_id: '',
    default_date_range_days: String(DEFAULT_DATE_RANGE_DAYS),
    last_refresh_at: '',
    last_refresh_status: '',
  };

  upsertConfigRecord_(configRecord);

  ui.alert(
    'Website Saved',
    'The website was saved. You can now run "Refresh Website" for audit data. Search Console and GA4 can be added later in the Config tab.',
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

function viewLogs() {
  ensureRequiredSheets_();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.LOGS);
  spreadsheet.setActiveSheet(sheet);
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
  const enrichment = buildGoogleEnrichment_(
    config,
    dateRange.startDate,
    dateRange.endDate
  );
  const refreshedAt = formatDateTime_(new Date());

  const rows = discovery.pages
    .map(function(pageEntry) {
      const pageUrl = normalizePageUrl_(pageEntry.page_url);
      const pagePath = extractPathFromUrl_(pageUrl);
      const audit = audits[pageUrl] || {};
      const comparisonUrl = audit.canonical_url
        ? normalizeComparableUrl_(audit.canonical_url)
        : normalizeComparableUrl_(pageUrl);
      const searchMetric =
        enrichment.searchConsole.metrics[comparisonUrl] ||
        enrichment.searchConsole.metrics[normalizeComparableUrl_(pageUrl)] ||
        createEmptySearchConsoleMetrics_();
      const gaMetric =
        enrichment.ga4.metrics[normalizePathKey_(pagePath)] || createEmptyGa4Metrics_();

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
      ];
    })
    .sort(function(left, right) {
      return left[2].localeCompare(right[2]);
    });

  return {
    rows: rows,
    discoveredCount: discovery.pages.length,
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

  if (isSitemapIndexXml_(xml)) {
    extractXmlTagValues_(xml, 'loc').forEach(function(childSitemapUrl) {
      const trimmedUrl = String(childSitemapUrl || '').trim();
      if (trimmedUrl) {
        collectSitemapPages_(trimmedUrl, allowedHost, visitedSitemaps, discoveredPages);
      }
    });
    return;
  }

  if (!isUrlSetXml_(xml)) {
    return;
  }

  extractXmlTagValues_(xml, 'loc').forEach(function(pageUrl) {
    const normalizedPageUrl = tryNormalizePageUrl_(pageUrl);
    if (normalizedPageUrl && getHost_(normalizedPageUrl) === allowedHost) {
      discoveredPages[normalizedPageUrl] = true;
    }
  });
}

function crawlWebsitePages_(websiteUrl, allowedHost) {
  const queue = [{ url: websiteUrl, depth: 0 }];
  const visited = {};
  const discovered = {};

  while (queue.length && Object.keys(discovered).length < CRAWL_MAX_PAGES) {
    const current = queue.shift();
    const comparableUrl = normalizeComparableUrl_(current.url);
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
      const comparableLink = normalizeComparableUrl_(linkUrl);
      if (!visited[comparableLink] && !discovered[linkUrl]) {
        queue.push({
          url: linkUrl,
          depth: current.depth + 1,
        });
      }
    });
  }

  return Object.keys(discovered).sort();
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
  const candidates = {};

  SITEMAP_CANDIDATE_PATHS.forEach(function(path) {
    candidates[origin + path] = true;
  });

  try {
    const robotsResponse = fetchWithRetry_(origin + '/robots.txt', {
      method: 'get',
      muteHttpExceptions: true,
    });
    if (isSuccessfulResponse_(robotsResponse.getResponseCode())) {
      extractRobotsSitemapUrls_(robotsResponse.getContentText()).forEach(function(sitemapUrl) {
        candidates[sitemapUrl] = true;
      });
    }
  } catch (error) {
    // Ignore robots.txt issues and continue with default sitemap candidates.
  }

  return Object.keys(candidates);
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

function valueOrEmpty_(value) {
  return value === null || value === undefined ? '' : value;
}

function truncateText_(value, maxLength) {
  const text = String(value || '');
  if (text.length <= maxLength) {
    return text;
  }
  return text.slice(0, maxLength - 3) + '...';
}
