/**
 * Adds placeholder FAQ rows to Map Sheet if they don't exist.
 */
function setupFAQInMapSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapSheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);
  if (!mapSheet) {
    SpreadsheetApp.getUi().alert('Map Sheet not found. Please run initial setup first.');
    return;
  }
  const data = mapSheet.getDataRange().getValues();
  const hasFAQ = data.some(row => row[0] === 'FAQ');
  if (!hasFAQ) {
    const placeholder = [
      ['FAQ', 'What are your requirements for a tourist visa?', 'The requirements include...'],
      ['FAQ', 'How do I register a corporation?', 'To register a corporation, you need to...'],
      ['FAQ', 'How can I schedule a consultation?', 'You can book a consultation with Marie Christine: https://tidycal.com/mariechristine or Atty. Mary Wendy: https://tidycal.com/attymarywendy']
    ];
    mapSheet.getRange(mapSheet.getLastRow() + 2, 1, placeholder.length, 3).setValues(placeholder);
  }
}

/**
 * One-time setup for AI addon (updated to not create new sheet).
 */
function setupAIAddon() {
  setupFAQInMapSheet();
  ensureCategoryRouting();  // already exists, but safe to call
  SpreadsheetApp.getUi().alert(
    'AI Addon Setup',
    'FAQ entries added to Map Sheet. Use "Configure Gemini Model" from the menu to configure Gemini.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Ensures that the Map Sheet has CategoryRouting entries (if not, creates a placeholder).
 */
function ensureCategoryRouting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapSheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);
  if (!mapSheet) return;
  const data = mapSheet.getDataRange().getValues();
  const hasCategoryRouting = data.some(row => row[0] === 'CategoryRouting');
  if (!hasCategoryRouting) {
    const placeholder = [
      ['CategoryRouting', 'Visa', 'visa@duranschulze.com'],
      ['CategoryRouting', 'Legal', 'legal@duranschulze.com'],
      ['CategoryRouting', 'Business Formation', 'corporate@duranschulze.com'],
      ['CategoryRouting', 'Trademark', 'trademark@duranschulze.com'],
      ['CategoryRouting', 'Accounting', 'accounting@duranschulze.com']
    ];
    mapSheet.getRange(mapSheet.getLastRow() + 2, 1, placeholder.length, 3).setValues(placeholder);
  }
}

/**
 * Opens the Gemini configuration dialog.
 * The model list is loaded from Gemini so users cannot save a misspelled or
 * unavailable model ID.
 */
function setAPIKeyAndModel() {
  const html = HtmlService.createHtmlOutputFromFile('GeminiModelSelector')
    .setWidth(720)
    .setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, 'Configure Gemini AI');
}

/**
 * Returns non-secret settings used when the model selector opens.
 */
function getGeminiSettingsForUi() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY') || '';

  return {
    hasApiKey: Boolean(apiKey),
    maskedApiKey: apiKey ? '••••••' + apiKey.slice(-4) : '',
    primaryModel: normalizeGeminiModelId(
      props.getProperty('GEMINI_MODEL') || AI_CONFIG.DEFAULT_GEMINI_MODEL
    ),
    fallbackModel: normalizeGeminiModelId(props.getProperty('GEMINI_FALLBACK_MODEL') || ''),
    validatedAt: props.getProperty('GEMINI_MODEL_VALIDATED_AT') || '',
    lastModelUsed: normalizeGeminiModelId(props.getProperty('GEMINI_LAST_MODEL_USED') || ''),
    lastFallbackAt: props.getProperty('GEMINI_LAST_FALLBACK_AT') || ''
  };
}

/**
 * Gets models available to either the supplied key or the saved key.
 * Only models that explicitly support generateContent are returned.
 */
function getAvailableGeminiModels(apiKeyInput) {
  const apiKey = resolveGeminiApiKey_(apiKeyInput);
  const response = UrlFetchApp.fetch(AI_CONFIG.GEMINI_API_BASE_URL + '/models?pageSize=100', {
    method: 'get',
    headers: { 'x-goog-api-key': apiKey },
    muteHttpExceptions: true
  });

  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();
  let data;

  try {
    data = JSON.parse(responseText);
  } catch (error) {
    throw new Error('Gemini returned an unreadable response. Please try again.');
  }

  if (statusCode < 200 || statusCode >= 300) {
    const apiMessage = data && data.error && data.error.message;
    throw new Error(apiMessage || 'Could not load Gemini models (HTTP ' + statusCode + ').');
  }

  const models = (data.models || [])
    .filter(function(model) {
      return (model.supportedGenerationMethods || []).indexOf('generateContent') !== -1;
    })
    .map(function(model) {
      return {
        id: normalizeGeminiModelId(model.name),
        label: model.displayName || normalizeGeminiModelId(model.name),
        description: model.description || '',
        inputTokenLimit: model.inputTokenLimit || null,
        outputTokenLimit: model.outputTokenLimit || null
      };
    })
    .filter(function(model) { return Boolean(model.id); })
    .sort(function(a, b) { return a.label.localeCompare(b.label); });

  if (!models.length) {
    throw new Error('No Gemini models supporting content generation were found for this API key.');
  }

  return models;
}

/**
 * Validates and saves the API key/model chosen in the configuration dialog.
 */
function saveGeminiSettings(settings) {
  settings = settings || {};
  const apiKey = resolveGeminiApiKey_(settings.apiKey);
  const primaryModel = normalizeGeminiModelId(settings.primaryModel || settings.model);
  const fallbackModel = normalizeGeminiModelId(settings.fallbackModel);

  if (!primaryModel) {
    throw new Error('Select a primary Gemini model before saving.');
  }
  if (fallbackModel && fallbackModel === primaryModel) {
    throw new Error('Choose a fallback model that is different from the primary model.');
  }

  // Fetch again at save time so a forged or stale browser value is never saved.
  const availableModels = getAvailableGeminiModels(apiKey);
  const availableIds = availableModels.map(function(model) { return model.id; });

  if (availableIds.indexOf(primaryModel) === -1) {
    throw new Error('The primary model is not available for this API key or does not support generateContent.');
  }
  if (fallbackModel && availableIds.indexOf(fallbackModel) === -1) {
    throw new Error('The fallback model is not available for this API key or does not support generateContent.');
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    GEMINI_API_KEY: apiKey,
    GEMINI_MODEL: primaryModel,
    GEMINI_FALLBACK_MODEL: fallbackModel,
    GEMINI_MODEL_VALIDATED_AT: new Date().toISOString()
  });

  return {
    success: true,
    primaryModel: primaryModel,
    fallbackModel: fallbackModel,
    validatedAt: props.getProperty('GEMINI_MODEL_VALIDATED_AT')
  };
}

/**
 * Runs a small, non-destructive generation request against the selected route.
 * This confirms that the key can actually call both models before settings are
 * saved. The API token is never returned to the browser.
 */
function testGeminiConfiguration(settings) {
  settings = settings || {};
  const apiKey = resolveGeminiApiKey_(settings.apiKey);
  const primaryModel = normalizeGeminiModelId(settings.primaryModel);
  const fallbackModel = normalizeGeminiModelId(settings.fallbackModel);

  if (!primaryModel) {
    throw new Error('Select a primary model before testing.');
  }
  if (fallbackModel && fallbackModel === primaryModel) {
    throw new Error('Primary and fallback models must be different.');
  }

  const models = getAvailableGeminiModels(apiKey);
  const availableIds = models.map(function(model) { return model.id; });
  if (availableIds.indexOf(primaryModel) === -1) {
    throw new Error('The primary model is not currently available to this API key.');
  }
  if (fallbackModel && availableIds.indexOf(fallbackModel) === -1) {
    throw new Error('The fallback model is not currently available to this API key.');
  }

  const results = [testGeminiModel_(apiKey, primaryModel, 'Primary')];
  if (fallbackModel) {
    results.push(testGeminiModel_(apiKey, fallbackModel, 'Fallback'));
  }

  return { success: true, results: results };
}

function testGeminiModel_(apiKey, model, role) {
  const url = AI_CONFIG.GEMINI_API_BASE_URL + '/models/' + encodeURIComponent(model) + ':generateContent';
  const startedAt = new Date().getTime();
  reserveAiRequest_('configuration', model);
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { 'x-goog-api-key': apiKey },
    contentType: 'application/json',
    payload: JSON.stringify({
      contents: [{ parts: [{ text: 'Reply with exactly: READY' }] }],
      generationConfig: { temperature: 0, maxOutputTokens: 8 }
    }),
    muteHttpExceptions: true
  });
  const statusCode = response.getResponseCode();

  if (statusCode < 200 || statusCode >= 300) {
    recordAiHttpFailure_(statusCode);
    let message = 'HTTP ' + statusCode;
    try {
      const data = JSON.parse(response.getContentText());
      message = data && data.error && data.error.message ? data.error.message : message;
    } catch (ignore) {}
    throw new Error(role + ' model test failed: ' + message);
  }

  recordAiRequestSuccess_();

  let data;
  try {
    data = JSON.parse(response.getContentText());
  } catch (error) {
    throw new Error(role + ' model test failed: Gemini returned an unreadable response.');
  }
  if (!data.candidates || !data.candidates.length ||
      !data.candidates[0].content || !data.candidates[0].content.parts ||
      !data.candidates[0].content.parts.length) {
    throw new Error(role + ' model test failed: no generated content was returned.');
  }

  return {
    role: role,
    model: model,
    latencyMs: new Date().getTime() - startedAt
  };
}

/**
 * Uses a newly entered key when present; otherwise safely falls back to the
 * saved Script Property without sending the saved key to the browser.
 */
function resolveGeminiApiKey_(apiKeyInput) {
  const suppliedKey = String(apiKeyInput || '').trim();
  const savedKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
  const apiKey = suppliedKey || savedKey;

  if (!apiKey) {
    throw new Error('Enter a Gemini API key before loading models.');
  }

  return apiKey;
}
