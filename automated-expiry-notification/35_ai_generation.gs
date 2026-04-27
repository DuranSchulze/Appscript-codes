// 35 Ai Generation

function isAiGenerationEnabled() {
  return getPropBoolean(PROP_KEYS.AI_ENABLED, false);
}

function getAiModel() {
  return getPropString(PROP_KEYS.AI_MODEL, DEFAULT_AI_MODEL);
}

function getAiApiKey() {
  return getPropString(PROP_KEYS.AI_API_KEY, "");
}

function getFallbackTemplateMode() {
  var mode = getPropString(
    PROP_KEYS.FALLBACK_TEMPLATE_MODE,
    FALLBACK_TEMPLATE_MODE.HARDCODED,
  ).toUpperCase();
  return mode === FALLBACK_TEMPLATE_MODE.PROPERTY
    ? FALLBACK_TEMPLATE_MODE.PROPERTY
    : FALLBACK_TEMPLATE_MODE.HARDCODED;
}

function getConfiguredFallbackTemplate() {
  return getPropString(PROP_KEYS.FALLBACK_TEMPLATE, "");
}

function maskSecret(value) {
  var text = String(value || "");
  if (!text) return "(not set)";
  if (text.length <= 8) return "********";
  return text.substring(0, 4) + "..." + text.substring(text.length - 4);
}

function tryGenerateAiEmailBody(clientName, expiryDate, docTypeText) {
  var apiKey = getAiApiKey();
  var model = getAiModel();
  if (!apiKey || !model) return null;

  var modelPath = model;
  if (modelPath.indexOf("models/") !== 0) {
    modelPath = "models/" + modelPath;
  }

  var prompt = [
    "Write a concise but professional visa/document expiry reminder email body.",
    "Tone: courteous, formal, and actionable.",
    "Do not include a subject line.",
    "Use plain text with short paragraphs.",
    "Client Name: " + clientName,
    "Document Type: " + docTypeText,
    "Expiry Date: " + formatDate(expiryDate),
  ].join("\n");

  var endpoint =
    "https://generativelanguage.googleapis.com/v1beta/" +
    modelPath +
    ":generateContent?key=" +
    encodeURIComponent(apiKey);

  var payload = {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.5,
      maxOutputTokens: 400,
    },
  };

  for (var attempt = 1; attempt <= 2; attempt++) {
    try {
      var response = UrlFetchApp.fetch(endpoint, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      });

      var statusCode = response.getResponseCode();
      var body = response.getContentText() || "";
      if (statusCode < 200 || statusCode >= 300) {
        if (attempt < 2) {
          Utilities.sleep(350);
          continue;
        }
        return null;
      }

      var parsed = JSON.parse(body);
      var text = extractGeminiText(parsed);
      if (!text) {
        if (attempt < 2) {
          Utilities.sleep(350);
          continue;
        }
        return null;
      }

      return {
        text: text.trim(),
        model: modelPath,
      };
    } catch (e) {
      if (attempt < 2) {
        Utilities.sleep(350);
        continue;
      }
      return null;
    }
  }

  return null;
}

function extractGeminiText(payload) {
  if (!payload || !payload.candidates || payload.candidates.length === 0) {
    return "";
  }

  var candidate = payload.candidates[0];
  var parts =
    candidate && candidate.content && candidate.content.parts
      ? candidate.content.parts
      : [];

  var textChunks = [];
  for (var i = 0; i < parts.length; i++) {
    var partText = String(parts[i].text || "").trim();
    if (partText) textChunks.push(partText);
  }
  return textChunks.join("\n\n");
}

function toggleAiGeneration() {
  var enabled = !isAiGenerationEnabled();
  setPropBoolean(PROP_KEYS.AI_ENABLED, enabled);
  SpreadsheetApp.getUi().alert(
    "AI Integration",
    "AI generation is now " + (enabled ? "ENABLED" : "DISABLED") + ".",
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

function setGeminiApiKey() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    "Set Gemini API Key",
    "Enter Gemini API key. Leave blank to clear.",
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var key = response.getResponseText().trim();
  setPropString(PROP_KEYS.AI_API_KEY, key);
  setPropString(PROP_KEYS.AI_PROVIDER, AI_PROVIDER.GEMINI);
  ui.alert(
    "AI Integration",
    key ? "Gemini key saved." : "Gemini key cleared.",
    ui.ButtonSet.OK,
  );
}

function selectGeminiModel() {
  var ui = SpreadsheetApp.getUi();
  var current = getAiModel();
  var apiKey = getAiApiKey();
  var models = apiKey ? listGeminiModels(apiKey) : [];

  var promptText =
    "Enter Gemini model name (e.g., models/gemini-1.5-flash).\n\nCurrent: " +
    current;
  if (models.length > 0) {
    promptText += "\n\nAvailable:\n- " + models.slice(0, 10).join("\n- ");
  }

  var response = ui.prompt(
    "Select Gemini Model",
    promptText,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var model = response.getResponseText().trim();
  if (!model) {
    ui.alert("Model cannot be empty.");
    return;
  }
  setPropString(PROP_KEYS.AI_MODEL, model);
  ui.alert("AI model saved: " + model);
}

function listGeminiModels(apiKey) {
  try {
    var response = UrlFetchApp.fetch(
      "https://generativelanguage.googleapis.com/v1beta/models?key=" +
        encodeURIComponent(apiKey),
      { muteHttpExceptions: true },
    );
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
      return [];
    }
    var payload = JSON.parse(response.getContentText() || "{}");
    var models = payload.models || [];
    return models
      .map(function (item) {
        return String(item.name || "").trim();
      })
      .filter(function (name) {
        return name.indexOf("models/gemini") === 0;
      });
  } catch (e) {
    return [];
  }
}

function testAiConnection() {
  var ui = SpreadsheetApp.getUi();
  var result = tryGenerateAiEmailBody("Test Client", new Date(), "Visa/Permit");
  if (result && result.text) {
    ui.alert(
      "AI Integration",
      "AI connection successful. Model: " + result.model,
      ui.ButtonSet.OK,
    );
  } else {
    ui.alert(
      "AI Integration",
      "AI test failed. Check API key/model and try again.",
      ui.ButtonSet.OK,
    );
  }
}

function showAiStatus() {
  var ui = SpreadsheetApp.getUi();
  var msg = [
    "AI Enabled: " + (isAiGenerationEnabled() ? "YES" : "NO"),
    "Provider: " + getPropString(PROP_KEYS.AI_PROVIDER, AI_PROVIDER.GEMINI),
    "Model: " + getAiModel(),
    "API Key: " + maskSecret(getAiApiKey()),
    "Fallback Source: " + getFallbackTemplateMode(),
    "Open Tracking URL: " + (getOpenTrackingBaseUrl() || "(not set)"),
  ].join("\n");
  ui.alert("AI Status", msg, ui.ButtonSet.OK);
}

function setFallbackTemplate() {
  var ui = SpreadsheetApp.getUi();
  var current = getConfiguredFallbackTemplate();
  var response = ui.prompt(
    "Set Fallback Template",
    "Enter fallback body template text.\n\nCurrent value (first 300 chars):\n" +
      (current ? current.substring(0, 300) : "(empty)"),
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  setPropString(PROP_KEYS.FALLBACK_TEMPLATE, response.getResponseText());
  ui.alert("Fallback template saved.");
}

function toggleFallbackTemplateSource() {
  var current = getFallbackTemplateMode();
  var next =
    current === FALLBACK_TEMPLATE_MODE.HARDCODED
      ? FALLBACK_TEMPLATE_MODE.PROPERTY
      : FALLBACK_TEMPLATE_MODE.HARDCODED;
  setPropString(PROP_KEYS.FALLBACK_TEMPLATE_MODE, next);
  SpreadsheetApp.getUi().alert(
    "Fallback Source",
    "Fallback template mode is now: " + next,
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}
