/**
 * ENHANCED NOTARY DOCUMENT PROCESSING WITH MONTHLY SHEET ORGANIZATION
 * Modified: Secure API Key Management + New Filename Format
 * Format: YYYY MM DD Doc. NUMBER Most Descriptive Title (Affiant).extension
 */

// ============================================================================
// CUSTOM MENU & API KEY MANAGEMENT
// ============================================================================

/**
 * Adds a custom menu to the spreadsheet interface with API key management
 * This function runs automatically when the spreadsheet is opened
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Notary Automation")
    .addItem("🚀 Process Uploaded Documents", "processUploads")
    .addSeparator()
    .addItem("🔐 Set API Key", "showApiKeyDialog")
    .addItem("🤖 Set Gemini Model", "showGeminiModelDialog")
    .addItem("✅ Test API Key", "testApiKey")
    .addItem("📊 API / Model Status", "showApiKeyStatus")
    .addItem("🗑️ Clear API Key", "clearApiKey")
    .addToUi();
  console.log(
    "Custom menu 'Notary Automation' created with API key management",
  );
}

/**
 * Shows dialog for API key input
 */
function showApiKeyDialog() {
  var ui = SpreadsheetApp.getUi();
  var currentKey = getApiKey();
  var keyHint = currentKey
    ? "Current key: " +
      currentKey.substring(0, 10) +
      "..." +
      currentKey.substring(currentKey.length - 4)
    : "No key configured";

  var promptResult = ui.prompt(
    "🔐 Set Gemini API Key",
    keyHint +
      "\n\nEnter your Gemini API key (starts with 'AIza'):\n" +
      "Get one at: https://aistudio.google.com/app/apikey",
    ui.ButtonSet.OK_CANCEL,
  );

  if (promptResult.getSelectedButton() !== ui.Button.OK) {
    console.log("API key input cancelled by user");
    return;
  }

  var apiKey = (promptResult.getResponseText() || "").trim();
  var result = saveApiKey(apiKey);

  if (result.success) {
    ui.alert("✅ API Key Saved", result.message, ui.ButtonSet.OK);
  } else {
    ui.alert("❌ API Key Error", result.message, ui.ButtonSet.OK);
  }
}

/**
 * Save API key to Script Properties (secure storage)
 */
function saveApiKey(apiKey) {
  try {
    if (!apiKey || apiKey.trim() === "") {
      return {
        success: false,
        message: "API key cannot be empty",
      };
    }

    apiKey = apiKey.trim();

    if (!apiKey.startsWith("AIza")) {
      return {
        success: false,
        message: 'Invalid API key format. Should start with "AIza"',
      };
    }

    if (apiKey.length < 30) {
      return {
        success: false,
        message: "API key seems too short",
      };
    }

    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty("GEMINI_API_KEY", apiKey);

    console.log("API key saved successfully");

    return {
      success: true,
      message: "API key saved successfully! You can now process documents.",
    };
  } catch (error) {
    console.error("Error saving API key: " + error.message);
    return {
      success: false,
      message: "Error saving API key: " + error.message,
    };
  }
}

/**
 * Get API key from Script Properties
 */
function getApiKey() {
  try {
    var scriptProperties = PropertiesService.getScriptProperties();
    var apiKey = scriptProperties.getProperty("GEMINI_API_KEY");

    if (!apiKey || apiKey === "") {
      console.log("No API key found in storage");
      return null;
    }

    return apiKey;
  } catch (error) {
    console.error("Error retrieving API key: " + error.message);
    return null;
  }
}

/**
 * Default Gemini model when no model is configured
 */
function getDefaultGeminiModel() {
  return "gemini-3.0-flash";
}

/**
 * Normalize model input (allows users to paste with optional "models/" prefix)
 */
function normalizeGeminiModelName(modelName) {
  if (!modelName) return "";

  var normalized = modelName.toString().trim();
  normalized = normalized.replace(/^models\//i, "");
  return normalized;
}

/**
 * Save Gemini model to Script Properties
 */
function saveGeminiModel(modelName) {
  try {
    var normalizedModel = normalizeGeminiModelName(modelName);

    if (!normalizedModel) {
      return {
        success: false,
        message: "Model name cannot be empty",
      };
    }

    if (!/^[-a-zA-Z0-9._]+$/.test(normalizedModel)) {
      return {
        success: false,
        message: "Invalid model name format. Use values like gemini-3.0-flash",
      };
    }

    if (!/^gemini-/i.test(normalizedModel)) {
      return {
        success: false,
        message: "Model name should start with 'gemini-'",
      };
    }

    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty("GEMINI_MODEL", normalizedModel);

    console.log("Gemini model saved successfully: " + normalizedModel);

    return {
      success: true,
      message: "Gemini model saved: " + normalizedModel,
    };
  } catch (error) {
    console.error("Error saving Gemini model: " + error.message);
    return {
      success: false,
      message: "Error saving model: " + error.message,
    };
  }
}

/**
 * Get Gemini model from Script Properties (falls back to default)
 */
function getGeminiModel() {
  try {
    var scriptProperties = PropertiesService.getScriptProperties();
    var model = scriptProperties.getProperty("GEMINI_MODEL");
    var normalizedModel = normalizeGeminiModelName(model);

    if (!normalizedModel) {
      return getDefaultGeminiModel();
    }

    return normalizedModel;
  } catch (error) {
    console.error("Error retrieving Gemini model: " + error.message);
    return getDefaultGeminiModel();
  }
}

/**
 * Show simple prompt for Gemini model selection
 */
function showGeminiModelDialog() {
  var ui = SpreadsheetApp.getUi();
  var currentModel = getGeminiModel();

  var promptResult = ui.prompt(
    "🤖 Set Gemini Model",
    "Current model: " +
      currentModel +
      "\n\nEnter Gemini model (example: gemini-3.0-flash):",
    ui.ButtonSet.OK_CANCEL,
  );

  if (promptResult.getSelectedButton() !== ui.Button.OK) {
    console.log("Gemini model input cancelled by user");
    return;
  }

  var modelInput = (promptResult.getResponseText() || "").trim();
  var result = saveGeminiModel(modelInput);

  if (result.success) {
    ui.alert("✅ Gemini Model Saved", result.message, ui.ButtonSet.OK);
  } else {
    ui.alert("❌ Gemini Model Error", result.message, ui.ButtonSet.OK);
  }
}

/**
 * Build Gemini generateContent URL using selected model
 */
function buildGeminiGenerateContentUrl(apiKey) {
  var model = getGeminiModel();
  return (
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    model +
    ":generateContent?key=" +
    apiKey
  );
}

/**
 * Test if the stored API key is valid
 */
function testApiKey() {
  var ui = SpreadsheetApp.getUi();
  try {
    var apiKey = getApiKey();
    var geminiModel = getGeminiModel();

    if (!apiKey) {
      ui.alert(
        "❌ No API Key Found",
        'Please set your Gemini API key first using "Notary Automation > Set API Key"',
        ui.ButtonSet.OK,
      );
      return;
    }

    ui.alert(
      "🔄 Testing API Key...",
      "Please wait while we verify your API key.",
      ui.ButtonSet.OK,
    );

    var url = buildGeminiGenerateContentUrl(apiKey);

    var payload = {
      contents: [
        {
          parts: [
            {
              text: "Hello, respond with 'OK' if you receive this.",
            },
          ],
        },
      ],
    };

    var options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };

    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();

    if (responseCode === 200) {
      ui.alert(
        "✅ API Key Valid",
        "Your Gemini API key is working correctly!\n\n" +
          "Model: " +
          geminiModel +
          "\n\n" +
          'You can now process documents using "Notary Automation > Process Uploaded Documents"',
        ui.ButtonSet.OK,
      );
      console.log("API key test successful");
    } else {
      var errorText = response.getContentText();
      ui.alert(
        "❌ API Key Invalid",
        "Your API key is not working.\n\n" +
          "Model: " +
          geminiModel +
          "\n\n" +
          "Error: " +
          errorText +
          "\n\n" +
          "Please check your API key/model and try again.",
        ui.ButtonSet.OK,
      );
      console.error("API key test failed: " + errorText);
    }
  } catch (error) {
    ui.alert(
      "❌ Test Failed",
      "Error testing API key: " + error.message,
      ui.ButtonSet.OK,
    );
    console.error("Error testing API key: " + error.message);
  }
}

/**
 * Show current API key status
 */
function showApiKeyStatus() {
  var ui = SpreadsheetApp.getUi();
  var apiKey = getApiKey();
  var geminiModel = getGeminiModel();
  if (!apiKey) {
    ui.alert(
      "📊 API / Model Status",
      "❌ No API key configured\n\n" +
        "Model: " +
        geminiModel +
        "\n\n" +
        'Please set your API key using "Notary Automation > Set API Key"',
      ui.ButtonSet.OK,
    );
  } else {
    var maskedKey =
      apiKey.substring(0, 10) + "..." + apiKey.substring(apiKey.length - 4);

    ui.alert(
      "📊 API / Model Status",
      "✅ API key is configured\n\n" +
        "Key: " +
        maskedKey +
        "\n" +
        "Model: " +
        geminiModel +
        "\n\n" +
        'Use "Test API Key" to verify it\'s working.',
      ui.ButtonSet.OK,
    );
  }
}

/**
 * Clear stored API key
 */
function clearApiKey() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert(
    "🗑️ Clear API Key",
    "Are you sure you want to remove the stored API key?\n\n" +
      "You will need to enter it again to process documents.",
    ui.ButtonSet.YES_NO,
  );

  if (response === ui.Button.YES) {
    try {
      var scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.deleteProperty("GEMINI_API_KEY");

      ui.alert(
        "✅ API Key Cleared",
        "Your API key has been removed from storage.",
        ui.ButtonSet.OK,
      );

      console.log("API key cleared successfully");
    } catch (error) {
      ui.alert(
        "❌ Error",
        "Error clearing API key: " + error.message,
        ui.ButtonSet.OK,
      );
      console.error("Error clearing API key: " + error.message);
    }
  }
}

// ============================================================================
// MAIN PROCESSING FUNCTION (MODIFIED FOR API KEY AND NEW FILENAME FORMAT)
// ============================================================================

/**
 * Main processing function - NOW USES STORED API KEY
 */
function processUploads() {
  // Configuration - Replace with your actual IDs
  var uploadFolderId = "1SGRUZS703zoGyhfCtUJYkGN47DHeWG8c";
  var logSheetId = "1cM36ob20ZDDkDzN6QRsKJ45iGsj1UA5CNhN9ASz9VxY";

  // MODIFIED: Get API key from secure storage instead of hardcoded value
  var geminiApiKey = getApiKey();

  if (!geminiApiKey) {
    console.log(
      "⚠️ WARNING: No Gemini API key found. Will use filename-based extraction only.",
    );
    console.log(
      "Set your API key using 'Notary Automation > Set API Key' for better accuracy.",
    );

    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      "⚠️ No API Key Found",
      "No Gemini API key is configured. The script will use basic filename extraction.\n\n" +
        'For better accuracy, set your API key using "Notary Automation > Set API Key"\n\n' +
        "Do you want to continue without AI analysis?",
      ui.ButtonSet.YES_NO,
    );

    if (response !== ui.Button.YES) {
      console.log("Processing cancelled by user");
      return;
    }
  }

  try {
    var uploadFolder = DriveApp.getFolderById(uploadFolderId);
    var masterSpreadsheet = SpreadsheetApp.openById(logSheetId);

    var files = uploadFolder.getFiles();
    var processedCount = 0;

    console.log("Starting enhanced notary document processing...");
    console.log(
      "API Key Status: " +
        (geminiApiKey ? "✅ Configured" : "❌ Not configured"),
    );
    console.log(
      "NEW Filename Format: YYYY MM DD Doc. NUMBER Title (Affiant).extension",
    );

    while (files.hasNext()) {
      var file = files.next();
      var originalName = file.getName();

      console.log("Processing file: " + originalName);

      if (isDuplicateFile(masterSpreadsheet, file)) {
        console.log(
          "Skipping duplicate/already processed file: " + originalName,
        );
        continue;
      }

      var extension = getFileExtension(originalName);
      if (!isValidExtension(extension)) {
        console.log("Skipping file with invalid extension: " + originalName);
        continue;
      }

      try {
        var aiAnalysis = null;
        var notaryDate = "";
        var docNumber = "";
        var documentTitle = "";
        var signatories = "";
        var documentSummary = "";
        var contentType = "";

        var competentEvidence = "";
        var idNumber = "";
        var notarialAct = "";

        if (geminiApiKey && geminiApiKey !== "") {
          try {
            console.log("Analyzing notary document with Gemini AI...");
            aiAnalysis = analyzeRefinedNotaryDocumentWithGemini(
              file,
              geminiApiKey,
            );

            if (aiAnalysis) {
              notaryDate = aiAnalysis.notaryDate || getCurrentDate();
              docNumber = aiAnalysis.docNumber || generateDocNumber();
              documentTitle =
                aiAnalysis.documentTitle ||
                extractTitleFromFilename(originalName);
              signatories = aiAnalysis.signatories || "unknown_signatory";
              documentSummary = aiAnalysis.summary || "";
              contentType = aiAnalysis.contentType || "";

              competentEvidence = aiAnalysis.competentEvidence || "";
              idNumber = aiAnalysis.idNumber || "";
              notarialAct = aiAnalysis.notarialAct || "";

              console.log("Enhanced AI analysis successful");
            }
          } catch (aiError) {
            console.log(
              "AI analysis failed, using fallback: " + aiError.message,
            );
            notaryDate = extractDateFromFilename(originalName);
            docNumber =
              extractDocNumberFromFilename(originalName) || generateDocNumber();
            documentTitle = extractTitleFromFilename(originalName);
            signatories = "Unknown Signatory";
          }
        } else {
          console.log(
            "Gemini API key not provided, using filename-based extraction",
          );
          notaryDate = extractDateFromFilename(originalName);
          docNumber =
            extractDocNumberFromFilename(originalName) || generateDocNumber();
          documentTitle = extractTitleFromFilename(originalName);
          signatories = "Unknown Signatory";
        }

        // MODIFIED: Use new formatting functions for the NEW filename convention
        var documentTitleForFile =
          formatDocumentTitleForFilename(documentTitle);
        var documentTitleForDisplay =
          formatDocumentTitleForDisplay(documentTitle);
        var signatoriesForFile = formatSignatoriesForFilename(signatories);
        var signatoriesForDisplay = formatSignatoriesForDisplay(signatories);

        // MODIFIED: Generate new filename with NEW format
        var newName = generateRefinedNotaryFilename(
          notaryDate,
          docNumber,
          documentTitleForFile,
          signatoriesForFile,
          extension,
        );

        console.log("Renaming '" + originalName + "' to '" + newName + "'");
        console.log("NEW Format: YYYY MM DD Doc. NUMBER Title (Affiant)");

        file.setName(newName);

        var monthlySheet = getOrCreateMonthlySheet(
          masterSpreadsheet,
          notaryDate,
        );

        logRefinedNotarySuccess(
          monthlySheet,
          file,
          newName,
          notaryDate,
          docNumber,
          documentTitleForDisplay,
          signatoriesForDisplay,
          competentEvidence,
          idNumber,
          notarialAct,
          documentSummary,
        );
        processedCount++;

        console.log("Successfully processed: " + newName);
      } catch (fileError) {
        console.error(
          "Error processing file '" + originalName + "': " + fileError.message,
        );
        var errorSheet = getOrCreateMonthlySheet(
          masterSpreadsheet,
          getCurrentDate(),
        );
        logError(errorSheet, file, originalName, fileError.message);
      }
    }

    console.log("Processing complete. Files processed: " + processedCount);

    var ui = SpreadsheetApp.getUi();
    ui.alert(
      "✅ Processing Complete",
      "Successfully processed " +
        processedCount +
        " file(s).\n\n" +
        "Check the monthly sheets for details.",
      ui.ButtonSet.OK,
    );
  } catch (mainError) {
    console.error("Main error: " + mainError.message);

    try {
      var masterSpreadsheet = SpreadsheetApp.openById(logSheetId);
      var errorSheet = getOrCreateMonthlySheet(
        masterSpreadsheet,
        getCurrentDate(),
      );
      errorSheet.appendRow([
        "",
        getCurrentDate(),
        "",
        "",
        "",
        "",
        "",
        "",
        "SYSTEM ERROR: " + mainError.message,
      ]);

      var ui = SpreadsheetApp.getUi();
      ui.alert(
        "❌ Processing Error",
        "An error occurred: " + mainError.message,
        ui.ButtonSet.OK,
      );
    } catch (logError) {
      console.error("Could not log error to sheet: " + logError.message);
    }
  }
}

// ============================================================================
// MODIFIED FILENAME GENERATION (NEW FORMAT)
// ============================================================================

/**
 * MODIFIED: Generate filename with NEW convention
 * Format: YYYY MM DD Doc. NUMBER Most Descriptive Title (Affiant).extension
 * Example: 2025 01 15 Doc. 12345 Affidavit Of Loss (John Doe).pdf
 */
function generateRefinedNotaryFilename(
  notaryDate,
  docNumber,
  documentTitle,
  signatories,
  extension,
) {
  // Convert date from YYYY-MM-DD to YYYY MM DD (spaces instead of hyphens)
  var dateForFilename = notaryDate.replace(/-/g, " "); // YYYY MM DD

  // Clean doc number (remove any non-alphanumeric)
  var cleanDocNumber = docNumber
    .toString()
    .replace(/[^a-zA-Z0-9]/g, "")
    .substring(0, 15);

  // Document title and signatories are already formatted by helper functions
  var cleanTitle = documentTitle;
  var affiant = signatories;

  // Build filename: "YYYY MM DD Doc. NUMBER Title (Affiant).extension"
  var baseFilename =
    dateForFilename +
    " Doc. " +
    cleanDocNumber +
    " " +
    cleanTitle +
    " (" +
    affiant +
    ")";

  // Google Drive filename limit is 255 characters including extension
  var maxLength = 250 - extension.length;

  if (baseFilename.length > maxLength) {
    // Truncate intelligently: preserve date, doc number, and affiant; shorten title
    var fixedPart = dateForFilename + " Doc. " + cleanDocNumber + " ";
    var suffixPart = " (" + affiant + ")";
    var availableForTitle = maxLength - fixedPart.length - suffixPart.length;

    if (availableForTitle > 20) {
      cleanTitle = cleanTitle.substring(0, availableForTitle);
      baseFilename = fixedPart + cleanTitle + suffixPart;
    } else {
      // If still too long, truncate affiant as well
      cleanTitle = cleanTitle.substring(0, 50);
      affiant = affiant.substring(0, 30);
      baseFilename = fixedPart + cleanTitle + " (" + affiant + ")";
      baseFilename = baseFilename.substring(0, maxLength);
    }
  }

  console.log("Generated filename: " + baseFilename + "." + extension);
  return baseFilename + "." + extension;
}

/**
 * NEW: Format document title for use in filename (with spaces, proper capitalization)
 * Converts underscored titles to readable format with title case
 */
function formatDocumentTitleForFilename(title) {
  if (!title || title === "" || title === null || title === undefined) {
    return "Unknown Document";
  }

  try {
    return title
      .toString()
      .replace(/[<>:"/\\|?*]/g, "") // Remove file system invalid characters
      .replace(/_+/g, " ") // Replace underscores with spaces
      .replace(/\s+/g, " ") // Normalize multiple spaces
      .trim() // Remove leading/trailing spaces
      .toLowerCase()
      .split(" ")
      .map(function (word) {
        // Capitalize each word for readability (Title Case)
        if (word.length > 0) {
          return word.charAt(0).toUpperCase() + word.slice(1);
        }
        return word;
      })
      .join(" ")
      .substring(0, 100); // Limit length but allow descriptive titles
  } catch (error) {
    console.log(
      "Error formatting document title for filename: " + error.message,
    );
    return "Unknown Document";
  }
}

/**
 * NEW: Format signatories for use in filename (parenthetical format)
 * Handles multiple signatories by joining with "and"
 */
function formatSignatoriesForFilename(signatories) {
  if (
    !signatories ||
    signatories === "" ||
    signatories === null ||
    signatories === undefined
  ) {
    return "Unknown Affiant";
  }

  try {
    // Clean and split signatories
    var cleanedSignatories = signatories
      .toString()
      .replace(/_and_/g, ", ") // Convert "_and_" to comma
      .replace(/_/g, " ") // Convert underscores to spaces
      .replace(/[^a-zA-Z0-9\s,.]/g, "") // Remove special chars except comma and period
      .replace(/\s+/g, " ") // Normalize spaces
      .replace(/,\s*/g, ", ") // Ensure proper comma spacing
      .trim();

    // Split by comma to get individual names
    var names = cleanedSignatories.split(/,\s*/);

    // Format each name with proper capitalization (Title Case)
    var formattedNames = names
      .map(function (name) {
        return name
          .trim()
          .toLowerCase()
          .split(/\s+/)
          .map(function (word) {
            if (word.length > 0) {
              return word.charAt(0).toUpperCase() + word.slice(1);
            }
            return word;
          })
          .join(" ");
      })
      .filter(function (name) {
        return name.length > 0;
      });

    // Handle multiple signatories
    if (formattedNames.length === 0) {
      return "Unknown Affiant";
    } else if (formattedNames.length === 1) {
      return formattedNames[0].substring(0, 50);
    } else if (formattedNames.length === 2) {
      return (formattedNames[0] + " and " + formattedNames[1]).substring(0, 50);
    } else {
      // More than 2 signatories: use first one with "et al."
      return (formattedNames[0] + " et al.").substring(0, 50);
    }
  } catch (error) {
    console.log("Error formatting signatories for filename: " + error.message);
    return "Unknown Affiant";
  }
}

/**
 * MODIFIED: Update pattern to match NEW filename format
 * New pattern: YYYY MM DD Doc. NUMBER Title (Affiant).extension
 */
function isAlreadyProcessed(filename) {
  if (!filename) return false;

  try {
    // Pattern for our NEW processed filename format: "YYYY MM DD Doc. NUMBER Title (Affiant).extension"
    // Example: 2025 01 15 Doc. 12345 Affidavit Of Loss (John Doe).pdf
    var processedPattern = /^\d{4} \d{2} \d{2} Doc\. \d+.+\(.+\)\.\w+$/i;

    var isProcessed = processedPattern.test(filename);

    if (isProcessed) {
      console.log("File matches NEW processed pattern: " + filename);
    }

    return isProcessed;
  } catch (error) {
    console.log("Error checking if file is processed: " + error.message);
    return false;
  }
}

// ============================================================================
// RETAIN ALL EXISTING FUNCTIONS BELOW - NO CHANGES
// ============================================================================

function getOrCreateMonthlySheet(spreadsheet, notaryDate) {
  try {
    var date = new Date(notaryDate);
    var monthNumber = date.getMonth();
    var year = date.getFullYear();
    var monthAbbr = getMonthAbbreviation(monthNumber);
    var sheetName = monthAbbr + "-" + year;

    console.log("Looking for or creating sheet: " + sheetName);

    var sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      console.log("Creating new monthly sheet: " + sheetName);
      sheet = spreadsheet.insertSheet(sheetName);

      sheet.appendRow([
        "Document No.",
        "Date",
        "Document Name",
        "Link",
        "Signatories",
        "Competent Evidence of Identity",
        "ID Number",
        "Notarial Act",
        "Summary of Description",
      ]);

      var headerRange = sheet.getRange(1, 1, 1, 9);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#e0e0e0");

      console.log("New monthly sheet created with headers: " + sheetName);
    }

    return sheet;
  } catch (error) {
    console.error("Error getting/creating monthly sheet: " + error.message);
    return spreadsheet.getActiveSheet();
  }
}

function getMonthNumber(monthName) {
  if (!monthName) return null;

  var months = {
    january: "01",
    jan: "01",
    february: "02",
    feb: "02",
    march: "03",
    mar: "03",
    april: "04",
    apr: "04",
    may: "05",
    june: "06",
    jun: "06",
    july: "07",
    jul: "07",
    august: "08",
    aug: "08",
    september: "09",
    sep: "09",
    sept: "09",
    october: "10",
    oct: "10",
    november: "11",
    nov: "11",
    december: "12",
    dec: "12",
  };

  return months[monthName.toLowerCase()] || null;
}

function analyzeRefinedNotaryDocumentWithGemini(file, apiKey) {
  try {
    var extension = getFileExtension(file.getName()).toLowerCase();
    var mimeType = file.getBlob().getContentType();

    var payload = {
      contents: [
        {
          parts: [
            {
              text: `Analyze this notary document thoroughly and extract the following detailed information:

1. Notary Date: CRITICAL - Find the EXACT date when the document was notarized. Look specifically in these locations in order of priority:
   a) Text containing "In witness whereof" - the date appears after this phrase
   b) Text containing "subscribed and sworn" or "subscribed" - the date appears before or after this phrase
   c) Text containing "before me" - the date appears before or after this phrase
   d) Below the affiant's name/signature but before the notary's name
   e) In the jurat or acknowledgment clause before the notary's signature
   The notary date is NOT the document creation date. Format as YYYY-MM-DD.

2. Document Number: Look for "Doc. No", "Document No.", "Doc No.", or similar numbering patterns.

3. Document Title: Extract the complete, official title/name of the document (e.g., "Affidavit of Loss", "Power of Attorney", "Deed of Sale", etc.).

4. Signatories: IMPORTANT - Find all people who signed the document as parties/affiants, but EXCLUDE the notary public's name.
   - Include: Document signers, affiants, parties, witnesses (but not the notary)
   - Exclude: The name that appears above "Notary Public" title - this is the notary's name, not a signatory
   - Look for signature lines, "Affiant:", party names, but skip the notary section entirely

5. Competent Evidence of Identity: What type of ID was used for verification (e.g., "Driver's License", "Passport", "Government ID", etc.).

6. ID Number: The identification number from the ID used.

7. Notarial Act: CRITICAL - Determine the type of notarial act based on the EXACT wording:
   **JURAT**: Look for phrases that start with "subscribed and sworn to" or "sworn to and subscribed" or contain "oath" or "affirmation"
   **ACKNOWLEDGMENT**: Look for phrases that start with "before me" or contain "acknowledged" or "appeared before me"
   Examples:
   - "Subscribed and sworn to before me this..." = JURAT
   - "Before me, the undersigned notary..." = ACKNOWLEDGMENT
   - "Acknowledged before me this..." = ACKNOWLEDGMENT
   - "Sworn to and subscribed..." = JURAT

8. Summary: Provide a brief 1-2 sentence description of what the document is about.

CRITICAL: When identifying signatories, DO NOT include the notary public's name. The notary's name typically appears directly above the "Notary Public" title or stamp. Only include the actual parties to the document (affiants, grantors, grantees, etc.).

Format your response as JSON:
{
  "notaryDate": "YYYY-MM-DD",
  "docNumber": "document_number_found",
  "documentTitle": "Official_Document_Title",
  "signatories": "parties_and_affiants_only_exclude_notary",
  "competentEvidence": "type_of_id_used",
  "idNumber": "id_number_from_document",
  "notarialAct": "Jurat or Acknowledgment",
  "summary": "Brief 1-2 sentence summary of the document purpose",
  "contentType": "notary_document_type"
}

Be thorough and accurate in extracting this information. Use underscores for multi-word entries where appropriate.`,
            },
          ],
        },
      ],
    };

    if (
      extension === "pdf" ||
      ["jpg", "jpeg", "png"].indexOf(extension) !== -1
    ) {
      var blob = file.getBlob();
      var base64Data = Utilities.base64Encode(blob.getBytes());

      payload.contents[0].parts.push({
        inline_data: {
          mime_type: mimeType,
          data: base64Data,
        },
      });
    } else if (["txt", "doc", "docx"].indexOf(extension) !== -1) {
      var textContent = file.getBlob().getDataAsString();
      if (textContent.length > 30000) {
        textContent = textContent.substring(0, 30000) + "...";
      }
      payload.contents[0].parts.push({
        text: "Document content:\n" + textContent,
      });
    } else {
      throw new Error("Unsupported file type for AI analysis: " + extension);
    }

    var options = {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      payload: JSON.stringify(payload),
    };

    var geminiModel = getGeminiModel();
    var url = buildGeminiGenerateContentUrl(apiKey);

    console.log(
      "Sending request to Gemini API for refined notary analysis using model: " +
        geminiModel,
    );
    var response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() !== 200) {
      throw new Error(
        "API request failed: " +
          response.getResponseCode() +
          " - " +
          response.getContentText(),
      );
    }

    var responseData = JSON.parse(response.getContentText());

    if (
      !responseData.candidates ||
      !responseData.candidates[0] ||
      !responseData.candidates[0].content
    ) {
      throw new Error("Invalid API response structure");
    }

    var aiText = responseData.candidates[0].content.parts[0].text;
    console.log("AI Response: " + aiText);

    try {
      var cleanText = aiText.replace(/``````\n?/g, "").trim();
      var aiResult = JSON.parse(cleanText);

      var notaryDate = aiResult.notaryDate || getCurrentDate();
      var docNumber = aiResult.docNumber || generateDocNumber();
      var documentTitle = cleanDocumentTitle(
        aiResult.documentTitle || "unknown_document",
      );
      var signatories = aiResult.signatories || "Unknown Signatory";
      var competentEvidence = aiResult.competentEvidence || "";
      var idNumber = aiResult.idNumber || "";
      var notarialAct = validateAndCategorizeNotarialAct(aiResult.notarialAct);

      return {
        notaryDate: validateDate(notaryDate) ? notaryDate : getCurrentDate(),
        docNumber:
          docNumber.replace(/[^a-zA-Z0-9_-]/g, "") || generateDocNumber(),
        documentTitle: documentTitle,
        signatories: signatories,
        competentEvidence: competentEvidence,
        idNumber: idNumber,
        notarialAct: notarialAct,
        summary: aiResult.summary || "",
        contentType: aiResult.contentType || "notary_document",
      };
    } catch (parseError) {
      console.log("Failed to parse AI response as JSON: " + parseError.message);

      var dateMatch = aiText.match(
        /notaryDate['":":\s]*['"]?(\d{4}-\d{2}-\d{2})['"]?/i,
      );
      var docMatch = aiText.match(/docNumber['":":\s]*['"]?([^'"]+)['"]?/i);
      var titleMatch = aiText.match(
        /documentTitle['":":\s]*['"]?([^'"]+)['"]?/i,
      );
      var signatoriesMatch = aiText.match(
        /signatories['":":\s]*['"]?([^'"]+)['"]?/i,
      );
      var evidenceMatch = aiText.match(
        /competentEvidence['":":\s]*['"]?([^'"]+)['"]?/i,
      );
      var idMatch = aiText.match(/idNumber['":":\s]*['"]?([^'"]+)['"]?/i);
      var actMatch = aiText.match(/notarialAct['":":\s]*['"]?([^'"]+)['"]?/i);

      return {
        notaryDate: dateMatch ? dateMatch[1] : getCurrentDate(),
        docNumber: docMatch
          ? docMatch[1].replace(/[^a-zA-Z0-9_-]/g, "") || generateDocNumber()
          : generateDocNumber(),
        documentTitle: titleMatch
          ? cleanDocumentTitle(titleMatch[1])
          : "unknown_document",
        signatories: signatoriesMatch
          ? signatoriesMatch[1]
          : "Unknown Signatory",
        competentEvidence: evidenceMatch ? evidenceMatch[1] : "",
        idNumber: idMatch ? idMatch[1] : "",
        notarialAct: validateAndCategorizeNotarialAct(
          actMatch ? actMatch[1] : "",
        ),
        summary: "",
        contentType: "notary_document",
      };
    }
  } catch (error) {
    console.error("Refined Gemini notary analysis error: " + error.message);
    return null;
  }
}

function validateAndCategorizeNotarialAct(notarialActText) {
  if (!notarialActText || notarialActText === "") {
    console.log("No notarial act text provided, defaulting to Acknowledgment");
    return "Acknowledgment";
  }

  var lowerText = notarialActText.toString().toLowerCase();

  console.log("Analyzing notarial act text: " + notarialActText);

  var juratPatterns = [
    /^subscribed\s+and\s+sworn/i,
    /^sworn\s+and\s+subscribed/i,
    /^sworn\s+to\s+and\s+subscribed/i,
    /subscribed\s+and\s+sworn\s+to/i,
    /oath/i,
    /affirmation/i,
    /swear/i,
    /sworn/i,
  ];

  var acknowledgmentPatterns = [
    /^before\s+me/i,
    /acknowledged/i,
    /appeared\s+before\s+me/i,
    /personally\s+appeared/i,
    /came\s+before\s+me/i,
  ];

  for (var i = 0; i < juratPatterns.length; i++) {
    if (juratPatterns[i].test(lowerText)) {
      console.log("Categorized as JURAT based on pattern: " + juratPatterns[i]);
      return "Jurat";
    }
  }

  for (var j = 0; j < acknowledgmentPatterns.length; j++) {
    if (acknowledgmentPatterns[j].test(lowerText)) {
      console.log(
        "Categorized as ACKNOWLEDGMENT based on pattern: " +
          acknowledgmentPatterns[j],
      );
      return "Acknowledgment";
    }
  }

  if (lowerText === "jurat" || lowerText === "jurats") {
    console.log("Direct match found: Jurat");
    return "Jurat";
  }

  if (
    lowerText === "acknowledgment" ||
    lowerText === "acknowledgement" ||
    lowerText === "acknowledgments" ||
    lowerText === "acknowledgements"
  ) {
    console.log("Direct match found: Acknowledgment");
    return "Acknowledgment";
  }

  console.log(
    "No clear pattern found, defaulting to Acknowledgment for: " +
      notarialActText,
  );
  return "Acknowledgment";
}

function cleanDocumentTitle(title) {
  if (!title || title === "") {
    return "unknown_document";
  }
  return title
    .toString()
    .replace(/[<>:"/\\|?*]/g, "")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "")
    .toLowerCase()
    .substring(0, 80);
}

function formatDocumentTitleForDisplay(documentTitle) {
  if (
    !documentTitle ||
    documentTitle === "" ||
    documentTitle === null ||
    documentTitle === undefined
  ) {
    return "Unknown Document";
  }
  try {
    return documentTitle
      .toString()
      .replace(/_+/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase()
      .split(" ")
      .map(function (word) {
        return word.charAt(0).toUpperCase() + word.slice(1);
      })
      .join(" ");
  } catch (error) {
    console.log(
      "Error formatting document title for display: " + error.message,
    );
    return "Unknown Document";
  }
}

function formatSignatoriesForDisplay(signatories) {
  if (
    !signatories ||
    signatories === "" ||
    signatories === null ||
    signatories === undefined
  ) {
    return "Unknown Signatory";
  }
  try {
    var cleanedSignatories = signatories
      .toString()
      .replace(/_and_/g, ", ")
      .replace(/_/g, " ")
      .replace(/[^a-zA-Z0-9\s,.]/g, "")
      .replace(/\s+/g, " ")
      .replace(/,\s*/g, ", ")
      .trim();

    return cleanedSignatories
      .split(/,\s*/)
      .map(function (name) {
        return name
          .trim()
          .toLowerCase()
          .split(/\s+/)
          .map(function (word) {
            if (word.length > 0) {
              return word.charAt(0).toUpperCase() + word.slice(1);
            }
            return word;
          })
          .join(" ");
      })
      .filter(function (name) {
        return name.length > 0;
      })
      .join(", ")
      .substring(0, 100);
  } catch (error) {
    console.log("Error formatting signatories for display: " + error.message);
    return "Unknown Signatory";
  }
}

function extractTitleFromFilename(filename) {
  var nameWithoutExt = filename.replace(/\.[^/.]+$/, "");

  var cleanName = nameWithoutExt
    .replace(/\d{4}[-_]\d{2}[-_]\d{2}/g, "")
    .replace(/\d{8}/g, "")
    .replace(/\d{2}[-_]\d{2}[-_]\d{4}/g, "")
    .replace(/docno[\s._-]*\d+/gi, "")
    .replace(/doc[\s._-]*no[\s._-]*\d+/gi, "")
    .replace(/document[\s._-]*no[\s._-]*\d+/gi, "")
    .replace(/doc[\s._-]*\d+/gi, "")
    .replace(/no[\s._-]*\d+/gi, "")
    .replace(/[^a-zA-Z\s\-_]/g, " ")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");

  return cleanName.toLowerCase() || "unknown_document";
}

function generateDocNumber() {
  return new Date().getTime().toString().slice(-6);
}

function validateDate(dateString) {
  var regex = /^\d{4}-\d{2}-\d{2}$/;
  if (!regex.test(dateString)) return false;

  var date = new Date(dateString);
  return date instanceof Date && !isNaN(date);
}

function extractDocNumberFromFilename(filename) {
  var docPatterns = [
    /docno[\s._-]*(\d+)/i,
    /doc[\s._-]*no[\s._-]*(\d+)/i,
    /document[\s._-]*no[\s._-]*(\d+)/i,
    /doc[\s._-]*(\d+)/i,
    /no[\s._-]*(\d+)/i,
  ];

  for (var i = 0; i < docPatterns.length; i++) {
    var match = filename.match(docPatterns[i]);
    if (match) {
      return match[1];
    }
  }

  return null;
}

function getFileExtension(filename) {
  var parts = filename.split(".");
  return parts.length > 1 ? parts.pop().toLowerCase() : "";
}

function isValidExtension(extension) {
  var validExtensions = ["pdf", "jpg", "jpeg", "png", "txt", "doc", "docx"];
  return validExtensions.indexOf(extension) !== -1;
}

function getMonthAbbreviation(monthNumber) {
  var monthAbbreviations = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];

  if (monthNumber >= 0 && monthNumber < 12) {
    return monthAbbreviations[monthNumber];
  }

  return monthAbbreviations[new Date().getMonth()];
}

function logRefinedNotarySuccess(
  sheet,
  file,
  newName,
  notaryDate,
  docNumber,
  documentTitle,
  signatories,
  competentEvidence,
  idNumber,
  notarialAct,
  summary,
) {
  try {
    var fileId = file.getId();
    var fileUrl = "https://drive.google.com/file/d/" + fileId + "/view";
    var hyperlinkFormula = '=HYPERLINK("' + fileUrl + '","view file")';

    var newRowData = [
      docNumber,
      notaryDate,
      documentTitle,
      hyperlinkFormula,
      signatories,
      competentEvidence,
      idNumber,
      notarialAct,
      summary,
    ];

    console.log(
      "Preparing to log document: " +
        docNumber +
        " to sheet: " +
        sheet.getName(),
    );

    var insertPosition = findInsertPosition(sheet, docNumber, notaryDate);

    console.log(
      "Determined insert position: " +
        insertPosition +
        " for Doc# " +
        docNumber,
    );

    if (insertPosition > sheet.getLastRow()) {
      sheet.appendRow(newRowData);
      console.log(
        "Appended row at the end (position " +
          insertPosition +
          ") for Doc# " +
          docNumber +
          " on " +
          notaryDate,
      );
    } else {
      sheet.insertRowAfter(insertPosition - 1);
      var range = sheet.getRange(insertPosition, 1, 1, newRowData.length);
      range.setValues([newRowData]);
      console.log(
        "Inserted new row at position " +
          insertPosition +
          " for Doc# " +
          docNumber +
          " on " +
          notaryDate,
      );
    }

    try {
      var verificationRange = sheet.getRange(insertPosition, 1, 1, 2);
      var verificationData = verificationRange.getValues()[0];
      console.log(
        "Verification - Row " +
          insertPosition +
          " now contains: Doc# " +
          verificationData[0] +
          ", Date: " +
          verificationData[1],
      );
    } catch (verifyError) {
      console.log("Could not verify insertion, but operation likely succeeded");
    }

    console.log(
      "Successfully logged document to sheet in numerical chronological order: " +
        newName,
    );
  } catch (error) {
    console.error("Failed to log success to sheet: " + error.message);
    try {
      sheet.appendRow(newRowData);
      console.log("Fallback: Appended row to end of sheet");
    } catch (fallbackError) {
      console.error("Fallback append also failed: " + fallbackError.message);
    }
  }
}

function findInsertPosition(sheet, docNumber, notaryDate) {
  try {
    var lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      console.log("Sheet is empty, inserting at row 2");
      return 2;
    }

    var dataRange = sheet.getRange(2, 1, lastRow - 1, 2);
    var existingData = dataRange.getValues();

    var newDate = new Date(notaryDate);
    var newDocNum = parseDocNumber(docNumber);

    console.log(
      "Finding insert position for Doc# " +
        docNumber +
        " (parsed: " +
        newDocNum +
        ") on " +
        notaryDate,
    );

    for (var i = 0; i < existingData.length; i++) {
      var existingDocNumber = existingData[i][0];
      var existingDate = existingData[i][1];

      if (!existingDocNumber || !existingDate) {
        console.log("Skipping empty row at position " + (i + 2));
        continue;
      }

      var existingDateObj = new Date(existingDate);
      var existingDocNum = parseDocNumber(existingDocNumber);

      console.log(
        "Comparing with existing Doc# " +
          existingDocNumber +
          " (parsed: " +
          existingDocNum +
          ") on " +
          existingDate,
      );

      if (newDocNum < existingDocNum) {
        console.log(
          "Inserting before row " +
            (i + 2) +
            " due to lower doc number (" +
            newDocNum +
            " < " +
            existingDocNum +
            ")",
        );
        return i + 2;
      } else if (newDocNum === existingDocNum) {
        if (newDate < existingDateObj) {
          console.log(
            "Inserting before row " +
              (i + 2) +
              " due to earlier date with same doc number",
          );
          return i + 2;
        } else if (newDate.getTime() === existingDateObj.getTime()) {
          console.log(
            "WARNING: Potential duplicate detected - same doc number and date",
          );
          return i + 2;
        }
      }
    }

    console.log(
      "Inserting at the end (row " +
        (lastRow + 1) +
        ") - highest doc number so far",
    );
    return lastRow + 1;
  } catch (error) {
    console.error("Error finding insert position: " + error.message);
    return sheet.getLastRow() + 1;
  }
}

function parseDocNumber(docNumber) {
  if (!docNumber) return 0;

  try {
    var docStr = docNumber.toString().trim();

    console.log("Parsing document number: '" + docStr + "'");

    var patterns = [
      /^Doc#(\d+)$/i,
      /^DOC(\d+)$/i,
      /^Document[\s#]?(\d+)$/i,
      /^(\d+)$/,
      /Doc[#\s]*(\d+)/i,
      /(\d+)/,
    ];

    for (var i = 0; i < patterns.length; i++) {
      var match = docStr.match(patterns[i]);
      if (match && match[1]) {
        var num = parseInt(match[1], 10);
        if (!isNaN(num)) {
          console.log("Successfully parsed '" + docStr + "' to number: " + num);
          return num;
        }
      }
    }

    var numStr = docStr.replace(/[^0-9]/g, "");
    if (numStr.length > 0) {
      var num = parseInt(numStr, 10);
      if (!isNaN(num)) {
        console.log("Extracted number from '" + docStr + "': " + num);
        return num;
      }
    }

    console.log(
      "Could not parse document number '" + docStr + "', defaulting to 0",
    );
    return 0;
  } catch (error) {
    console.log(
      "Error parsing doc number '" + docNumber + "': " + error.message,
    );
    return 0;
  }
}

function logError(sheet, file, originalName, errorMessage) {
  try {
    var hyperlinkFormula = "";
    if (file) {
      var fileId = file.getId();
      var fileUrl = "https://drive.google.com/file/d/" + fileId + "/view";
      hyperlinkFormula = '=HYPERLINK("' + fileUrl + '","view file")';
    }

    var errorRowData = [
      "",
      getCurrentDate(),
      "",
      hyperlinkFormula,
      "",
      "",
      "",
      "",
      "ERROR: " + errorMessage,
    ];

    var insertPosition = findInsertPosition(sheet, "", getCurrentDate());

    if (insertPosition > sheet.getLastRow()) {
      sheet.appendRow(errorRowData);
    } else {
      sheet.insertRowAfter(insertPosition - 1);
      var range = sheet.getRange(insertPosition, 1, 1, errorRowData.length);
      range.setValues([errorRowData]);
    }

    console.log(
      "Successfully logged error to sheet in chronological order: " +
        originalName,
    );
  } catch (error) {
    console.error("Failed to log error to sheet: " + error.message);
    try {
      sheet.appendRow(errorRowData);
    } catch (finalError) {
      console.error("Final fallback append failed: " + finalError.message);
    }
  }
}

function getCurrentDate() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd",
  );
}

function isDuplicateFile(masterSpreadsheet, file) {
  var filename = file.getName();
  var fileId = file.getId();

  if (isAlreadyProcessed(filename)) {
    console.log("Duplicate detected by filename pattern: " + filename);
    return true;
  }

  if (isFileInSpreadsheet(masterSpreadsheet, fileId)) {
    console.log(
      "Duplicate detected by file ID in spreadsheet: " +
        filename +
        " (ID: " +
        fileId +
        ")",
    );
    return true;
  }

  return false;
}

function extractDateFromFilename(filename) {
  var datePatterns = [
    /\d{4}[-_]\d{2}[-_]\d{2}/,
    /\d{8}/,
    /\d{2}[-_]\d{2}[-_]\d{4}/,
  ];

  for (var i = 0; i < datePatterns.length; i++) {
    var match = filename.match(datePatterns[i]);
    if (match) {
      var dateStr = match[0];

      if (dateStr.length === 8 && /\d{8}/.test(dateStr)) {
        return (
          dateStr.substring(0, 4) +
          "-" +
          dateStr.substring(4, 6) +
          "-" +
          dateStr.substring(6, 8)
        );
      } else if (dateStr.length === 10) {
        return dateStr.replace(/_/g, "-");
      }
    }
  }

  console.log(
    "Warning: No date found in filename, using current date as fallback",
  );
  return getCurrentDate();
}

function isMonthlySheet(sheetName) {
  if (!sheetName) return false;

  var monthlyPattern =
    /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{4}$/;
  return monthlyPattern.test(sheetName);
}

function isFileInSpreadsheet(masterSpreadsheet, fileId) {
  try {
    var sheets = masterSpreadsheet.getSheets();

    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      var sheetName = sheet.getName();

      if (!isMonthlySheet(sheetName)) {
        continue;
      }

      var lastRow = sheet.getLastRow();
      if (lastRow <= 1) continue;

      var linkColumn = sheet.getRange(2, 4, lastRow - 1, 1).getValues();

      for (var j = 0; j < linkColumn.length; j++) {
        var linkFormula = linkColumn[j][0];
        if (linkFormula && typeof linkFormula === "string") {
          var fileIdMatch = linkFormula.match(/\/file\/d\/([a-zA-Z0-9_-]+)\//);
          if (fileIdMatch && fileIdMatch[1] === fileId) {
            console.log("File ID " + fileId + " found in sheet: " + sheetName);
            return true;
          }
        }
      }
    }

    return false;
  } catch (error) {
    console.log(
      "Error checking if file exists in spreadsheet: " + error.message,
    );
    return false;
  }
}

// Test function to verify configuration
function testEnhancedNotaryConfiguration() {
  var uploadFolderId = "1dWNkT_ujKtRN1LdVFLLmQCIPQokIiBSE";
  var logSheetId = "1cM36ob20ZDDkDzN6QRsKJ45iGsj1UA5CNhN9ASz9VxY";

  try {
    var uploadFolder = DriveApp.getFolderById(uploadFolderId);
    console.log("Upload folder accessible: " + uploadFolder.getName());

    var masterSpreadsheet = SpreadsheetApp.openById(logSheetId);
    console.log(
      "Master spreadsheet accessible: " + masterSpreadsheet.getName(),
    );

    var testDate = getCurrentDate();
    var testSheet = getOrCreateMonthlySheet(masterSpreadsheet, testDate);
    console.log("Test monthly sheet created/accessed: " + testSheet.getName());

    console.log(
      "Testing NEW filename format: YYYY MM DD Doc. NUMBER Title (Affiant)",
    );

    var testFilename = generateRefinedNotaryFilename(
      "2025-10-19",
      "12345",
      formatDocumentTitleForFilename("Affidavit of Loss"),
      formatSignatoriesForFilename("John Doe"),
      "pdf",
    );
    console.log("Generated test filename: " + testFilename);
    console.log(
      "Expected format: 2025 10 19 Doc. 12345 Affidavit Of Loss (John Doe).pdf",
    );

    console.log("Testing pattern matching...");
    console.log("Pattern matches: " + isAlreadyProcessed(testFilename));

    var geminiApiKey = getApiKey();
    if (geminiApiKey) {
      console.log("API Key found: " + geminiApiKey.substring(0, 10) + "...");
    } else {
      console.log("No API Key configured - set using menu");
    }

    console.log("Enhanced notary configuration test completed!");
  } catch (error) {
    console.error(
      "Enhanced notary configuration test failed: " + error.message,
    );
  }
}

function testNewFormatAndApiKey() {
  console.log("=== TESTING NEW FORMAT & API KEY ===\n");

  // Test 1: API Key Status
  console.log("Test 1: API Key Status");
  var apiKey = getApiKey();
  console.log(apiKey ? "✅ API Key configured" : "❌ No API Key");

  // Test 2: New Filename Format
  console.log("\nTest 2: New Filename Format");
  var test1 = generateRefinedNotaryFilename(
    "2025-10-19",
    "12345",
    formatDocumentTitleForFilename("Affidavit of Loss"),
    formatSignatoriesForFilename("John Doe"),
    "pdf",
  );
  console.log("Generated: " + test1);
  console.log(
    "Expected: 2025 10 19 Doc. 12345 Affidavit Of Loss (John Doe).pdf",
  );
  console.log(
    test1 === "2025 10 19 Doc. 12345 Affidavit Of Loss (John Doe).pdf"
      ? "✅ PASS"
      : "❌ FAIL",
  );

  // Test 3: Pattern Matching
  console.log("\nTest 3: Pattern Matching");
  console.log(
    "Matches pattern: " + (isAlreadyProcessed(test1) ? "✅ PASS" : "❌ FAIL"),
  );

  console.log("\n=== TESTS COMPLETE ===");
}
