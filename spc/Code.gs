/**
 * ═══════════════════════════════════════════════════════════════
 * AUTOMATED LIQUIDATION PROCESSING SYSTEM
 * Version: 3.0 (Production Ready)
 * Last Updated: October 29, 2025
 * Author: Atty. Mary Wendy Duran
 *
 * Features:
 * - Dynamic configuration (no hardcoded IDs)
 * - Monthly subfolder organization (MM Month format)
 * - Gemini AI-powered voucher extraction
 * - Multi-voucher support
 * - Rate limiting for API quota management
 * - Comprehensive error handling and logging
 * ═══════════════════════════════════════════════════════════════
 */

/**
 * Get system configuration from Script Properties
 * @returns {Object} Configuration object
 */
function getConfig() {
  const scriptProps = PropertiesService.getScriptProperties();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const deprecatedModelReplacements = {
    "gemini-2.0-flash": "gemini-3.5-flash",
    "gemini-2.0-flash-001": "gemini-3.5-flash",
    "gemini-2.0-flash-lite": "gemini-3.1-flash-lite",
    "gemini-2.0-flash-lite-001": "gemini-3.1-flash-lite",
    "gemini-3.1-flash-lite-preview": "gemini-3.1-flash-lite",
    "gemini-3-pro-preview": "gemini-3.1-pro-preview",
  };

  ["GEMINI_MODEL", "GEMINI_FALLBACK_MODEL"].forEach((propertyName) => {
    const currentModel = scriptProps.getProperty(propertyName);
    const replacementModel = deprecatedModelReplacements[currentModel];
    if (replacementModel) {
      scriptProps.setProperty(propertyName, replacementModel);
      Logger.log(
        `Updated ${propertyName} from deprecated ${currentModel} to ${replacementModel}`,
        "CONFIG",
      );
    }
  });

  return {
    DRIVE_FOLDER_ID: scriptProps.getProperty("DRIVE_FOLDER_ID") || "",
    SHEET_ID: spreadsheet.getId(),
    SHEET_TAB_NAME:
      scriptProps.getProperty("SHEET_TAB_NAME") ||
      "Liquidation (unreplenished)",
    LOG_SHEET_NAME: "Processing Log",
    SUPPORTED_FORMATS: [
      "pdf",
      "jpg",
      "jpeg",
      "png",
      "doc",
      "docx",
      "xls",
      "xlsx",
    ],
    MAX_VOUCHERS_PER_BATCH: 30,
    MULTI_VOUCHER_NOTE: "Part of multi-voucher file",
    DATE_SUBFOLDER_FORMAT: "MM Month",
    MAX_FILES_PER_RUN: 10,
    RATE_LIMIT_DELAY_MS: 5000, // Existing delay between files

    // NEW: Enhanced rate limiting parameters
    BATCH_SIZE: 3, // Process 3 files per batch
    DELAY_BETWEEN_BATCHES_MS: 10000, // 10 seconds between batches
    ENABLE_RATE_LIMITER: true, // Enable pre-request quota checking
    MAX_RETRIES: 5, // Maximum retry attempts for 429 errors

    DEFAULT_GEMINI_MODEL: "gemini-3.1-flash-lite",
    DEFAULT_GEMINI_FALLBACK_MODEL: "gemini-3.5-flash",
    GEMINI_MODELS: [
      "gemini-3.1-flash-lite",
      "gemini-3.5-flash",
      "gemini-3-flash-preview",
      "gemini-3.1-pro-preview",
      "gemini-2.5-flash",
      "gemini-2.5-flash-lite",
      "gemini-2.5-pro",
    ],
  };
}

function showToast(message, title = "Liquidation System", timeoutSeconds = 5) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      message,
      title,
      timeoutSeconds,
    );
  } catch (error) {
    console.log(`Toast skipped: ${message}`);
  }
}

/**
 * Create custom menu on spreadsheet open
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();

    const mainMenu = ui
      .createMenu("Main Functions")
      .addItem("▶️ Process New Files", "processNewFiles")
      .addItem("✍️ Manual Voucher Entry", "manualVoucherEntry");

    const settingsMenu = ui
      .createMenu("Settings & Monitoring")
      .addItem("⚙️ Setup System", "setupSystemWizard")
      .addItem("🔑 Set Gemini API Key", "setGeminiApiKey")
      .addItem("📁 Configure Drive Folder", "configureDriveFolder")
      .addSeparator()
      .addItem("📋 View Processing Log", "showProcessingLog")
      .addItem("ℹ️ System Info", "showSystemInfo")
      .addItem("📊 Rate Limiter Status", "testRateLimiter");

    const tweakingMenu = ui
      .createMenu("Tweaking & Maintenance")
      .addItem("🤖 Set Gemini Models", "setGeminiModel")
      .addItem("🔄 Reset Rate Limiter", "resetRateLimiter")
      .addSeparator()
      .addItem("🔄 Fresh Setup (Reset)", "freshSetup");

    ui.createMenu("Liquidation System")
      .addSubMenu(mainMenu)
      .addSubMenu(settingsMenu)
      .addSubMenu(tweakingMenu)
      .addToUi();
  } catch (error) {
    console.error("Failed to create menu:", error);
  }
}

function setGeminiModel() {
  const CONFIG = getConfig();
  const ui = SpreadsheetApp.getUi();
  const scriptProps = PropertiesService.getScriptProperties();

  const availableModels = CONFIG.GEMINI_MODELS || [];
  if (!availableModels.length) {
    ui.alert(
      "❌ No Models Configured",
      "No Gemini models are configured in the script.",
      ui.ButtonSet.OK,
    );
    return;
  }

  const currentModel =
    scriptProps.getProperty("GEMINI_MODEL") || CONFIG.DEFAULT_GEMINI_MODEL;
  const currentFallbackModel =
    scriptProps.getProperty("GEMINI_FALLBACK_MODEL") ||
    CONFIG.DEFAULT_GEMINI_FALLBACK_MODEL;
  const modelListText = availableModels
    .map((m, i) => `${i + 1}. ${m}`)
    .join("\n");

  const response = ui.prompt(
    "🤖 Select Primary Gemini Model",
    `Current primary model: ${currentModel}\nCurrent fallback model: ${currentFallbackModel}\n\nSelect PRIMARY model by number (1-${availableModels.length}):\n\n${modelListText}`,
    ui.ButtonSet.OK_CANCEL,
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const selected = response.getResponseText().trim();
  const index = parseInt(selected, 10);
  if (!index || index < 1 || index > availableModels.length) {
    ui.alert(
      "❌ Invalid Selection",
      `Please enter a number between 1 and ${availableModels.length}.`,
      ui.ButtonSet.OK,
    );
    return;
  }

  const model = availableModels[index - 1];
  scriptProps.setProperty("GEMINI_MODEL", model);

  const fallbackResponse = ui.prompt(
    "🛟 Select Fallback Gemini Model",
    `Primary model set to: ${model}\n\nSelect FALLBACK model by number (1-${availableModels.length}) or leave blank to keep current fallback:\n\n${modelListText}`,
    ui.ButtonSet.OK_CANCEL,
  );

  let fallbackModel = currentFallbackModel;
  if (fallbackResponse.getSelectedButton() === ui.Button.OK) {
    const fallbackSelected = fallbackResponse.getResponseText().trim();
    if (fallbackSelected) {
      const fallbackIndex = parseInt(fallbackSelected, 10);
      if (
        fallbackIndex &&
        fallbackIndex >= 1 &&
        fallbackIndex <= availableModels.length
      ) {
        fallbackModel = availableModels[fallbackIndex - 1];
        scriptProps.setProperty("GEMINI_FALLBACK_MODEL", fallbackModel);
      } else {
        ui.alert(
          "⚠️ Invalid Fallback Selection",
          `Keeping fallback model: ${fallbackModel}`,
          ui.ButtonSet.OK,
        );
      }
    }
  }

  ui.alert(
    "✅ Success",
    `Primary Gemini model: ${model}\nFallback Gemini model: ${fallbackModel}`,
    ui.ButtonSet.OK,
  );
  Logger.log(
    `Gemini models updated: primary=${model}, fallback=${fallbackModel}`,
    "CONFIG",
  );
}

/**
 * Manual fallback when a voucher cannot be parsed reliably by Gemini.
 */
function manualVoucherEntry() {
  const ui = SpreadsheetApp.getUi();

  try {
    const fields = [
      {
        key: "voucherNo",
        label: "Voucher Reference",
        required: true,
        example: "2025-001-C",
      },
      {
        key: "company",
        label: "Client / Company",
        required: true,
        example: "DDS",
      },
      {
        key: "errandDate",
        label: "Errand Date",
        required: true,
        example: "2025-08-17",
      },
      {
        key: "errandBy",
        label: "Staff / Errand By",
        required: false,
        example: "Shaira",
      },
      {
        key: "service",
        label: "Service",
        required: false,
        example: "Admin",
      },
      {
        key: "details",
        label: "Details",
        required: false,
        example: "Transmit letter to RTC",
      },
      {
        key: "mainLocation",
        label: "Main Location",
        required: false,
        example: "Taguig",
      },
      {
        key: "total",
        label: "Amount",
        required: true,
        example: "100",
      },
      {
        key: "expenseClassification",
        label: "Expense Classification",
        required: false,
        example: "Gas-allowance",
      },
    ];

    const rawData = {};
    for (const field of fields) {
      const response = ui.prompt(
        "✍️ Manual Voucher Entry",
        `${field.label}${field.required ? " *" : ""}\nExample: ${field.example}`,
        ui.ButtonSet.OK_CANCEL,
      );

      if (response.getSelectedButton() !== ui.Button.OK) {
        ui.alert("Manual Entry Cancelled", "No record was added.", ui.ButtonSet.OK);
        return;
      }

      const value = response.getResponseText().trim();
      if (field.required && !value) {
        ui.alert(
          "Required Field",
          `${field.label} is required. Please start manual entry again.`,
          ui.ButtonSet.OK,
        );
        return;
      }

      rawData[field.key] = value;
    }

    const validatedData = DataValidator.validateAndClean(rawData);
    const processedDate = new Date();
    const rowNumber = SheetManager_v2.addLiquidationRecord(
      validatedData,
      "Manual Entry",
      "",
      "",
      false,
      "Manual Entry",
      processedDate,
    );

    Logger.log(
      `Manual voucher entry added: ${validatedData.voucherNo} at row ${rowNumber}`,
      "SUCCESS",
    );
    ui.alert(
      "✅ Manual Entry Added",
      `Voucher ${validatedData.voucherNo} was added at row ${rowNumber}.`,
      ui.ButtonSet.OK,
    );
  } catch (error) {
    Logger.log(`Manual voucher entry failed: ${error.message}`, "ERROR");
    ui.alert("❌ Manual Entry Failed", error.message, ui.ButtonSet.OK);
  }
}

/**
 * Initial setup wizard for first-time configuration
 */
function setupSystemWizard() {
  try {
    const ui = SpreadsheetApp.getUi();
    const scriptProps = PropertiesService.getScriptProperties();
    const CONFIG = getConfig();

    const welcomeResponse = ui.alert(
      "🎉 Welcome to Liquidation System Setup",
      "This wizard will configure:\n\n1. ✅ Google Sheet (auto-detected)\n2. 📁 Drive Folder ID\n3. 🔑 Gemini API Key\n\nReady to begin?",
      ui.ButtonSet.YES_NO,
    );

    if (welcomeResponse !== ui.Button.YES) {
      return;
    }

    // Configure Drive Folder
    const currentDriveFolderId =
      scriptProps.getProperty("DRIVE_FOLDER_ID") || "";

    const drivePromptMessage = currentDriveFolderId
      ? `✅ Current Drive Folder ID:\n${currentDriveFolderId}\n\n📝 Enter NEW Drive Folder ID (or leave blank to keep current):\n\nHow to get Folder ID:\n1. Open folder in Google Drive\n2. Right-click > Get link\n3. Copy ID from URL:\n   drive.google.com/drive/folders/YOUR_FOLDER_ID`
      : `📁 Enter Google Drive Folder ID:\n\nHow to get it:\n1. Open your voucher folder in Google Drive\n2. Right-click folder > Get link\n3. Copy the ID from URL:\n   drive.google.com/drive/folders/YOUR_FOLDER_ID\n\nPaste only the ID:`;

    const driveFolderResponse = ui.prompt(
      "📁 Step 1: Configure Drive Folder",
      drivePromptMessage,
      ui.ButtonSet.OK_CANCEL,
    );

    if (driveFolderResponse.getSelectedButton() === ui.Button.CANCEL) {
      ui.alert(
        "❌ Setup Cancelled",
        "Setup was cancelled. No changes made.",
        ui.ButtonSet.OK,
      );
      return;
    }

    const driveFolderId = driveFolderResponse.getResponseText().trim();

    if (driveFolderId && driveFolderId.length > 10) {
      try {
        const folder = DriveApp.getFolderById(driveFolderId);
        scriptProps.setProperty("DRIVE_FOLDER_ID", driveFolderId);
        ui.alert(
          "✅ Drive Folder Configured",
          `Folder: ${folder.getName()}\nID: ${driveFolderId}`,
          ui.ButtonSet.OK,
        );
      } catch (error) {
        ui.alert(
          "❌ Invalid Folder ID",
          `Cannot access folder: ${driveFolderId}\n\nError: ${error.message}`,
          ui.ButtonSet.OK,
        );
        return;
      }
    } else if (!currentDriveFolderId) {
      ui.alert(
        "❌ Required",
        "Drive Folder ID is required for first-time setup!",
        ui.ButtonSet.OK,
      );
      return;
    }

    // Configure API Key
    const currentApiKey = scriptProps.getProperty("GEMINI_API_KEY") || "";
    const apiKeyStatus = currentApiKey
      ? "✅ Already configured"
      : "❌ Not configured";

    const apiKeyResponse = ui.alert(
      "🔑 Step 2: Gemini API Key",
      `Current status: ${apiKeyStatus}\n\nDo you want to ${currentApiKey ? "update" : "set"} the Gemini API Key now?`,
      ui.ButtonSet.YES_NO,
    );

    if (apiKeyResponse === ui.Button.YES) {
      const apiKeyPrompt = ui.prompt(
        "🔑 Enter Gemini API Key",
        "Get your API key from:\nhttps://aistudio.google.com/app/apikey\n\nPaste your API key below:",
        ui.ButtonSet.OK_CANCEL,
      );

      if (apiKeyPrompt.getSelectedButton() === ui.Button.OK) {
        const apiKey = apiKeyPrompt.getResponseText().trim();
        if (apiKey && apiKey.length > 20) {
          scriptProps.setProperty("GEMINI_API_KEY", apiKey);
          ui.alert("✅ Success", "Gemini API Key saved!", ui.ButtonSet.OK);
        }
      }
    }

    const availableModels = CONFIG.GEMINI_MODELS || [];
    if (availableModels.length) {
      const currentModel =
        scriptProps.getProperty("GEMINI_MODEL") || CONFIG.DEFAULT_GEMINI_MODEL;
      const currentFallbackModel =
        scriptProps.getProperty("GEMINI_FALLBACK_MODEL") ||
        CONFIG.DEFAULT_GEMINI_FALLBACK_MODEL;
      const modelListText = availableModels
        .map((m, i) => `${i + 1}. ${m}`)
        .join("\n");

      const modelResponse = ui.prompt(
        "🤖 Step 3: Primary Gemini Model",
        `Current primary model: ${currentModel}\nCurrent fallback model: ${currentFallbackModel}\n\nSelect PRIMARY model by number (1-${availableModels.length}) or leave blank to keep current:\n\n${modelListText}`,
        ui.ButtonSet.OK_CANCEL,
      );

      if (modelResponse.getSelectedButton() === ui.Button.CANCEL) {
        ui.alert(
          "❌ Setup Cancelled",
          "Setup was cancelled. No further changes made.",
          ui.ButtonSet.OK,
        );
        return;
      }

      const selected = modelResponse.getResponseText().trim();
      if (selected) {
        const index = parseInt(selected, 10);
        if (index && index >= 1 && index <= availableModels.length) {
          scriptProps.setProperty("GEMINI_MODEL", availableModels[index - 1]);
        } else {
          ui.alert(
            "⚠️ Invalid Model Selection",
            `Keeping current model: ${currentModel}`,
            ui.ButtonSet.OK,
          );
        }
      }

      const fallbackResponse = ui.prompt(
        "🛟 Step 4: Fallback Gemini Model",
        `Current fallback model: ${currentFallbackModel}\n\nSelect FALLBACK model by number (1-${availableModels.length}) or leave blank to keep current:\n\n${modelListText}`,
        ui.ButtonSet.OK_CANCEL,
      );

      if (fallbackResponse.getSelectedButton() === ui.Button.CANCEL) {
        ui.alert(
          "❌ Setup Cancelled",
          "Setup was cancelled. No further changes made.",
          ui.ButtonSet.OK,
        );
        return;
      }

      const fallbackSelected = fallbackResponse.getResponseText().trim();
      if (fallbackSelected) {
        const fallbackIndex = parseInt(fallbackSelected, 10);
        if (
          fallbackIndex &&
          fallbackIndex >= 1 &&
          fallbackIndex <= availableModels.length
        ) {
          scriptProps.setProperty(
            "GEMINI_FALLBACK_MODEL",
            availableModels[fallbackIndex - 1],
          );
        } else {
          ui.alert(
            "⚠️ Invalid Fallback Model Selection",
            `Keeping fallback model: ${currentFallbackModel}`,
            ui.ButtonSet.OK,
          );
        }
      }
    }

    // Create sheets
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const liquidationSheet =
      SheetManager_v2.createLiquidationSheet(spreadsheet);
    SheetManager_v2.setupDropdownValidation(liquidationSheet);
    Logger.setupLogSheet(spreadsheet);

    // Final status
    const finalDriveFolderId = scriptProps.getProperty("DRIVE_FOLDER_ID");
    const finalApiKey = scriptProps.getProperty("GEMINI_API_KEY");
    const finalModel =
      scriptProps.getProperty("GEMINI_MODEL") || CONFIG.DEFAULT_GEMINI_MODEL;
    const finalFallbackModel =
      scriptProps.getProperty("GEMINI_FALLBACK_MODEL") ||
      CONFIG.DEFAULT_GEMINI_FALLBACK_MODEL;

    let folderName = "Unknown";
    try {
      if (finalDriveFolderId) {
        folderName = DriveApp.getFolderById(finalDriveFolderId).getName();
      }
    } catch (e) {
      folderName = "Error accessing folder";
    }

    const setupComplete = finalDriveFolderId && finalApiKey;

    const statusMessage = `✅ SETUP COMPLETE!\n\n📊 Spreadsheet: ${spreadsheet.getName()}\n📁 Drive Folder: ${folderName}\n🔑 API Key: ${finalApiKey ? "✅ Configured" : "❌ Not set"}\n🤖 Primary Gemini Model: ${finalModel}\n🛟 Fallback Gemini Model: ${finalFallbackModel}\n📋 Sheet Structure: 21 columns (A-U)\n\n${setupComplete ? '🎉 System ready! Upload vouchers and click "Process New Files"' : "⚠️ Complete configuration required!"}`;

    ui.alert("Setup Complete", statusMessage, ui.ButtonSet.OK);
    Logger.log("System setup completed via wizard", "SETUP");
  } catch (error) {
    console.error("Setup wizard error:", error);
    SpreadsheetApp.getUi().alert(
      "❌ Setup Error",
      `Error: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

/**
 * Fresh setup - completely resets system with new configuration
 */
function freshSetup() {
  try {
    const ui = SpreadsheetApp.getUi();

    const confirmation = ui.alert(
      "⚠️ Fresh Setup",
      "This will:\n\n1. DELETE the current Liquidation sheet\n2. CREATE a new sheet with 21 columns (A-U)\n3. Reconfigure Drive Folder & API Key\n\nOld data will be LOST!\n\nContinue?",
      ui.ButtonSet.YES_NO,
    );

    if (confirmation !== ui.Button.YES) {
      return;
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const scriptProps = PropertiesService.getScriptProperties();

    // Delete old sheet
    const oldSheet = spreadsheet.getSheetByName("Liquidation (unreplenished)");
    if (oldSheet) {
      spreadsheet.deleteSheet(oldSheet);
    }

    // Get Drive Folder ID
    const driveFolderPrompt = ui.prompt(
      "📁 Step 1/2: Configure Drive Folder",
      "Enter Google Drive Folder ID:\n\nHow to get it:\n1. Open folder in Google Drive\n2. Right-click > Get link\n3. Copy ID from URL:\n   drive.google.com/drive/folders/YOUR_ID\n\nPaste ID:",
      ui.ButtonSet.OK_CANCEL,
    );

    if (driveFolderPrompt.getSelectedButton() === ui.Button.CANCEL) {
      ui.alert("❌ Setup Cancelled", "Fresh setup cancelled.", ui.ButtonSet.OK);
      return;
    }

    const driveFolderId = driveFolderPrompt.getResponseText().trim();

    if (!driveFolderId || driveFolderId.length < 10) {
      ui.alert(
        "❌ Error",
        "Invalid Drive Folder ID. Please try again.",
        ui.ButtonSet.OK,
      );
      return;
    }

    try {
      const folder = DriveApp.getFolderById(driveFolderId);
      scriptProps.setProperty("DRIVE_FOLDER_ID", driveFolderId);
    } catch (error) {
      ui.alert(
        "❌ Access Error",
        `Cannot access folder: ${driveFolderId}\n\nError: ${error.message}`,
        ui.ButtonSet.OK,
      );
      return;
    }

    // Get API Key
    const apiKeyPrompt = ui.prompt(
      "🔑 Step 2/2: Set Gemini API Key",
      "Enter Gemini API Key:\n(Get from: https://aistudio.google.com/app/apikey)",
      ui.ButtonSet.OK_CANCEL,
    );

    if (apiKeyPrompt.getSelectedButton() === ui.Button.OK) {
      const apiKey = apiKeyPrompt.getResponseText().trim();
      if (apiKey && apiKey.length > 20) {
        scriptProps.setProperty("GEMINI_API_KEY", apiKey);
      }
    }

    // Create new sheets
    const liquidationSheet =
      SheetManager_v2.createLiquidationSheet(spreadsheet);
    SheetManager_v2.setupDropdownValidation(liquidationSheet);
    Logger.setupLogSheet(spreadsheet);

    const finalFolderId = scriptProps.getProperty("DRIVE_FOLDER_ID");
    const finalApiKey = scriptProps.getProperty("GEMINI_API_KEY");

    let folderName = "Unknown";
    try {
      folderName = DriveApp.getFolderById(finalFolderId).getName();
    } catch (e) {
      folderName = "Error";
    }

    const statusMessage = `✅ FRESH SETUP COMPLETE!\n\n📊 Spreadsheet: ${spreadsheet.getName()}\n📁 Drive Folder: ${folderName}\n🔑 API Key: ${finalApiKey ? "✅ Set" : "❌ Not set"}\n📋 Sheet Columns: 21 (A-U)\n\n${finalFolderId && finalApiKey ? "🎉 System ready! Process files now." : "⚠️ Configuration incomplete!"}`;

    ui.alert("✅ Fresh Setup Complete", statusMessage, ui.ButtonSet.OK);
    Logger.log("Fresh setup completed successfully", "SETUP");
  } catch (error) {
    console.error("Fresh setup error:", error);
    SpreadsheetApp.getUi().alert(
      "❌ Error",
      `Fresh setup failed: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

/**
 * Configure Drive Folder (can be changed anytime)
 */
function configureDriveFolder() {
  try {
    const ui = SpreadsheetApp.getUi();
    const scriptProps = PropertiesService.getScriptProperties();
    const currentFolderId =
      scriptProps.getProperty("DRIVE_FOLDER_ID") || "Not configured";

    let currentFolderName = "Unknown";
    if (currentFolderId !== "Not configured") {
      try {
        currentFolderName = DriveApp.getFolderById(currentFolderId).getName();
      } catch (e) {
        currentFolderName = "Invalid/No access";
      }
    }

    const response = ui.prompt(
      "📁 Configure Drive Folder",
      `Current folder: ${currentFolderName}\nID: ${currentFolderId}\n\n📝 Enter NEW Google Drive Folder ID:\n\nHow to get ID:\n1. Open folder in Google Drive\n2. Right-click > Get link\n3. Copy ID from URL\n\nPaste ID:`,
      ui.ButtonSet.OK_CANCEL,
    );

    if (response.getSelectedButton() === ui.Button.OK) {
      const newFolderId = response.getResponseText().trim();

      if (!newFolderId || newFolderId.length < 10) {
        ui.alert(
          "❌ Error",
          "Please enter a valid Folder ID.",
          ui.ButtonSet.OK,
        );
        return;
      }

      try {
        const folder = DriveApp.getFolderById(newFolderId);
        scriptProps.setProperty("DRIVE_FOLDER_ID", newFolderId);

        ui.alert(
          "✅ Success",
          `Drive folder updated!\n\nFolder: ${folder.getName()}\nID: ${newFolderId}`,
          ui.ButtonSet.OK,
        );
        Logger.log(`Drive folder configured: ${folder.getName()}`, "CONFIG");
      } catch (error) {
        ui.alert(
          "❌ Access Error",
          `Cannot access folder: ${newFolderId}\n\nError: ${error.message}`,
          ui.ButtonSet.OK,
        );
      }
    }
  } catch (error) {
    console.error("Configure drive folder failed:", error);
    SpreadsheetApp.getUi().alert(
      "❌ Error",
      error.message,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

/**
 * Set or update Gemini API Key
 */
function setGeminiApiKey() {
  const ui = SpreadsheetApp.getUi();
  const scriptProps = PropertiesService.getScriptProperties();
  const currentKey = scriptProps.getProperty("GEMINI_API_KEY");
  const promptText = currentKey
    ? "🔑 Current API Key is set.\n\nEnter NEW Gemini API Key:\n(Get from: https://aistudio.google.com/app/apikey)"
    : "🔑 Enter Gemini API Key:\n(Get from: https://aistudio.google.com/app/apikey)";
  const response = ui.prompt(
    "Gemini API Key Setup",
    promptText,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() === ui.Button.OK) {
    const apiKey = response.getResponseText().trim();
    if (apiKey && apiKey.length > 20) {
      scriptProps.setProperty("GEMINI_API_KEY", apiKey);
      ui.alert(
        "✅ Success",
        "Gemini API Key saved successfully!",
        ui.ButtonSet.OK,
      );
      Logger.log("API Key updated", "CONFIG");
    } else {
      ui.alert("❌ Error", "Please enter a valid API key.", ui.ButtonSet.OK);
    }
  }
}

/**
 * Process new files from Drive folder
 * Main processing function with rate limiting and error handling
 */
function processNewFiles() {
  try {
    const CONFIG = getConfig();
    const ui = SpreadsheetApp.getUi();
    const scriptProps = PropertiesService.getScriptProperties();

    const driveFolderId = scriptProps.getProperty("DRIVE_FOLDER_ID");
    const apiKey = scriptProps.getProperty("GEMINI_API_KEY");

    if (!driveFolderId) {
      showToast("Drive folder is not configured yet.", "Setup Required", 8);
      ui.alert(
        "❌ Configuration Required",
        "Drive Folder not configured!\n\nPlease run:\nLiquidation System > Configure Drive Folder",
        ui.ButtonSet.OK,
      );
      return;
    }

    if (!apiKey) {
      showToast("Gemini API key is not set yet.", "Setup Required", 8);
      ui.alert(
        "❌ API Key Required",
        "Gemini API Key not set!\n\nPlease run:\nLiquidation System > Set Gemini API Key",
        ui.ButtonSet.OK,
      );
      return;
    }

    const confirmation = ui.alert(
      "📁 Process Files",
      "Start processing new files?\n\nFiles will be organized into monthly subfolders (MM Month format).",
      ui.ButtonSet.YES_NO,
    );

    if (confirmation !== ui.Button.YES) {
      return;
    }

    const files = DriveManager.getUnprocessedFiles();

    if (files.length === 0) {
      showToast("No unprocessed files were found.", "Nothing To Process", 6);
      ui.alert(
        "📄 No New Files",
        "No unprocessed files found in Drive folder.",
        ui.ButtonSet.OK,
      );
      return;
    }

    const filesToProcess = Math.min(files.length, CONFIG.MAX_FILES_PER_RUN);

    if (files.length > CONFIG.MAX_FILES_PER_RUN) {
      showToast(
        `Found ${files.length} files. Processing first ${filesToProcess}.`,
        "Batch Started",
        8,
      );
      ui.alert(
        "⚠️ Batch Processing",
        `Found ${files.length} files.\n\nProcessing first ${filesToProcess} to avoid API quota limits.\n\nRun again for remaining files.`,
        ui.ButtonSet.OK,
      );
    } else {
      showToast(
        `Processing ${filesToProcess} file(s).`,
        "Processing Started",
        6,
      );
      ui.alert(
        "⏳ Processing Started",
        `Processing ${filesToProcess} file(s)...`,
        ui.ButtonSet.OK,
      );
    }

    const processedDate = new Date();
    showToast("Preparing monthly Drive folder.", "Organizing Files", 5);
    const dateSubfolder = DriveManager.createOrGetDateSubfolder(processedDate);
    const folderUrl = DriveManager.getFolderUrl(dateSubfolder);

    let totalProcessedVouchers = 0;
    let totalErrorVouchers = 0;
    let filesProcessed = 0;
    let filesWithErrors = 0;
    let abortRun = false;

    for (let i = 0; i < filesToProcess; i++) {
      const file = files[i];
      let fileProcessedSuccessfully = false;
      let renamedFileName = file.getName();

      try {
        filesProcessed++;
        const originalFileName = file.getName();
        showToast(
          `File ${filesProcessed}/${filesToProcess}: ${originalFileName}`,
          "Processing File",
          8,
        );

        // Rate limiting
        if (filesProcessed > 1) {
          showToast("Waiting briefly to avoid Gemini quota errors.", "Rate Limit", 5);
          Utilities.sleep(CONFIG.RATE_LIMIT_DELAY_MS);
        }

        showToast(`Reading vouchers from ${originalFileName}.`, "Gemini Parsing", 8);
        const vouchersData = GeminiParser.parseMultipleVouchers(file, apiKey);

        if (vouchersData.length === 0) {
          showToast(
            `No vouchers found in ${originalFileName}.`,
            "Skipped File",
            8,
          );
          Logger.log(`No vouchers found: ${originalFileName}`, "WARNING");
          continue;
        }

        showToast(
          `Found ${vouchersData.length} voucher(s) in ${originalFileName}.`,
          "Gemini Parsing Complete",
          8,
        );

        const batches = MultiVoucherProcessor.createBatches(
          vouchersData,
          CONFIG.MAX_VOUCHERS_PER_BATCH,
        );

        let fileProcessedVouchers = 0;
        let fileErrorVouchers = 0;
        let fileUrl = "";
        let movedFile = null;

        for (let batchIndex = 0; batchIndex < batches.length; batchIndex++) {
          const batch = batches[batchIndex];

          for (
            let voucherIndex = 0;
            voucherIndex < batch.length;
            voucherIndex++
          ) {
            const absoluteVoucherIndex =
              batchIndex * CONFIG.MAX_VOUCHERS_PER_BATCH + voucherIndex;

            try {
              const voucherData = batch[voucherIndex];
              const validatedData = DataValidator.validateAndClean(voucherData);
              const uniqueVoucherData =
                MultiVoucherProcessor.handleDuplicateVouchers(
                  validatedData,
                  vouchersData,
                  absoluteVoucherIndex,
                );

              if (fileProcessedVouchers === 0) {
                showToast(
                  `Organizing source file for ${uniqueVoucherData.voucherNo || "voucher"}.`,
                  "Drive Update",
                  6,
                );
                renamedFileName = DriveManager.renameProcessedFile(
                  file,
                  uniqueVoucherData,
                  false,
                );

                try {
                  movedFile = DriveManager.moveFileToSubfolder(
                    file,
                    dateSubfolder,
                  );
                  fileUrl = DriveManager.getFileUrl(movedFile);

                  DriveManager.markFileAsProcessed(movedFile);
                  renamedFileName = movedFile.getName();

                  fileProcessedSuccessfully = true;
                } catch (moveError) {
                  throw new Error(`File move failed: ${moveError.message}`);
                }
              }

              const isMultiVoucher = vouchersData.length > 1;

              showToast(
                `Saving voucher ${absoluteVoucherIndex + 1}/${vouchersData.length}: ${uniqueVoucherData.voucherNo || "No voucher no."}`,
                "Writing To Sheet",
                6,
              );
              SheetManager_v2.addLiquidationRecord(
                uniqueVoucherData,
                renamedFileName,
                fileUrl,
                folderUrl,
                isMultiVoucher,
                originalFileName,
                processedDate,
              );

              fileProcessedVouchers++;
              showToast(
                `Saved ${uniqueVoucherData.voucherNo || "voucher"} successfully.`,
                "Voucher Saved",
                5,
              );
              Logger.log(
                `Processed voucher: ${uniqueVoucherData.voucherNo}`,
                "SUCCESS",
              );
            } catch (voucherError) {
              if (!fileProcessedSuccessfully) {
                throw voucherError;
              }

              showToast(
                `A voucher had an issue. Adding a review row.`,
                "Voucher Needs Review",
                8,
              );
              const partialData =
                MultiVoucherProcessor.createPartialVoucherData(
                  batch[voucherIndex],
                  voucherError.message,
                );

              SheetManager_v2.addLiquidationRecord(
                partialData,
                renamedFileName,
                fileUrl,
                folderUrl,
                true,
                originalFileName,
                processedDate,
              );

              fileErrorVouchers++;
              Logger.log(`Voucher error: ${voucherError.message}`, "ERROR");
            }
          }
        }

        totalProcessedVouchers += fileProcessedVouchers;
        totalErrorVouchers += fileErrorVouchers;

        showToast(
          `Completed ${originalFileName}: ${fileProcessedVouchers} saved, ${fileErrorVouchers} needs review.`,
          "File Complete",
          8,
        );
        Logger.log(`Completed: ${originalFileName}`, "INFO");
      } catch (error) {
        filesWithErrors++;
        showToast(
          `Failed: ${file.getName()}. Check Processing Log.`,
          "File Error",
          10,
        );
        Logger.log(`Failed: ${file.getName()}: ${error.message}`, "ERROR");

        const errorMessage = String(
          error && error.message ? error.message : "",
        );
        if (
          errorMessage.includes("QUOTA_ZERO") ||
          errorMessage.includes("Daily API limit reached")
        ) {
          abortRun = true;
          showToast(
            "Processing stopped early due to Gemini quota/limit issue.",
            "Quota Limit",
            10,
          );
          Logger.log(
            "Aborting run due to Gemini quota/limit issue. Please check API plan/quota and try again later.",
            "WARNING",
          );
          break;
        }
      }
    }

    const message = `✅ Processing Complete!\n\n📊 Results:\n• Files processed: ${filesProcessed}\n• Files with errors: ${filesWithErrors}\n• Vouchers processed: ${totalProcessedVouchers}\n• Voucher errors: ${totalErrorVouchers}\n• Folder: ${dateSubfolder.getName()}${abortRun ? "\n\n⚠️ Processing stopped early due to Gemini quota/limit issue." : ""}`;

    showToast(
      `Done: ${totalProcessedVouchers} voucher(s), ${filesWithErrors} file error(s).`,
      "Processing Complete",
      10,
    );
    ui.alert("Complete", message, ui.ButtonSet.OK);
  } catch (error) {
    console.error("Process error:", error);
    showToast(error.message, "Processing Error", 10);
    SpreadsheetApp.getUi().alert(
      "❌ Error",
      error.message,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

/**
 * Show processing log sheet
 */
function showProcessingLog() {
  const CONFIG = getConfig();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = spreadsheet.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (logSheet) {
    spreadsheet.setActiveSheet(logSheet);
    SpreadsheetApp.getUi().alert(
      "📋 Log",
      "Switched to Processing Log sheet.",
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  } else {
    SpreadsheetApp.getUi().alert(
      "⚠️ No Log",
      "Processing log not found. Run Setup System first.",
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

/**
 * Display system information and configuration
 */
function showSystemInfo() {
  const CONFIG = getConfig();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProps = PropertiesService.getScriptProperties();

  const driveFolderId =
    scriptProps.getProperty("DRIVE_FOLDER_ID") || "Not configured";
  const apiKeySet = scriptProps.getProperty("GEMINI_API_KEY")
    ? "✅ Set"
    : "❌ Not set";
  const modelSet =
    scriptProps.getProperty("GEMINI_MODEL") || CONFIG.DEFAULT_GEMINI_MODEL;
  const fallbackModelSet =
    scriptProps.getProperty("GEMINI_FALLBACK_MODEL") ||
    CONFIG.DEFAULT_GEMINI_FALLBACK_MODEL;

  let folderName = "Unknown";
  let folderLink = "Not available";
  if (driveFolderId !== "Not configured") {
    try {
      const folder = DriveApp.getFolderById(driveFolderId);
      folderName = folder.getName();
      folderLink = folder.getUrl();
    } catch (e) {
      folderName = "Invalid/No access";
    }
  }

  // Get rate limiter stats
  const rateStats = RateLimiterManager.getUsageStats();

  const info = `🔧 System Information

📊 Current Sheet ID: ${CONFIG.SHEET_ID}
📁 Drive Folder ID: ${driveFolderId}
🔑 Gemini API Key: ${apiKeySet}
🤖 Primary Gemini Model: ${modelSet}
🛟 Fallback Gemini Model: ${fallbackModelSet}
📋 Log Sheet: ${CONFIG.LOG_SHEET_NAME}
📑 Main Sheet: ${CONFIG.SHEET_TAB_NAME}
📅 Subfolder Format: ${CONFIG.DATE_SUBFOLDER_FORMAT}

⚙️ Processing Configuration:
⏱️ Rate Limit Delay: ${CONFIG.RATE_LIMIT_DELAY_MS}ms
📦 Max Files Per Run: ${CONFIG.MAX_FILES_PER_RUN}
📦 Batch Size: ${CONFIG.BATCH_SIZE}
⏳ Delay Between Batches: ${CONFIG.DELAY_BETWEEN_BATCHES_MS}ms
🔄 Max Retries: ${CONFIG.MAX_RETRIES}

📊 Current API Usage (Rate Limiter):
📈 Requests this minute: ${rateStats.requestsThisMinute}/${rateStats.maxRPM}
📈 Requests today: ${rateStats.requestsToday}/${rateStats.maxRPD}
📈 Tokens this minute: ${rateStats.tokensThisMinute}/${rateStats.maxTPM}

🔗 Drive Folder Link:
${folderLink}`;

  SpreadsheetApp.getUi().alert(
    "System Info",
    info,
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}
