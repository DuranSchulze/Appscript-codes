/**
 * Gemini AI Parser for Petty Cash Vouchers
 * Version: 4.0 (Enhanced with Exponential Backoff & Rate Limiting)
 * Last Updated: November 10, 2025
 */

var GeminiParser = {
  /**
   * Configuration for retry logic
   */
  getRetryConfig() {
    return {
      MAX_RETRIES: 5,
      INITIAL_BACKOFF_MS: 1000, // 1 second
      MAX_BACKOFF_MS: 60000, // 60 seconds
      BACKOFF_MULTIPLIER: 2,
    };
  },

  getModelName() {
    const modelFromProps =
      PropertiesService.getScriptProperties().getProperty("GEMINI_MODEL");
    const model = (modelFromProps || "gemini-3.1-flash-lite").trim();
    return this.replaceDeprecatedModel(this.normalizeModelName(model));
  },

  getFallbackModelName() {
    const modelFromProps =
      PropertiesService.getScriptProperties().getProperty(
        "GEMINI_FALLBACK_MODEL",
      );
    const model = (modelFromProps || "gemini-3.5-flash").trim();
    return this.replaceDeprecatedModel(this.normalizeModelName(model));
  },

  normalizeModelName(model) {
    return model.startsWith("models/") ? model.replace(/^models\//, "") : model;
  },

  replaceDeprecatedModel(model) {
    const replacements = {
      "gemini-2.0-flash": "gemini-3.5-flash",
      "gemini-2.0-flash-001": "gemini-3.5-flash",
      "gemini-2.0-flash-lite": "gemini-3.1-flash-lite",
      "gemini-2.0-flash-lite-001": "gemini-3.1-flash-lite",
      "gemini-3.1-flash-lite-preview": "gemini-3.1-flash-lite",
      "gemini-3-pro-preview": "gemini-3.1-pro-preview",
    };
    return replacements[model] || model;
  },

  getModelSequence() {
    const primaryModel = this.getModelName();
    const fallbackModel = this.getFallbackModelName();
    return [primaryModel, fallbackModel].filter((model, index, models) => {
      return model && models.indexOf(model) === index;
    });
  },

  getExtractionFields() {
    return [
      {
        key: "voucherNo",
        description:
          "voucher reference number from the top-right voucher header; keep partial values exactly as shown",
      },
      {
        key: "company",
        description: "text after COMPANY",
      },
      {
        key: "errandDate",
        description: "date after ERRAND DATE; return YYYY-MM-DD",
      },
      {
        key: "errandBy",
        description: "person in the ERRAND BY column",
      },
      {
        key: "service",
        description: "SERVICE column value",
      },
      {
        key: "details",
        description: "DETAILS column value; keep concise",
      },
      {
        key: "mainLocation",
        description: "text after Main Location of Errand",
      },
      {
        key: "total",
        description: "AMOUNT column value only; no currency symbol or comma",
      },
      {
        key: "expenseClassification",
        description: "EXPENSE CLASSIFICATION field value",
      },
    ];
  },

  getVoucherResponseSchema() {
    const voucherProperties = {};
    const fields = this.getExtractionFields();
    fields.forEach((field) => {
      voucherProperties[field.key] = {
        type: "string",
        description: field.description,
      };
    });

    return {
      type: "array",
      minItems: 1,
      items: {
        type: "object",
        properties: voucherProperties,
        required: fields.map((field) => field.key),
        propertyOrdering: fields.map((field) => field.key),
        additionalProperties: false,
      },
    };
  },

  buildTargetedExtractionPrompt() {
    const fieldList = this.getExtractionFields()
      .map((field) => `- ${field.key}: ${field.description}`)
      .join("\n");

    return `Extract petty cash voucher rows only.

Target fields:
${fieldList}

Rules:
- Search only for the target fields above.
- Do not transcribe, summarize, or describe unrelated page content.
- If a page has a summary/list and no actual voucher form rows, ignore that page.
- Return one JSON object per voucher row.
- Use empty strings for missing fields.
- Return an empty JSON array if no voucher rows are found.`;
  },

  logFileSizeGuidance(file, fileContent) {
    const sizeMb = (fileContent.sizeBytes || 0) / (1024 * 1024);
    if (!sizeMb) return;

    if (fileContent.type === "image" && sizeMb > 18) {
      Logger.log(
        `Large image detected (${sizeMb.toFixed(1)} MB): ${file.getName()}. Compress or split this file if Gemini returns size or quota errors.`,
        "WARNING",
      );
      showToast(
        `${file.getName()} is a large image. Compress/split if parsing fails.`,
        "Large File Warning",
        8,
      );
      return;
    }

    if (fileContent.type === "pdf" && sizeMb > 25) {
      Logger.log(
        `Large PDF detected (${sizeMb.toFixed(1)} MB): ${file.getName()}. Multi-page PDFs use document tokens per page; split very large batches for faster processing.`,
        "WARNING",
      );
      showToast(
        `${file.getName()} is a large PDF. Split huge batches for faster parsing.`,
        "Large File Warning",
        8,
      );
    }
  },

  maskApiKeyInText(text) {
    if (!text) return text;
    return String(text).replace(/key=([^&\s]+)/g, "key=***");
  },

  isZeroQuotaError(apiError) {
    const message =
      apiError && apiError.message ? String(apiError.message) : "";
    if (message.includes("limit: 0")) return true;
    if (apiError && apiError.status === "RESOURCE_EXHAUSTED" && message) {
      return message.includes("Quota exceeded") && message.includes("limit: 0");
    }
    return false;
  },

  /**
   * Parse multiple vouchers from a document using Gemini 2.0 Flash
   */
  parseMultipleVouchers(file, apiKey) {
    const fileContent = DriveManager.getFileContent(file);
    this.logFileSizeGuidance(file, fileContent);
    const modelSequence = this.getModelSequence();
    const errors = [];
    const prompt = this.buildTargetedExtractionPrompt();

    for (const modelName of modelSequence) {
      try {
        const result = this.callGeminiVision(
          fileContent,
          prompt,
          apiKey,
          0,
          modelName,
          true,
          true,
        );

        if (result.includes("SUMMARY_SECTION_DETECTED")) {
          throw new Error(
            "SUMMARY_SECTION_DETECTED: This file contains a summary/list, not individual vouchers",
          );
        }

        const cleanedResult = this.cleanJsonResponse(result);
        const parsedData = JSON.parse(cleanedResult);

        const vouchers = Array.isArray(parsedData) ? parsedData : [parsedData];

        console.log(
          `Gemini extracted ${vouchers.length} voucher(s) from ${file.getName()} using ${modelName}`,
        );
        Logger.log(
          `Gemini extracted ${vouchers.length} voucher(s) from ${file.getName()} using ${modelName}`,
          "INFO",
        );

        return vouchers;
      } catch (error) {
        const safeMessage = this.maskApiKeyInText(error && error.message);
        errors.push(`${modelName}: ${safeMessage}`);

        if (
          safeMessage.includes("QUOTA_ZERO") ||
          safeMessage.includes("Daily API limit")
        ) {
          break;
        }

        Logger.log(
          `Gemini model ${modelName} failed, trying fallback if available: ${safeMessage}`,
          "WARNING",
        );
      }
    }

    const message = errors.join(" | ");
    console.error("Gemini parsing error:", message);
    throw new Error(`Multi-voucher Gemini parsing failed: ${message}`);
  },

  /**
   * Call Gemini 2.0 Flash with vision capabilities + exponential backoff
   * @param {Object} fileContent - File content with type and data
   * @param {string} prompt - Prompt for Gemini
   * @param {string} apiKey - Gemini API key
   * @param {number} retryCount - Current retry attempt (internal use)
   * @returns {string} - Gemini response text
   */
  callGeminiVision(
    fileContent,
    prompt,
    apiKey,
    retryCount = 0,
    modelNameOverride = null,
    useStructuredOutput = true,
    useMediaResolution = true,
  ) {
    const modelName = modelNameOverride || this.getModelName();

    try {
      const retryConfig = this.getRetryConfig();

      // Estimate tokens (rough approximation: 1 token ≈ 4 characters)
      const estimatedTokens = Math.ceil((prompt.length + 1000) / 4); // Add buffer for image tokens

      // PRE-REQUEST RATE LIMIT CHECK
      const limitCheck = RateLimiterManager.checkLimit(estimatedTokens);
      if (!limitCheck.canProceed) {
        if (limitCheck.reason.includes("DAILY_LIMIT")) {
          throw new Error(
            `Daily API limit reached (200 requests). Quota resets at midnight PST. Wait time: ${(limitCheck.waitTime / 3600000).toFixed(1)} hours`,
          );
        }

        console.log(
          `⚠️ Rate limit pre-check: ${limitCheck.reason}, waiting ${(limitCheck.waitTime / 1000).toFixed(1)}s`,
        );
        Logger.log(`Rate limit pre-check: ${limitCheck.reason}`, "INFO");
        rateLimiterSleep(
          limitCheck.waitTime,
          `Pre-request rate limiting (${limitCheck.reason})`,
        );
      }

      // Prepare API request
      const base64Data = Utilities.base64Encode(fileContent.data.getBytes());

      const generationConfig = {
        temperature: 0,
        topK: 16,
        topP: 0.8,
        maxOutputTokens: 2048,
      };

      if (useMediaResolution) {
        generationConfig.mediaResolution = "MEDIA_RESOLUTION_MEDIUM";
      }

      if (useStructuredOutput) {
        generationConfig.responseFormat = {
          text: {
            mimeType: "application/json",
            schema: this.getVoucherResponseSchema(),
          },
        };
      }

      const payload = {
        contents: [
          {
            parts: [
              { text: prompt },
              {
                inline_data: {
                  mime_type: fileContent.mimeType,
                  data: base64Data,
                },
              },
            ],
          },
        ],
        generationConfig,
      };

      const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      };

      const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;

      console.log(
        `🌐 Gemini API request to ${modelName} (Attempt ${retryCount + 1}/${retryConfig.MAX_RETRIES + 1})`,
      );

      // Make request
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      // SUCCESS CASE
      if (responseCode === 200) {
        const jsonResponse = JSON.parse(responseText);

        // Record successful request with actual token usage
        const tokensUsed =
          jsonResponse.usageMetadata?.totalTokenCount || estimatedTokens;
        RateLimiterManager.recordRequest(tokensUsed);

        console.log(`✓ Gemini API success. Tokens used: ${tokensUsed}`);

        if (
          !jsonResponse.candidates ||
          !jsonResponse.candidates[0] ||
          !jsonResponse.candidates[0].content
        ) {
          throw new Error("Invalid response structure from Gemini API");
        }

        const textContent = jsonResponse.candidates[0].content.parts
          .filter((part) => part.text)
          .map((part) => part.text)
          .join("");

        return textContent;
      }

      // ERROR CASE - Parse error response
      let errorResponse;
      try {
        errorResponse = JSON.parse(responseText);
      } catch (parseError) {
        throw new Error(`HTTP ${responseCode}: ${responseText}`);
      }

      // HANDLE 429 (RATE LIMIT) ERRORS
      if (responseCode === 429) {
        if (this.isZeroQuotaError(errorResponse.error)) {
          const safeServerMessage = this.maskApiKeyInText(
            errorResponse.error?.message || "Quota exceeded",
          );
          throw new Error(`HTTP 429: QUOTA_ZERO: ${safeServerMessage}`);
        }

        if (retryCount >= retryConfig.MAX_RETRIES) {
          throw new Error(
            `Max retries (${retryConfig.MAX_RETRIES}) exceeded. Last error: ${errorResponse.error?.message || "Rate limit exceeded"}`,
          );
        }

        // Extract suggested retry delay from API error
        const suggestedDelay = this.extractRetryDelay(errorResponse.error);

        // Calculate exponential backoff delay
        const exponentialDelay =
          retryConfig.INITIAL_BACKOFF_MS *
          Math.pow(retryConfig.BACKOFF_MULTIPLIER, retryCount);

        // Use the larger of suggested delay or exponential backoff
        const backoffDelay = suggestedDelay
          ? Math.max(suggestedDelay, exponentialDelay)
          : exponentialDelay;

        const finalDelay = Math.min(backoffDelay, retryConfig.MAX_BACKOFF_MS);

        console.log(
          `⚠️ HTTP 429: Quota exceeded. Retry ${retryCount + 1}/${retryConfig.MAX_RETRIES}`,
        );
        console.log(
          `   Suggested: ${suggestedDelay ? (suggestedDelay / 1000).toFixed(1) : "none"}s, Using: ${(finalDelay / 1000).toFixed(1)}s`,
        );
        Logger.log(
          `HTTP 429 detected. Retry ${retryCount + 1}/${retryConfig.MAX_RETRIES} after ${(finalDelay / 1000).toFixed(1)}s`,
          "WARNING",
        );

        rateLimiterSleep(
          finalDelay,
          `Exponential backoff (attempt ${retryCount + 1})`,
        );

        // RECURSIVE RETRY
        return this.callGeminiVision(
          fileContent,
          prompt,
          apiKey,
          retryCount + 1,
          modelName,
          useStructuredOutput,
          useMediaResolution,
        );
      }

      if (responseCode === 400 && useStructuredOutput) {
        const errorMessage = errorResponse.error?.message || responseText;
        if (
          errorMessage.includes("responseFormat") ||
          errorMessage.includes("schema") ||
          errorMessage.includes("JSON mode")
        ) {
          Logger.log(
            `Structured output unavailable for ${modelName}; retrying legacy JSON prompt mode`,
            "WARNING",
          );
          return this.callGeminiVision(
            fileContent,
            prompt,
            apiKey,
            retryCount,
            modelName,
            false,
            useMediaResolution,
          );
        }
      }

      if (responseCode === 400 && useMediaResolution) {
        const errorMessage = errorResponse.error?.message || responseText;
        if (
          errorMessage.includes("mediaResolution") ||
          errorMessage.includes("media_resolution") ||
          errorMessage.includes("MediaResolution")
        ) {
          Logger.log(
            `Media resolution setting unavailable for ${modelName}; retrying without it`,
            "WARNING",
          );
          return this.callGeminiVision(
            fileContent,
            prompt,
            apiKey,
            retryCount,
            modelName,
            useStructuredOutput,
            false,
          );
        }
      }

      // HANDLE 503 (SERVICE UNAVAILABLE) AND 500 (INTERNAL SERVER ERROR)
      if (responseCode === 503 || responseCode === 500) {
        if (retryCount >= retryConfig.MAX_RETRIES) {
          throw new Error(
            `Max retries (${retryConfig.MAX_RETRIES}) exceeded. Status: ${responseCode}, Error: ${errorResponse.error?.message || "Server error"}`,
          );
        }

        const backoffDelay = Math.min(
          retryConfig.INITIAL_BACKOFF_MS *
            Math.pow(retryConfig.BACKOFF_MULTIPLIER, retryCount),
          retryConfig.MAX_BACKOFF_MS,
        );

        console.log(
          `⚠️ HTTP ${responseCode}: Transient error. Retry ${retryCount + 1}/${retryConfig.MAX_RETRIES} in ${(backoffDelay / 1000).toFixed(1)}s`,
        );
        Logger.log(
          `HTTP ${responseCode} transient error. Retry ${retryCount + 1}`,
          "WARNING",
        );

        rateLimiterSleep(backoffDelay, "Transient error backoff");

        return this.callGeminiVision(
          fileContent,
          prompt,
          apiKey,
          retryCount + 1,
          modelName,
          useStructuredOutput,
          useMediaResolution,
        );
      }

      // NON-RETRYABLE ERROR (400, 401, 403, etc.)
      const errorMessage = errorResponse.error?.message || responseText;
      throw new Error(
        `HTTP ${responseCode}: ${this.maskApiKeyInText(errorMessage)}`,
      );
    } catch (error) {
      // Re-throw our custom errors
      if (
        error.message.includes("Max retries") ||
        error.message.includes("HTTP") ||
        error.message.includes("Daily API limit")
      ) {
        error.message = this.maskApiKeyInText(error.message);
        throw error;
      }

      // UNEXPECTED ERRORS - Retry if attempts remain
      console.error(
        "✗ Unexpected error:",
        this.maskApiKeyInText(error.message),
      );

      const retryConfig = this.getRetryConfig();
      if (retryCount < retryConfig.MAX_RETRIES) {
        const backoffDelay = Math.min(
          retryConfig.INITIAL_BACKOFF_MS *
            Math.pow(retryConfig.BACKOFF_MULTIPLIER, retryCount),
          retryConfig.MAX_BACKOFF_MS,
        );

        console.log(
          `⚠️ Unexpected error. Retry ${retryCount + 1}/${retryConfig.MAX_RETRIES} in ${(backoffDelay / 1000).toFixed(1)}s`,
        );
        Logger.log(
          `Unexpected error: ${this.maskApiKeyInText(error.message)}. Retrying...`,
          "ERROR",
        );

        rateLimiterSleep(backoffDelay, "Error recovery");

        return this.callGeminiVision(
          fileContent,
          prompt,
          apiKey,
          retryCount + 1,
          modelName,
          useStructuredOutput,
          useMediaResolution,
        );
      }

      throw new Error(
        `Request failed after ${retryConfig.MAX_RETRIES} retries: ${this.maskApiKeyInText(error.message)}`,
      );
    }
  },

  /**
   * Extract retry delay from Gemini API error response
   * @param {Object} error - Error object from API response
   * @returns {number|null} - Delay in milliseconds, or null if not found
   */
  extractRetryDelay(error) {
    try {
      // Check error message for retry delay
      if (error.message && error.message.includes("Please retry in")) {
        const match = error.message.match(/retry in ([\d.]+)s/);
        if (match && match[1]) {
          return Math.ceil(parseFloat(match[1]) * 1000);
        }
      }

      // Check for RetryInfo in error details
      if (error.details && Array.isArray(error.details)) {
        const retryInfo = error.details.find(
          (d) => d["@type"] === "type.googleapis.com/google.rpc.RetryInfo",
        );
        if (retryInfo && retryInfo.retryDelay) {
          // retryDelay format: "42s" or "42.5s"
          const seconds = parseFloat(retryInfo.retryDelay.replace("s", ""));
          return Math.ceil(seconds * 1000);
        }
      }
    } catch (e) {
      console.warn("Could not extract retry delay:", e.message);
    }

    return null;
  },

  /**
   * Clean JSON response from Gemini
   */
  cleanJsonResponse(text) {
    let cleaned = text.trim();

    // Remove markdown code fences
    cleaned = cleaned.replace(/```\s*/g, "");

    // Extract JSON array
    const jsonStart = cleaned.indexOf("[");
    const jsonEnd = cleaned.lastIndexOf("]");

    if (jsonStart !== -1 && jsonEnd !== -1 && jsonEnd > jsonStart) {
      cleaned = cleaned.substring(jsonStart, jsonEnd + 1);
    }

    return cleaned;
  },

  parseMultiVoucherJsonResponse(text) {
    const cleanedResult = this.cleanJsonResponse(text);
    const parsedData = JSON.parse(cleanedResult);
    return Array.isArray(parsedData) ? parsedData : [parsedData];
  },
};
