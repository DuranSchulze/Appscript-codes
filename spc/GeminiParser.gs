/**
 * Gemini AI Parser for Petty Cash Vouchers
 * Version: 5.1 (Truncation detection + malformed JSON recovery)
 * Last Updated: August 4, 2026
 */

var GeminiParser = {
  /**
   * Configuration for retry logic
   */
  getRetryConfig() {
    return {
      MAX_RETRIES: 2,
      INITIAL_BACKOFF_MS: 1000, // 1 second
      MAX_BACKOFF_MS: 60000, // 60 seconds
      BACKOFF_MULTIPLIER: 2,
    };
  },

  getModelName() {
    const modelFromProps =
      PropertiesService.getScriptProperties().getProperty("GEMINI_MODEL");
    const model = (modelFromProps || "gemini-2.5-flash-lite").trim();
    return this.normalizeModelName(model);
  },

  getFallbackModelName() {
    const modelFromProps =
      PropertiesService.getScriptProperties().getProperty(
        "GEMINI_FALLBACK_MODEL",
      );
    const model = (modelFromProps || "gemini-3.5-flash").trim();
    return this.normalizeModelName(model);
  },

  normalizeModelName(model) {
    return model.startsWith("models/") ? model.replace(/^models\//, "") : model;
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
        description:
          "date printed after ERRAND DATE; preserve the exact calendar day and return YYYY-MM-DD",
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
        description:
          "transcribe the DETAILS column value verbatim; preserve names, reference numbers, wording, and meaningful punctuation; do not summarize, correct, or rewrite",
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
      minItems: 0,
      items: {
        type: "object",
        properties: voucherProperties,
        required: fields.map((field) => field.key),
        propertyOrdering: fields.map((field) => field.key),
      },
    };
  },

  buildTargetedExtractionPrompt() {
    const fieldList = this.getExtractionFields()
      .map((field) => `- ${field.key}: ${field.description}`)
      .join("\n");

    return `Extract liquidation/errand detail rows only.

Target fields:
${fieldList}

Rules:
- Search only for the target fields above.
- A valid source section must contain the detailed liquidation/errand fields, such as ERRAND DATE, ERRAND BY, SERVICE, DETAILS, Main Location of Errand, AMOUNT, and EXPENSE CLASSIFICATION.
- Do not extract a PCF Voucher, Petty Cash Fund Voucher, request, cover sheet, replenishment sheet, receipt, or summary/list page merely because it contains a voucher number, company, date, or amount.
- If a PDF contains both a PCF Voucher/summary page and detailed liquidation/errand rows, ignore the PCF Voucher/summary page and extract only the detailed rows.
- Copy DETAILS exactly as printed. Do not shorten, paraphrase, translate, correct grammar, expand abbreviations, or combine it with SERVICE or other fields.
- Preserve the calendar day printed under ERRAND DATE. Do not apply timezone conversion or substitute the document/processing date.
- Do not transcribe, summarize, or describe unrelated page content.
- Return one JSON object per voucher row.
- Use empty strings for missing fields.
- Return an empty JSON array if no voucher rows are found.`;
  },

  hasDetailedLiquidationMarkers(text) {
    const sourceText = String(text || "");
    const hasErrandDate = /\berrand\s+date\b/i.test(sourceText);
    const hasDetails = /\bdetails?\b/i.test(sourceText);
    const supportingMarkers = [
      /\berrand\s+by\b/i,
      /\bservice\b/i,
      /\bmain\s+location(?:\s+of\s+errand)?\b/i,
      /\bexpense\s+classification\b/i,
    ].filter((pattern) => pattern.test(sourceText)).length;

    return hasErrandDate && hasDetails && supportingMarkers >= 1;
  },

  isSummaryOnlyPcfText(text) {
    const sourceText = String(text || "");
    const hasPcfVoucherHeading =
      /\b(?:pcf|petty\s+cash\s+fund)\s+voucher\b/i.test(sourceText);
    return (
      hasPcfVoucherHeading &&
      !this.hasDetailedLiquidationMarkers(sourceText)
    );
  },

  normalizeForSourceEvidence(text) {
    return String(text || "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "");
  },

  annotateVoucherSections(text) {
    let sectionNumber = 0;
    return String(text).replace(
      /(^|\n)(\s*(?:petty\s+cash\s+voucher|voucher\s*(?:no\.?|number|reference)\s*[:#]?))/gi,
      (match, lineStart, heading) => {
        sectionNumber++;
        return `${lineStart}\n--- DETECTED VOUCHER SECTION ${sectionNumber} ---\n${heading}`;
      },
    );
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

  isQuotaError(error) {
    const message = String(
      error && error.message ? error.message : error || "",
    ).toLowerCase();
    return (
      message.includes("generate_content_free_tier_requests") ||
      message.includes("resource_exhausted") ||
      message.includes("exceeded your current quota") ||
      message.includes("quota exceeded") ||
      message.includes("quota_exhausted") ||
      message.includes("daily api limit")
    );
  },

  /**
   * Parse multiple vouchers from a document using Gemini 2.0 Flash
   */
  parseMultipleVouchers(file, apiKey) {
    const fileContent = DriveManager.getContentForVoucherParsing(file);
    if (fileContent.type === "text") {
      fileContent.text = this.annotateVoucherSections(fileContent.text);
      if (!this.hasDetailedLiquidationMarkers(fileContent.text)) {
        const skipReason = this.isSummaryOnlyPcfText(fileContent.text)
          ? "summary-only PCF Voucher"
          : "document without a detailed liquidation/errand section";
        Logger.log(
          `Skipped ${skipReason}: ${file.getName()}.`,
          "WARNING",
        );
        return [];
      }
    }
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

        const vouchers = this.parseMultiVoucherJsonResponse(result);
        this.validateVoucherArray(vouchers, fileContent);

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

        if (this.isQuotaError(error)) {
          throw new Error(`QUOTA_EXHAUSTED: ${safeMessage}`);
        }

        const shouldUseStrongerModel =
          (error && error.name === "SyntaxError") ||
          safeMessage.includes("VALIDATION_FAILED") ||
          safeMessage.includes("Invalid response structure");
        if (!shouldUseStrongerModel) {
          throw error;
        }

        Logger.log(
          `Gemini output from ${modelName} failed validation; trying the stronger model once: ${safeMessage}`,
          "WARNING",
        );
      }
    }

    // A model can occasionally return malformed JSON even when responseSchema
    // is enabled (for example, an unterminated string). Make one final request
    // with explicit prompt-only JSON instructions. This also gives a deployment
    // with the same primary/fallback model a real recovery attempt instead of
    // failing after the single unique model in getModelSequence().
    const recoveryModel =
      modelSequence[modelSequence.length - 1] || this.getModelName();
    const recoveryPrompt = `${prompt}

CRITICAL OUTPUT FORMAT:
- Return ONLY one complete, valid JSON array.
- Do not use Markdown code fences or add an explanation.
- Escape quotation marks and line breaks inside string values.
- Keep every object and close the final JSON array.`;

    try {
      Logger.log(
        `Structured JSON parsing failed; retrying ${file.getName()} once with prompt-only JSON recovery`,
        "WARNING",
      );
      const recoveryResult = this.callGeminiVision(
        fileContent,
        recoveryPrompt,
        apiKey,
        0,
        recoveryModel,
        false,
        true,
        8192,
      );
      const recoveredVouchers =
        this.parseMultiVoucherJsonResponse(recoveryResult);
      this.validateVoucherArray(recoveredVouchers, fileContent);
      Logger.log(
        `Gemini JSON recovery extracted ${recoveredVouchers.length} voucher(s) from ${file.getName()} using ${recoveryModel}`,
        "INFO",
      );
      return recoveredVouchers;
    } catch (recoveryError) {
      const safeRecoveryMessage = this.maskApiKeyInText(
        recoveryError && recoveryError.message,
      );
      errors.push(`${recoveryModel} recovery: ${safeRecoveryMessage}`);
      if (this.isQuotaError(recoveryError)) {
        throw new Error(`QUOTA_EXHAUSTED: ${safeRecoveryMessage}`);
      }
    }

    const message = errors.join(" | ");
    console.error("Gemini parsing error:", message);
    throw new Error(`Multi-voucher Gemini parsing failed: ${message}`);
  },

  validateVoucherArray(vouchers, fileContent) {
    if (!Array.isArray(vouchers)) {
      throw new Error("VALIDATION_FAILED: Gemini did not return a JSON array");
    }

    const allowedFields = this.getExtractionFields().map((field) => field.key);
    vouchers.forEach((voucher, index) => {
      if (!voucher || typeof voucher !== "object" || Array.isArray(voucher)) {
        throw new Error(
          `VALIDATION_FAILED: Voucher ${index + 1} is not an object`,
        );
      }
      const hasIdentity = ["voucherNo", "company", "errandDate", "details"].some(
        (key) => String(voucher[key] || "").trim(),
      );
      if (!hasIdentity) {
        throw new Error(
          `VALIDATION_FAILED: Voucher ${index + 1} has no identifying fields`,
        );
      }

      if (fileContent.type === "text" && voucher.details) {
        const sourceEvidence = this.normalizeForSourceEvidence(fileContent.text);
        const detailsEvidence = this.normalizeForSourceEvidence(voucher.details);
        if (
          detailsEvidence.length >= 6 &&
          !sourceEvidence.includes(detailsEvidence)
        ) {
          throw new Error(
            `VALIDATION_FAILED: Voucher ${index + 1} DETAILS was rewritten or is not present in the source text`,
          );
        }
      }

      allowedFields.forEach((key) => {
        if (voucher[key] === null || voucher[key] === undefined) {
          voucher[key] = "";
        } else if (typeof voucher[key] !== "string") {
          voucher[key] = String(voucher[key]);
        }
      });
    });

    // An empty array is valid unless the extracted text contains several
    // detailed liquidation/errand labels. A PCF heading by itself is not
    // evidence of a row that should be extracted.
    if (
      vouchers.length === 0 &&
      fileContent.type === "text" &&
      this.hasDetailedLiquidationMarkers(fileContent.text)
    ) {
      throw new Error(
        "VALIDATION_FAILED: Voucher labels were present but no vouchers were returned",
      );
    }
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
    maxOutputTokens = 4096,
  ) {
    const modelName = modelNameOverride || this.getModelName();

    try {
      const retryConfig = this.getRetryConfig();

      const inputLength =
        fileContent.type === "text" ? fileContent.text.length : 1000;
      const estimatedTokens = Math.ceil((prompt.length + inputLength) / 4);

      const generationConfig = {
        temperature: 0,
        topK: 16,
        topP: 0.8,
        maxOutputTokens,
      };

      if (useMediaResolution && fileContent.type !== "text") {
        generationConfig.mediaResolution = "MEDIA_RESOLUTION_MEDIUM";
      }

      if (useStructuredOutput) {
        generationConfig.responseMimeType = "application/json";
        generationConfig.responseSchema = this.getVoucherResponseSchema();
      }

      const parts = [{ text: prompt }];
      if (fileContent.type === "text") {
        parts.push({ text: `\n\nSOURCE DOCUMENT:\n${fileContent.text}` });
      } else {
        parts.push({
          inline_data: {
            mime_type: fileContent.mimeType,
            data: Utilities.base64Encode(fileContent.data.getBytes()),
          },
        });
      }

      const payload = {
        contents: [
          {
            parts,
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

      RateLimiterManager.acquireRequestSlot();
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

        const finishReason = String(
          jsonResponse.candidates?.[0]?.finishReason || "",
        ).toUpperCase();
        if (finishReason === "MAX_TOKENS") {
          if (maxOutputTokens < 8192) {
            Logger.log(
              `Gemini response from ${modelName} was truncated at ${maxOutputTokens} output tokens; retrying with 8192`,
              "WARNING",
            );
            return this.callGeminiVision(
              fileContent,
              prompt,
              apiKey,
              retryCount,
              modelName,
              useStructuredOutput,
              useMediaResolution,
              8192,
            );
          }
          throw new Error(
            "VALIDATION_FAILED: Gemini response was truncated at the maximum output-token limit",
          );
        }

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

      // Quota/rate errors are not transient. Retrying only burns the free tier.
      if (responseCode === 429) {
        const safeServerMessage = this.maskApiKeyInText(
          errorResponse.error?.message || "Quota exceeded",
        );
        throw new Error(`HTTP 429: QUOTA_EXHAUSTED: ${safeServerMessage}`);
      }

      // HANDLE 503 (SERVICE UNAVAILABLE) AND 500 (INTERNAL SERVER ERROR)
      if ([500, 502, 503, 504].includes(responseCode)) {
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
          maxOutputTokens,
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

      // Only network/temporary failures are retried. Validation, auth, quota,
      // and ordinary programming errors fail immediately.
      console.error(
        "✗ Unexpected error:",
        this.maskApiKeyInText(error.message),
      );

      const retryConfig = this.getRetryConfig();
      const transientMessage = String(error && error.message ? error.message : error);
      const isTransient = /timed?\s*out|temporary|connection reset|dns|network/i.test(
        transientMessage,
      );
      if (isTransient && retryCount < retryConfig.MAX_RETRIES) {
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
          maxOutputTokens,
        );
      }

      throw error;
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
    let parsedData;
    try {
      parsedData = JSON.parse(cleanedResult);
    } catch (error) {
      throw new Error(
        `VALIDATION_FAILED: Gemini returned malformed JSON (${error.message})`,
      );
    }
    if (!Array.isArray(parsedData)) {
      throw new Error("VALIDATION_FAILED: Gemini response must be a JSON array");
    }
    return parsedData;
  },
};
