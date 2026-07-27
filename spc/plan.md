# Implementation Plan — SPC Liquidation System

## Overview

This plan addresses the limitations identified in the current Gemini-powered voucher parsing system. The fixes are organized by priority and dependency order.

---

## 1. High Priority

### 1.1 Add Response Truncation Detection

**File:** `GeminiParser.gs` — `callGeminiVision()`

**Problem:** `maxOutputTokens: 2048` can truncate multi-voucher responses with no detection. A truncated JSON array silently fails in `parseMultiVoucherJsonResponse()`.

**Fix:**
- After a 200 response, check `jsonResponse.candidates[0].finishReason`.
- If `finishReason === "MAX_TOKENS"`, log a warning and either:
  - Increase `maxOutputTokens` for a retry (up to `8192`), or
  - Throw a specific error so the caller knows the response was cut off.

```javascript
// In callGeminiVision(), after responseCode === 200:
const finishReason = jsonResponse.candidates[0]?.finishReason;
if (finishReason === "MAX_TOKENS") {
  Logger.log(`Response truncated (MAX_TOKENS) on ${modelName}`, "WARNING");
  // Retry with higher output tokens if not already at max
  if (generationConfig.maxOutputTokens < 8192) {
    return this.callGeminiVision(
      fileContent, prompt, apiKey, retryCount, modelName,
      useStructuredOutput, useMediaResolution,
      /* increasedMaxTokens */ true
    );
  }
}
```

---

### 1.2 Add Schema-less JSON Fallback

**File:** `GeminiParser.gs` — `parseMultipleVouchers()`

**Problem:** When both primary and fallback models fail on structured output with schema, the entire file fails. Some Gemini models (especially newer ones) may not support `responseSchema` but do support `responseMimeType: "application/json"` with in-prompt format instructions.

**Fix:**
- After the model sequence loop exhausts (line 269-273), if all attempts failed, try a **third pass** with `useStructuredOutput = false` and an augmented prompt that explicitly requests JSON array format.

```javascript
// In parseMultipleVouchers(), after the for loop:
if (errors.length === modelSequence.length) {
  Logger.log("Schema-mode failed on all models; attempting prompt-only JSON", "WARNING");
  const jsonPrompt = prompt + "\n\nCRITICAL: Return ONLY a valid JSON array of objects. No markdown, no explanation.";
  const result = this.callGeminiVision(
    fileContent, jsonPrompt, apiKey, 0,
    modelSequence[0], /* useStructuredOutput */ false, true
  );
  const vouchers = this.parseMultiVoucherJsonResponse(result);
  this.validateVoucherArray(vouchers, fileContent);
  return vouchers;
}
```

---

### 1.3 Handle File-in-Limbo After Timeout/Crash

**File:** `DriveManager.gs` — `renameProcessedFile()` & `markFileAsProcessed()`

**Problem:** `renameProcessedFile()` runs *before* `markFileAsProcessed()`. If the execution crashes between rename and mark, the renamed file (without `[PROCESSED]`) is invisible to `getUnprocessedFiles()` because the rename changes the filename to a vouchernumber-based pattern that doesn't match the original search logic.

**Fix:**
- Swap the order: mark as `[PROCESSED]` first, then do the descriptive rename.
- Or: always mark first, and make `renameProcessedFile` optional (skip on failure).

```javascript
// In Code.gs processFiles(), move markFileAsProcessed before renameProcessedFile:
DriveManager.markFileAsProcessed(file);  // Mark first — this is the "safety" tag
renamedFileName = DriveManager.renameProcessedFile(file, uniqueVoucherData, false);
```

---

## 2. Medium Priority

### 2.1 Make OCR Wait Time Configurable and Increase Default

**File:** `DriveManager.gs` — `extractTextWithGoogleOcr()`

**Problem:** Fixed 3-second default (`ocrWaitMs`) is too short for large scanned PDFs. Google Drive OCR can take 10–30 seconds.

**Fix:**
- Increase default from `3000` to `8000` ms.
- Add retry logic: if extracted text is empty, wait longer and retry once.

```javascript
const ocrWaitMs = Number(scriptProps.getProperty("DRIVE_OCR_WAIT_MS")) || 8000;
// ... after first read attempt:
if (!text || text.length < 40) {
  Utilities.sleep(ocrWaitMs); // Wait again
  const text2 = DocumentApp.openById(temporaryDocumentId).getBody().getText().trim();
  if (text2 && text2.length >= 40) return text2;
}
```

---

### 2.2 Remove Dead Code: `extractRetryDelay()`

**File:** `GeminiParser.gs` — lines 543–569

**Problem:** The function is defined but never called. It adds maintenance burden.

**Fix:** Delete the entire `extractRetryDelay()` method.

---

### 2.3 Remove 400 Fallback That No Longer Matches

**File:** `GeminiParser.gs` — `parseMultipleVouchers()`, error handling around line 257

**Problem:** The old 400 fallback checked for `responseFormat` in error messages (now dead code since we changed to `responseMimeType`). It also checked for `VALIDATION_FAILED` — this one is still useful.

**Fix:** Clean up the condition to only retry on validation errors, not on `responseFormat`:

```javascript
// Before:
if (errorMessage.includes("responseFormat") || errorMessage.includes("schema") || ...)

// After:
if (errorMessage.includes("schema") || errorMessage.includes("JSON mode"))
```

Wait — actually this code was removed in the current version. Let me re-check. Actually no — looking at lines 445-456 of the old version, this was in `callGeminiVision`. In the current version I read, this 400 fallback doesn't exist anymore (lines 478-481 just throw). So this is already addressed — confirm and remove any residual references.

---

## 3. Low Priority

### 3.1 Move API Key from Query String to Header

**File:** `GeminiParser.gs` — `callGeminiVision()`, line 390

**Problem:** API key in URL query params is exposed in execution logs and Cloud Logging.

**Fix:**
```javascript
// Before:
const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;

// After:
const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent`;
options.headers = { "x-goog-api-key": apiKey };
```

---

### 3.2 Remove `propertyOrdering` from Response Schema

**File:** `GeminiParser.gs` — `getVoucherResponseSchema()`, line 119

**Problem:** `propertyOrdering` is not part of the standard JSON Schema spec. Some Gemini API versions may silently ignore it or reject it.

**Fix:** Remove the `propertyOrdering` line:
```javascript
// Delete this line:
propertyOrdering: fields.map((field) => field.key),
```

---

### 3.3 Increase Default `maxOutputTokens`

**File:** `GeminiParser.gs` — `callGeminiVision()`, line 305 (updated line)

**Problem:** `2048` tokens is tight for 10+ vouchers.

**Fix:** Increase default to `4096` and allow dynamic increase on truncation (see 1.1):
```javascript
maxOutputTokens: increasedMaxTokens ? 8192 : 4096,
```

---

## 4. Future Considerations (Not in Scope)

| Item | Rationale |
|------|-----------|
| Migrate from `v1beta` to `v1` endpoint | v1beta will eventually be deprecated |
| PDF page-range splitting | Enable processing of >1000 page PDFs |
| Stateful resume via PropertiesService | Allow crash recovery for large batches |
| Webhook-based trigger chain | Break 10-file batches into chained executions to avoid timeout |
| Service account auth | Higher rate limits than API key free tier |

---

## Implementation Order

```
                  ┌─────────────────────┐
                  │  1.3  File Limbo    │  (prevents data loss on crash)
                  └────────┬────────────┘
                           │
                  ┌────────▼────────────┐
                  │  1.1  Truncation    │  (detects silent data loss)
                  └────────┬────────────┘
                           │
                  ┌────────▼────────────┐
                  │  1.2  JSON Fallback │  (increases parse success rate)
                  └────────┬────────────┘
                           │
           ┌───────────────┼───────────────┐
           │                               │
  ┌────────▼────────────┐    ┌─────────────▼──────────────┐
  │  2.1  OCR Wait Time │    │  2.2 / 2.3  Cleanup       │
  └────────┬────────────┘    └─────────────┬──────────────┘
           │                               │
           └───────────────┬───────────────┘
                           │
           ┌───────────────┼───────────────┐
           │               │               │
  ┌────────▼──────┐ ┌──────▼──────┐ ┌──────▼──────────┐
  │ 3.1  API Key  │ │ 3.2 Schema │ │ 3.3 Token Limit │
  │    Header     │ │  Cleanup   │ │    Increase     │
  └───────────────┘ └─────────────┘ └─────────────────┘
```

---

## File Change Summary

| File | Items |
|------|-------|
| `GeminiParser.gs` | 1.1 (truncation), 1.2 (schema fallback), 2.2 (dead code), 2.3 (400 cleanup), 3.1 (key header), 3.2 (propertyOrdering), 3.3 (token limit) |
| `DriveManager.gs` | 2.1 (OCR wait) |
| `Code.gs` | 1.3 (mark-before-rename order) |
