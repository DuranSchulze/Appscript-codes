# Client Document Monitoring System — Documentation

**Script:** `Code-Client-Document-Management.gs`  
**Manifest:** `appscript.json`  
**Type:** Google Apps Script (Bound to a Google Spreadsheet)  
**Authored by:** Atty. Mary Wendy Duran  
**Last Modified:** 21 September 2025

---

## Overview

This is a **Google Apps Script automation system** bound to a Google Spreadsheet. It monitors two Google Drive folders used by a law firm (Duran Schulze) for legal case management:

1. **"01 Client Documents"** — Files uploaded by clients
2. **"5.02 Scanned Pleadings"** — Scanned legal pleadings filed by or against the firm

The system automatically:

- Detects new, modified, moved, or deleted files in both drives
- Tracks **who uploaded or modified** each file (by email)
- **AI-renames pleading files** using Gemini into a clean `YYYY-MM-DD Descriptive Title` format
- Logs all activity to a dedicated Logs sheet
- Runs on a **daily schedule** (6 PM Manila time, Mon–Fri)

---

## File Structure

```
litigation/
├── Code-Client-Document-Management.gs   ← Main script
├── appscript.json                        ← Manifest (OAuth scopes, Advanced APIs)
└── DOCUMENTATION.md                      ← This file
```

---

## Spreadsheet Sheets Created

| Sheet Name             | Purpose                                                             |
| ---------------------- | ------------------------------------------------------------------- |
| `Uploaded By Client`   | Tracks all files in the Client Documents drive                      |
| `Summary of Pleadings` | Tracks all scanned pleading files                                   |
| `Annexes`              | Manual annex registry                                               |
| `Checklist for Client` | Task checklist                                                      |
| `Hearing Dates`        | Hearing schedule                                                    |
| `Entries`              | Journal entries                                                     |
| `Logs`                 | System event log (max 1000 rows, auto-trimmed)                      |
| `System Config`        | Hidden key-value store for Drive IDs, scan timestamps, file indexes |

---

## Configuration Constants

### `CONFIG` object (top of script)

| Key                       | Value                 | Purpose                                  |
| ------------------------- | --------------------- | ---------------------------------------- |
| `SHEETS.*`                | Sheet name strings    | References to all sheet tabs             |
| `DRIVE_TYPES.CLIENT_DOCS` | `"CLIENT_DOCUMENTS"`  | Identifier for client docs drive         |
| `DRIVE_TYPES.PLEADINGS`   | `"SCANNED_PLEADINGS"` | Identifier for pleadings drive           |
| `BATCH_SIZE`              | `25`                  | Files per batch (for future batch UI)    |
| `TIMEZONE`                | `"Asia/Manila"`       | Used for scheduled trigger               |
| `SCHEDULE_HOUR`           | `18`                  | 6 PM daily trigger                       |
| `DOMAIN`                  | `".duranschulze.com"` | Used as fallback for user identification |
| `MAX_RETRIES`             | `3`                   | Drive API retry attempts                 |
| `RETRY_DELAY_MS`          | `1000`                | Delay between retries (ms)               |
| `API_DELAY_MS`            | `100`                 | Throttle delay every 10 files            |

### `GEMINI_CONFIG` object

| Key                  | Value                                                     | Purpose                                          |
| -------------------- | --------------------------------------------------------- | ------------------------------------------------ |
| `FALLBACK_MODEL`     | `"gemini-2.0-flash"`                                      | Default model used when no selection is saved    |
| `FALLBACK_ENDPOINT`  | `gemini-2.0-flash` REST URL                               | Default endpoint when no model is selected       |
| `MODELS_LIST_URL`    | `https://generativelanguage.googleapis.com/v1beta/models` | URL for fetching live model list                 |
| `USER_PROP_KEY`      | `"GEMINI_API_KEY"`                                        | Key used to store API key in User Properties     |
| `SELECTED_MODEL_KEY` | `"GEMINI_SELECTED_MODEL"`                                 | Key in System Config for the user-selected model |
| `TEMPERATURE`        | `0.2`                                                     | Low randomness for consistent legal titles       |
| `MAX_TOKENS`         | `150`                                                     | Max output tokens for AI responses               |

> **Important:** The API key is stored in **Google Apps Script User Properties** (per-user, not shared). It is never written to the spreadsheet. The selected model is stored in the **System Config** sheet (shared across the spreadsheet).

---

## Column Headers

### `Uploaded By Client` sheet

| Col | Header                   |
| --- | ------------------------ |
| A   | Date Uploaded            |
| B   | File Name (hyperlink)    |
| C   | File URL (folder path)   |
| D   | Uploaded by (Email)      |
| E   | Date Modified            |
| F   | Last Modified by (Email) |
| G   | Remarks                  |
| H   | Change Log               |

### `Summary of Pleadings` sheet

| Col | Header                            |
| --- | --------------------------------- |
| A   | Date Uploaded                     |
| B   | File Name (hyperlink, AI-renamed) |
| C   | File URL (folder path)            |
| D   | Filed / Issued By _(dropdown)_    |
| E   | Uploaded by (Email)               |
| F   | Last Modified by (Email)          |
| G   | Action Needed _(dropdown)_        |
| H   | Assigned DDS Party _(dropdown)_   |
| I   | Status                            |

### Dropdown Options

**Filed / Issued By (col D):** `DDS`, `Court`, `Opposing Party`

**Action Needed (col G):** `For Response`, `For Compliance`, `For Filing/Release`, `Filed/Completed`, `No Action Needed`

**Assigned DDS Party (col H):** `Partner`, `Senior Lawyer`, `Associate Lawyer`, `Paralegal`, `Legal Assistants`, `Liaison Officers`, `Other Support Personnel`

---

## Menu System (`onOpen`)

When the spreadsheet opens, a custom menu **"📁 Client Document Monitor"** is added with the following items:

| Menu Item                         | Function Called                | Purpose                                                           |
| --------------------------------- | ------------------------------ | ----------------------------------------------------------------- |
| 🔧 Initial Setup                  | `setupSystem()`                | Creates all sheets and initializes config                         |
| 📄 Set Client Documents Drive ID  | `setClientDocumentsDriveId()`  | Prompts for and saves the Drive folder ID                         |
| ⚖️ Set Scanned Pleadings Drive ID | `setScannedPleadingsDriveId()` | Prompts for and saves the Drive folder ID                         |
| 🔑 Set Gemini API Key             | `setGeminiApiKey()`            | Shows current key status (masked), validates, and saves new key   |
| 🤖 Select AI Model                | `selectGeminiModel()`          | Fetches live model list, highlights newest, saves selection       |
| 🗑️ Clear Gemini API Key           | `clearGeminiApiKey()`          | Deletes the stored Gemini key                                     |
| 📄 Full Scan - Client Documents   | `manualFullScanClientDocs()`   | Rebuilds the client docs sheet from scratch                       |
| ⚖️ Full Scan - Scanned Pleadings  | `manualFullScanPleadings()`    | Rebuilds the pleadings sheet from scratch                         |
| 🔄 Scan Both Drives               | `scanBothDrives()`             | Runs full scans on both drives sequentially                       |
| ⏰ Setup Daily Schedule           | `setupDailySchedule()`         | Creates a daily 6 PM trigger                                      |
| 🛑 Remove Daily Schedule          | `removeDailySchedule()`        | Deletes the scheduled trigger                                     |
| 📄 Test Client Documents Access   | `testClientDocsAccess()`       | Verifies Drive folder access                                      |
| ⚖️ Test Pleadings Access          | `testPleadingsAccess()`        | Verifies Drive folder access                                      |
| 👥 Test User Tracking             | `testUserTracking()`           | Tests email extraction on sample files                            |
| � Test Gemini Connection          | `testGeminiConnection()`       | Tests stored key + selected model, shows latency and response     |
| 🧪 Test Gemini AI Integration     | `testGeminiIntegration()`      | Full end-to-end AI pipeline test with sample filename             |
| �📊 View System Status            | `viewSystemStatus()`           | Shows current config, scan timestamps, API status                 |
| 🚨 View Error Report              | `viewErrorReport()`            | Shows last 20 ERROR log entries formatted for developer reporting |
| 🗑️ Clear Error Logs               | `clearErrorLogs()`             | Deletes only ERROR rows from the Logs sheet                       |

---

## Function Reference

### Setup & Initialization

#### `setupSystem()`

Creates all required sheets if they don't exist, initializes all config keys in the hidden `System Config` sheet, and shows a setup summary alert.

#### `createClientDocsSheet(spreadsheet)`

Creates the `Uploaded By Client` sheet with headers, frozen row, blue header formatting, and column widths.

#### `createPleadingsSheet(spreadsheet)`

Creates the `Summary of Pleadings` sheet with headers, formatting, and calls `setupPleadingsDropdowns()`.

#### `createSheet(spreadsheet, name, headers, hidden)`

Generic sheet creator used for Annexes, Checklist, Hearing Dates, Entries, Logs, and System Config.

#### `setupPleadingsDropdowns(sheet)`

Applies data validation dropdowns to columns D (Filed/Issued By), G (Action Needed), and H (Assigned DDS Party) for rows 2–1000.

---

### Drive Configuration

#### `setClientDocumentsDriveId()`

Prompts the user to enter a Google Drive folder ID. Validates access via `DriveApp.getFolderById()`, then saves to `System Config` as `CLIENT_DOCUMENTS_DRIVE_ID`.

#### `setScannedPleadingsDriveId()`

Same as above but saves as `SCANNED_PLEADINGS_DRIVE_ID`.

---

### AI / Gemini Functions

#### `getGeminiEndpoint()`

Returns the active Gemini `generateContent` REST endpoint. Reads `GEMINI_SELECTED_MODEL` from System Config and constructs the URL dynamically. Falls back to `FALLBACK_ENDPOINT` (`gemini-2.0-flash`) if no model has been selected. All AI calls use this function — no hardcoded endpoint anywhere.

#### `generateAiPleadingTitle(fileName)`

The **core AI function**. Given a raw filename, calls the Gemini REST API (via `getGeminiEndpoint()`) to generate a clean, formal legal title in proper title case.

- Reads the API key from User Properties (`GEMINI_API_KEY`)
- If no key is set, returns `null` immediately (graceful fallback)
- Uses the dynamically selected model via `getGeminiEndpoint()`
- Returns the AI-generated title string, or `null` on any error
- All errors are logged to console (non-fatal)

**Prompt used:**

> "You are a legal document classifier. Given the raw filename of a scanned legal pleading, produce a clean, formal, human-readable descriptive title (no date prefix, no file extension). Use proper title case. Output ONLY the title, nothing else."

#### `setGeminiApiKey()` _(improved)_

Shows the **current key status** before prompting (masked as `AIza...xYZ` if set, or "No key set"). Validates the new key via `testGeminiApiKey()` before saving. Shows a masked confirmation on success. Key is NOT saved if validation fails.

#### `testGeminiApiKey(apiKey)`

Sends a minimal test prompt to the active Gemini endpoint. Throws an error if the HTTP response is not 200 or if the response structure is invalid.

#### `clearGeminiApiKey()`

Deletes the stored `GEMINI_API_KEY` from User Properties after confirmation.

#### `fetchGeminiModels(apiKey)`

Calls `https://generativelanguage.googleapis.com/v1beta/models?key=API_KEY` and returns a sorted array of model name strings that support `generateContent`. Throws on network or API error.

#### `pickNewestModel(models)`

Scores each model name by version number, penalizes `exp`/`preview`/`legacy` suffixes, and returns the highest-scoring model name. Used to dynamically determine the `★ RECOMMENDED` model in the selector.

#### `selectGeminiModel()`

Fetches the live Gemini model list, presents a numbered menu to the user with the newest model marked `★ RECOMMENDED` and the current selection marked `◀ CURRENT`. Saves the chosen model to `System Config` as `GEMINI_SELECTED_MODEL`. All subsequent AI calls use this model automatically.

#### `testGeminiConnection()`

Standalone diagnostic that tests the currently stored API key against the currently selected model. Shows: key (masked), model name, HTTP response, AI reply text, and latency in ms. Logs result to the Logs sheet.

#### `testGeminiIntegration()`

Full end-to-end integration test of the AI pipeline:

1. Checks API key is set
2. Reads selected model from config
3. Calls `generateAiPleadingTitle()` with sample filename `"motion_to_dismiss_complaint_2024.pdf"`
4. Validates the AI returned a non-empty result
5. Simulates `generatePleadingName()` to show the full final filename
6. Shows a step-by-step pass/fail report with latency
7. Logs result as `TEST_GEMINI_INTEGRATION` event

---

### File Naming

#### `generatePleadingName(originalName, dateCreated)`

Generates the final display/Drive name for a pleading file.

**Logic:**

1. Formats the date as `YYYY-MM-DD`
2. Calls `generateAiPleadingTitle(originalName)` — if AI returns a title, uses it
3. If AI is unavailable or fails, falls back to cleaning up the original filename:
   - Removes file extension
   - Replaces `_` and `-` with spaces
   - Title-cases each word
4. Returns: `YYYY-MM-DD <Title>.<extension>`

**Example:**

- Input: `motion_to_dismiss_2024.pdf`, date `2024-03-15`
- AI output: `"Motion to Dismiss"`
- Final name: `2024-03-15 Motion to Dismiss.pdf`

---

### Scanning

#### `scanSharedDriveEnhanced(driveId, driveType, isFullScan, sheetName)`

Recursively scans all files in a Drive folder and its subfolders.

- Skips system/temp files via `isSystemOrTempFile()`
- Calls `getFileInfoEnhanced()` for each file
- On full scan: clears the sheet and calls `addFileToClientDocsSheet()` or `addFileToPleadingsSheet()` for each file
- Returns a `fileIndex` object keyed by file ID
- Throttles every 10 files with a 100ms sleep

#### `detectChangesEnhanced(driveId, driveType, sheetName, indexKey)`

Compares the current Drive state against the stored file index to detect:

- **NEW** files → adds to sheet (and renames pleadings)
- **USER_MODIFIED** files → logs the modification
- **NAME_UPDATED** files → renames existing pleading files in Drive and updates the sheet hyperlink

Saves the updated index back to `System Config`.

#### `manualFullScanClientDocs()` / `manualFullScanPleadings()`

Wrappers that call `manualFullScan()` with the correct config keys for each drive type.

#### `manualFullScan(configKey, driveType, displayName, sheetName, indexKey, lastScanKey)`

Shows a confirmation dialog, runs `scanSharedDriveEnhanced()` in full-scan mode, saves the updated index and timestamp, and shows a completion summary.

#### `scanBothDrives()`

Runs full scans on both drives sequentially after a single confirmation dialog.

#### `scheduledScan()`

The function called by the daily time-based trigger. Skips weekends. Calls `detectChangesEnhanced()` for both drives (only if they have been previously full-scanned).

---

### File Sheet Operations

#### `addFileToClientDocsSheet(fileInfo, remarks)`

Appends a row to `Uploaded By Client` with a `HYPERLINK` formula for the file name, upload/modification emails, and timestamps.

#### `addFileToPleadingsSheet(fileInfo)`

Appends a row to `Summary of Pleadings`. Before appending:

1. Calls `generatePleadingName()` to get the AI-renamed title
2. Renames the actual Drive file via `DriveApp.getFileById().setName()`
3. Auto-detects default values for Filed/Issued By and Action Needed based on filename keywords
4. Special case: `judgment`/`decision` files default to `"No Action Needed"` and `"Completed"`

#### `updateExistingPleadingName(sheet, rowIndex, newName, fileId)`

Updates the hyperlink formula in an existing pleadings row to reflect a renamed file.

#### `findExistingRowByFileId(sheet, fileId)`

Searches the pleadings sheet for a row whose hyperlink URL contains the given file ID. Returns `{ rowIndex, data }` or `null`.

---

### User Tracking

#### `getFileInfoEnhanced(file, folderPath, driveType, rootFolderName)`

Builds a `fileInfo` object for a Drive file, including calling `getEnhancedUserInfo()` to populate `uploadedByEmail` and `lastModifiedByEmail`.

#### `getEnhancedUserInfo(file)`

Attempts to extract uploader and last-modifier emails using three methods in order:

1. **Advanced Drive API** (`Drive.Files.get()` with `owners` and `lastModifyingUser` fields) — most accurate
2. **Standard DriveApp** (`file.getOwners()`) — fallback
3. **Domain inference** — checks if filename contains the firm domain as a last resort

Returns `{ uploadedByEmail, lastModifiedByEmail, fromAdvancedApi }`.

---

### Scheduling

#### `setupDailySchedule()`

Deletes any existing `scheduledScan` triggers, then creates a new time-based trigger: daily at `CONFIG.SCHEDULE_HOUR` (18:00) in `Asia/Manila` timezone.

#### `removeDailySchedule()`

Finds and deletes all triggers whose handler function is `scheduledScan`.

---

### Diagnostics

#### `testClientDocsAccess()` / `testPleadingsAccess()`

Calls `testDriveAccess()` for the respective drive.

#### `testDriveAccess(configKey, driveType, displayName)`

Opens the configured folder, counts up to 5 files, and shows an access report.

#### `testUserTracking()`

Calls `testDriveUserTracking()` for each configured drive and shows a summary of how many files had successful email extraction and whether the Advanced Drive API is working.

#### `testDriveUserTracking(driveId, driveName)`

Tests `getEnhancedUserInfo()` on up to 3 files in the given folder.

#### `viewSystemStatus()`

Shows a full status alert including: Drive IDs configured, last scan timestamps, Gemini API key presence, trigger status, and feature flags.

---

### Utility Functions

#### `logEvent(eventType, driveType, fileId, fileName, userEmail, details, status, errorMessage)`

Appends a row to the `Logs` sheet. Automatically trims the log to the most recent 1000 entries. **ERROR rows are automatically highlighted with a red background (`#f4cccc`)** for easy visual identification.

#### `viewErrorReport()`

Reads the Logs sheet, filters rows where Status = `"ERROR"`, and displays the last 20 errors (most recent first) in a formatted alert showing timestamp, event type, details, and error message. Designed for copy-pasting to a developer for reporting.

#### `clearErrorLogs()`

Deletes only `ERROR`-status rows from the Logs sheet after confirmation. Non-error rows (INFO, SUCCESS) are preserved. Iterates from the bottom up to avoid row-index shifting during deletion.

#### `getConfigValue(key)` / `setConfigValue(key, value)`

Read/write key-value pairs in the hidden `System Config` sheet.

#### `retryDriveOperation(operation, maxRetries)`

Wraps any Drive API call with retry logic (up to `CONFIG.MAX_RETRIES` attempts, with exponential-ish delay).

#### `isSystemOrTempFile(fileName)`

Returns `true` for temp/system files: `~$*`, `.tmp`, `.crdownload`, `.part`, `desktop.ini`, `thumbs.db`, `.DS_Store`, LibreOffice lock files.

#### `detectPleadingType(fileName)`

Matches the filename against `PLEADING_TYPES` regex patterns and returns the type string (e.g., `"Motion"`, `"Order"`, `"Affidavit"`). Defaults to `"Document"`.

#### `getExistingFileData(sheet, driveType)`

Reads all rows from a tracking sheet and returns a map of `{ displayName → { rowIndex, data } }` by parsing the `HYPERLINK` formula in the File Name column.

#### `updateActionNeededDropdown()`

Standalone function to re-apply the Action Needed dropdown validation to `G2:G100` of the `Summary of Pleadings` sheet. Useful for fixing validation on existing sheets without running a full setup.

#### `testSheetAccess()`

Simple diagnostic: checks if the `Summary of Pleadings` sheet exists and shows an alert.

---

## `appscript.json` Manifest

```json
{
  "timeZone": "Asia/Manila",
  "dependencies": {
    "enabledAdvancedServices": [
      { "userSymbol": "Drive", "serviceId": "drive", "version": "v2" }
    ]
  },
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive.metadata.readonly",
    "https://www.googleapis.com/auth/script.external_request"
  ],
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```

| Field                     | Purpose                                                                                                                        |
| ------------------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| `timeZone`                | Sets script timezone to Manila (UTC+8)                                                                                         |
| `enabledAdvancedServices` | Enables the **Advanced Drive API v2** required for `Drive.Files.get()` (owner/modifier email extraction)                       |
| `oauthScopes`             | Declares all required permissions: Sheets, Drive read/write, Drive metadata, and external HTTP requests (for Gemini API calls) |
| `exceptionLogging`        | Sends unhandled errors to Google Cloud Stackdriver                                                                             |
| `runtimeVersion`          | Uses V8 JavaScript engine (supports modern JS syntax)                                                                          |

> **Critical:** The `enabledAdvancedServices` entry must also be enabled in the Google Cloud Console for the Apps Script project under **APIs & Services → Enable APIs → Google Drive API**.

---

## Bugs Fixed

| #   | Location                       | Bug                                                                                                                                                                                   | Fix Applied                                                                                                     |
| --- | ------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------- |
| 1   | `appscript.json`               | `dependencies` was empty `{}` — Advanced Drive API not declared, causing `Drive.Files.get()` to throw `ReferenceError: Drive is not defined`                                          | Added `enabledAdvancedServices` with Drive API v2                                                               |
| 2   | `appscript.json`               | `oauthScopes` missing — Gemini API calls (`script.external_request`) and Drive metadata access not authorized                                                                         | Added full `oauthScopes` array                                                                                  |
| 3   | `GEMINI_CONFIG.MODEL_ENDPOINT` | Used deprecated `gemini-2.0-flash-exp` model endpoint                                                                                                                                 | Updated to stable `gemini-2.0-flash` (now dynamic via `getGeminiEndpoint()`)                                    |
| 4   | `createClientDocsSheet()`      | `setFrontColor("white")` — typo, not a valid Spreadsheet method, silently fails leaving header text invisible                                                                         | Fixed to `setFontColor("white")`                                                                                |
| 5   | `testGeminiApiKey()`           | `generationConfig` did not pass `maxOutputTokens` — `MAX_TOKENS` constant was defined but never used in any API call                                                                  | Added `maxOutputTokens: GEMINI_CONFIG.MAX_TOKENS`                                                               |
| 6   | `generatePleadingName()`       | AI was configured (Gemini key, endpoint, config) but `generateAiPleadingTitle()` was never implemented — the renaming was purely regex-based despite the header claiming "AI-powered" | Implemented `generateAiPleadingTitle()` and wired it into `generatePleadingName()` with filename-based fallback |

---

## Setup Instructions (First Time)

1. Open the bound Google Spreadsheet
2. Go to **Extensions → Apps Script** and paste the script (or it should already be there)
3. In Apps Script editor, go to **Project Settings → Google Cloud Platform project** and ensure the **Google Drive API** is enabled in the linked GCP project
4. Run **📁 Client Document Monitor → 🔧 Initial Setup** — this creates all sheets
5. Run **Configure Drives → 📄 Set Client Documents Drive ID** — paste the folder ID from the Drive URL
6. Run **Configure Drives → ⚖️ Set Scanned Pleadings Drive ID** — paste the folder ID
7. Run **🔑 Set Gemini API Key** — paste your key from [https://aistudio.google.com/app/apikey](https://aistudio.google.com/app/apikey)
8. Run **🤖 Select AI Model** — fetches the live model list; pick the `★ RECOMMENDED` one
9. Run **🩺 Diagnostics → 🧪 Test Gemini AI Integration** to verify the full AI pipeline works
10. Run **🩺 Diagnostics → 👥 Test User Tracking** to verify email extraction works
11. Run **🔍 Full Scanning → 🔄 Scan Both Drives** for the initial population
12. Run **⏰ Setup Daily Schedule** to enable automatic daily scans

---

## How to Get a Gemini API Key

1. Go to [https://aistudio.google.com/app/apikey](https://aistudio.google.com/app/apikey)
2. Sign in with your Google account
3. Click **"Create API key"**
4. Copy the key
5. In the spreadsheet menu: **📁 Client Document Monitor → 🔑 Set Gemini API Key**
6. Paste the key and click OK — it will be tested automatically before saving

> The key is stored in **your personal User Properties** for this script only. It is not visible to other users of the spreadsheet.

---

## Data Flow Diagram

```
Google Drive Folder
       │
       ▼
scanSharedDriveEnhanced()
       │
       ├── getFileInfoEnhanced()
       │       └── getEnhancedUserInfo()  [Advanced Drive API → DriveApp → Domain fallback]
       │
       ├── [PLEADINGS] generatePleadingName()
       │       ├── generateAiPleadingTitle()  [Gemini API → null on failure]
       │       └── Filename regex fallback
       │
       ├── addFileToClientDocsSheet()  OR  addFileToPleadingsSheet()
       │       └── DriveApp.getFileById().setName()  [renames actual Drive file]
       │
       └── logEvent()  →  Logs sheet
```

---

## Scheduled Trigger Flow

```
Daily 6 PM (Mon–Fri, Asia/Manila)
       │
       ▼
scheduledScan()
       ├── detectChangesEnhanced()  [Client Documents]
       │       ├── New files → addFileToClientDocsSheet()
       │       └── Modified files → logEvent()
       └── detectChangesEnhanced()  [Scanned Pleadings]
               ├── New files → addFileToPleadingsSheet() + rename in Drive
               └── Existing files with old names → rename in Drive + updateExistingPleadingName()
```
