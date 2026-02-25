# DDS + FilePino System Memory

_Last updated: 2026-02-25_

This document is the operational memory for both Apps Script systems:

- `code-dds.gs`
- `code-filepino.gs`

It explains how the system works **today**, what safeguards are already in place, and how it is expected to operate moving forward.

---

## 1) Purpose of the two scripts

Both scripts are automation engines for document operations across multiple departments:

1. Scan configured Google Drive department folders
2. Extract document titles using Gemini AI (with fallback logic)
3. Rename and reorganize files into date-based folder structures
4. Record processing logs and audit trails
5. Send daily/weekly summary emails with batching and safeguards
6. Send alert emails for no-files and error conditions

They are near-identical systems deployed for two different business contexts/brands:

- **DDS**
- **FilePino**

---

## 2) System architecture (high level)

```text
Google Sheet (Dashboard config)
        |
        v
Department list + Drive folder IDs + recipient config
        |
        v
Drive scan + file filters + de-duplication
        |
        v
AI title extraction (Gemini -> OCR -> Filename fallback)
        |
        v
Rename + move to YYYY/MMMM/DD MMMM folders
        |
        +--> Scanned Files Log (row records)
        +--> Script Properties (email payload index)
        +--> Audit Trail (success/error telemetry)
        |
        v
Email subsystem
   - Daily summary
   - Weekly digest
   - Alerts (no files / processing errors)
```

---

## 3) Primary data stores

## 3.1 Spreadsheet sheets

### Dashboard

Main configuration table per department.

Expected columns:

1. Department Names
2. Google Drive ID
3. Emails
4. Status
5. Manager Emails
6. Alert Emails (Errors/No Files)

### Scanned Files Log

Operational record of each processed file:

- Date Processed
- Department
- Original File Name
- New File Name
- View Folder
- File Path
- File Size
- Process Time
- AI Used
- Status

### Audit Trail

System telemetry and diagnostics:

- Timestamp
- Action
- Details
- Status
- Error Message

### Flagged Files

Review queue for duplicates or items requiring manual review.

## 3.2 Script Properties (state + config)

Common keys used by both systems:

- `TARGET_SPREADSHEET_ID`
- `GEMINI_API_KEY`
- `GEMINI_MODEL`
- `processedFileIds`
- `processed_rows_YYYY-MM-DD`
- `email_YYYY-MM-DD_row_X`
- `GEMINI_DAILY_CALL_LIMIT`
- `GEMINI_CALLS_YYYY-MM-DD`
- `GEMINI_LAST_TEST_MS`

---

## 4) End-to-end lifecycle

## 4.1 Initialization

- `onOpen()` builds the custom menu
- `initializeSetup()` creates/updates required sheets
- Setup optionally opens Gemini model selection from live API model list
- If API key is missing, the user is prompted to set key first
- System fetches current Gemini models from Google API, shows selectable list, and user picks one
- Selected model is saved to Script Properties (`GEMINI_MODEL`) and becomes the active model for all Gemini processing calls

## 4.2 Manual operation

Typical manual day flow:

1. `Scan Files Now` → runs processing
2. `Send Daily Emails` → sends current day summary
3. If needed, `Send Weekly Digest` for week summary

## 4.3 Scheduled operation

- Trigger runs `scheduledProcessing()` around 5:15 PM daily
- Function skips weekends (Mon–Fri behavior)
- Weekly digest auto-send runs on configured day when enabled

> Important current behavior: scheduled processing scans files, but daily summary email is still a separate send flow.

---

## 5) File processing flow

For each active department row in Dashboard:

1. Open department root folder by Drive ID
2. Iterate files
3. Skip if already in `processedFileIds`
4. Skip files not directly in department root (ignore nested items)
5. Skip files not created "today" (timezone-aware)
6. Skip unsupported MIME types
7. Process eligible files:
   - Extract title
   - Build new filename
   - Move file to structured subfolder
   - Log to Scanned Files Log
   - Mark file as processed

Supported MIME types:

- PDF
- JPEG
- PNG
- GIF
- BMP

---

## 6) AI extraction strategy and fallback chain

Title extraction is resilient by design:

1. **Primary:** Gemini extraction (`extractTitleWithGemini`)
2. **Fallback 1:** OCR extraction (PDF/image)
3. **Fallback 2:** filename sanitization fallback

Fallback reason tagging is embedded in `AI Used`, e.g.:

- `OCR Fallback (Gemini API 429)`
- `OCR Fallback (Gemini quota reached)`
- `Filename Fallback (Gemini key missing)`

This makes post-incident diagnosis possible from logs and email summaries.

---

## 7) Gemini controls and safety controls

Both scripts include safeguards for API usage:

1. **Model control**
   - Gemini model is stored in Script Properties (`GEMINI_MODEL`)
   - Standard flow uses API-driven model picker (not hardcoded model typing)
   - User can select by list number or model name from fetched models
   - Saved model is used by all `generateContent` calls through centralized URL builder

2. **Daily budget**
   - `GEMINI_DAILY_CALL_LIMIT` with per-day counter
   - Calls blocked when budget is exhausted

3. **Test cooldown**
   - Prevents repeated connection tests consuming quota too quickly

## 7.1 Gemini model selection workflow (current behavior)

Current expected behavior in both scripts:

1. User opens **Set Gemini Model** from menu (or setup prompt)
2. System checks `GEMINI_API_KEY`
3. System requests latest model catalog from Gemini Models API
4. System filters for valid `gemini-*` models supporting `generateContent`
5. User selects desired model from displayed list
6. System saves selection to `GEMINI_MODEL`
7. All runtime Gemini requests use selected model automatically

Notes:

- The old manual-entry helper still exists for legacy fallback/testing, but normal user flow is API-list selection.
- If model fetch fails (key/project/quota/network issue), the system keeps current saved model and logs/alerts the issue.

---

## 8) Email subsystem (daily/weekly)

## 8.1 Recipient planning

Before sending:

- Parse and normalize recipient CSVs
- De-duplicate addresses
- Validate email format
- Separate TO and CC safely
- Log recipient issues by department

## 8.2 Quota enforcement

- Check `MailApp.getRemainingDailyQuota()` before sending
- Block sends if quota is insufficient
- Write explicit quota-block audit logs

## 8.3 Batch delivery

- If recipients exceed batch limit, split sends
- Delay between batches (`EMAIL_BATCH_DELAY`)
- Capture failed batch details:
  - batch number
  - recipient count
  - sample recipients
  - exact error message

## 8.4 Sender identity

Emails are sent with `MailApp.sendEmail` under the account context running the script (owner/authorized execution context), with branded display names:

- DDS: `DDS Document Automation`
- FilePino: `FilePino Inc Document Automation`

---

## 9) Alert subsystem

Two critical alerts were added:

1. **No-files alert**
   - Triggered when daily email flow finds zero processed files for the day
   - Sent to `Alert Emails (Errors/No Files)` recipients

2. **Processing error summary alert**
   - Triggered when processing records errors
   - Includes:
     - processing totals
     - AI fallback statistics
     - Gemini 429 / quota indicator counts
     - latest error rows from Audit Trail

This turns the system into a self-reporting operations workflow instead of silent failure.

---

## 10) Diagnostics: how to know why something did not send

Use this sequence:

1. **Audit Trail first**
   - Look for actions:
     - `No email addresses`
     - `Email blocked - quota exhausted`
     - `Batch email failed`
     - `Email sending incomplete`

2. **Check recipient quality**
   - Invalid emails are excluded and logged
   - Missing department emails are surfaced as warnings

3. **Check quota**
   - Confirm remaining daily quota

4. **Check failed batch metadata**
   - Sample recipient preview is logged

5. **Check AI/fallback behavior**
   - Look at `AI Used` in Scanned Files Log for 429/quota patterns

---

## 11) DDS vs FilePino differences

Core logic is intentionally mirrored. Differences are mostly deployment-level:

1. Branding strings (subjects, sender display name)
2. Default spreadsheet IDs
3. Some constants (for example filename-length limits)

Operational recommendation: keep both files synchronized function-for-function unless a business rule explicitly differs.

---

## 12) How the system is expected to work moving forward

Target operating model:

1. Setup once via `Initialize Setup`
2. Keep Dashboard maintained (Drive IDs + recipients + alert recipients)
3. Use scheduled processing for daily ingestion
4. Use daily email flow for daily summary distribution
5. Let weekly digest run on configured day
6. Monitor Audit Trail and alerts for issues
7. Re-check and switch Gemini model using live model picker when needed
8. Adjust Gemini budget settings as operational demand changes

In short: **automated ingestion + transparent diagnostics + controlled AI cost + reliable stakeholder reporting**.

---

## 13) Operational runbook (quick)

### Daily

- Verify schedule is active
- Check Audit Trail for overnight/day errors
- Run/send daily summary as needed

### Weekly

- Verify weekly digest delivered
- Review fallback rates (Gemini vs OCR)
- Investigate recurring 429/quota events

### Incident response

- Step 1: Audit Trail action + error details
- Step 2: Recipient + quota status
- Step 3: Gemini key/model/budget validation
- Step 4: Re-run targeted manual flow

---

## 14) Documentation maintenance rule

Any change to one script's logic that affects behavior must be evaluated for the other script and, if applicable, mirrored and documented here.
