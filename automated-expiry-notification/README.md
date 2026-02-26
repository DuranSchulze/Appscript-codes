# Automated Expiry Notification (Google Apps Script)

Google Apps Script automation that sends expiry reminder emails from a Google Sheet, tracks send outcomes, and provides menu-based diagnostics for operations staff.

---

## Project Goal

Provide a reliable, auditable reminder system that:

- sends reminder emails based on `Expiry Date` + `Notice Date`
- allows both scheduled and manual runs
- writes operational outcomes back to the sheet (`Status`, `Staff Email`)
- provides clear diagnostics and run history via `LOGS`

---

## Purpose and Business Value

This script helps teams avoid missed renewals by automating outreach while keeping spreadsheet operations simple:

- non-technical users manage rows directly in Google Sheets
- the system enforces required data and logs errors for correction
- reminder scheduling is deterministic and traceable

---

## Current Feature Set (Implemented)

### 1) Automation and Scheduling

- Daily trigger support (8:00 AM Asia/Manila)
- Manual run from menu (`Run Manual Check Now`)
- Trigger status visibility (`Show Automation Status`)
- Trigger activate/deactivate actions

### 2) Data Processing Rules

- Dynamic header mapping from row 2 (column order can change)
- Required fields validation (`Client Name`, `Client Email`, `Expiry Date`, `Notice Date`, `Status`)
- Optional row-level send control via `Send Mode`:
  - `Auto` -> normal processing
  - `Hold` -> skipped with reason
  - `Manual Only` -> skipped in bulk runs (still sendable via test send)
- Status handling:
  - `Active` rows are processed
  - blank `Status` rows are auto-set to `Active` and processed
  - `Sent`/`Error` rows are skipped in normal runs

### 3) Date and Eligibility Logic

- Notice string parser supports:
  - `On expiry date`
  - `N days before`
  - `N weeks before`
  - `N months before`
- Send eligibility is **due/overdue**:
  - row sends when `targetDate <= today`

### 4) Email Composition and Sending

- Subject uses document type + expiry date + client name
- Body source priority:
  1. `Remarks` template (with placeholders)
  2. built-in fallback template if `Remarks` is empty
- Placeholder replacement (case-insensitive):
  - `[Client Name]`
  - `[Date of Expiration]`, `[Date of Expiry]`, `[Expiry Date]`
  - `[Document Type]`
- Optional CC via `Staff Email`
- Drive attachments from `Attached Files` (URLs or file IDs)

### 5) Post-Send Updates and Logging

- On success:
  - set `Status = Sent`
  - write sender account email to `Staff Email`
  - write sent metadata when columns exist (`Sent At`, `Sent Thread Id`, `Sent Message Id`, `Open Tracking Token`)
  - set `Reply Status = Pending` (when column exists)
  - write `SENT` log
- On failure:
  - set `Status = Error`
  - write `ERROR` log
- `SKIPPED` logs are written for Send Mode holds/manual-only
- End-of-run `SUMMARY` log includes totals, auto-activated count, and skip counters

### 6) Operations and Diagnostics

- Initialize/select working tab via menu
- Preview target dates without sending
- Inspect row details by row number
- Send a real test email by `No.` value
- Preview effective fallback template rendering

### 7) Reply Tracking (Implemented)

- Configurable reply keywords (default: `ACK, RECEIVED, OK`)
- Reply scan supports manual run and daily scheduled trigger
- Scan matches replies in sent threads and updates sheet fields:
  - `Reply Status`
  - `Replied At`
  - `Reply Keyword`

### 8) AI Integration (Implemented)

- AI menu actions:
  - toggle AI generation
  - set Gemini API key
  - select model
  - test connection
  - show AI status
- Body source resolution when `Remarks` is empty:
  1. AI-generated body (if enabled and configured)
  2. fallback template

### 9) Open/Read Tracking (Best Effort Implemented)

- Optional open tracking via pixel token in outgoing email HTML
- `doGet` web app endpoint records open/click events
- When tracking columns exist, updates:
  - `First Opened At`
  - `Last Opened At`
  - `Open Count`

---

## Implementation Audit Snapshot

### Implemented from current scope

- [x] Configurable sheet tab initialization
- [x] Manual + scheduled run parity
- [x] Due/overdue catch-up send behavior
- [x] Blank status auto-activation
- [x] Send Mode support (`Auto`, `Hold`, `Manual Only`)
- [x] Placeholder support including `[Document Type]`
- [x] `Staff Email` update after successful send
- [x] `No.`-based diagnostic test sending
- [x] LOGS sheet with summary and colored actions

- [x] Reply keyword configuration + reply scan (manual and scheduled)
- [x] Sent metadata persistence for reply linkage
- [x] AI integration menu + Gemini generation path
- [x] Configurable fallback source (hardcoded/property template)
- [x] Open tracking token + `doGet` event capture

### Remaining considerations (operational)

- [ ] Deploy Apps Script as a Web App and set tracking URL for open tracking to work
- [ ] Add new optional columns to the sheet if reply/open metadata should be stored
- [ ] Monitor Gmail/App Script quotas for daily reply scans and send volume

---

## Spreadsheet Contract

### Required structure

- Row 1: optional title row (ignored)
- Row 2: header row (used for mapping)
- Row 3+: data rows

### Required headers

- `Client Name`
- `Client Email`
- `Expiry Date`
- `Notice Date`
- `Status`

### Optional headers used by the script

- `No.`
- `Type of ID/Document`
- `Remarks`
- `Attached Files`
- `Staff Email`
- `Send Mode`
- `Sent At`
- `Sent Thread Id`
- `Sent Message Id`
- `Reply Status`
- `Replied At`
- `Reply Keyword`
- `Open Tracking Token`
- `First Opened At`
- `Last Opened At`
- `Open Count`

Header alias supported:

- `Assigned Staff Email` -> `Staff Email`

---

## Menu Map

### Expiry Notifications

- Show Automation Status
- Initialize / Select Working Sheet
- Run Manual Check Now
- Check Schedule Status
- Activate Daily Schedule (8 AM)
- Deactivate Daily Schedule

### Diagnostics

- Preview Target Dates (no emails sent)
- Inspect Row...
- Send Test Email by No....
- Preview Effective Fallback Body

### Reply Tracking

- Run Reply Scan Now
- Set Reply Keywords
- View Reply Tracking Status
- Activate Reply Scan Schedule
- Deactivate Reply Scan Schedule

### AI Integration

- Toggle AI Generation
- Set Gemini API Key
- Select Gemini Model
- Test AI Connection
- View AI Status
- Set Fallback Template
- Toggle Fallback Source
- Set Open Tracking URL

---

## Runtime Flow (Current)

1. Resolve configured sheet tab
2. Ensure `LOGS` exists
3. Build/validate column map
4. Read rows from row 3+
5. Filter processable statuses (`Active` or blank)
6. Apply `Send Mode` gate (`Auto` runs, `Hold`/`Manual Only` skipped)
7. Auto-set blank status to `Active`
8. Validate required values
9. Parse notice offset and compute target date
10. Skip if target date is still in the future
11. Resolve attachments
12. Build body via Remarks/AI/fallback pipeline
13. Send email
14. Update status/staff/sent metadata
15. Append summary log

---

## Known Constraints and Limitations

- Open tracking is best effort (depends on recipient client/privacy settings)
- Reply matching relies on sent thread metadata; legacy rows without metadata may not match
- AI generation depends on valid Gemini API key/model and external API availability

---

## Security and Operations Notes

- Script permissions must allow Gmail, Drive, and Spreadsheet access
- Daily trigger should be checked periodically to avoid duplicate triggers
- API keys must be stored in Script/Document Properties, not worksheet cells
- For open tracking, deploy as Web App and set `Set Open Tracking URL` to the deployed endpoint

---

## Future Roadmap

Roadmap and feature checklist are maintained in:

- `plans.md` (PM-style TODO/checklist board)

---

## Key Files

- `code.gs` - implementation
- `plans.md` - roadmap / delivery checklist
- `README.md` - product + operational documentation (this file)
