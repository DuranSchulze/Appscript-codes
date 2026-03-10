# Automated Expiry Notification (Google Apps Script)

> **System Version:** March 4, 2026

Google Apps Script automation that sends expiry reminder emails from a Google Sheet, tracks reply acknowledgements, logs all outcomes, and provides a guided setup wizard and menu-based diagnostics for operations staff.

---

## Project Goal

Provide a reliable, auditable reminder system that:

- sends reminder emails based on `Expiry Date` + `Notice Date` per row
- collects reply acknowledgements from clients using configurable keywords
- allows both scheduled and manual runs across multiple configured sheet tabs
- writes all outcomes back to the sheet and to a `LOGS` tab
- is easy to set up via a guided wizard — no scripting knowledge required

---

## Purpose and Business Value

- Non-technical users manage client rows directly in Google Sheets
- The system enforces required data and logs errors for correction
- Clients are prompted to reply with a keyword phrase to confirm receipt
- Reply scans run twice daily to keep acknowledgement status current
- The daily send time is configurable per team preference

---

## Feature Set

### 1. Guided Setup Wizard

A step-by-step first-time setup accessible from the top of the menu.

**Steps:**

1. **Select or create a sheet tab** — creates a new tab with correct headers pre-filled (bold, frozen row 1), or registers an existing tab
2. **Verify column headers** — auto-detects required columns; offers manual mapping if any are missing
3. **Apply dropdowns** — sets `Status`, `Send Mode`, and `Notice Date` dropdown validation
4. **Activate daily schedule** — offers to enable the daily trigger at the configured send time

The wizard auto-registers the tab in Script Properties and sets `headerRow = 1`, `dataStartRow = 2` for new tabs.

---

### 2. Multi-Tab Automation

- Multiple sheet tabs can be configured for automation simultaneously
- Each tab has its own stored column map, header row, data start row, and notice date options
- `LOGS` sheet is shared across all tabs; each log entry records the source tab name

---

### 3. Daily Email Schedule

- Configurable send time — default is **8:00 AM** (Asia/Manila), user-adjustable via menu
- Send time is stored in Script Properties (`DAILY_TRIGGER_HOUR`) and survives script updates
- `Activate Daily Schedule` installs the trigger using the stored time
- `Set Send Time` prompts for a new hour (0–23) without reinstalling the trigger

---

### 4. Data Processing Rules

- **Dynamic header mapping** — column order can change freely; script maps by header name
- **Flexible column aliases** supported (e.g. `Seller Email` → `Client Email`, `Due Date` → `Expiry Date`)
- **Required fields:** `Client Name`, `Client Email`, `Expiry Date`, `Notice Date`, `Status`
- **Send Mode** per row:
  - `Auto` — normal processing
  - `Hold` — skipped, logged as SKIPPED
  - `Manual Only` — skipped in bulk runs (still sendable via Inspect/Test)
- **Status handling:**
  - `Active` rows are processed
  - Blank `Status` rows are auto-activated to `Active` and processed
  - `Sent` / `Error` rows are skipped in normal runs
- **Guard for empty tabs** — tabs with no columns are skipped gracefully (no crash)

---

### 5. Date and Eligibility Logic

- Notice string parser supports any of:
  - `On expiry date` → sends on the expiry date itself
  - `N days before` / `N weeks before` / `N months before` / `N years before`
  - Plain integer (treated as days before)
  - ISO date string (`YYYY-MM-DD`) for a fixed target date
- Send eligibility: row sends when `targetDate <= today` (due or overdue catch-up)

---

### 6. Email Composition and Sending

- **Subject:** `[Document Type] – Expiry: [Expiry Date] – [Client Name]`
- **Body source priority:**
  1. `Remarks` column template (with placeholder substitution)
  2. Configurable fallback template (stored in Script Properties)
  3. Built-in hardcoded fallback template
- **Placeholder substitution** (case-insensitive):
  - `[Client Name]`
  - `[Date of Expiration]`, `[Date of Expiry]`, `[Expiry Date]`
  - `[Document Type]`
- **Reply acknowledgement phrase** — every outgoing email includes a styled callout:

  > Kindly reply **"RECEIVED"** to confirm that you have received and read this reminder.

  The phrase uses the first configured reply keyword (bold + underlined, blue callout box). This prompts the client to reply, which the reply scan detects automatically.

- **Sender name** — derives from the running account's email domain (for example, `@duranschulze.com` → `Duranschulze`, `@filepino.com` → `Filepino`); falls back to `Office`
- **Recipients:** primary from `Client Email`; optional CC from `Staff Email` column and/or default CC list
- **Drive attachments** from `Attached Files` column (Google Drive URLs or file IDs)
- **Attachment fallback** — if a file cannot be fetched, a clickable Drive link is appended to the email body instead of failing silently
- **Final notice emails** — second-stage emails for overdue rows; stored in separate metadata columns

---

### 7. Post-Send Updates and Logging

**On success:**

- `Status` → `Sent`
- `Staff Email` → running account's email address
- `Sent At`, `Sent Thread Id`, `Sent Message Id` written (when columns exist)
- `Open Tracking Token` written (when column exists and open tracking is configured)
- `Reply Status` → `Pending` (when column exists)
- `SENT` entry appended to LOGS

**On failure:**

- `Status` → `Error`
- `ERROR` entry appended to LOGS

**Other:**

- `SKIPPED` log for Hold / Manual Only rows
- End-of-run `SUMMARY` log includes totals: sent, errors, skipped, auto-activated

---

### 8. Reply Tracking

- **Configurable keywords** — default: `ACK`, `RECEIVED`, `OK`; stored in Script Properties
- **2x daily reply scan** — runs at **9:00 AM and 3:00 PM** (Asia/Manila) when activated
- Manual run available via `Run & Diagnostics → Run Reply Scan Now`
- Scan looks up sent Gmail threads by `Sent Thread Id`, reads replies from the client's address after the send timestamp, and checks body/subject for configured keywords
- On match, updates:
  - `Reply Status` → `Replied`
  - `Replied At` → reply timestamp
  - `Reply Keyword` → matched keyword
- After a row is marked `Replied`, later non-final reminder emails are suppressed for that row; the final expiry-day email can still send
- Rows without `Sent Thread Id` are skipped

---

### 9. Open/Read Tracking (Best Effort)

- Outgoing emails contain a 1×1 tracking pixel linked to a deployed Web App endpoint
- `doGet` handler records open/click events and updates the sheet:
  - `First Opened At`
  - `Last Opened At`
  - `Open Count`
- Requires the script to be deployed as a Web App and the URL set via `Set Open Tracking URL`
- Open tracking is best effort — email clients that block remote images will not trigger it

---

### 10. Diagnostics and Testing

- **Preview Target Dates** — shows what would send today without sending anything
- **Inspect Row** — displays all field values for a given row number
- **Send Test Email by No.** — sends a real email to the client of a specific row (by `No.` value)
- **Test Gmail Send** — prompts for any recipient address, sends a test email to verify Gmail connectivity
- **Test Drive Access** — verifies Drive API access and shows root folder name
- **Test All Connections** — checks Gmail + Drive + configured tabs in one report
- **Check Column Mappings** — shows detected columns across all configured tabs
- **Validate Tab Structure** — checks required columns are present on the working tab
- **View Tab Configuration** — shows stored header row, data start row, column mappings, notice options
- **Check Reply Tracking Setup** — reports whether thread ID / reply status columns exist and trigger is active
- **Preview Fallback Template** — renders the fallback template with sample data
- **System Diagnostics** — full system report: tabs, triggers, column maps, property values

---

## Spreadsheet Structure

### Default layout (new tabs created by wizard)

| Row | Purpose                   |
| --- | ------------------------- |
| 1   | Header row (bold, frozen) |
| 2+  | Data rows                 |

> Existing tabs can use a different header row — configurable via `Tab Management → Set Tab Header Row`.

### Required column headers

| Header         | Aliases accepted                                                         |
| -------------- | ------------------------------------------------------------------------ |
| `Client Name`  | `Seller`, `Seller Name`, `Buyer`, `Buyer Name`, `Applicant Name`         |
| `Client Email` | `Seller Email`, `Buyer Email`, `Email`, `Email Address`, `Contact Email` |
| `Expiry Date`  | `Due Date`                                                               |
| `Notice Date`  | `Remaining Days`, `Reminder Days`                                        |
| `Status`       | `Project Status`                                                         |

### Optional column headers used by the script

| Header                    | Aliases                                   | Purpose                                    |
| ------------------------- | ----------------------------------------- | ------------------------------------------ |
| `No.`                     | `#`, `ID`, `Ref No`                       | Row identifier for diagnostics             |
| `Type of ID/Document`     | `Services`, `Service`, `Visa Type`        | Used in email subject and placeholders     |
| `Remarks`                 | `Reminder (Email Content)`                | Email body template; overrides fallback    |
| `Attached Files`          | `Attached File`, `GSheet`, `Google Sheet` | Drive URLs / file IDs for attachments      |
| `Staff Email`             | `Assigned Staff Email`                    | Overrides CC; also written back after send |
| `Send Mode`               | `Send Option`, `Mode`                     | `Auto` / `Hold` / `Manual Only`            |
| `Sent At`                 | —                                         | Timestamp of last successful send          |
| `Sent Thread Id`          | `Sent Thread ID`                          | Gmail thread ID for reply scan             |
| `Sent Message Id`         | `Sent Message ID`                         | Gmail message ID                           |
| `Reply Status`            | —                                         | `Pending` or `Replied`                     |
| `Replied At`              | —                                         | Timestamp of detected reply                |
| `Reply Keyword`           | —                                         | Keyword matched in client reply            |
| `Open Tracking Token`     | `Open Token`                              | Pixel token for open tracking              |
| `First Opened At`         | `First Open At`                           | Timestamp of first tracked open            |
| `Last Opened At`          | `Last Open At`                            | Timestamp of most recent open              |
| `Open Count`              | —                                         | Number of tracked opens                    |
| `Final Notice Sent At`    | `Final Sent At`                           | Timestamp of final notice email            |
| `Final Notice Thread Id`  | `Final Thread ID`                         | Thread ID of final notice                  |
| `Final Notice Message Id` | `Final Message ID`                        | Message ID of final notice                 |

---

## Menu Structure

```
🔔 Expiry Notifications
├── 🚀 Setup This Sheet for Automation
├── ────────────────────────────────
├── Status & Overview
│   ├── Show System Status
│   ├── View Last Run Summary
│   └── Show All Tabs Info
├── ────────────────────────────────
├── Tab Management
│   ├── Configure Automation Tabs
│   ├── Select Working Tab
│   ├── Setup Tab Dropdowns
│   ├── Map Tab Columns
│   └── Set Tab Header Row
├── ────────────────────────────────
├── Automation Settings
│   ├── Activate Daily Schedule
│   ├── Set Send Time
│   ├── Deactivate Daily Schedule
│   ├── Set Default CC Emails
│   ├── ────────────────────────
│   ├── Set Reply Keywords
│   ├── Activate Reply Scan (2x Daily)
│   ├── Deactivate Reply Scan
│   ├── ────────────────────────
│   └── Set Open Tracking URL
├── ────────────────────────────────
├── Run & Diagnostics
│   ├── Run Manual Check Now
│   ├── Preview Target Dates
│   ├── ────────────────────────
│   ├── Inspect Row...
│   ├── Send Test Email by No....
│   ├── ────────────────────────
│   ├── Test Gmail Send
│   ├── Test Drive Access
│   ├── Test All Connections
│   ├── ────────────────────────
│   ├── Check Column Mappings
│   ├── Validate Tab Structure
│   ├── View Tab Configuration
│   ├── Check Reply Tracking Setup
│   ├── Preview Fallback Template
│   ├── ────────────────────────
│   └── System Diagnostics
├── ────────────────────────────────
└── Help
    ├── Want to integrate this to another google sheet?
    ├── ────────────────────────
    ├── View Documentation
    └── About
```

---

## Runtime Flow — Daily Check (`runDailyCheck`)

1. Resolve all configured automation tabs
2. Ensure `LOGS` sheet exists
3. For each configured tab:
   a. Build and validate column map (skip tab if no columns)
   b. Read data rows from configured data start row
   c. For each row:
   - Skip non-processable statuses (`Sent`, `Error`, `Skipped`)
   - Apply Send Mode gate (`Hold` / `Manual Only` → skip with log)
   - Auto-activate blank Status to `Active`
   - Validate required fields
   - Parse Notice Date offset → compute target date
   - Skip if target date is in the future
   - Skip later non-final reminders when `Reply Status = Replied` (but still allow the final expiry-day send)
   - Resolve Drive attachments (failed attachments → fallback links in email body)
   - Build email body (Remarks → fallback template)
   - Inject reply acknowledgement phrase (first configured keyword, bold + underlined)
   - Send email via GmailApp
   - On success: update Status, Staff Email, sent metadata, Reply Status
   - On failure: update Status to Error
   - Append SENT / ERROR log
4. Append SUMMARY log with run totals

---

## Runtime Flow — Reply Scan (`runReplyScan`)

Runs at **9:00 AM and 3:00 PM** daily (Asia/Manila) when activated.

1. Resolve all configured automation tabs
2. For each tab:
   a. Find rows where `Reply Status = Pending` and `Sent Thread Id` is populated
   b. For each such row, look up the Gmail thread
   c. Find reply messages from the client address, sent after the original send timestamp
   d. Check message body/subject for configured reply keywords
   e. On match: set `Reply Status = Replied`, write `Replied At` and `Reply Keyword`
3. Append scan summary to LOGS

---

## Script Properties Reference

All settings are stored in **Script Properties** (not in the sheet). View/manage via Apps Script editor → Project Settings → Script Properties.

| Property key             | Set via                                       | Purpose                                               |
| ------------------------ | --------------------------------------------- | ----------------------------------------------------- |
| `SPREADSHEET_ID`         | Auto (on open)                                | Locks the script to the source spreadsheet            |
| `AUTOMATION_SHEET_NAMES` | Setup wizard / Tab Management                 | Comma-separated list of configured tab names          |
| `DAILY_TRIGGER_HOUR`     | Automation Settings → Set Send Time           | Hour (0–23) for the daily trigger; default 8          |
| `REPLY_KEYWORDS`         | Automation Settings → Set Reply Keywords      | Comma-separated keywords; default `ACK, RECEIVED, OK` |
| `DEFAULT_CC_EMAILS`      | Automation Settings → Set Default CC Emails   | Default CC list for all outgoing emails               |
| `FALLBACK_TEMPLATE_MODE` | Run & Diagnostics → Preview Fallback Template | `HARDCODED` or `PROPERTY`                             |
| `FALLBACK_TEMPLATE`      | Run & Diagnostics → Preview Fallback Template | Custom fallback template text                         |
| `OPEN_TRACKING_BASE_URL` | Automation Settings → Set Open Tracking URL   | Web App URL for open tracking                         |
| `TAB_CONFIG_[tabName]_*` | Tab Management                                | Per-tab header row, column map, notice options        |

---

## Security and Operations Notes

- Script requires **Gmail**, **Drive**, and **Spreadsheet** permissions (granted on first run)
- API keys and settings are stored in **Script Properties**, never in sheet cells
- Duplicate triggers can accumulate if `Activate` is run multiple times — always `Deactivate` first, then `Activate` to reset
- For open tracking, deploy the script as a **Web App** (Execute as: Me, Access: Anyone) and paste the URL into `Set Open Tracking URL`
- Gmail sending quotas apply: ~100 emails/day for consumer accounts, ~1,500/day for Workspace accounts

---

## Known Constraints

- Open tracking is best effort — email clients that block remote images will not register opens
- Reply matching requires `Sent Thread Id` to be populated on the row; rows sent before the column existed cannot be scanned
- Tabs with no column headers are skipped gracefully during automation runs

---

## Key Files

| File        | Purpose                                                                |
| ----------- | ---------------------------------------------------------------------- |
| `code.gs`   | Full implementation — all functions, constants, menu, wizard, triggers |
| `README.md` | Product and operational documentation (this file)                      |
