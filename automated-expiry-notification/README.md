# Automated Expiry Notification

Google Apps Script automation that sends staged expiry-reminder emails from a
Google Sheet, tracks Gmail replies, and tracks email opens via a 1×1 pixel.
Each row's outgoing email is sent FROM the row's `Assigned Staff Email`, so
the client sees the staff member as the sender.

## Column contract

### Team A — required user-input columns

The setup wizard validates that all eight are present (header aliases are
honored — e.g. an existing "Staff Email" column matches "Assigned Staff
Email"). If any are missing, the wizard offers to add them.

| Column                     | Notes                                              |
| -------------------------- | -------------------------------------------------- |
| Company/Client Name        |                                                    |
| Client Email               |                                                    |
| Expiry Date / Renewal Date |                                                    |
| Notice Date                |                                                    |
| Description                | (was "Remarks" — alias preserved)                  |
| Attached Files             |                                                    |
| Name of Staff              | display name on outgoing emails                    |
| Assigned Staff Email       | From-address; must be a Gmail "Send mail as" alias |

### Team V — code-managed columns

Headers are auto-created only when missing. Existing user-renamed variants are
matched via aliases and preserved.

`Status`, `Reply Status`, `Final Notice Sent At`, `Final Notice Thread Id`,
`Final Notice Message Id`, `Send Mode`, `Sent At`, `Sent Thread Id`,
`Sent Message Id`, `Replied At`, `Reply Keyword`, `Open Tracking Token`,
`First Opened At`, `Last Opened At`, `Open Count`.

Both lists are exported from [Foundation.gs](Foundation.gs)
as `REQUIRED_USER_COLUMNS` and `MANAGED_COLUMNS`.

## File layout

Seven `.gs` files at the project root, each one responsibility. Apps Script
is a single namespace, so identifiers in any file are visible to every other
file. Each file uses chapter banners (`// ═══ ChapterName ═══`) to mark the
former module boundaries.

| File                                       | Owns                                                                                                                                                                              |
| ------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| [Foundation.gs](Foundation.gs)             | Constants, enums, header aliases, column contract, pure helpers, properties access, configured-tab resolution, sheet/row I/O, ensure-column logic. Everything else builds on this. |
| [Menu.gs](Menu.gs)                         | The `🔔 Expiry Notifications` menu (`onOpen`) and read-only info dialogs (status, docs, about, integration link).                                                                  |
| [Setup.gs](Setup.gs)                       | Everything the user clicks while configuring a sheet: multi-tab picker, column mapping (detection + UI), dropdowns, user-column validation, and the guided setup wizard.          |
| [Send.gs](Send.gs)                         | Outbound-email pipeline: send-mode rules → alias verification → AI body generation → attachments → subject/body composition → Gmail delivery + CC.                                |
| [Tracking.gs](Tracking.gs)                 | Post-send observation: open-pixel handler (`doGet`) and reply-scan orchestration (`runReplyScan`).                                                                                |
| [Orchestration.gs](Orchestration.gs)       | Runtime loop (`runDailyCheck`, `manualRunNow`), trigger install/remove, LOGS sheet bootstrap and append.                                                                          |
| [Diagnostics.gs](Diagnostics.gs)           | Test/inspect utilities surfaced through the menu (preview, inspect row, send test, connectivity tests, system check).                                                             |
| [appsscript.json](appsscript.json)         | Apps Script manifest.                                                                                                                                                             |

## Import into Google Sheet Apps Script

1. In the target spreadsheet open **Extensions → Apps Script**.
2. For each `.gs` file in this repo, create a matching file in the editor
   (**Files → ＋ → Script**) using the same name (without the `.gs`
   suffix — the editor adds it). Paste the file contents.
3. Open **Project Settings → Show "appsscript.json" manifest file in editor**.
4. Replace the editor's `appsscript.json` with this repo's
   [appsscript.json](appsscript.json).
5. Reload the spreadsheet — the **🔔 Expiry Notifications** menu appears.
   Run **🚀 Setup This Sheet for Automation** to begin.

## Stable entrypoints

These names are kept for menu, trigger, and web-app deployment compatibility:
`onOpen`, `runDailyCheck`, `manualRunNow`, `runReplyScan`, `doGet`, plus all
menu-exposed setup and diagnostics functions.

## Per-row sender

[Send.gs](Send.gs) calls `GmailApp.sendEmail` with
`{ from: staffEmail, name: staffName }`. Gmail only accepts the `from`
option when the address is a verified "Send mail as" alias on the
script-runner's Gmail account.

The AliasResolver chapter inside `Send.gs` checks `GmailApp.getAliases()`
(cached per execution) and exposes `canSendAs(email)`. The daily run
(in `Orchestration.gs`) skips any row whose staff email isn't verified,
marks the row Status = `Error`, and writes a clear message to the LOGS
sheet. The setup wizard runs the same check across every distinct staff
email and warns the user before scheduling.

## Multi-tab menu

Every per-tab setup operation ("Map Tab Columns", "Setup Tab Dropdowns",
"Set Tab Header Row", "Select Working Tab") accepts a comma-separated tab
list (`1,3` or `1-3` or names) via the `promptSelectTabs` helper inside
the UiPrompts chapter of [Setup.gs](Setup.gs). The selected tabs are
processed in a loop and a per-tab result summary is shown at the end.

## Main flows

### Daily run ([Orchestration.gs](Orchestration.gs))

1. `runDailyCheck` resolves configured tabs and resets the alias cache.
2. For each tab, ensure managed columns → validate Team A columns → load
   data rows.
3. Per row: validate Team A fields, check send mode, decide notice vs final,
   verify the staff email is a valid alias, compose body (manual remarks → AI
   → fallback template), inject the open-tracking pixel, send via Gmail with
   per-row `from`, write metadata, log.

### Reply scan ([Tracking.gs](Tracking.gs))

Triggered 9 AM and 3 PM Asia/Manila. For each row with a thread id, loads
the Gmail thread and matches keyword regex. Self-message filter covers ALL
verified aliases (not just the runner) since outgoing mail can come from
any staff alias.

### Open tracking ([Tracking.gs](Tracking.gs))

`doGet(e)` handles `?mode=open&t=TOKEN`, finds the row across configured
tabs, increments open count, writes timestamps, marks Reply Status as
`Replied` with keyword `OPEN_TRACKED`.

### Setup wizard ([Setup.gs](Setup.gs))

Tab pick/create → user-column validation (offer to auto-add missing) →
ensure managed columns → dropdowns → staff alias pre-flight → schedule
install.
