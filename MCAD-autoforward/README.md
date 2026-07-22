# MCAD AutoForward

MCAD AutoForward is a Google Sheets-bound Apps Script that lets each Gmail user maintain an account-owned forwarding rule tab. Each participating user authorizes and activates their own Gmail automation; that automation reads only the rule tab registered to that Gmail account.

## What changed

The script no longer uses a fixed spreadsheet ID or a fixed tab such as `Rules - Code.gs`.

For example, when `seo@filepino.com` chooses **AutoForward → Create or open my rule tab**, the script:

1. Detects `seo@filepino.com` as the effective user.
2. Creates `Rules - seo@filepino.com`.
3. Registers the tab by its permanent numeric sheet ID.
4. Protects the tab so `seo@filepino.com`, the spreadsheet owner, and configured administrators can edit it.
5. Uses that registered sheet ID whenever a trigger owned by `seo@filepino.com` runs.

The tab name is a display label only. Renaming or reordering the tab does not redirect the automation.

Processed-message history, activation time, summary records, cleanup state, locks, and triggers are isolated per Google user.

## Required deployment model

Place `code.gs` in the Apps Script project bound to the rules spreadsheet:

1. Open the Google Sheet.
2. Select **Extensions → Apps Script**.
3. Replace the editor contents with `code.gs`.
4. Save the project.
5. Reload the Google Sheet.

The **AutoForward** menu should appear after the reload. A standalone Apps Script project will not receive the spreadsheet `onOpen` menu in the intended way.

## First-time user workflow

Every Gmail account must complete these steps while signed in as that account:

1. Open the AutoForward Google Sheet.
2. Select **AutoForward → Create or open my rule tab**.
3. Accept the Google authorization prompt when shown.
4. Enter forwarding rules in the generated tab.
5. Select **AutoForward → Validate my rules**.
6. Select **AutoForward → Preview matching emails** if a read-only preview is desired.
7. Select **AutoForward → Activate or repair my automation**.

Activation creates two installable triggers owned by the current Gmail user:

- Gmail monitoring every five minutes.
- A daily summary during the 11 PM hour in `Asia/Manila`.

Messages received before the user's first activation time are not forwarded.

## Rule table

| Column | Description |
|---|---|
| Enabled | Checkbox controlling whether the row is active. |
| Sender | One exact sender email address. |
| Match Mode | `All messages` or `Any keyword`. |
| Keywords | One keyword or phrase per line. Used only for `Any keyword`. |
| Recipients | One forwarding address per line. Commas and semicolons are also accepted. |
| Notes | Optional operator notes; ignored by automation. |

`Any keyword` checks the email subject and plain-text body. Matching is case-insensitive. Short uppercase codes such as `AFS`, `OTP`, and `MC28` use whole-word matching.

Rules are evaluated from top to bottom. The first matching row is used. Put sender-specific keyword rows above an `All messages` row for the same sender. Validation rejects rows hidden below an earlier `All messages` rule.

## AutoForward menu

- **Create or open my rule tab** — creates one generated tab for the signed-in account or opens its existing tab.
- **Adopt my active legacy rule tab** — converts and registers an existing six-column `Enabled / Sender / Match All / Keywords / Recipients / Notes` tab. It validates first and does not activate forwarding.
- **Validate my rules** — validates all enabled rows without reading or changing Gmail.
- **Preview matching emails** — reads Gmail and logs recent matches without forwarding or labeling.
- **Preview pending emails** — shows only messages the next live run would attempt. The account must already be activated.
- **Activate or repair my automation** — validates rules, creates Gmail labels, and replaces only the current user's AutoForward triggers.
- **Show my status** — shows the assigned account, rule tab, activation, trigger count, last run, and last error.
- **Pause my automation** — removes only the current user's AutoForward triggers and keeps rules/history.
- **Reset my processed history** — confirms, pauses, then clears only the current user's processed and summary history. Reactivation can forward recent messages again.

## Existing rule-tab migration

To migrate `Rules - Code.gs` or `Rules - Code2.gs`:

1. Decide which Gmail account owns the tab.
2. Sign in as that Gmail account.
3. Open the shared spreadsheet and select the legacy tab.
4. Select **AutoForward → Adopt my active legacy rule tab**.
5. Review the converted `Match Mode` values.
6. Validate and preview.
7. Activate only after the results are correct.

The legacy columns must be in this order: Enabled, Sender, Match All, Keywords, Recipients, and optional Notes. Adoption adds the account/status area above the table and changes `Match All` booleans into the `Match Mode` dropdown.

Do not activate the old and new script versions for the same Gmail account simultaneously.

## Gmail behavior

The monitor searches enabled senders in the inbox, spam (when enabled in `CONFIG`), and previously failed threads. It processes the oldest messages first and forwards the original Gmail message with its attachments.

Gmail thread labels are:

- `AutoForward/Detected`
- `AutoForward/Forwarded`
- `AutoForward/Failed`

A message is marked processed only after forwarding succeeds. Failed messages remain eligible for retry. Processed IDs are retained for 60 days by default.

## Privacy and protection limitation

A protected tab controls who can edit it, not who can see it. Anyone who can view the shared workbook may still see `Rules - seo@filepino.com`.

If rule senders and recipients must be visible only to the assigned user and administrators, use a separate private spreadsheet and a separate bound copy of this Apps Script for each Gmail account. Each user must still authorize and activate their own copy because Google does not copy installable triggers or Gmail authorization.

## Administrator configuration

Optional administrators who should be able to edit every generated tab can be listed in `CONFIG.ADMIN_EMAILS`:

```javascript
ADMIN_EMAILS: [
  "automation-admin@filepino.com"
]
```

The spreadsheet owner is added automatically when Google exposes an owner address. Shared Drive files may not expose an individual owner, so configure at least one administrator for recovery in that deployment model.

## Important Google constraints

- A rule tab cannot grant access to a Gmail inbox.
- The Gmail account that creates the installable trigger is the mailbox that the trigger searches.
- Each participating user must authorize the Apps Script personally.
- Workspace administrators may block Gmail scopes, external forwarding, or recipient domains.
- Time-based triggers are approximate; the daily summary runs sometime during the configured hour.
- Gmail and Apps Script sending/runtime quotas still apply.

## Safe rollout

1. Back up the existing Google Sheet and Apps Script project.
2. Configure administrator addresses.
3. Pilot with one noncritical Gmail account.
4. Create/adopt, validate, and preview its rules.
5. Activate and confirm one controlled test message.
6. Confirm labels, duplicate prevention, and the daily summary.
7. Enroll remaining Gmail accounts one at a time.
