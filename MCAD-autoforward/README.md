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

On open, the script also creates a managed **AutoForward Info** tab. It contains onboarding steps, rule-column explanations, menu guidance, account-isolation details, privacy limitations, and troubleshooting. The tab is versioned and is refreshed only when its built-in guide version changes. If an unrelated tab already uses that name, the script creates a uniquely named **AutoForward Guide** tab instead of overwriting user data.

## First-time user workflow

Every Gmail account must complete these steps while signed in as that account:

1. Open the AutoForward Google Sheet.
2. Select **AutoForward → Create or open my rule tab**.
3. Accept the Google authorization prompt when shown.
4. Enter forwarding rules in the generated tab.
5. Select **AutoForward → Validate my rules**.
6. Select **AutoForward → Preview matching emails** if a read-only preview is desired.
7. Select **⚡ AutoForward → ▶️ Start or repair my automation**.

Activation creates three installable triggers owned by the current Gmail user:

- Gmail monitoring every five minutes.
- A daily summary during the 11 PM hour in `Asia/Manila`.
- A self-repair watchdog every six hours.

Messages received before the user's first activation time are not forwarded.

The monitor, summary, and watchdog check the same trigger set. If one trigger
remains, it can recreate the other missing AutoForward triggers. If all three
triggers are deleted, authorization is revoked, or Google disables execution,
the user must choose **▶️ Start or repair my automation** once.

## Distributing private copies

Give every recipient their own Google Sheets copy. Installable triggers always
run as the account that created them, so each recipient must start their copy
once using their own Gmail account.

On first Start, AutoForward removes copied registration metadata that points to
the source workbook. If the private copy contains exactly one valid
`Rules - ...` tab and has no local registrations yet, that tab is reassigned to
the signed-in user automatically. If no single safe starter tab can be
identified, AutoForward creates a new account-owned rule tab instead of
selecting one ambiguously.

If a saved registration points to a rule tab that was deleted, the next
**▶️ Start or repair my automation** action removes that broken local
registration and safely rebuilds it. The source workbook and previously
recorded successful message IDs are not modified.

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

- **📖 Open usage guide** — opens the managed AutoForward Info tab and creates it if the automatic open-time setup could not do so.
- **👤 Set up or open my rule tab** — creates or safely reuses one account-owned tab for the signed-in user.
- **♻️ Adopt my active legacy rule tab** — converts and registers an existing six-column `Enabled / Sender / Match All / Keywords / Recipients / Notes` tab.
- **✅ Validate my rules** — validates enabled rows without reading or changing Gmail.
- **🔎 Preview matching emails** — reads Gmail and logs recent matches without forwarding or labeling.
- **📬 Preview pending emails** — shows only messages the next live run would attempt.
- **▶️ Start or repair my automation** — validates rules and installs or repairs the current user's three AutoForward triggers.
- **📊 Show automation status** — checks trigger health, performs repair when active, and shows account, rule tab, last run, and last error.
- **⏸️ Pause automation** — removes the current user's three triggers and keeps rules/history.
- **🧹 Reset my processed history** — confirms, pauses, and clears the current user's processed and summary history.

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

The monitor searches enabled senders in the inbox, spam (when enabled in `CONFIG`), and previously failed threads. It processes the oldest messages first and forwards the original Gmail message with its attachments. Successfully forwarded message IDs and retry state are stored separately for each Gmail user; those message-level records, not Gmail labels alone, control duplicate prevention and retries.

Gmail thread labels are:

- `AutoForward/Detected` — the conversation contains at least one matching message.
- `AutoForward/Forwarded` — Gmail accepted a forward request for at least one message in the conversation; this does not prove final recipient delivery.
- `AutoForward/Failed` — at least one message is waiting for another retry.
- `AutoForward/Retry Exhausted` — at least one message reached the retry limit and needs manual review.

Gmail labels apply to whole conversations. A conversation can therefore carry both `Forwarded` and `Failed`, or both `Forwarded` and `Retry Exhausted`, when its individual messages have different outcomes. The `Forwarded` label does not by itself mean every message in that conversation was sent.

Forwarding and recording the processed message ID are separate Google service operations. A rare interruption after Gmail accepts the forward but before Apps Script saves the processed ID can cause the message to be retried and forwarded more than once.

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
