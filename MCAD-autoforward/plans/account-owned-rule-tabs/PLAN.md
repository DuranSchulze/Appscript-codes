# Account-Owned Rule Tabs

## 1. Goal
Create a future-proof Google Sheets experience where each Gmail user can use an AutoForward custom menu to generate, open, validate, and activate a standardized rule tab assigned to that Gmail account. The forwarding trigger must run only against the Gmail account that authorized it and must load only that account's assigned rule tab.

## 2. Context Summary
The current project has one `code.gs` file. It reads a fixed spreadsheet ID and fixed tab name, searches the executing user's Gmail, forwards matching messages, stores processed state in Script Properties, applies Gmail labels, and sends a daily summary. The existing workbook examples use six columns: Enabled, Sender, Match All, Keywords, Recipients, and Notes.

The requested workflow is multi-user: for example, `info@filepino.com` opens the shared Google Sheet, chooses a menu command, receives a generated rule tab, enters senders/matching rules/recipients, and activates automation for `info@filepino.com` only.

Confirmed technical boundary: a Google Sheet does not grant access to a person's Gmail. Each Gmail user must personally authorize the Apps Script and create that user's installable Gmail/time triggers. A tab-to-account mapping can select rules, but the trigger's creator determines which Gmail mailbox is searched.

Assumptions:

- The Apps Script will be container-bound to the shared Google Sheet so `onOpen` can add a custom menu.
- All participating accounts are allowed to authorize Gmail and Spreadsheet scopes under the organization's Google Workspace policies.
- Rule tabs may be visible to other spreadsheet viewers. Google Sheets can restrict editing per tab, but it cannot securely hide one tab from selected editors in the same workbook.
- The spreadsheet owner/administrators remain emergency editors of protected tabs.
- The first version preserves the current rule semantics: match every message from a sender, or match when any listed keyword appears in the subject/plain-text body.

Open decisions to confirm before implementation:

- Whether users may forward to any valid address or only approved domains/addresses.
- Which administrator accounts should be allowed to repair every protected tab.
- Whether existing `Rules - Code.gs` and `Rules - Code2.gs` data should be migrated into named account tabs and which Gmail account owns each existing tab.
- Whether rule visibility between spreadsheet users is acceptable. Google Sheets does not support per-user tab visibility. If `seo@filepino.com` must be the only ordinary user who can see its rules, use one private spreadsheet per Gmail account instead of account tabs in a shared workbook.

## 3. Scope

- Add an `AutoForward` custom menu when the spreadsheet opens.
- Let the signed-in user create or open exactly one account-owned rule tab.
- Generate a polished, consistent sheet layout with instructions, account/status information, filters, frozen headers, checkboxes, dropdowns, wrapping, validation, and conditional formatting.
- Maintain an internal account registry using immutable sheet IDs rather than relying only on editable tab names.
- Protect each generated rule tab so its assigned Gmail user and configured administrators can edit it.
- Make activation, triggers, processed-message state, summary records, and configuration account-specific.
- Ensure a user's automation resolves and reads only the tab registered to that executing Gmail account.
- Add menu actions for validation, preview, activation/status, pausing, and safe reset.
- Preserve existing Gmail forwarding, retry, labels, duplicate prevention, and daily summary behavior.
- Document onboarding, authorization, ownership limits, recovery, and migration.

## 4. Out of Scope

- Building a separate web application or external database.
- Allowing one user to authorize Gmail access on behalf of another user.
- Providing true per-tab confidentiality inside one shared spreadsheet.
- Sending or forwarding email during tab creation.
- Supporting Outlook or non-Gmail source mailboxes.
- Automatically bypassing Google Workspace OAuth, Gmail forwarding, or administrator restrictions.
- Changing keyword semantics beyond the existing Match All/keyword-any behavior in the first release.

## 5. Affected Files and Folders

```txt
MCAD-autoforward/
├── code.gs
├── README.md
├── appsscript.json                    (candidate, if the deployed project manifest is checked in)
├── AccountContext.gs                  (candidate)
├── AutoForwardMenu.gs                 (candidate)
├── RuleSheetManager.gs                (candidate)
├── ForwardingService.gs               (candidate extraction from code.gs)
└── plans/
    └── account-owned-rule-tabs/
        └── PLAN.md
```

- `code.gs`: confirmed current implementation; its fixed tab configuration, shared properties, locks, setup, reset, and trigger behavior must be refactored.
- `README.md`: confirmed current requirements file; update it into an operator/user guide and retain the original forwarding requirements as migration reference.
- `appsscript.json`: optional manifest to make scopes, time zone, and runtime explicit if the Apps Script project is managed through source control.
- Candidate `.gs` files: recommended separation of account identity/state, menu handling, sheet generation/protection, and forwarding logic. Apps Script files share one project/global namespace, so this is organizational rather than a runtime module system.
- The live Google Sheet will gain a protected internal registry tab and one generated rule tab per enrolled Gmail account.

## 6. Step-by-Step Implementation Plan

1. Define the account identity and tenancy contract.
   - Treat the effective user who manually enrolls/activates as the Gmail account owner.
   - Normalize the email address and create a stable account key that is safe for property names and logs.
   - Fail closed when Google does not return a usable account email; do not guess from the tab name or active cell.
   - Do not require a configured rule-tab name. Resolve the executing identity first, then discover that identity's registered sheet by immutable sheet ID.
   - Document that the Gmail mailbox is determined by the trigger creator, not the spreadsheet tab.
   - Affects `AccountContext.gs` or the equivalent section of `code.gs`; this must precede registry and trigger changes.

2. Add an internal account registry.
   - Create a protected administrative tab such as `_AutoForward Accounts`.
   - Store normalized Gmail address, immutable rule sheet ID, display tab name, enrollment state, activation state, trigger/setup timestamps, last successful run, last error summary, and schema version.
   - Use the numeric sheet ID as the durable reference so renaming a visible tab does not disconnect automation.
   - Reject duplicate account registrations and duplicate sheet assignments.
   - Hide the registry tab for convenience, while documenting that hidden/protected tabs are not a confidentiality boundary.
   - Affects `RuleSheetManager.gs` and `AccountContext.gs`; depends on Step 1.

3. Add the spreadsheet custom menu.
   - Add an `AutoForward` menu from a lightweight `onOpen` handler.
   - Include: `Create or open my rule tab`, `Validate my rules`, `Preview pending emails`, `Activate or repair automation`, `Show my status`, `Pause automation`, and `Reset my processed history`.
   - Require an explicit confirmation for reset because it can cause duplicate forwarding.
   - Keep menu creation authorization-free; request OAuth only when the user selects an action that needs Gmail, trigger, protection, or spreadsheet access.
   - Affects `AutoForwardMenu.gs`; depends on Step 1.

4. Generate the account-owned rule tab.
   - If the account is already registered, open its existing tab instead of creating a duplicate.
   - Otherwise, generate a sanitized human-readable name from the detected account, such as `Rules - seo@filepino.com`, handling the 100-character tab limit and name collisions.
   - Treat the generated name as a display label only. Store and subsequently resolve the immutable numeric sheet ID so a rename cannot redirect the automation.
   - Add a compact status/header area showing the assigned Gmail account, automation state, last validation, last run, and brief instructions.
   - Add the rule table below the status area with columns: Enabled, Sender, Match Mode, Keywords, Recipients, and Notes.
   - Use checkboxes for Enabled, a dropdown for Match Mode (`All messages` or `Any keyword`), wrapped multiline keyword/recipient cells, alternating row colors, frozen table headers, a filter, intentional column widths, and conditional formatting for disabled or invalid-looking rows.
   - Protect title/status/formula cells while leaving only intended rule-entry cells editable by the account owner and administrators.
   - Register the tab only after creation, formatting, validation rules, and protection succeed; cleanly report partial failures.
   - Affects `RuleSheetManager.gs`; depends on Steps 1 and 2.

5. Implement account-scoped protection and ownership checks.
   - Protect the generated sheet/ranges and set the account owner plus configured administrators as allowed editors where Workspace permissions permit.
   - Before every menu mutation and every automation run, verify that the executing account matches the registry owner of the target sheet.
   - Never select rules from the active tab alone; resolve account to registered sheet ID and confirm the sheet still belongs to the registry entry.
   - Provide an administrator repair path for deleted sheets, renamed tabs, changed protections, and departed users without silently reassigning ownership.
   - Affects `AccountContext.gs` and `RuleSheetManager.gs`; depends on Steps 2 and 4.

6. Refactor rule loading to use the current account's registered tab.
   - Remove the fixed `RULES_SHEET_NAME` dependency completely and use the executing account's registry entry.
   - During a time-trigger run, detect the trigger creator/effective Gmail account, normalize it, look up that account's registered sheet ID, verify ownership metadata, and load only that sheet.
   - Never scan for a tab by partial email-name matching during an automation run; names are editable and are not safe identifiers.
   - Keep the spreadsheet ID only if the deployment remains standalone; prefer the bound spreadsheet identity for a container-bound project.
   - Read the new table's header row from its defined location and map display labels to internal normalized fields.
   - Preserve compatibility with current Match All rows during migration, translating them to the new Match Mode representation.
   - Continue validating all enabled rows before activation or forwarding.
   - Affects `code.gs`/`ForwardingService.gs` and `RuleSheetManager.gs`; depends on Steps 2 and 4.

7. Isolate runtime state per Gmail account.
   - Replace shared Script Properties for start time, processed message IDs, summary records, summary recipient, and cleanup timestamp with User Properties, or prefix every shared key with a verified account key where administrative reporting requires shared visibility.
   - Prefer User Properties for Gmail processing state because triggers run as individual users.
   - Replace the project-wide script lock with an account/user lock where appropriate so one mailbox's long run does not unnecessarily block another mailbox.
   - Keep message IDs and summary records isolated even if Gmail identifiers happen to overlap or multiple users enroll concurrently.
   - Affects `code.gs`/`ForwardingService.gs` and `AccountContext.gs`; depends on Step 1 and must be completed before multi-account activation.

8. Make trigger lifecycle account-specific.
   - Activation must be manually invoked by the Gmail account owner and must validate the account's rules before creating triggers.
   - Manage only the current user's project triggers; avoid deleting or interpreting another user's trigger state.
   - Save activation metadata in the account's registry entry and user-scoped properties.
   - Add a repair action that detects missing/duplicate triggers for the current user and recreates only that user's expected monitoring and summary triggers.
   - Pause must remove only the current user's AutoForward triggers while retaining rules and processed history.
   - Affects `code.gs`/`ForwardingService.gs`, `AutoForwardMenu.gs`, and `AccountContext.gs`; depends on Steps 6 and 7.

9. Improve validation and status feedback.
   - Validate the full rule set from the registered sheet before activation.
   - Write non-sensitive status information to the tab's protected status area and provide detailed validation errors in a dialog/sidebar or execution log.
   - Include row numbers and field names in errors, and do not expose email body content in shared status cells.
   - Track last successful run and a sanitized last-error summary in the registry/status area.
   - Affects `RuleSheetManager.gs`, `AutoForwardMenu.gs`, and forwarding error handling; depends on Steps 4, 6, and 8.

10. Migrate existing rules safely.
    - Obtain an explicit owner mapping for `Rules - Code.gs` and `Rules - Code2.gs`.
    - Back up/copy existing tabs before transformation.
    - Convert Match All values to Match Mode, preserve keywords/recipients/notes, register immutable sheet IDs, and validate each migrated account.
    - Do not activate triggers or forward messages as part of migration.
    - Affects the live spreadsheet and migration documentation; depends on Steps 2, 4, and 6 and on the unresolved owner mapping.

11. Update documentation and operational guidance.
    - Document user enrollment, OAuth prompts, rule entry examples, validation, preview, activation, pausing, reset risks, ownership/protection limits, and administrator recovery.
    - Document that each Gmail account must enroll itself and that copying a tab does not copy Gmail authorization or triggers.
    - Include a rollout checklist and a small pilot recommendation before enrolling all accounts.
    - Affects `README.md`; finalize after behavior and UI labels are stable.

## 7. Database Changes

No database changes required.

## 8. Backend Changes

Apps Script acts as the backend. It needs an account-context layer, protected account registry, account-specific trigger lifecycle, account-specific properties/locks, dynamic rule resolution, and fail-closed ownership checks.

The current forwarding engine can remain substantially intact after its dependencies are changed from global configuration to an account context containing the executing Gmail address, registered sheet ID, rule set, and user-scoped state store. Trigger handlers should reconstruct this context on every execution rather than trusting active-sheet state, because time triggers have no active tab.

Administrative status stored in the shared registry must be sanitized. Detailed Gmail message content, matched bodies, and recipient history should remain in the user's Gmail/log context rather than a shared sheet unless explicitly requested.

## 9. Frontend Changes

The Google Sheet is the frontend. The custom menu becomes the entry point, while each generated rule tab acts as an account dashboard and editor.

The recommended tab layout is:

- A title and assigned-account banner.
- A small read-only status block for activation, validation, last run, and last error.
- A short instruction strip with examples.
- A frozen, filterable rule table with constrained inputs and multiline fields.

Dialogs or a sidebar should handle confirmations and detailed errors. The main sheet should show concise state and avoid cluttering rule rows with system data. Formatting must remain readable on typical laptop widths and tolerate long recipient lists through wrapping and controlled row heights.

Two deployment experiences must be distinguished:

- Shared-workbook mode: `Rules - seo@filepino.com` can be protected so only that user and administrators edit it, but other users with workbook access can still see it.
- Private-workbook mode: each Gmail account receives a separate restricted spreadsheet generated/copied from the managed template. This is the required mode when rules and recipients must be visible only to the assigned user and administrators. The user must still authorize and activate that workbook's Gmail automation personally; triggers are not transferred by copying a template.

## 10. Validation Rules

- The enrolling/activating account email must be available and normalize to a valid email-like value.
- One normalized Gmail account may map to only one active rule sheet.
- One rule sheet ID may belong to only one account.
- The executing account must match the registered owner before reading rules for automation or changing account configuration.
- Enabled must be a checkbox/boolean.
- Sender must contain one normalized email address, not a display-name list.
- Match Mode must be one of the supported values.
- `Any keyword` requires at least one nonblank keyword; `All messages` must not depend on keyword contents.
- Recipients must contain at least one address; duplicates should be normalized and removed before sending.
- Recipient count and total address length must remain within Gmail/Apps Script sending constraints.
- Empty rows are ignored; partially completed enabled rows are errors.
- Duplicate or shadowing rules for the same sender should be detected and clearly warned, especially an `All messages` rule above keyword-specific rules.
- Missing/deleted registered sheets, modified headers, missing protections, and schema-version mismatches must stop activation or forwarding with a repair-oriented error.
- Rule updates should be fully validated as a set before they are used by a new activation; malformed live edits should fail closed for affected runs.

## 11. Security Considerations

- Each user must explicitly authorize Gmail access; no account can authorize another user's mailbox through a tab assignment.
- Installable triggers execute as their creator. Activation must show and record the exact Gmail account being enrolled.
- Sheet protection prevents ordinary edits but is not a security boundary against the spreadsheet owner or authorized administrators.
- Tabs in a shared spreadsheet are not confidential. Use separate spreadsheets if users must not see each other's rules or recipient lists.
- Never trust a tab name, active tab, or user-editable owner cell for authorization. Use normalized execution identity plus a protected registry and immutable sheet ID.
- Use least-privilege OAuth scopes in the manifest where practical and review Workspace administrator restrictions before rollout.
- Store processed IDs and operational state per user. Avoid placing message bodies, authentication data, or unnecessary personal data in shared registry cells or logs.
- Sanitize tab names, status text, notes displayed in dialogs, and logged errors to reduce formula injection and accidental data exposure.
- Decide whether external forwarding recipients are allowed. If forwarding is organization-only, enforce an approved domain/address policy in both validation and runtime checks.
- Retain an administrator recovery process and an audit-friendly record of enrollment, activation, pause, and ownership changes.

## 12. Testing Plan

- Happy path: a new user opens the sheet, creates a tab, adds valid rules, validates, previews, activates, receives one matching forward, and receives the daily summary.
- Idempotency: choosing create/open repeatedly returns the same registered tab and does not create duplicates.
- Multi-account isolation: two users enroll, create triggers, and confirm that each trigger reads only its own tab and searches only its own Gmail mailbox.
- State isolation: processed IDs, start times, summaries, cleanup, pause, and reset for one account do not affect another.
- Permission tests: the account owner can edit rule-entry cells; another ordinary editor cannot; administrators can repair; protected status cells cannot be overwritten by the owner.
- Identity failure: unavailable session email prevents enrollment/activation and produces actionable guidance.
- Rule validation: invalid sender, missing keyword, invalid match mode, empty recipients, duplicates, excessive recipients, malformed headers, and partially completed enabled rows.
- Rule ordering: same-sender rules and an All Messages shadowing rule produce a warning or deterministic documented behavior.
- Trigger tests: repeated activation does not duplicate current-user triggers; pause removes only current-user triggers; repair restores missing triggers.
- Migration tests: existing six-column rows retain sender, keywords, recipients, enabled state, and notes after conversion, without triggering forwarding.
- Failure/retry: forwarding failure adds Failed state and retries later without marking processed; successful retry records only once.
- Summary tests: zero activity, multiple records, send failure, malformed summary record, and records isolated by account.
- Sheet lifecycle: owner renames the tab, another user copies it, the tab is deleted, protection is changed, and the registry contains a stale sheet ID.
- Regression: spam inclusion, attachments, old-message cutoff, daily cleanup, previews, labels, and stop/reset behavior continue to work.
- Load/concurrency: simultaneous enrollment and simultaneous monitor runs do not create duplicate registry entries, tabs, triggers, or forwards.

## 13. Rollback Plan

- Before rollout, copy the live rule tabs and export the Apps Script project/version.
- Pilot with one noncritical Gmail account before migrating all users.
- Keep existing rules unchanged until the new account registry and generated tab validate successfully.
- If the new workflow fails, pause/remove only the pilot user's new triggers, restore the prior script version, and point that account back to its backed-up fixed-name rule tab.
- Do not clear processed history during rollback; preserving it minimizes duplicate forwarding.
- Restore tab protections and registry data from the backup if sheet generation or migration partially succeeds.
- Re-enable accounts individually after rollback validation rather than recreating triggers for every user at once.

## 14. Final Checklist

- [ ] Plan reviewed
- [ ] Files identified
- [ ] Database changes checked
- [ ] Backend changes checked
- [ ] Frontend changes checked
- [ ] Validation rules checked
- [ ] Security considerations checked
- [ ] Tests planned
- [ ] Rollback plan reviewed
- [ ] Assumptions and open questions resolved
