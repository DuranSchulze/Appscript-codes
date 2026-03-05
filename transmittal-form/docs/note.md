# Smart Transmittal Feature Roadmap

## Purpose

This document is the working roadmap for upcoming Smart Transmittal improvements. It replaces the previous scratch-note backlog with a more structured planning spec so future work can be implemented with clearer intent, scope, and validation criteria.

The goal of this file is to:

- translate feature ideas into implementation-ready planning checklists
- capture user problems and expected outcomes
- define scope boundaries before coding starts
- identify dependencies, risks, and validation scenarios early

This is a planning document only. It does not mean every item is immediately ready to build, and it should not be treated as proof that a feature already exists.

## Implementation Priorities

Recommended implementation order:

1. User AI API key support
2. In-app onboarding guided tour
3. Column width / position control improvements
4. Project Number UX clarity
5. Project + Sign-Off suggestion dropdowns (type + select)

This order prioritizes system reliability, user self-service, and reduced support burden before lower-risk UI refinements.

## Feature 1: User-Saved AI API Key Support ✅ FINISHED

### Feature Objective

Reduce shared-key quota exhaustion and give users a more reliable AI parsing path by allowing each signed-in user to save and use a personal Gemini API key.

### User Problem

Today, AI parsing depends on environment-level Gemini keys only. When usage is high, the shared key can hit quota limits or become a bottleneck for all users. Users currently have no way to provide their own key to keep working.

### Current State

- `services/geminiService.ts` resolves Gemini keys from environment variables only.
- There is no per-user key storage.
- There is no server preference layer for selecting a user-specific key first.
- There is no UI for saving, replacing, testing, or deleting a personal AI key.

### Scope

Include:

- account-level user setting for Gemini API key
- save, update, and remove personal key
- secure server-side storage
- parser preference order of:
  1. user key
  2. shared system key
  3. existing fallback parsing behavior
- clear user feedback about whether a personal key is active

Do not include in phase 1:

- multiple AI providers
- per-request manual key input
- organization-shared AI key pools

### Detailed TODO Checklist

- [x] Introduce a user settings data model for AI credentials.
- [x] Add a server-side storage strategy for user Gemini keys.
- [x] Store keys encrypted at rest instead of plain text.
- [x] Add an app-level encryption secret, such as `APP_SETTINGS_ENCRYPTION_KEY`.
- [x] Define a dedicated persistence model for this feature.
- [x] Recommended model: `UserAiSettings` keyed by `userId`.
- [x] Include fields at minimum for:
  - [x] `userId`
  - [x] encrypted Gemini API key
  - [x] active/exists indicator (stored or derived)
  - [x] timestamps
- [x] Add authenticated API routes for AI settings:
  - [x] `GET /api/user-ai-settings`
  - [x] `PUT /api/user-ai-settings`
  - [x] `DELETE /api/user-ai-settings`
- [x] Define route behavior:
  - [x] `GET` returns metadata only and never the raw key
  - [x] `PUT` accepts a new key, validates it, and stores the encrypted value
  - [x] `DELETE` removes the stored key and restores shared-key fallback behavior
- [x] Add a user settings entry point in the UI.
- [x] Recommended location: floating account menu or a dedicated settings modal.
- [x] Add a compact settings form that lets the user:
  - [x] paste a Gemini API key
  - [x] save it
  - [x] replace it
  - [x] remove it
  - [x] see whether a personal key is currently active
- [x] Add a `Test key` flow or validate the key during save using a lightweight server-side Gemini call.
- [x] Refactor parser key resolution so the main parsing path can prefer the authenticated user’s saved key.
- [x] Ensure all primary AI parsing flows use the server-aware key resolution path.
- [x] Keep the existing fallback document-number logic unchanged as the final safety net.
- [x] Add user-facing messaging explaining:
  - [x] when the personal key is active
  - [x] when the shared system key is used instead
  - [x] when AI still falls back due to missing or failed keys
- [x] Add post-login prompt for users without a personal key:
  - [x] `Yes` opens AI Key Settings modal
  - [x] `No` dismisses prompt and keeps system-key fallback messaging
  - [x] prompt is not shown again after user declines

Implemented with additive-only schema and API changes; no existing transmittal, agency, or auth data paths were removed.

### Acceptance Criteria

- A signed-in user can save a personal Gemini API key.
- The raw stored key is never returned to the client after save.
- AI parsing prefers the user key when it exists.
- Deleting the personal key returns behavior to shared-key fallback.
- If both user and shared keys fail or are missing, existing fallback parsing still works.

### Risks / Dependencies

- Requires Prisma schema changes and a migration.
- Requires an encryption key management approach on the server.
- Validation must be strict enough to catch obvious bad keys without blocking valid ones unnecessarily.
- Logging and API responses must not leak secrets.

## Feature 2: In-App Onboarding Guided Tour ✅ FINISHED

### Feature Objective

Help first-time and occasional users understand the full workflow, from importing documents to saving and exporting transmittals, without needing external training material first.

### User Problem

The current interface is feature-rich, but there is no structured onboarding. New users must infer the workflow from the UI, which increases confusion, errors, and support needs.

### Current State

- There is no first-run onboarding flow.
- There is no guided tour.
- There is no built-in “Help / Tour” relaunch path.

### Scope

Include:

- first-run guided walkthrough
- manual relaunch option
- contextual explanation of the key workflow
- skip, dismiss, and finish behavior

Do not include in phase 1:

- separate help center page
- video tutorials
- role-based onboarding variants

### Detailed TODO Checklist

- [x] Define onboarding trigger rules.
- [x] Auto-open on the first successful login per browser.
- [x] Do not auto-open again after completion or dismissal.
- [x] Persist onboarding completion or dismissal in `localStorage`.
- [x] Recommended key: `transmittal_onboarding_state_v1`.
- [x] Add a manual `Help / Tour` relaunch control.
- [x] Recommended location: `SidebarHeader` or the floating account menu.
- [x] Create a multi-step guided tour covering:
  - [x] sign-in / account context
  - [x] intelligent import input
  - [x] upload files
  - [x] browse Drive
  - [x] transmission settings
  - [x] sender tab
  - [x] recipient tab
  - [x] project tab
  - [x] signatories tab
  - [x] live preview
  - [x] save transmittal
  - [x] export actions
  - [x] reopen saved transmittals
- [x] For each step, define:
  - [x] target UI element
  - [x] short explanation
  - [x] expected user outcome
- [x] Add `Next`, `Back`, `Skip`, and `Finish` controls.
- [x] Make the tour resilient when a target element is temporarily hidden or unavailable.
- [x] Add a final summary step explaining:
  - [x] transmittals are saved to the database
  - [x] Drive access depends on linked Google sign-in
  - [x] exports support local download and Drive upload
- [x] Ensure the tour works in both desktop and mobile layouts.
- [x] Ensure the tour is dismissible and never traps the user in a blocked state.

### Acceptance Criteria

- A first-time user sees the guided tour automatically after entering the app.
- A returning user does not see it again unless they manually relaunch it.
- The tour clearly explains the import, edit, save, and export flow.
- Users can skip the tour without breaking app behavior.
- The tour works across sidebar and preview-based layouts.

### Risks / Dependencies

- Responsive UI and hidden elements may make target-highlighting brittle.
- Mobile layouts may need fallback positioning or simplified steps.
- This should be implemented with minimal dependency overhead unless a library clearly improves reliability.

## Feature 3: Column Width / Position Control Improvements

### Feature Objective

Reduce friction when users work with wide or uneven document data by giving them stronger control over how the item table columns are sized and visually balanced.

### User Problem

Users often work with varying document titles, long reference numbers, and inconsistent remarks. The current table supports some resizing, but the experience is incomplete and resets after refresh.

### Current State

The preview template already supports resizing these columns:

- `No. of Items`
- `QTY`
- `Document # / Ref #`
- `Description`
- `Remarks`

Current limitations:

- Width settings are intentionally stored only in component state (session-only).
- User adjustments intentionally reset after refresh.

### Scope

This is a column layout refinement, not column reordering.

Include:

- resizable control for all visible item-table columns, including `Description`
- clearer resize affordances
- Docs-style divider behavior where resizing one column adjusts the immediate right column
- fixed total table width so resizing never overflows the form
- reset-to-default action

Do not include:

- persisting column layout between sessions
- changing the left-to-right order of columns
- reordering larger page sections
- per-transmittal server-side saved layouts
- custom width sync to PDF/DOCX exports

### Detailed TODO Checklist

- [x] Define a normalized column layout state object covering all item-table columns.
- [x] Expand the current width model to include the `Description` column.
- [x] Replace the partial width-only state with a full layout config used consistently by the preview table.
- [x] Add a resize handle to the `Description` header.
- [x] Improve the resize affordance styling so users can clearly see that headers are adjustable.
- [x] Implement Docs-style divider adjustment (left column changes, immediate right column compensates).
- [x] Keep total column width fixed so table stays inside form bounds.
- [x] Stop drag at min/max boundaries to prevent overflow.
- [x] Add a `Reset column layout` action in the preview toolbar or a nearby control surface.
- [x] Clamp widths to safe minimum and maximum bounds to protect print readability.
- [x] Ensure the chosen widths affect:
  - [x] live preview
  - [x] fixed-width behavior inside form (no overflow)
  - [x] visible editing layout
- [x] Keep custom width behavior session-only (reset on refresh, no `localStorage` persistence).
- [x] Keep PDF output using default stable widths.
- [x] Keep DOCX generation unchanged.

Shipped in this phase as a preview-only, temporary viewing enhancement. Users can widen columns while reviewing data in the live form, reset widths during the session, and the PDF path continues to use the default stable layout. Divider behavior now follows Docs-style neighbor adjustment: expanding one column shrinks the immediate right column, total table width stays fixed, and drag stops at min/max boundaries to prevent overflow outside the form.

### Acceptance Criteria

- Users can resize all item-table columns, including `Description`.
- Resizing behaves like Docs: the immediate right column compensates automatically.
- Table width remains inside the form while dragging (no overflow beyond form bounds).
- A reset action restores the default layout.
- Width changes are temporary and reset on browser refresh.
- The table remains readable in normal desktop usage.
- PDF output keeps default stable widths.
- No normal interaction should result in overlapping columns or unreadable clipping.

### Risks / Dependencies

- Overly narrow widths can make the table unreadable if clamping is too loose.
- Session-only behavior means settings do not follow users across refreshes/devices.
- DOCX output remains layout-independent in this phase.

## Feature 4: Project Number UX Clarity Improvement

### Feature Objective

Make it obvious to users that the Project Number field is editable and clearly separate it from the transmittal number.

### User Problem

The capability already exists, but users may not immediately understand the distinction between project-level identifiers and the transmittal identifier. This creates confusion during data entry.

### Current State

- `project.projectNumber` already exists in the state model.
- It is editable today.
- It is saved and loaded correctly.
- The confusion is primarily UX and labeling-related.

The current form also contains an editable `Transmittal ID`, which may contribute to confusion between:

- project number
- transmittal number
- project title
- engagement reference

### Scope

This is a UX clarity improvement, not a data-model change.

Include:

- clearer labels
- better helper text
- stronger distinction between identifiers
- better field visibility and comprehension

Do not include:

- new schema fields
- changes to uniqueness rules
- changes to transmittal number generation

### Detailed TODO Checklist

- [x] Review and clarify labels used in the Project tab for:
  - [x] project title
  - [x] project number
  - [x] engagement reference
  - [x] transmittal number
- [x] Add or refine helper text so users understand:
  - [x] `Project Number` is editable and project-specific
  - [x] `Transmittal ID` is auto-generated but still editable when needed
- [x] Ensure `Project Number` is visually clear and not treated as a secondary field.
- [x] Review the preview label `Project No:` and align wording if needed with the editing UI.
- [x] Review import merge behavior and document it clearly in the future implementation spec.
- [x] Lock this phase as UX-only:
  - [x] no overwrite-behavior changes
  - [x] no backend changes
- [x] Add subtle inline guidance to explain the difference between:
  - [x] project number
  - [x] transmittal number
  - [x] engagement reference

Implemented in this phase by adding a dedicated `Project Number` input in the Project tab and updating helper copy so users can clearly distinguish `Project Number` from `Transmittal ID`.

### Acceptance Criteria

- A user can immediately identify where and how to edit the project number.
- The distinction between project number and transmittal number is clearer than today.
- No persistence, API, or schema behavior changes are required.
- Save/load behavior remains unchanged.

### Risks / Dependencies

- Too much helper text could clutter the sidebar.
- This should remain a lightweight UX improvement, not grow into a form redesign.

## Feature 5: Project + Sign-Off Suggestion Dropdowns

### Feature Objective

Speed up repetitive entry in Project and Sign-Off fields by allowing users to either type freely or select from their own previously used values.

### User Problem

Users repeatedly type the same project titles and signatory names/roles. This is time-consuming and error-prone when values are long or frequently reused.

### Current State

- Project and Sign-Off fields supported free typing only.
- There was no history-based suggestion source.
- Users could not quickly reuse values from previous transmittals.

### Scope

Include:

- suggestion dropdowns for:
  - `Project Title` (`project.projectName`)
  - `Prepared By Name` (`signatories.preparedBy`)
  - `Prepared By Role` (`signatories.preparedByRole`)
  - `Noted By Name` (`signatories.notedBy`)
  - `Noted By Role` (`signatories.notedByRole`)
- user-scoped suggestions (current signed-in user only)
- free typing remains fully allowed
- shared API endpoint and reusable combobox UI pattern

Do not include:

- cross-user/global suggestions
- enforced dropdown-only selection
- schema/model changes
- save payload changes

### Detailed TODO Checklist

- [x] Add authenticated `GET /api/transmittal-suggestions`.
- [x] Query transmittal history with user scope only (`where: { userId }`).
- [x] Return only the 5 suggestion lists needed by Project/Sign-Off inputs.
- [x] Deduplicate values case-insensitively while preserving most-recent casing.
- [x] Trim values and exclude empty strings from response.
- [x] Add one-time suggestion loading in `App.tsx` after session is ready.
- [x] Keep graceful degradation: if fetch fails, tabs still work as normal inputs.
- [x] Wire `Project Title` to combobox type+select behavior.
- [x] Wire 4 Sign-Off fields to combobox type+select behavior.
- [x] Keep existing update callbacks and save flow unchanged.
- [x] Keep `Project Number`, `Transmittal ID`, `Purpose`, `Department`, and time fields unchanged.
- [x] Keep feature additive with no schema changes and no destructive operations.

Implemented in this phase with a new authenticated read-only suggestions API and combobox inputs in Project/Sign-Off tabs. Free typing remains the default behavior, selection is optional, and suggestions are scoped to each user’s own transmittal history only.

### Acceptance Criteria

- Users can type freely in all 5 targeted fields without selecting suggestions.
- Matching suggestions are available for quick selection in those fields.
- Suggestions are scoped to the signed-in user only.
- Save/load behavior and payload shape remain unchanged.
- If suggestions fail to load, inputs remain usable with plain typing.

### Risks / Dependencies

- Suggestion quality depends on historical data quality.
- Large history sets are bounded server-side to protect UI responsiveness.
- Privacy relies on strict user-scoped query filters on the API route.

## Closed Items

### Feature 6: Project & Sign-Off Extra Rows (Closed / Superseded)

### Feature Objective

Preserve the idea for future expansion without forcing a premature implementation before the real user workflow is better defined.

### User Problem

Some users may eventually need more flexible project metadata rows or additional sign-off fields, but the exact requirement is still unclear.

### Current State

- The project metadata block uses fixed rows.
- The sign-off and received-by areas use fixed fields.
- There is currently no custom-row or repeatable-row model for these sections.

### Scope

This item is closed. The requested UX outcome for Project and Sign-Off tabs is satisfied by Feature 5 (type + select suggestion dropdowns sourced from user history), so no additional extra-row implementation is planned in this roadmap.

### Detailed TODO Checklist

- [x] Confirm user intent for Project and Sign-Off tab improvement.
- [x] Deliver suggestion dropdowns for Project/Sign-Off fields via Feature 5.
- [x] Keep free typing behavior (no forced selection).
- [x] Keep implementation additive with no schema changes.
- [x] Mark extra-rows concept as superseded for current roadmap scope.

### Acceptance Criteria

- The roadmap clearly states Feature 6 is closed/superseded.
- Engineers should use Feature 5 as the implemented Project/Sign-Off UX enhancement.
- No additional extra-row work is implied in this roadmap version.

### Risks / Dependencies

- If a future requirement explicitly asks for repeatable extra rows, a new separate spec should be created rather than reopening this closed item implicitly.

## Important Changes Or Additions To Public APIs / Interfaces / Types

This roadmap introduces expected future interface changes for planning purposes.

### New / Changed Client State

- Column layout state should evolve from the current partial `columnWidths` object into a fuller layout config that includes `description`.
- Onboarding will need local UI state for:
  - current step
  - completion / dismissal status
- `ProjectInfo` does not require schema changes for the Project Number UX item because the field already exists.
- AI key support should avoid storing raw keys in client state after save.

Recommended client state for AI key UI:

- whether a custom key exists
- loading status
- saving status
- deleting status
- validation result

### New Local Storage Keys

Recommended local keys:

- `transmittal_column_layout_v1`
- `transmittal_onboarding_state_v1`

Existing linked-sheet storage can remain unchanged.

### New Server APIs For AI Key Support

Planned authenticated routes:

- `GET /api/user-ai-settings`
- `PUT /api/user-ai-settings`
- `DELETE /api/user-ai-settings`

Expected high-level behavior:

- `GET`
  - returns whether a custom Gemini key exists
  - returns metadata only
- `PUT`
  - accepts a raw user-supplied key
  - stores the encrypted version
  - returns success and active-status metadata
- `DELETE`
  - removes the stored key
  - returns success and fallback-to-system-key metadata

### New Persistence Model For AI Key Support

Recommended model:

- `UserAiSettings`
  - `id`
  - `userId` (unique)
  - `geminiApiKeyEncrypted`
  - `createdAt`
  - `updatedAt`

This is preferred over placing a plaintext API key field directly on `User`.

## Test Cases And Scenarios

The following scenarios should be considered mandatory when each feature is later implemented.

### Column Layout

- [x] User resizes each column, including `Description`.
- [x] Divider drag adjusts immediate neighbor inversely.
- [x] Table remains within form bounds (no overflow).
- [x] Reset restores defaults.
- [x] Refresh resets layout to defaults (session-only behavior).
- [x] PDF generation remains readable using default widths.

### Onboarding

- [x] First login shows the guided tour.
- [x] User skips the tour and it stays dismissed.
- [x] User finishes the tour and it stays completed.
- [x] User can manually reopen the tour later.
- [x] The tour does not break when a target UI element is hidden or unavailable.
- [x] The tour remains usable on smaller screens.

### Project Number UX

- [x] Users can clearly identify the editable project number field.
- [x] Users understand the difference between project number and transmittal number.
- [x] Saving and reopening transmittals behaves exactly as it does today.
- [x] No backend payload or schema changes are required.

### Project + Sign-Off Suggestions

- [x] Project Title supports type + optional select from suggestions.
- [x] Prepared/Noted name and role fields support type + optional select.
- [x] Suggestions are user-history only (no cross-user exposure).
- [x] Unauthenticated suggestion request is blocked (`401`).
- [x] Free typing still works when suggestion lists are empty or fetch fails.
- [x] Save/load behavior remains unchanged.

### User AI API Key

- [x] User saves a valid Gemini key.
- [x] The system confirms the key is active without returning the raw secret.
- [x] Parsing prefers the user key when present.
- [x] Removing the key restores shared-key behavior.
- [x] Invalid key save attempts return user-friendly errors.
- [x] If the user key fails during parsing, fallback parsing still completes safely.
- [x] One user can never access another user’s stored key or raw status details beyond their own account context.

### Closed Project & Sign-Off Rows Item

- [x] The item is marked closed/superseded by Feature 5.
- [x] The roadmap does not treat it as an active build target.

## Notes And Assumptions

- This document is a roadmap and checklist, not an implementation commit.
- `docs/note.md` is now intended to be a structured planning file instead of a scratch note.
- “Movable columns” is locked to width / position control only, not true column reordering.
- Column layout persistence is local-only in phase 1.
- “Editable Project Number” is already functionally present; the future work is UX clarity only.
- The previous “Project & Sign Off add row” note is closed in this roadmap revision and superseded by the implemented suggestion-dropdown UX in Feature 5.
- AI API key support is implemented as a server-backed per-user setting with encrypted storage and shared-key fallback.
- The current fallback document-number logic remains the baseline safety net and is not being redesigned as part of this roadmap document.
