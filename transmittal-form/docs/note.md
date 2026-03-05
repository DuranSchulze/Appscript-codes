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
5. Project & Sign-Off extra rows (`TO BE DEVELOPED`)

This order prioritizes system reliability, user self-service, and reduced support burden before lower-risk UI refinements.

## Feature 1: User-Saved AI API Key Support

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

## Feature 2: In-App Onboarding Guided Tour

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

- [ ] Define onboarding trigger rules.
- [ ] Auto-open on the first successful login per browser.
- [ ] Do not auto-open again after completion or dismissal.
- [ ] Persist onboarding completion or dismissal in `localStorage`.
- [ ] Recommended key: `transmittal_onboarding_state_v1`.
- [ ] Add a manual `Help / Tour` relaunch control.
- [ ] Recommended location: `SidebarHeader` or the floating account menu.
- [ ] Create a multi-step guided tour covering:
  - [ ] sign-in / account context
  - [ ] intelligent import input
  - [ ] upload files
  - [ ] browse Drive
  - [ ] transmission settings
  - [ ] sender tab
  - [ ] recipient tab
  - [ ] project tab
  - [ ] signatories tab
  - [ ] live preview
  - [ ] save transmittal
  - [ ] export actions
  - [ ] reopen saved transmittals
- [ ] For each step, define:
  - [ ] target UI element
  - [ ] short explanation
  - [ ] expected user outcome
- [ ] Add `Next`, `Back`, `Skip`, and `Finish` controls.
- [ ] Make the tour resilient when a target element is temporarily hidden or unavailable.
- [ ] Add a final summary step explaining:
  - [ ] transmittals are saved to the database
  - [ ] Drive access depends on linked Google sign-in
  - [ ] exports support local download and Drive upload
- [ ] Ensure the tour works in both desktop and mobile layouts.
- [ ] Ensure the tour is dismissible and never traps the user in a blocked state.

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

- [ ] Review and clarify labels used in the Project tab for:
  - [ ] project title
  - [ ] project number
  - [ ] engagement reference
  - [ ] transmittal number
- [ ] Add or refine helper text so users understand:
  - [ ] `Project Number` is editable and project-specific
  - [ ] `Transmittal ID` is auto-generated but still editable when needed
- [ ] Ensure `Project Number` is visually clear and not treated as a secondary field.
- [ ] Review the preview label `Project No:` and align wording if needed with the editing UI.
- [ ] Review import merge behavior and document it clearly in the future implementation spec.
- [ ] Lock this phase as UX-only:
  - [ ] no overwrite-behavior changes
  - [ ] no backend changes
- [ ] Add subtle inline guidance to explain the difference between:
  - [ ] project number
  - [ ] transmittal number
  - [ ] engagement reference

### Acceptance Criteria

- A user can immediately identify where and how to edit the project number.
- The distinction between project number and transmittal number is clearer than today.
- No persistence, API, or schema behavior changes are required.
- Save/load behavior remains unchanged.

### Risks / Dependencies

- Too much helper text could clutter the sidebar.
- This should remain a lightweight UX improvement, not grow into a form redesign.

## Deferred / TO BE DEVELOPED

### Feature 5: Project & Sign-Off Extra Rows

### Feature Objective

Preserve the idea for future expansion without forcing a premature implementation before the real user workflow is better defined.

### User Problem

Some users may eventually need more flexible project metadata rows or additional sign-off fields, but the exact requirement is still unclear.

### Current State

- The project metadata block uses fixed rows.
- The sign-off and received-by areas use fixed fields.
- There is currently no custom-row or repeatable-row model for these sections.

### Scope

This item is intentionally deferred and should not be treated as build-ready.

### Detailed TODO Checklist

- [ ] Mark this feature as `TO BE DEVELOPED`.
- [ ] Document clearly that the final behavior is not yet defined.
- [ ] Capture the open design question:
  - [ ] should this become custom key/value metadata rows
  - [ ] should this become additional fixed built-in fields
  - [ ] should this become repeatable sign-off blocks
- [ ] Defer schema design until real user examples are collected.
- [ ] Revisit only after the onboarding and AI-key features are completed or better defined.

### Acceptance Criteria

- The roadmap preserves the idea for future discussion.
- The roadmap makes it clear this is not implementation-ready.
- No engineer should treat this item as a build task without a separate follow-up spec.

### Risks / Dependencies

- This feature touches printable layout, state shape, persistence, and exports.
- Choosing the wrong abstraction too early would create future migration cost.

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

- [ ] First login shows the guided tour.
- [ ] User skips the tour and it stays dismissed.
- [ ] User finishes the tour and it stays completed.
- [ ] User can manually reopen the tour later.
- [ ] The tour does not break when a target UI element is hidden or unavailable.
- [ ] The tour remains usable on smaller screens.

### Project Number UX

- [ ] Users can clearly identify the editable project number field.
- [ ] Users understand the difference between project number and transmittal number.
- [ ] Saving and reopening transmittals behaves exactly as it does today.
- [ ] No backend payload or schema changes are required.

### User AI API Key

- [ ] User saves a valid Gemini key.
- [ ] The system confirms the key is active without returning the raw secret.
- [ ] Parsing prefers the user key when present.
- [ ] Removing the key restores shared-key behavior.
- [ ] Invalid key save attempts return user-friendly errors.
- [ ] If the user key fails during parsing, fallback parsing still completes safely.
- [ ] One user can never access another user’s stored key or raw status details beyond their own account context.

### Deferred Project & Sign-Off Rows

- [ ] The item remains clearly marked as deferred.
- [ ] The roadmap wording does not make it appear build-ready.

## Notes And Assumptions

- This document is a roadmap and checklist, not an implementation commit.
- `docs/note.md` is now intended to be a structured planning file instead of a scratch note.
- “Movable columns” is locked to width / position control only, not true column reordering.
- Column layout persistence is local-only in phase 1.
- “Editable Project Number” is already functionally present; the future work is UX clarity only.
- “Project & Sign Off add row” remains intentionally deferred and should stay marked `TO BE DEVELOPED` until the workflow is better defined.
- AI API key support is planned as a server-backed per-user setting with encrypted storage and shared-key fallback.
- The current fallback document-number logic remains the baseline safety net and is not being redesigned as part of this roadmap document.
