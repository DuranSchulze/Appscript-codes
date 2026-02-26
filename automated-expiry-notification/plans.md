# Automated Expiry Notification — Product Roadmap Checklist

> **Last Updated**: February 2026  
> **Primary Implementation**: `code.gs`  
> **Product/Operations Documentation**: `README.md`

This file is the delivery roadmap and TODO checklist for upcoming features, with PM-style milestones, acceptance criteria, and execution notes.

---

## 1) Current Baseline Audit (Already Implemented)

Validated against `code.gs`:

- [x] Custom menu with status, initialization, schedule controls, diagnostics
- [x] Configurable working sheet tab via `AUTOMATION_SHEET_NAME`
- [x] Manual + scheduled runs share same processing logic
- [x] Header mapping by row-2 names (with alias support)
- [x] Active + blank status eligibility (blank auto-set to Active)
- [x] Due/overdue sending (`targetDate <= today`)
- [x] Placeholder-aware body template from `Remarks`
- [x] Built-in fallback body when `Remarks` is empty
- [x] Drive attachment parsing from links or file IDs
- [x] Success updates: `Status = Sent` + sender written to `Staff Email`
- [x] Diagnostics: preview dates, inspect row, send test by `No.`
- [x] LOGS sheet creation + summary/error/sent/info logs

---

## 2) Future Feature Backlog (Requested)

### Epic A — Row-Level Send Control Upgrade

**Goal**: Allow row-level control of whether sending is automatic, held, or manual.

#### Proposed model

- New column (recommended): `Send Mode`
- Suggested values:
  - `Auto` (normal scheduled/manual processing)
  - `Hold` (never send until changed)
  - `Manual Only` (exclude from schedule, allow explicit test/manual command)

#### Tasks

- [x] Define final dropdown values and default behavior
- [x] Add column contract to docs (`README.md`)
- [x] Update eligibility logic in `runDailyCheck()`
- [x] Add per-row reason logging when skipped by mode
- [x] Update diagnostics to display send mode and eligibility reason
- [ ] Add regression test checklist for each mode

#### Acceptance criteria

- [x] Schedule run respects send mode consistently
- [x] Manual run behavior is explicit and documented
- [x] Logs clearly explain mode-based skip/send outcomes

---

### Epic B — Reply Keyword Tracking

**Goal**: Detect recipient replies containing configured keywords and reflect that in the sheet.

#### Proposed model

- New columns:
  - `Reply Status` (e.g., Pending, Replied)
  - `Replied At`
  - `Reply Keyword`

#### Tasks

- [x] Define keyword list (e.g., `ACK`, `RECEIVED`, `OK`)
- [x] Store keywords in configurable script/document property
- [x] Save outbound linkage metadata needed for matching replies
- [x] Build periodic reply-scanner function
- [x] Update sheet on keyword match
- [x] Add logs for reply-detected events
- [x] Add menu item to run reply scan manually

#### Acceptance criteria

- [x] Reply scan updates the correct row reliably
- [x] Duplicate matches are handled safely
- [x] False positives minimized through sender/thread checks

---

### Epic C — Email Open/Read Visibility (Best Effort)

**Goal**: Provide visibility into whether reminder emails were opened/looked at.

> Note: true open tracking is client-dependent and privacy-limited.

#### Proposed model

- New columns:
  - `First Opened At`
  - `Last Opened At`
  - `Open Count` (optional)

#### Tasks

- [x] Decide tracking method (tracking pixel and/or tracked links)
- [x] Implement tokenized per-row tracking metadata
- [x] Capture and store open events when available
- [x] Add fallback behavior when tracking is blocked
- [x] Document privacy/accuracy limitations in README

#### Acceptance criteria

- [x] Open data is written when events are captured
- [x] No runtime failures if tracking signals are unavailable

---

### Epic D — Fallback Body Strategy Hardening

**Goal**: Improve resilience and maintainability of fallback email content.

#### Current state

- Implemented: hardcoded fallback in `buildEmailBody()` when `Remarks` is empty.

#### Enhancements

- [x] Decide fallback source:
  - [x] Keep hardcoded
  - [ ] Move to configurable sheet template
  - [x] Move to script property template
- [x] Add placeholder support parity with remarks template
- [x] Add versioned fallback template note in docs
- [x] Add diagnostics to preview resolved fallback body

#### Acceptance criteria

- [x] Empty `Remarks` always produces a valid body
- [x] Fallback content can be maintained without breaking sends

---

### Epic E — AI Integration Menu + Content Generation

**Goal**: Add optional AI-generated email content when `Remarks` is empty.

#### Proposed UX

- New top-level menu group: `AI Integration`
  - `Set API Key`
  - `Select Model`
  - `Test AI Connection`
  - `View AI Status`

#### Runtime behavior (target)

1. If `Remarks` exists -> use template as-is.
2. If `Remarks` is empty and AI configured -> generate body with selected model.
3. If AI is not configured or generation fails -> use fallback body.

#### Tasks

- [x] Add secure API key storage in script properties
- [x] Add Gemini model selection storage
- [x] Add AI settings/status helpers
- [x] Add model list retrieval and validation flow
- [x] Add AI generation path in body builder
- [x] Add timeout/retry/fallback handling
- [x] Add logs for AI-generated vs fallback-used outcomes

#### Acceptance criteria

- [x] AI path is optional and safe-by-default
- [x] No send failures caused by missing/invalid AI settings
- [x] Fallback path always works

---

## 3) Milestone-Based Delivery Plan

### Milestone 1 (Now) — Control + Stability

- [x] Epic A (Send Mode)
- [x] Epic D enhancements (fallback configurability)
- [x] Documentation refresh for new data contract

### Milestone 2 (Next) — Reply Intelligence

- [x] Epic B (Reply keyword tracking)
- [x] Initial event audit logs and monitoring

### Milestone 3 (Later) — Engagement + AI

- [x] Epic C (open/read best effort)
- [x] Epic E (AI integration)

---

## 4) Dependencies, Risks, and Constraints

- [ ] Gmail/Apps Script quotas evaluated for added scanners/triggers
- [ ] Privacy/compliance review for open tracking approach
- [x] Secure handling of API keys (never in sheet cells)
- [x] Backward compatibility with existing sheets
- [x] Duplicate trigger protection retained after new jobs are added

---

## 5) Definition of Done (Roadmap Items)

For each epic, all must be true:

- [x] Functional behavior implemented in `code.gs`
- [x] Menu/UI flow implemented where required
- [x] Logging and diagnostics updated
- [x] `README.md` updated to match behavior
- [x] `plans.md` checklist status updated
- [ ] Manual test scenarios documented and executed

---

## 6) Open Product Decisions

- [x] Send mode final values and semantics confirmed
- [x] Reply keywords finalized
- [x] Open tracking strategy chosen (or deferred)
- [x] AI scope confirmed (Gemini-only vs multi-provider)
- [x] Fallback template source finalized

---

## 7) Backlog Parking Lot (Nice-to-Have)

- [ ] Per-client localization of email templates
- [ ] Business-hour send window constraints
- [ ] Department/owner-based sender naming rules
- [ ] Weekly KPI summary email for operations managers

---
