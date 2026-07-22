# Mini-CRM — Plan & Task Tracker

> Status legend:  
> `[ ]` To Do &nbsp; `[~]` In Progress &nbsp; `[x]` Done &nbsp; `[-]` Cancelled

---

## System Snapshot — v5.7

_Last assessed: 2026-07-22_

### ✅ What's Complete & Working

[x] **Core Gmail sync** — Pulls message-level Inbox emails into monthly sheets every 4h with classification  
[x] **Inbound-only collection** — Skips messages sent by the monitored mailbox or any Gmail send-as alias, even inside matching Inbox threads  
[x] **Mailbox execution lock** — Gmail reads, drafts, sends, web actions, initial setup, and trigger creation are restricted to `info@filepino.com` as the effective Google account  
[x] **Gmail message-ID deduplication** — Checks IDs across every monthly sheet; the sender/date/subject key is used only for pre-ID legacy rows  
[x] **Ingestion regression checks** — Covers owner/alias replies, non-Inbox messages, authoritative Gmail IDs, legacy fallback rows, and in-run identity recording  
[x] **Email classification** — Meeting, Conversion, Payment, Unsubscribe, Trash, Internal keywords  
[x] **AI Auto-Replies (Gemini)** — FAQ matching → immediate reply, complex → draft + Chat approval  
[x] **Gemini fallback model** — Primary model fails → tries fallback before giving up  
[x] **Shared AI request governor** — All Gemini generation routes share per-minute, daily, and purpose-specific limits  
[x] **AI cooldown and backoff** — Rate-limit/auth failures stop fallback spam and defer repeated work with exponential delays  
[x] **AI response caching** — Repeated qualification of identical sanitized content reuses the six-hour cached decision  
[x] **Duplicate-draft protection** — Auto-reply draft checkpoints prevent repeated Gemini calls and duplicate Gmail drafts after notification failures  
[x] **AI usage visibility** — The AI menu reports total, qualification, auto-reply, and cooldown status  
[x] **Master AI pause control** — Users can stop qualification and draft-generation calls from the AI menu without removing their API key  
[x] **Pre-AI bulk-mail filtering** — List-Unsubscribe, Auto-Submitted, and bulk/list/junk headers reject obvious automation before Gemini  
[x] **Google Chat notifications** — Interactive cards with Approve/Reject buttons  
[x] **Secure web app endpoint** — Chat approve/reject links are signed, action-bound, expire after 24 hours, and can only be used once  
[x] **Follow-up reminders** — Daily check for stale leads, creates drafts + Chat notifications  
[x] **Round-robin team assignment** — Cycles through team members via script properties  
[x] **Archive system** — Moves sheets from prior years using the current year dynamically and safely trims old engagement rows  
[x] **Dashboard** — Auto-refreshes daily with recent-month metrics  
[x] **Initial setup wizard** — One-click full workbook creation, headers, formatting, trigger config  
[x] **Conversion tracking** — Keyword-based detection of signed retainers, payments, etc.  
[x] **AI Sales Qualification** — New `Qualification.gs` qualifying leads via AI into Potential Clients queue  
[x] **Potential Clients Sheet** — Review queue with 20-column schema (Candidate ID → Engagement Row)  
[x] **On-edit automation** (`onEditInstallable`) — Reacts to status, payment, department, date changes  
[x] **Gemini Model Selector UI** — Beautiful HTML dialog (`GeminiModelSelector.html`) for model/key config  
[x] **Quote generation** — Creates docs from Google Docs templates, replaces `{{placeholders}}`  
[x] **Historical sync** — Backfill emails from Jan 2024+ or choose an inclusive custom date range from the menu  
[x] **Service → Template mapping** — Map Sheet stores Google Docs template IDs per service  
[x] **Quick start guide** — Shown after initial setup  
[x] **Category routing emails** — Routes replies by category (Visa → visa@duranschulze.com, etc.)  
[x] **Financial formulas** — Engagement sheet formulas for fees, VAT, balances  
[x] **Dropdown validations** — Data validation set up on key columns  
[x] **Idempotent setup** — Rerunning setup preserves existing data  

---

## Current Sprint

### 🔴 Bugs to Fix

[x] **Fix `sendChatApprovalCard` category bug** — Chat routing now receives the generated category instead of the draft ID.
[x] **Secure the web app endpoint** — Approve/reject links now require an HMAC-SHA256 signature tied to the action, draft, nonce, and expiry; GET only confirms, while the POST action is single-use.
[x] **Prevent data loss on Gemini failure** — Failed work remains unread/labeled, enters backoff, and retries later without creating duplicate drafts.
[x] **Hardcoded year in `Archive.gs`** — The archive cutoff now follows the current year in the configured CRM timezone.
[x] **Archive chunk loop can stall forever** — Chunk traversal now advances even when no rows in the current chunk are deleted.
[x] **Hardcoded date range in `Backfill.gs`** — The core accepts `startDate` and `endDate`, with a menu flow for validated `YYYY-MM-DD` input, paginated Gmail searches, and a thread safety limit.

### 🟡 Quality Improvements

[ ] **Add structured logging** — Log to a dedicated sheet with timestamps, severity, module. Currently relies on `Logger.log` / `console.log` which are hard to review in production.
[ ] **Unify Map Sheet caching** — `MAP_CACHE` exists in `Code.gs` but `AutoReply.gs`, `ChatNotifications.gs`, `FollowUpReminders.gs`, and `Qualification.gs` all re-read the Map Sheet independently. Use the cache everywhere.
[ ] **Follow-up dedup** — `processFollowUpReminders` creates a new Gmail draft every run with no check if one already exists for that contact. Could create dozens of duplicate drafts.
[ ] **Use semantic FAQ matching** — `checkFAQ` uses `string.includes()` which is fragile ("visa" matches "advisable"). Use Gemini embeddings or token-based matching.
[x] **Add API rate-limiting** — Shared local quotas, bounded batches, five-minute auto-reply polling, cooldowns, and deferred retry protect the Gemini project.
[ ] **Thread context for Gemini** — Only the last message is sent to Gemini; earlier thread context is lost. Send a thread summary.

### 🟢 Enhancements

[ ] **Add dry-run mode** — Preview what AI would generate without creating actual drafts or modifying state.
[ ] **Weekly summary email** — Auto-generate and email a weekly digest: new leads, conversions, response times.
[ ] **Response-time tracking** — Measure time from email received → reply sent.
[ ] **Attachment awareness** — Note or extract attachments (PDFs, DOCs) from tracked emails.
[ ] **Config validation on startup** — Validate API key, webhook URLs, required sheets exist. Fail early instead of silently.
[x] **Parameterize `Backfill.gs`** — Accepts inclusive `startDate` and `endDate` inputs and exposes a custom-range menu action.

---

## Backlog

_Add new tasks here before prioritizing them into a sprint._

[ ] **Incremental sync cursor** — Replace the fixed 30-day retrieval query with a last-successful-sync cursor and safe overlap window; keep Gmail message ID as the idempotency key.

---

## Icebox

_Ideas and nice-to-haves for the future._

[ ] Add unit tests for classification & FAQ matching logic
[ ] Integrate with Clerk for authentication on the web app
[ ] Auto-generate engagement letters from templates on conversion
[ ] SMS/WhatsApp notifications as alternative to Google Chat
[ ] Client portal via Google Sites embedding the web app
[ ] Multi-language auto-reply support (English + Filipino)

---

## Legend

| Mark | Meaning |
|---|---|
| `[ ]` | Not started |
| `[~]` | In progress |
| `[x]` | Done & verified |
| `[-]` | Cancelled / Won't do |
