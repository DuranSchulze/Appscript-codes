# Automated Expiry Notification - System Audit & Findings Report

**Document Owner:** @zafajardo  
**Audit Date:** April 28, 2026  
**System Version:** April 27, 2026  
**Classification:** Internal Technical Review  

---

## Executive Summary

This report presents a comprehensive technical audit of the Google Apps Script automation for expiry notification emails. The system automates sending reminder emails to clients based on document expiry dates and configurable notice periods. This audit evaluates the system's logic correctness, email delivery mechanisms, status tracking, template substitution capabilities, and error handling robustness.

**Overall Assessment:** The system is **production-ready with caveats**. Core logic is sound, error handling is comprehensive (33 try-catch blocks), but three critical issues require attention before full deployment.

---

## 1. Email Sender Logic Analysis

### 1.1 Current Implementation

The system implements a **per-row sender delegation** model where emails are sent from the "Assigned Staff Email" column rather than the automation runner's account.

**Key Code References:**
- `Send.gs:821-837` - `sendReminderEmail()` function applies `fromEmail` parameter
- `Orchestration.gs:637` - Retrieves staff email from row data
- `Orchestration.gs:618-632` - Alias verification gate

### 1.2 Gmail "Send As" Alias Requirement (CRITICAL)

**Finding:** The system enforces strict sender verification through `canSendAs(staffEmail)` check (`Send.gs:80-85`). If the assigned staff email is not registered as a "Send mail as" alias in the Gmail account running the automation, **the entire row is skipped with an error status**.

**Code Behavior:**
```javascript
// Orchestration.gs:618-632
if (!canSendAs(staffEmail)) {
  setResolvedStatus(visaSheet, rowIndex, colMap, tabName, STATUS.ERROR);
  appendLog(logsSheet, tabName, clientName, "ERROR",
    'Assigned Staff Email "' + staffEmail +
    '" is not a verified Gmail "Send mail as" alias...'
  );
  errors++;
  continue;  // ROW IS SKIPPED ENTIRELY
}
```

**Impact Assessment:**
- **High Risk:** If staff have not configured their emails as Gmail aliases, all their assigned rows will fail
- **Silent Failure Pattern:** Errors are logged to the LOGS sheet but no real-time notification is sent
- **Operational Dependency:** Requires Gmail Settings → Accounts → "Send mail as" configuration for each staff member

**Recommendation:**
1. **Immediate Action:** Verify all staff emails are configured as Gmail aliases before production deployment
2. **Enhancement:** Consider implementing a fallback mechanism to send from the automation runner's email with the staff member in the Reply-To header
3. **Documentation:** Create setup guide for staff alias configuration

### 1.3 Sender Display Name Logic

**Finding:** Display name is derived from the "Name of Staff" column, with domain-based fallback.

**Code Reference:** `Send.gs:809-819`
```javascript
function getSenderDisplayName(email) {
  var raw = String(email || "").trim();
  var atIndex = raw.indexOf("@");
  if (atIndex < 0) return CONFIG.SENDER_NAME;
  var domain = raw.slice(atIndex + 1);
  var domainBase = domain.split(".")[0];
  return domainBase.charAt(0).toUpperCase() + domainBase.slice(1).toLowerCase();
}
```

**Assessment:** Logic is functional but inflexible. Consider allowing custom sender names per staff member via configuration.

---

## 2. Email Content & Template System

### 2.1 Remarks Column as Email Body (CRITICAL)

**Finding:** The "Description" column (aliased as "Remarks") is used as the **primary source for email body content**. If Remarks is empty, a fallback template is used.

**Code Reference:** `Send.gs:521-571`
```javascript
function buildEmailContent(remarks, clientName, expiryDate, docType, openToken, templateContext) {
  var bodyText = "";
  var source = "";

  if (remarks) {
    bodyText = applyTemplatePlaceholders(remarks, clientName, expiryStr, docTypeText, templateContext);
    source = "Remarks";  // Uses Remarks column content!
  } else {
    var fallback = resolveFallbackTemplateText();
    bodyText = applyTemplatePlaceholders(fallback.text, ...);
    source = fallback.source;
  }
  // ... converts to HTML and injects tracking pixel
}
```

**Security Concern:** The Remarks content is inserted **directly into HTML without sanitization**. This creates potential for:
- HTML injection if users include unescaped HTML characters
- Email layout breakage from malformed HTML
- Potential XSS vectors (though mitigated by Gmail's sanitization)

**Recommendation:**
1. Implement HTML escaping for Remarks content before insertion
2. Add validation to warn users if Remarks contains HTML-like content
3. Consider supporting both plain-text and HTML modes

### 2.2 Template Substitution System (ADVANCED FEATURE)

**Finding:** The system supports **dynamic token substitution from ANY column** using the `[Column Name]` syntax.

**Built-in Tokens:**
- `[Client Name]` → Client's name
- `[Expiry Date]` / `[Date of Expiration]` → Formatted expiry date
- `[Document Type]` → Document type label

**Dynamic Tokens:** Any column header can be referenced as `[Header Name]` and will be replaced with that row's cell value.

**Code References:**
- `Send.gs:613-627` - `applyTemplatePlaceholders()` with built-in tokens
- `Send.gs:639-662` - `buildRowTemplateContext()` builds context from all headers
- `Send.gs:665-678` - `replaceGenericTemplateTokens()` handles dynamic substitution

**Example Usage:**
If your sheet has columns: "Client Name", "Staff Email", "Project ID"
Then Remarks can contain: "Dear [Client Name], please contact [Staff Email] regarding project [Project ID]"

**Risk Assessment:**
- **Low-Moderate Risk:** Unknown tokens remain literal (e.g., "[Price]" stays as "[Price]" if no such column exists)
- **User Education Required:** Users must understand this substitution happens
- **Data Exposure Risk:** All column data is potentially exposed in email content if referenced

**Recommendation:**
1. Document available tokens for end users
2. Add validation to warn about unsubstituted tokens
3. Consider implementing a token whitelist for security

### 2.3 AI-Generated Email Content (Optional Feature)

**Finding:** The system can use Google Gemini AI to generate email content when no Remarks are provided.

**Code Reference:** `Send.gs:137-214`
```javascript
function tryGenerateAiEmailBody(clientName, expiryDate, docTypeText) {
  var apiKey = getAiApiKey();
  var model = getAiModel();
  if (!apiKey || !model) return null;
  // ... calls Gemini API ...
}
```

**Assessment:** Properly implemented with fallback to hardcoded template. API failures gracefully degrade to fallback.

---

## 3. Status Tracking & Updates

### 3.1 State Machine Design

**Finding:** The system implements a **well-designed state machine** for email lifecycle management.

**Status Values:**
- `Active` / `(blank)` → Eligible for processing
- `Notice Sent` → First reminder sent, waiting for final reminder date
- `Sent` → Final reminder sent (or both sent same day)
- `Error` → Processing failed, requires manual review
- `Skipped` → Intentionally skipped per configuration

**State Transitions:**
```
Active/Blank → Notice Sent (first reminder sent)
Notice Sent → Sent (final reminder sent)
Active/Blank → Sent (if notice+final sent same day)
Any → Error (on failure)
```

**Code Reference:** `Orchestration.gs:676-700`, `Orchestration.gs:790`

### 3.2 Post-Send Metadata Tracking

**Finding:** When emails are sent, the system records comprehensive metadata.

**Columns Updated:**
- `Sent At` - Timestamp of email send
- `Sent Thread Id` - Gmail thread identifier
- `Sent Message Id` - Gmail message identifier
- `Open Tracking Token` - UUID for open tracking
- `Reply Status` - Set to "Pending"

**Code Reference:** `Foundation.gs:1059-1080` - `writePostSendMetadata()`

**Assessment:** Excellent audit trail. Thread/Message IDs enable reply tracking.

### 3.3 Final Notice Tracking

**Finding:** Separate tracking exists for final/expiry-day reminders.

**Columns:**
- `Final Notice Sent At`
- `Final Notice Thread Id`
- `Final Notice Message Id`

**Code Reference:** `Orchestration.gs:766-770`, `Foundation.gs` (similar pattern)

---

## 4. Email Open Tracking

### 4.1 Tracking Mechanism

**Finding:** Uses **1x1 transparent pixel webhook** technique for open tracking.

**Implementation:** `Tracking.gs:42-56`
```javascript
function injectOpenTrackingPixel(htmlBody, openToken) {
  var baseUrl = getOpenTrackingBaseUrl();
  if (!baseUrl || !openToken) return htmlBody;
  var separator = baseUrl.indexOf("?") >= 0 ? "&" : "?";
  var openUrl = baseUrl + separator + "mode=open&t=" + encodeURIComponent(openToken);
  return htmlBody + '<img src="' + openUrl + '" width="1" height="1" style="display:none;" />';
}
```

### 4.2 Web App Deployment Requirement (CRITICAL)

**Finding:** Open tracking **requires the Apps Script to be deployed as a web app** with a publicly accessible URL.

**Code Reference:** `Tracking.gs:13-23`
```javascript
function getOpenTrackingBaseUrl() {
  var configured = getPropString(PROP_KEYS.OPEN_TRACKING_BASE_URL, "");
  if (configured) return configured;
  try {
    var serviceUrl = ScriptApp.getService().getUrl();
    if (serviceUrl) return serviceUrl;
  } catch (e) {}
  return "";
}
```

**Impact:** If not deployed as web app:
- `ScriptApp.getService().getUrl()` returns null/throws
- `getOpenTrackingBaseUrl()` returns empty string
- `injectOpenTrackingPixel()` returns HTML unchanged (no pixel injected)
- **Open tracking silently fails** - no errors, just no tracking

**Columns Affected When Tracking Enabled:**
- `First Opened At` - First email open timestamp
- `Last Opened At` - Most recent open timestamp
- `Open Count` - Cumulative open count

**Recommendation:**
1. Deploy as web app if tracking is needed
2. Add diagnostic check to warn if tracking expected but URL not configured
3. Document deployment requirements

### 4.3 Open Event Processing

**Finding:** When pixel fires, `doGet()` handler updates sheet and sets Reply Status.

**Code Reference:** `Tracking.gs:58-166`
- Updates `First Opened At`, `Last Opened At`, `Open Count`
- Sets `Reply Status` to "Replied" (note: this is an open, not a reply - naming is confusing)
- Sets `Replied At` timestamp
- Sets `Reply Keyword` to "OPEN_TRACKED"

**Note:** The naming is misleading - an "open" event sets status to "Replied".

---

## 5. Reply Tracking

### 5.1 Reply Scan Mechanism

**Finding:** Separate system for detecting actual email replies via Gmail API.

**Implementation:** `Tracking.gs:233-420`
- Scans Gmail threads for replies after send timestamp
- Looks for keywords in subject/body (default: "ACK", "RECEIVED", "OK")
- Updates `Reply Status`, `Replied At`, `Reply Keyword` columns

**Schedule:** Runs twice daily (9 AM, 3 PM) via triggers

**Code Reference:** `Orchestration.gs:253-281` - `installReplyScanTrigger()`

### 5.2 Reply Detection Logic

**Finding:** Sophisticated matching to distinguish outgoing from incoming messages.

**Code Reference:** `Tracking.gs:422-478`
- Filters out messages from verified sender aliases
- Matches sender email to client email
- Searches subject + body for configured keywords
- Only processes messages after the send timestamp

---

## 6. Error Handling Assessment

### 6.1 Error Handling Coverage

**Finding:** **Comprehensive error handling** with 33 try-catch blocks across 6 files.

| File | Try-Catch Count | Primary Coverage |
|------|-----------------|------------------|
| `Send.gs` | 11 | AI generation, attachment resolution, Gmail API |
| `Diagnostics.gs` | 7 | Test functions, connection validation |
| `Orchestration.gs` | 6 | Main processing loop, row-level isolation |
| `Foundation.gs` | 4 | Spreadsheet access, properties |
| `Tracking.gs` | 4 | Webhook handling, reply scanning |
| `Setup.gs` | 1 | Configuration functions |

### 6.2 Row-Level Isolation (EXCELLENT)

**Finding:** Each row is processed in its own try-catch block, ensuring one row's failure doesn't affect others.

**Code Reference:** `Orchestration.gs:634-837`
```javascript
for (var i = 0; i < data.length; i++) {
  var row = data[i];
  var rowIndex = dataStartRow + i;
  // ... extract row data ...
  
  try {
    // ... send email logic ...
  } catch (e) {
    setResolvedStatus(visaSheet, rowIndex, colMap, tabName, STATUS.ERROR);
    appendLog(logsSheet, tabName, clientName, "ERROR", "Send failed: " + e.message);
    errors++;
    // Continue to next row!
  }
}
```

**Assessment:** This is a **best practice implementation** - failures are isolated and logged without crashing the entire run.

### 6.3 Silent Failures (ATTENTION NEEDED)

**Finding:** Some failures fail silently without user notification.

**Examples:**
- Open tracking fails silently if web app not deployed
- Attachment warnings are logged but don't stop sending
- AI generation failure falls back to template without warning

**Recommendation:** Consider adding a summary notification for runs with warnings (not just errors).

### 6.4 Validation Chain

**Finding:** Multi-layer validation before sending.

**Validation Steps:**
1. **Required fields check** - Client Name, Client Email, Expiry Date, Notice Date, Staff Name, Staff Email
2. **Date parsing validation** - Expiry date must be valid date
3. **Notice offset parsing** - Must match supported formats
4. **Email format validation** - Regex check on client emails
5. **Alias verification** - Staff email must be Gmail alias
6. **Attachment resolution** - Drive files must be accessible

---

## 7. Date & Time Logic

### 7.1 Date Calculation Accuracy

**Finding:** **Correct date arithmetic** using proper midnight normalization.

**Supported Notice Formats:**
- `"7 days before"`, `"2 weeks before"`
- `"1 month before"`, `"1 year before"`
- `"On expiry date"` (0 days)
- Numeric: `"7"` (treated as days)
- Absolute: `"2026-12-31"` (explicit date)

**Code Reference:** `Foundation.gs:519-569` - `parseNoticeOffset()`

### 7.2 Timezone Handling

**Finding:** System uses `Asia/Manila` timezone explicitly.

**Code Reference:** `Orchestration.gs:236`
```javascript
.inTimezone("Asia/Manila")
```

**Assessment:** Hardcoded timezone may need configuration for other regions.

---

## 8. Attachment Handling

### 8.1 Google Drive Integration

**Finding:** Supports attaching files from Google Drive via file ID or URL.

**Code Reference:** `Send.gs:383-464`
- Parses Drive URLs to extract file IDs
- Fetches files via `DriveApp.getFileById()`
- Converts to blobs for email attachment

### 8.2 Fallback for Failed Attachments

**Finding:** If a file can't be accessed, system adds a fallback link section to the email.

**Code Reference:** `Send.gs:467-493` - `buildFallbackLinksHtml()`

**Assessment:** Good user experience - recipient still gets access to files even if attachment fails.

---

## 9. Recommendations & Action Items

### 9.1 Critical Issues (Must Fix Before Production)

| Priority | Issue | Action |
|----------|-------|--------|
| **P0** | Gmail alias requirement causes row skipping | Verify all staff have configured aliases; document setup process |
| **P0** | HTML injection risk from Remarks column | Sanitize Remarks content before HTML insertion |
| **P1** | Open tracking silently fails without web app deployment | Add diagnostic warning; document deployment steps |

### 9.2 Enhancements (Recommended)

| Priority | Enhancement | Rationale |
|----------|-------------|-----------|
| **P2** | Add sender fallback mechanism | If alias check fails, send from runner's email with staff in Reply-To |
| **P2** | Token substitution validation | Warn users about unsubstituted `[tokens]` in sent emails |
| **P3** | Configurable timezone | Currently hardcoded to Asia/Manila |
| **P3** | Summary email on completion | Notify automation owner of run results |

### 9.3 Documentation Needs

1. **Setup Guide:** How to configure Gmail "Send mail as" aliases
2. **Template Guide:** Document available `[tokens]` for email content
3. **Deployment Guide:** Web app deployment for open tracking
4. **Troubleshooting Guide:** Common errors and resolutions

---

## 10. Testing Recommendations

### 10.1 Pre-Deployment Testing

1. **Test Gmail Send** (via Diagnostics menu)
2. **Test Drive Access** (verify file attachments work)
3. **Validate Date Parsing** (all notice formats)
4. **Preview Target Dates** (verify correct date calculations)
5. **Diagnostic Inspect Row** (verify column mapping)
6. **Send Test Email** (end-to-end test with real email)

### 10.2 Production Monitoring

1. Review LOGS sheet after each run
2. Monitor "Error" status rows
3. Track open rates if tracking enabled
4. Verify reply detection accuracy

---

## 11. Conclusion

The Automated Expiry Notification system is a **well-architected, production-ready solution** with robust error handling and comprehensive feature coverage. The codebase demonstrates professional software engineering practices including:

- ✅ Modular design with clear separation of concerns
- ✅ Comprehensive error handling with graceful degradation
- ✅ State machine for email lifecycle management
- ✅ Flexible template system with dynamic substitution
- ✅ Multiple tracking mechanisms (open + reply)

**However, the Gmail alias requirement is a critical deployment blocker** that must be addressed before full production use. Once staff aliases are configured and the HTML sanitization issue is resolved, the system is ready for production deployment.

**Final Rating: 8.5/10** (Excellent with noted caveats)

---

**Report Prepared By:** @zafajardo  
**Review Status:** Complete  
**Next Review:** Upon system updates or issue resolution
