# Google Apps Script Web App Implementation Guide

_Last updated: 2026-02-25_

This is a **general implementation guide** for building an Apps Script Web App that uses Google Sheets as its data store.

> Scope: This is intentionally independent from your existing DDS/FilePino scripts. It is a reusable blueprint for future projects.

---

## Table of Contents

1. [What a GAS Web App is](#1-what-a-gas-web-app-is)
2. [When to use it (and when not to)](#2-when-to-use-it-and-when-not-to)
3. [Core architecture patterns](#3-core-architecture-patterns)
4. [Security and access model (most important)](#4-security-and-access-model-most-important)
5. [Data modeling in Google Sheets](#5-data-modeling-in-google-sheets)
6. [Project setup from scratch](#6-project-setup-from-scratch)
7. [Backend implementation patterns](#7-backend-implementation-patterns)
8. [Frontend implementation patterns](#8-frontend-implementation-patterns)
9. [Row-level access control pattern](#9-row-level-access-control-pattern)
10. [CRUD implementation example](#10-crud-implementation-example)
11. [Deployment and environment strategy](#11-deployment-and-environment-strategy)
12. [Performance, quotas, and scale limits](#12-performance-quotas-and-scale-limits)
13. [Reliability and observability](#13-reliability-and-observability)
14. [Testing strategy](#14-testing-strategy)
15. [Hardening checklist before go-live](#15-hardening-checklist-before-go-live)
16. [Troubleshooting guide](#16-troubleshooting-guide)
17. [Starter template (copy/paste)](#17-starter-template-copypaste)
18. [Implementation roadmap (practical)](#18-implementation-roadmap-practical)

---

## 1) What a GAS Web App is

A Google Apps Script (GAS) Web App is a web endpoint hosted by Google that can:

- Render pages (`HtmlService`)
- Expose HTTP handlers (`doGet`, `doPost`)
- Read/write Google services (Sheets, Drive, Gmail, etc.)

In your use case, the sheet acts like a lightweight database.

---

## 2) When to use it (and when not to)

Use GAS + Sheets when:

- You need an internal app quickly
- Data volume is moderate
- Users are mostly Google Workspace users
- You want low ops overhead

Avoid or reconsider when:

- You need high throughput and low latency at scale
- You need strict enterprise auth beyond Workspace constraints
- You need complex relational queries/transactions
- You need full API gateway features (CORS policies, strict token middleware, etc.)

---

## 3) Core architecture patterns

### Pattern A: Full Web App (recommended for internal tools)

- Frontend served by Apps Script (`HtmlService`)
- Frontend talks to server functions using `google.script.run`
- No CORS headaches

### Pattern B: JSON API endpoint

- Use `doGet/doPost` returning JSON via `ContentService`
- External frontend can consume endpoint
- You must design authentication and request validation carefully

### Pattern C: Hybrid

- Internal UI in Apps Script
- Limited public API endpoints for integration jobs

---

## 4) Security and access model (most important)

When deploying Web App you choose:

1. **Execute as**
   - `Me (script owner)`
   - `User accessing the web app`

2. **Who has access**
   - Only myself
   - Anyone in domain
   - Anyone with Google account
   - Anyone (public)

### Recommended baseline for private business apps

- Execute as: **User accessing** (preferable for per-user context)
- Access: **Anyone in your Workspace domain**
- Add server-side authorization checks anyway

### Non-negotiable rules

1. Never trust the browser for authorization
2. Enforce permissions server-side on every data operation
3. Return only needed fields (data minimization)
4. Keep sensitive configuration in Script Properties
5. Log all read/write operations for audit

---

## 5) Data modeling in Google Sheets

Treat each sheet like a table with strict schema.

## 5.1 Suggested sheets

1. `Clients`
2. `AccessControl`
3. `AuditLog`
4. `Config`

## 5.2 Example schema

### Clients

| clientId | clientName | status | amount | ownerEmail | updatedAt |
|---|---|---|---:|---|---|

### AccessControl

| userEmail | clientId | role | canRead | canEdit |
|---|---|---|---|---|

### AuditLog

| timestamp | actorEmail | action | targetType | targetId | result | details |
|---|---|---|---|---|---|---|

## 5.3 Modeling rules

- Keep IDs stable (`clientId`) — never rely on row index as primary identity
- Include `updatedAt` and optional `updatedBy`
- Use controlled values for enums (`status`, `role`)
- Add header validation at startup

---

## 6) Project setup from scratch

1. Create Google Sheet (database)
2. Add tabs with headers (`Clients`, `AccessControl`, `AuditLog`)
3. Open Extensions → Apps Script
4. Create files:
   - `Code.gs` (server)
   - `Auth.gs` (authorization)
   - `Data.gs` (sheet read/write)
   - `Api.gs` (request handlers)
   - `Index.html` (UI)
   - `App.html` (client JS)
5. Set Script Properties:
   - `DB_SPREADSHEET_ID`
   - optional feature flags
6. Implement `doGet` and backend service functions
7. Deploy as Web App

---

## 7) Backend implementation patterns

## 7.1 Configuration helpers

```javascript
const SHEETS = {
  CLIENTS: 'Clients',
  ACCESS: 'AccessControl',
  AUDIT: 'AuditLog',
};

function getDb() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_SPREADSHEET_ID');
  if (!id) throw new Error('Missing DB_SPREADSHEET_ID');
  return SpreadsheetApp.openById(id);
}

function getSheet(name) {
  const sh = getDb().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name);
  return sh;
}
```

## 7.2 Identity helper

```javascript
function getCurrentUserEmailOrThrow() {
  const email = (Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email) {
    throw new Error('Unable to resolve user identity. Check deployment access settings.');
  }
  return email;
}
```

> Note: `Session.getActiveUser().getEmail()` behavior depends on Workspace context and deployment settings.

## 7.3 Central authorization gate

```javascript
function assertCanReadClient(userEmail, clientId) {
  const rules = getAccessRulesForUser(userEmail);
  const allowed = rules.some(r => r.clientId === clientId && r.canRead === true);
  if (!allowed) throw new Error('Access denied for client ' + clientId);
}
```

## 7.4 Standard response envelope

```javascript
function ok(data) {
  return { ok: true, data: data, error: null };
}

function fail(message, code) {
  return { ok: false, data: null, error: { message, code: code || 'UNKNOWN' } };
}
```

---

## 8) Frontend implementation patterns

For Apps Script-hosted UI, use `google.script.run`:

```html
<script>
  function loadMyClients() {
    google.script.run
      .withSuccessHandler(renderClients)
      .withFailureHandler(showError)
      .listMyClients();
  }
</script>
```

Frontend rules:

1. Never embed secrets in HTML/JS
2. Never assume frontend filters are security
3. Keep UI state independent from raw sheet row numbers
4. Sanitize user input before sending to server

---

## 9) Row-level access control pattern

This is the pattern for “share only specific rows with specific users”.

## 9.1 Access model

- `AccessControl` maps user emails to `clientId`
- Every read operation checks user → allowed `clientId`s
- Every update operation checks user edit permission (`canEdit`)

## 9.2 Read flow

1. Resolve current user email
2. Pull allowed client IDs
3. Query `Clients` rows
4. Return only matching rows

## 9.3 Write flow

1. Resolve current user email
2. Verify `canEdit` for target client
3. Validate payload
4. Apply update with lock (`LockService`) for conflict safety
5. Write audit row

---

## 10) CRUD implementation example

## 10.1 Read: list current user records

```javascript
function listMyClients() {
  try {
    const user = getCurrentUserEmailOrThrow();
    const allowed = new Set(getReadableClientIds(user));
    const rows = readClientsTable();
    const filtered = rows.filter(r => allowed.has(r.clientId));
    audit(user, 'READ_CLIENTS', 'bulk', 'OK', { count: filtered.length });
    return ok(filtered);
  } catch (e) {
    return fail(e.message, 'READ_FAILED');
  }
}
```

## 10.2 Update: one client record

```javascript
function updateClientStatus(input) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const user = getCurrentUserEmailOrThrow();
    const clientId = String(input.clientId || '').trim();
    const status = String(input.status || '').trim();

    assertCanEditClient(user, clientId);
    assertValidStatus(status);

    const changed = updateClientById(clientId, { status, updatedAt: new Date(), updatedBy: user });
    audit(user, 'UPDATE_CLIENT_STATUS', clientId, 'OK', { changed });
    return ok({ changed });
  } catch (e) {
    return fail(e.message, 'UPDATE_FAILED');
  } finally {
    lock.releaseLock();
  }
}
```

---

## 11) Deployment and environment strategy

## 11.1 Recommended environments

- `dev` spreadsheet + dev deployment
- `prod` spreadsheet + prod deployment

Keep IDs in Script Properties per deployment.

## 11.2 Deploy steps

1. Click **Deploy** → **New deployment**
2. Select **Web app**
3. Set execute-as and access scope
4. Deploy and authorize scopes
5. Test URL as intended user roles

## 11.3 Versioning

- Create deployment notes every release
- Maintain `CHANGELOG.md`
- Do not edit production behavior without versioned deployment

---

## 12) Performance, quotas, and scale limits

## 12.1 Optimization basics

1. Read ranges in bulk (`getDataRange().getValues()`) once
2. Transform data in memory
3. Write back in batches (`setValues`) not per-cell loops
4. Cache stable lookup data with `CacheService`

## 12.2 Contention control

- Use `LockService` for writes
- Keep lock window short
- Avoid long processing inside lock

## 12.3 Quota awareness

Apps Script has platform quotas and execution limits that vary by account type. Design for:

- short server execution paths
- retry-safe operations
- idempotent writes

---

## 13) Reliability and observability

## 13.1 Audit logging

Every critical action should log:

- timestamp
- actor
- action
- target
- result
- message/details

## 13.2 Error categorization

Use consistent error codes:

- `AUTH_FAILED`
- `ACCESS_DENIED`
- `VALIDATION_FAILED`
- `READ_FAILED`
- `UPDATE_FAILED`
- `SYSTEM_ERROR`

## 13.3 User-safe error messages

- Show generic but useful message to user
- Keep stack traces internal in logs

---

## 14) Testing strategy

## 14.1 Test matrix

1. user with access to one client
2. user with access to multiple clients
3. user with no access
4. unauthorized/public attempt
5. malformed payload
6. concurrent update scenario

## 14.2 Practical tests

- Verify each user only sees expected rows
- Verify forbidden updates are blocked
- Verify audit rows are written for reads/updates/errors

---

## 15) Hardening checklist before go-live

- [ ] Web app deployment access is restricted correctly
- [ ] Server-side auth checks exist on every endpoint/function
- [ ] No sensitive data leaked to frontend
- [ ] Row-level access is validated against `AccessControl`
- [ ] Input validation exists for all writes
- [ ] `LockService` used for conflicting writes
- [ ] Audit logging is complete
- [ ] Backup/export strategy defined
- [ ] Rollback plan documented
- [ ] At least 2-role user acceptance tests completed

---

## 16) Troubleshooting guide

### Problem: User sees blank data

Common causes:

- `Session.getActiveUser()` email unavailable for current deployment mode
- no matching access rows in `AccessControl`
- case mismatch in email strings

### Problem: User sees too much data

Common causes:

- client-side filtering only (missing server auth gate)
- forgot `assertCanReadClient` on endpoint
- exposed full table endpoint

### Problem: Updates collide or overwrite

Common causes:

- no `LockService`
- write-by-row-index instead of write-by-id
- missing optimistic checks (`updatedAt` comparison)

### Problem: Slow response

Common causes:

- per-row read/writes in loops
- repeated sheet open calls
- no caching for access rules

---

## 17) Starter template (copy/paste)

## 17.1 `Code.gs`

```javascript
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Client Portal');
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}
```

## 17.2 `Data.gs`

```javascript
function readTableAsObjects(sheetName) {
  const sh = getSheet(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => String(h).trim());
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => (obj[h] = row[i]));
    return obj;
  });
}
```

## 17.3 `Auth.gs`

```javascript
function getReadableClientIds(userEmail) {
  const rows = readTableAsObjects('AccessControl');
  return rows
    .filter(r => String(r.userEmail || '').toLowerCase() === userEmail.toLowerCase())
    .filter(r => r.canRead === true || String(r.canRead).toUpperCase() === 'TRUE')
    .map(r => String(r.clientId || '').trim())
    .filter(Boolean);
}
```

## 17.4 `Api.gs`

```javascript
function listMyClients() {
  const user = getCurrentUserEmailOrThrow();
  const allowed = new Set(getReadableClientIds(user));
  const rows = readTableAsObjects('Clients');
  return rows.filter(r => allowed.has(String(r.clientId || '').trim()));
}
```

## 17.5 `Index.html`

```html
<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width,initial-scale=1" />
    <title>Client Portal</title>
  </head>
  <body>
    <h1>My Clients</h1>
    <pre id="out">Loading...</pre>
    <script>
      google.script.run
        .withSuccessHandler(function(data) {
          document.getElementById('out').textContent = JSON.stringify(data, null, 2);
        })
        .withFailureHandler(function(err) {
          document.getElementById('out').textContent = 'Error: ' + err.message;
        })
        .listMyClients();
    </script>
  </body>
</html>
```

---

## 18) Implementation roadmap (practical)

### Phase 1: Foundation (Day 1)

1. Build sheet schema
2. Implement read-only web app
3. Enforce row-level read authorization
4. Deploy to restricted test users

### Phase 2: Controlled writes (Day 2)

1. Add update endpoints
2. Add validation + lock + audit
3. Add role-based edit permissions

### Phase 3: Hardening (Day 3)

1. Add caching + performance tuning
2. Add full error taxonomy
3. Add monitoring checklist and runbook
4. UAT with real users

### Phase 4: Production

1. Production deployment with version tag
2. Access review and permission freeze
3. Go-live and incident support window

---

## Final guidance

If your requirement is: **"specific user should only see specific client rows"**, then this stack works well when implemented with strict server-side authorization.

Use this mindset:

- **Sheet = storage**
- **Apps Script backend = security boundary**
- **Web UI = presentation only**

That separation is what makes the system safe and maintainable.
