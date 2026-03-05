# System Overview — Smart Transmittal

## What This System Is

Smart Transmittal is an authenticated web application for creating, saving, reopening, updating, deleting, and exporting formal transmittal forms.

The UI runs in the browser, but the system is not browser-only. Authentication, transmittal persistence, agency persistence, token exchange, and export data APIs are handled through Next.js server routes backed by Prisma and PostgreSQL.

## Primary User Workflow

1. The user signs in through Better Auth using either the `google` or `google-dds` provider.
2. The app restores the session, loads saved agencies, and requests the next transmittal number.
3. The user builds a transmittal by importing or manually editing line items and header fields.
4. The user saves the transmittal to the database.
5. The user exports the transmittal as PDF, DOCX, or CSV, or uploads the export to Google Drive.
6. The user can reopen and update previously saved transmittals later.

## Supported Import Sources

- Local uploads:
  - PDF files
  - image files
- Google Drive file selection from the in-app Drive browser
- Google Drive folder link
- Google Drive single file link
- Google Sheets link

## Review And Editing Surface

The app uses a split interface:

- Sidebar tabs for:
  - content
  - sender
  - recipient
  - project
  - signatories
- A live printable transmittal preview rendered from the current `AppData`

The user can:

- edit transmittal header fields
- edit item rows
- reorder rows
- adjust quantities
- remove rows
- bulk import rows from supported sources
- manage reusable agency presets

## Export And Delivery Options

- PDF export
- DOCX export
- CSV export of item rows
- Upload exported files to Google Drive
- `mailto:` helper for opening a prefilled email draft

## Persistence Model

Primary persistence is server-backed:

- transmittals are stored in PostgreSQL
- agencies are stored in PostgreSQL
- transmittal items and recipients are stored relationally

Client-side browser storage is still used for a small amount of local configuration:

- the linked Google Sheet ID is stored in `localStorage`

## Important Current Constraints

- Gemini API access is currently environment-driven, not per-user.
- Parsing can fall back when Gemini fails, returns invalid output, or returns a weak document number.
- Document-number fallback uses deterministic extraction and placeholder generation from file names, descriptions, and detected tokens.
- Google Drive actions depend on an authenticated session with a linked Google account and a valid access token from `/api/google-token`.
