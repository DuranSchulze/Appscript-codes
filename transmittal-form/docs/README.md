# Smart Transmittal Documentation

This `docs/` folder describes the current Smart Transmittal system as implemented in this repository.

## Quick Context

- Framework: Next.js 16, React 19, TypeScript
- App structure: Next.js App Router
- Auth: Better Auth with Google OAuth providers (`google`, `google-dds`)
- Persistence: Prisma + PostgreSQL
- Integrations: Google Drive and Google Sheets
- Parsing: Gemini-based document analysis with built-in document-number fallback generation
- Primary data flow: create, save, reopen, update, delete, and export transmittals from the database

## Documents

- `SYSTEM_OVERVIEW.md`: product-level overview of what the app does and how users work with it
- `ARCHITECTURE.md`: engineering reference for runtime flow, modules, APIs, data types, and persistence
- `RUNBOOK.md`: local setup, environment variables, database commands, and troubleshooting
- `GOOGLE_SETUP.md`: Google OAuth, Drive, and Sheets setup for this implementation
- `note.md`: backlog / planned enhancements; not guaranteed to reflect the current implementation

## Source Of Truth

These docs should follow the code, not older project notes. The key entrypoints are:

- `app/page.tsx`
- `components/main/App.tsx`
- `server/auth.ts`
- `prisma/schema.prisma`

If a doc conflicts with the code, treat the code as correct and update the doc.
