# Runbook

## Prerequisites

- Node.js installed
- access to a PostgreSQL database
- Google OAuth credentials for the enabled providers
- a Gemini API key if AI parsing should be enabled

## Install

```bash
npm install
```

## Required Environment Variables

Create an `.env` file in the project root and set the values required for your environment.

### Core App And Auth

- `DATABASE_URL`
  - PostgreSQL connection string
  - also used by the Prisma client configuration in `server/auth.ts`
- `BETTER_AUTH_SECRET`
  - secret used by Better Auth
- `BETTER_AUTH_URL`
  - server-side canonical auth base URL
  - example: `http://localhost:3000`
- `NEXT_PUBLIC_BETTER_AUTH_URL`
  - browser-visible auth base URL used by the client
  - if omitted, the client falls back to `window.location.origin`
- `BETTER_AUTH_TRUSTED_ORIGINS`
  - optional comma-separated trusted origins

### Google OAuth

- `GOOGLE_CLIENT_ID`
- `GOOGLE_CLIENT_SECRET`
- `GOOGLE_DDS_CLIENT_ID`
- `GOOGLE_DDS_CLIENT_SECRET`

These power the `google` and `google-dds` Better Auth providers and are also used by `/api/google-token` when refreshing access tokens.

### Gemini Parsing

`services/geminiService.ts` checks these names in this order:

- `NEXT_PUBLIC_GEMINI_API_KEY`
- `NEXT_PUBLIC_GOOGLE_API_KEY`
- `GEMINI_API_KEY`
- `GOOGLE_API_KEY`

Optional:

- `GEMINI_MODEL`
  - defaults to `gemini-2.5-flash`

### User AI Key Encryption

- `APP_SETTINGS_ENCRYPTION_KEY`
  - required for storing and reading user-saved Gemini API keys
  - used by `server/user-ai-settings.ts` for encryption/decryption
  - should be a strong secret and kept server-side only

### Export API

- `SEND_API_TOKEN`
  - required for `GET /api/export-transmittals`

## Database Setup

Generate the Prisma client and sync the schema:

```bash
npm run db:sync
```

Reset the database schema if needed:

```bash
npm run db:reset
```

## Run Locally

Start the development server:

```bash
npm run dev
```

By default, the app is typically served from:

```text
http://localhost:3000
```

## Build And Start

Production build:

```bash
npm run build
```

Run the production server:

```bash
npm run start
```

## Google Token Flow

The browser does not directly manage long-lived Google OAuth tokens.

Instead:

1. The user signs in through Better Auth.
2. The client calls `GET /api/google-token`.
3. The server loads the linked `Account` record.
4. If the access token is expired and a refresh token exists, the server refreshes it.
5. The server returns a usable access token to the client.

`services/googleDriveService.ts` uses this route for Drive and Sheets actions.

## Better Auth Redirect Notes

For local development, make sure your Google OAuth applications allow the callback URLs used by Better Auth under:

```text
/api/auth/callback/google
/api/auth/callback/google-dds
```

The full URL depends on `BETTER_AUTH_URL`.

Example local callback URLs:

```text
http://localhost:3000/api/auth/callback/google
http://localhost:3000/api/auth/callback/google-dds
```

## Available NPM Scripts

- `npm run dev`
- `npm run dev:clean`
- `npm run build`
- `npm run start`
- `npm run db:sync`
- `npm run db:reset`

## Troubleshooting

### Authentication Fails

- verify `BETTER_AUTH_SECRET`
- verify `BETTER_AUTH_URL`
- verify Google client IDs, secrets, and callback URLs
- verify trusted origins if `BETTER_AUTH_TRUSTED_ORIGINS` is set

### Google Drive Is Unavailable

- verify the signed-in user has a linked Google account
- verify `GET /api/google-token` succeeds
- verify the relevant Google client ID and secret are present
- if the token expired without a refresh token, sign in again

### Parsing Falls Back Instead Of Using AI

- verify one of the supported Gemini env vars is set
- verify the key is valid and not quota-limited
- if using personal AI keys, verify `APP_SETTINGS_ENCRYPTION_KEY` is configured
- verify `GET /api/user-ai-settings` and `PUT /api/user-ai-settings` work for the signed-in user
- remember that weak or generic document numbers are intentionally replaced by fallback logic

### Export API Returns 401

- verify `SEND_API_TOKEN` is configured
- verify the request supplies the same token in `x-api-token` or the `token` query parameter
