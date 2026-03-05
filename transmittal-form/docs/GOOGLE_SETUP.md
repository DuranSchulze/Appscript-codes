# Google Setup Guide

This app uses Better Auth with Google OAuth providers and then reuses the linked Google account for Drive and Sheets access.

## Auth Model

The current implementation uses two Better Auth generic OAuth providers:

- `google`
- `google-dds`

The user signs in through one of those providers. After sign-in, the app can retrieve a Google access token for the current session through:

- `GET /api/google-token`

That route reads the linked Better Auth `Account` record and can refresh an expired access token server-side when a refresh token is available.

## Required Google Scopes

The configured scopes in `server/auth.ts` are:

- `openid`
- `https://www.googleapis.com/auth/userinfo.email`
- `https://www.googleapis.com/auth/userinfo.profile`
- `https://www.googleapis.com/auth/drive.file`
- `https://www.googleapis.com/auth/drive.readonly`
- `https://www.googleapis.com/auth/spreadsheets`

## Required Environment Variables

Set the credentials for the providers you use:

```env
GOOGLE_CLIENT_ID=your-google-client-id
GOOGLE_CLIENT_SECRET=your-google-client-secret
GOOGLE_DDS_CLIENT_ID=your-google-dds-client-id
GOOGLE_DDS_CLIENT_SECRET=your-google-dds-client-secret
BETTER_AUTH_URL=http://localhost:3000
NEXT_PUBLIC_BETTER_AUTH_URL=http://localhost:3000
```

## Google Cloud Console Setup

### 1. Configure The OAuth Consent Screen

In Google Cloud Console:

1. Open **APIs & Services** > **OAuth consent screen**
2. Configure the app details
3. Add the scopes listed above
4. Add test users if the app is still in testing mode

### 2. Create OAuth 2.0 Client Credentials

Create web application credentials for the providers you intend to use.

For local development, add the authorized JavaScript origin:

```text
http://localhost:3000
```

Add authorized redirect URIs that match your Better Auth callback routes:

```text
http://localhost:3000/api/auth/callback/google
http://localhost:3000/api/auth/callback/google-dds
```

If you use a different local host or port, update both the origins and callbacks to match your `BETTER_AUTH_URL`.

## What The App Uses Google For

### Google Drive

The app uses the session-backed access token for:

- file metadata lookup
- file download as content for parsing
- listing files for Drive import
- listing folders for export target selection
- uploading generated exports to Drive

### Google Sheets

The app uses the same Google account for:

- reading rows from `Sheet1`
- appending transmittal summary rows to `Sheet1`

The linked spreadsheet ID itself is stored locally in browser `localStorage`, but Google API access is still session-backed.

## Important Implementation Notes

- The current code does not use a separate client-side Google Identity Services token flow for Drive access.
- The current code does not require a separate Drive Picker API key.
- The active implementation does not use Vite-prefixed environment variables.

## Troubleshooting

### `GET /api/google-token` Returns 404

The current session does not have a linked Google account for either `google` or `google-dds`. Sign in again with a configured Google provider.

### `GET /api/google-token` Returns 401

The session is missing, expired, or the stored Google token expired without a refresh token. Sign in again.

### Token Refresh Fails

- verify the correct client ID and client secret are set for the provider that was used
- verify the Google OAuth app is allowed to issue refresh tokens for this flow

### Drive Or Sheets Requests Fail

- verify the required scopes were granted
- verify the APIs are enabled in Google Cloud Console
- verify the signed-in account can access the target Drive files or spreadsheet
