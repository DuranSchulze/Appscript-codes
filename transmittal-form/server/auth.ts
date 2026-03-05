import "dotenv/config";
import { betterAuth } from "better-auth";
import { prismaAdapter } from "better-auth/adapters/prisma";
import { nextCookies } from "better-auth/next-js";
import { genericOAuth } from "better-auth/plugins";
import { PrismaClient } from "@prisma/client";
import { withAccelerate } from "@prisma/extension-accelerate";

const parseOrigins = (value: string | undefined, fallback: string[]) => {
  const items = String(value || "")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
  return items.length ? items : fallback;
};

const GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token";

type GoogleTokenExchangeInput = {
  providerId: "google" | "google-dds";
  clientId: string;
  clientSecret: string;
  code: string;
  redirectURI: string;
  codeVerifier?: string;
};

const isTransientNetworkError = (error: unknown): boolean => {
  const message = String(
    typeof error === "object" && error && "message" in error
      ? (error as { message?: string }).message
      : error,
  );

  return (
    /aborterror/i.test(message) ||
    /fetch failed/i.test(message) ||
    /timeout/i.test(message) ||
    /und_err_connect_timeout/i.test(message) ||
    /econnreset/i.test(message) ||
    /enotfound/i.test(message) ||
    /eai_again/i.test(message)
  );
};

const parsePositiveInt = (value: string | undefined, fallback: number): number => {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? Math.floor(parsed) : fallback;
};

const ACCELERATE_MAX_ATTEMPTS = parsePositiveInt(
  process.env.PRISMA_ACCELERATE_MAX_ATTEMPTS,
  3,
);

const ACCELERATE_RETRY_BASE_MS = parsePositiveInt(
  process.env.PRISMA_ACCELERATE_RETRY_BASE_MS,
  300,
);

const isTransientAccelerateError = (error: unknown): boolean => {
  const message = String(
    typeof error === "object" && error && "message" in error
      ? (error as { message?: string }).message
      : error,
  );

  return (
    /fetch failed/i.test(message) ||
    /timeout/i.test(message) ||
    /und_err_connect_timeout/i.test(message) ||
    /econnreset/i.test(message) ||
    /enotfound/i.test(message) ||
    /eai_again/i.test(message)
  );
};

const accelerateFetchWithRetry = async (
  input: RequestInfo | URL,
  init?: RequestInit,
): Promise<Response> => {
  let lastError: unknown;

  for (let attempt = 1; attempt <= ACCELERATE_MAX_ATTEMPTS; attempt++) {
    try {
      return await fetch(input, init);
    } catch (error) {
      lastError = error;

      if (!isTransientAccelerateError(error) || attempt === ACCELERATE_MAX_ATTEMPTS) {
        throw error;
      }

      console.warn(
        `[prisma-accelerate] fetch failed, retrying (${attempt}/${ACCELERATE_MAX_ATTEMPTS})`,
        error,
      );

      await new Promise((resolve) =>
        setTimeout(resolve, ACCELERATE_RETRY_BASE_MS * attempt),
      );
    }
  }

  throw lastError || new Error("Prisma Accelerate request failed");
};

const exchangeGoogleAuthorizationCode = async ({
  providerId,
  clientId,
  clientSecret,
  code,
  redirectURI,
  codeVerifier,
}: GoogleTokenExchangeInput) => {
  const maxAttempts = 3;
  let lastError: unknown;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 20000);

    try {
      const body = new URLSearchParams({
        code,
        client_id: clientId,
        client_secret: clientSecret,
        redirect_uri: redirectURI,
        grant_type: "authorization_code",
      });

      if (codeVerifier) {
        body.set("code_verifier", codeVerifier);
      }

      const response = await fetch(GOOGLE_TOKEN_URL, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body,
        signal: controller.signal,
      });

      const data = await response.json().catch(() => ({} as Record<string, any>));

      if (!response.ok) {
        const errorCode =
          typeof data.error === "string" ? data.error : "token_exchange_failed";
        const errorDescription =
          typeof data.error_description === "string"
            ? data.error_description
            : response.statusText;

        throw new Error(
          `[${providerId}] Google token exchange failed (${response.status} ${errorCode}): ${errorDescription}`,
        );
      }

      return {
        tokenType:
          typeof data.token_type === "string" ? data.token_type : undefined,
        accessToken:
          typeof data.access_token === "string" ? data.access_token : undefined,
        refreshToken:
          typeof data.refresh_token === "string" ? data.refresh_token : undefined,
        idToken: typeof data.id_token === "string" ? data.id_token : undefined,
        accessTokenExpiresAt:
          typeof data.expires_in === "number"
            ? new Date(Date.now() + data.expires_in * 1000)
            : undefined,
        scopes:
          typeof data.scope === "string"
            ? data.scope.split(/\s+/).filter(Boolean)
            : undefined,
        raw: data,
      };
    } catch (error) {
      lastError = error;

      if (!isTransientNetworkError(error) || attempt === maxAttempts) {
        console.error(
          `[${providerId}] OAuth token exchange error (attempt ${attempt}/${maxAttempts})`,
          error,
        );
        break;
      }

      console.warn(
        `[${providerId}] Transient OAuth token exchange error, retrying (${attempt}/${maxAttempts})`,
        error,
      );
      await new Promise((resolve) => setTimeout(resolve, 500 * attempt));
    } finally {
      clearTimeout(timeout);
    }
  }

  throw (
    lastError ||
    new Error(`[${providerId}] Google token exchange failed for unknown reason`)
  );
};

export const prisma = new PrismaClient({
  accelerateUrl: process.env.DATABASE_URL,
}).$extends(
  withAccelerate({
    fetch: accelerateFetchWithRetry,
  }),
);

const globalForPrisma = globalThis as unknown as {
  prisma?: typeof prisma;
};

export const db = globalForPrisma.prisma || prisma;

if (!globalForPrisma.prisma) {
  globalForPrisma.prisma = db;
}

export const auth = betterAuth({
  baseURL: process.env.BETTER_AUTH_URL,
  trustedOrigins: parseOrigins(process.env.BETTER_AUTH_TRUSTED_ORIGINS, [
    "http://localhost:3000",
  ]),
  database: prismaAdapter(db, { provider: "postgresql" }),
  plugins: [
    nextCookies(),
    genericOAuth({
      config: [
        {
          providerId: "google",
          clientId: process.env.GOOGLE_CLIENT_ID as string,
          clientSecret: process.env.GOOGLE_CLIENT_SECRET as string,
          authorizationUrl: "https://accounts.google.com/o/oauth2/v2/auth",
          tokenUrl: "https://oauth2.googleapis.com/token",
          userInfoUrl: "https://www.googleapis.com/oauth2/v3/userinfo",
          scopes: [
            "openid",
            "https://www.googleapis.com/auth/userinfo.email",
            "https://www.googleapis.com/auth/userinfo.profile",
            "https://www.googleapis.com/auth/drive.file",
            "https://www.googleapis.com/auth/drive.readonly",
            "https://www.googleapis.com/auth/spreadsheets",
          ],
          accessType: "offline",
          prompt: "select_account",
          pkce: false,
          getToken: ({ code, redirectURI, codeVerifier }) =>
            exchangeGoogleAuthorizationCode({
              providerId: "google",
              clientId: process.env.GOOGLE_CLIENT_ID as string,
              clientSecret: process.env.GOOGLE_CLIENT_SECRET as string,
              code,
              redirectURI,
              codeVerifier,
            }),
        },
        {
          providerId: "google-dds",
          clientId: process.env.GOOGLE_DDS_CLIENT_ID as string,
          clientSecret: process.env.GOOGLE_DDS_CLIENT_SECRET as string,
          authorizationUrl: "https://accounts.google.com/o/oauth2/v2/auth",
          tokenUrl: "https://oauth2.googleapis.com/token",
          userInfoUrl: "https://www.googleapis.com/oauth2/v3/userinfo",
          scopes: [
            "openid",
            "https://www.googleapis.com/auth/userinfo.email",
            "https://www.googleapis.com/auth/userinfo.profile",
            "https://www.googleapis.com/auth/drive.file",
            "https://www.googleapis.com/auth/drive.readonly",
            "https://www.googleapis.com/auth/spreadsheets",
          ],
          accessType: "offline",
          prompt: "select_account",
          pkce: false,
          getToken: ({ code, redirectURI, codeVerifier }) =>
            exchangeGoogleAuthorizationCode({
              providerId: "google-dds",
              clientId: process.env.GOOGLE_DDS_CLIENT_ID as string,
              clientSecret: process.env.GOOGLE_DDS_CLIENT_SECRET as string,
              code,
              redirectURI,
              codeVerifier,
            }),
        },
      ],
    }),
  ],
  account: {
    accountLinking: {
      enabled: true,
    },
  },
});
