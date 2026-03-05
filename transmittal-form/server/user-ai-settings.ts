import { createCipheriv, createDecipheriv, createHash, randomBytes } from "crypto";
import { db } from "@/server/auth";

const ENCRYPTION_ENV_KEY = "APP_SETTINGS_ENCRYPTION_KEY";
const ALGORITHM = "aes-256-gcm";

const deriveKey = (): Buffer => {
  const secret = String(process.env[ENCRYPTION_ENV_KEY] || "").trim();
  if (!secret) {
    throw new Error(
      `${ENCRYPTION_ENV_KEY} is not configured. Cannot encrypt/decrypt user AI keys.`,
    );
  }
  return createHash("sha256").update(secret).digest();
};

const encodePayload = (iv: Buffer, tag: Buffer, encrypted: Buffer): string =>
  `${iv.toString("base64")}.${tag.toString("base64")}.${encrypted.toString("base64")}`;

const decodePayload = (payload: string): { iv: Buffer; tag: Buffer; encrypted: Buffer } => {
  const [ivB64, tagB64, dataB64] = String(payload || "").split(".");
  if (!ivB64 || !tagB64 || !dataB64) {
    throw new Error("Invalid encrypted payload format.");
  }
  return {
    iv: Buffer.from(ivB64, "base64"),
    tag: Buffer.from(tagB64, "base64"),
    encrypted: Buffer.from(dataB64, "base64"),
  };
};

export const encryptUserGeminiApiKey = (plainText: string): string => {
  const value = String(plainText || "").trim();
  if (!value) throw new Error("Missing API key.");

  const key = deriveKey();
  const iv = randomBytes(12);
  const cipher = createCipheriv(ALGORITHM, key, iv);
  const encrypted = Buffer.concat([cipher.update(value, "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  return encodePayload(iv, tag, encrypted);
};

export const decryptUserGeminiApiKey = (payload: string): string => {
  const { iv, tag, encrypted } = decodePayload(payload);
  const key = deriveKey();
  const decipher = createDecipheriv(ALGORITHM, key, iv);
  decipher.setAuthTag(tag);
  const decrypted = Buffer.concat([decipher.update(encrypted), decipher.final()]);
  const value = decrypted.toString("utf8").trim();
  if (!value) {
    throw new Error("Decrypted API key is empty.");
  }
  return value;
};

export const isLikelyGeminiApiKey = (value: string): boolean => {
  const trimmed = String(value || "").trim();
  if (!trimmed) return false;
  // Soft validation: basic length + allowed token chars.
  return /^[A-Za-z0-9_\-]{20,200}$/.test(trimmed);
};

export const getUserAiKeyMetadata = async (userId: string) => {
  const settings = await db.userAiSettings.findUnique({
    where: { userId },
    select: { updatedAt: true },
  });

  return {
    hasCustomGeminiKey: Boolean(settings),
    updatedAt: settings?.updatedAt ?? null,
  };
};

export const getDecryptedUserGeminiApiKey = async (
  userId: string,
): Promise<string | null> => {
  const settings = await db.userAiSettings.findUnique({
    where: { userId },
    select: { geminiApiKeyEncrypted: true },
  });

  if (!settings?.geminiApiKeyEncrypted) return null;

  try {
    return decryptUserGeminiApiKey(settings.geminiApiKeyEncrypted);
  } catch (error) {
    console.error("Failed to decrypt user Gemini API key:", error);
    return null;
  }
};

