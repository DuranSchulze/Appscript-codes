import { NextResponse } from "next/server";
import { GoogleGenAI } from "@google/genai";
import { auth, db } from "@/server/auth";
import {
  encryptUserGeminiApiKey,
  getUserAiKeyMetadata,
  isLikelyGeminiApiKey,
} from "@/server/user-ai-settings";

export const runtime = "nodejs";

const resolveGeminiModel = (): string => {
  return process.env.GEMINI_MODEL || "gemini-2.5-flash";
};

const authenticate = async (request: Request) => {
  const session = await auth.api.getSession({
    headers: request.headers as any,
  });

  if (!session?.user?.id) {
    return null;
  }

  return session.user;
};

export async function GET(request: Request) {
  try {
    const user = await authenticate(request);
    if (!user) {
      return NextResponse.json({ error: "Not authenticated" }, { status: 401 });
    }

    const metadata = await getUserAiKeyMetadata(user.id);
    return NextResponse.json(metadata);
  } catch (error: any) {
    console.error("Get user AI settings error:", error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}

export async function PUT(request: Request) {
  try {
    const user = await authenticate(request);
    if (!user) {
      return NextResponse.json({ error: "Not authenticated" }, { status: 401 });
    }

    const body = await request.json().catch(() => ({}));
    const geminiApiKey = String(body?.geminiApiKey || "").trim();

    if (!geminiApiKey) {
      return NextResponse.json(
        { error: "Gemini API key is required." },
        { status: 400 },
      );
    }

    if (!isLikelyGeminiApiKey(geminiApiKey)) {
      return NextResponse.json(
        { error: "Gemini API key format looks invalid." },
        { status: 400 },
      );
    }

    const encrypted = encryptUserGeminiApiKey(geminiApiKey);

    let warning: string | undefined;
    try {
      const ai = new GoogleGenAI({ apiKey: geminiApiKey });
      await ai.models.generateContent({
        model: resolveGeminiModel(),
        contents: "Ping",
        config: { maxOutputTokens: 8, temperature: 0 },
      });
    } catch {
      warning =
        "Key was saved, but live validation failed. It may be invalid, restricted, or temporarily unavailable.";
    }

    const saved = await db.userAiSettings.upsert({
      where: { userId: user.id },
      create: {
        userId: user.id,
        geminiApiKeyEncrypted: encrypted,
      },
      update: {
        geminiApiKeyEncrypted: encrypted,
      },
      select: { updatedAt: true },
    });

    return NextResponse.json({
      ok: true,
      hasCustomGeminiKey: true,
      updatedAt: saved.updatedAt,
      warning,
    });
  } catch (error: any) {
    console.error("Save user AI settings error:", error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}

export async function DELETE(request: Request) {
  try {
    const user = await authenticate(request);
    if (!user) {
      return NextResponse.json({ error: "Not authenticated" }, { status: 401 });
    }

    await db.userAiSettings.deleteMany({
      where: { userId: user.id },
    });

    return NextResponse.json({ ok: true, hasCustomGeminiKey: false });
  } catch (error: any) {
    console.error("Delete user AI settings error:", error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}

