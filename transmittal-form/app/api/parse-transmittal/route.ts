import { NextResponse } from "next/server";
import { parseTransmittalDocument } from "@/services/geminiService";
import { auth } from "@/server/auth";
import { getDecryptedUserGeminiApiKey } from "@/server/user-ai-settings";

export const runtime = "nodejs";

const MAX_CONTENT_LENGTH = 25 * 1024 * 1024;

export async function POST(request: Request) {
  try {
    let userGeminiApiKey: string | undefined;
    try {
      const session = await auth.api.getSession({
        headers: request.headers as any,
      });
      if (session?.user?.id) {
        const decrypted = await getDecryptedUserGeminiApiKey(session.user.id);
        if (decrypted) {
          userGeminiApiKey = decrypted;
        }
      }
    } catch (sessionError) {
      // Allow parser to continue with env-key fallback when session resolution fails.
      console.warn("Session lookup failed for parse route:", sessionError);
    }

    const body = await request.json();
    const content = String(body?.content || "");
    const mimeType = String(body?.mimeType || "");
    const isText = Boolean(body?.isText);
    const fileName = body?.fileName ? String(body.fileName) : undefined;

    if (!content || !mimeType) {
      return NextResponse.json(
        { error: "Missing content or mimeType." },
        { status: 400 },
      );
    }

    if (content.length > MAX_CONTENT_LENGTH) {
      return NextResponse.json(
        { error: "File content is too large to analyze." },
        { status: 413 },
      );
    }

    const result = await parseTransmittalDocument(
      content,
      mimeType,
      isText,
      fileName,
      userGeminiApiKey,
    );

    return NextResponse.json(result);
  } catch (error: any) {
    console.error("Parse transmittal error:", error);
    return NextResponse.json(
      { error: String(error?.message || "Failed to analyze document.") },
      { status: 500 },
    );
  }
}
