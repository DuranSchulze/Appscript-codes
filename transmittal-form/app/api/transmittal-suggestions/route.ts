import { NextResponse } from "next/server";
import { auth, db } from "@/server/auth";

export const runtime = "nodejs";

const MAX_HISTORY_ROWS = 500;

type SuggestionField =
  | "projectName"
  | "department"
  | "preparedBy"
  | "preparedByRole"
  | "notedBy"
  | "notedByRole";

const buildUniqueSuggestions = (
  rows: Array<Record<SuggestionField, string | null>>,
  field: SuggestionField,
): string[] => {
  const seen = new Set<string>();
  const items: string[] = [];

  for (const row of rows) {
    const value = String(row[field] || "").trim();
    if (!value) continue;

    const normalized = value.toLowerCase();
    if (seen.has(normalized)) continue;

    seen.add(normalized);
    items.push(value);
  }

  return items;
};

export async function GET(request: Request) {
  try {
    const session = await auth.api.getSession({
      headers: request.headers as any,
    });

    if (!session?.user?.id) {
      return NextResponse.json({ error: "Not authenticated" }, { status: 401 });
    }

    const rows = await db.transmittal.findMany({
      where: { userId: session.user.id },
      select: {
        projectName: true,
        department: true,
        preparedBy: true,
        preparedByRole: true,
        notedBy: true,
        notedByRole: true,
      },
      orderBy: { updatedAt: "desc" },
      take: MAX_HISTORY_ROWS,
    });

    return NextResponse.json({
      projectNames: buildUniqueSuggestions(rows, "projectName"),
      departments: buildUniqueSuggestions(rows, "department"),
      preparedByNames: buildUniqueSuggestions(rows, "preparedBy"),
      preparedByRoles: buildUniqueSuggestions(rows, "preparedByRole"),
      notedByNames: buildUniqueSuggestions(rows, "notedBy"),
      notedByRoles: buildUniqueSuggestions(rows, "notedByRole"),
    });
  } catch (error: any) {
    console.error("Load transmittal suggestions error:", error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}
