import { NextResponse } from "next/server";
import { auth, db } from "@/server/auth";

export const runtime = "nodejs";

const TRANSMITTAL_PREFIX = "TR-FP-";

/**
 * Generate the next unique transmittal number for the current month.
 * Format: YYYYMM-XXXX (e.g. 202601-0001)
 * The DB stores it with the TR-FP- prefix, but the API returns it without.
 */
export async function GET(request: Request) {
  try {
    const session = await auth.api.getSession({
      headers: request.headers as any,
    });

    if (!session?.user) {
      return NextResponse.json({ error: "Not authenticated" }, { status: 401 });
    }

    const now = new Date();
    const yearMonth = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}`;
    const prefix = `${TRANSMITTAL_PREFIX}${yearMonth}-`;

    // Find the highest existing number for this month across ALL users
    // to ensure global uniqueness
    const existing = await db.transmittal.findMany({
      where: {
        transmittalNumber: { startsWith: prefix },
      },
      select: { transmittalNumber: true },
    });

    let maxSeq = 0;
    for (const row of existing) {
      if (!row.transmittalNumber) continue;
      const suffix = row.transmittalNumber.slice(prefix.length);
      const num = Number(suffix);
      if (!Number.isNaN(num) && num > maxSeq) {
        maxSeq = num;
      }
    }

    const nextSeq = maxSeq + 1;
    const nextNumber = `${yearMonth}-${String(nextSeq).padStart(4, "0")}`;

    return NextResponse.json({ transmittalNumber: nextNumber });
  } catch (error: any) {
    console.error("Generate next transmittal number error:", error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}
