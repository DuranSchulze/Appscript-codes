import { NextResponse } from "next/server";
import { auth, db } from "@/server/auth";

export const runtime = "nodejs";

export async function GET(request: Request) {
  try {
    const session = await auth.api.getSession({
      headers: request.headers as any,
    });

    if (!session?.user) {
      return NextResponse.json({ error: "Not authenticated" }, { status: 401 });
    }

    const agencies = await db.agency.findMany({
      where: { userId: session.user.id },
      orderBy: { updatedAt: "desc" },
    });

    return NextResponse.json({ agencies });
  } catch (error: any) {
    console.error("Load agencies error:", error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}

export async function POST(request: Request) {
  try {
    const session = await auth.api.getSession({
      headers: request.headers as any,
    });

    if (!session?.user) {
      return NextResponse.json({ error: "Not authenticated" }, { status: 401 });
    }

    const body = await request.json();
    const agency = body?.agency;

    if (!agency || !String(agency.name || "").trim()) {
      return NextResponse.json({ error: "Agency name is required" }, { status: 400 });
    }

    const name = String(agency.name).trim();

    const upserted = await db.agency.upsert({
      where: {
        userId_name: {
          userId: session.user.id,
          name,
        },
      },
      create: {
        userId: session.user.id,
        name,
        addressLine1: agency.addressLine1 || null,
        addressLine2: agency.addressLine2 || null,
        website: agency.website || null,
        telephoneNumber: agency.telephoneNumber || null,
        contactNumber: agency.contactNumber || null,
        email: agency.email || null,
        logoBase64: agency.logoBase64 || null,
      },
      update: {
        addressLine1: agency.addressLine1 || null,
        addressLine2: agency.addressLine2 || null,
        website: agency.website || null,
        telephoneNumber: agency.telephoneNumber || null,
        contactNumber: agency.contactNumber || null,
        email: agency.email || null,
        logoBase64: agency.logoBase64 || null,
      },
    });

    return NextResponse.json({ agency: upserted });
  } catch (error: any) {
    console.error("Create agency error:", error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}
