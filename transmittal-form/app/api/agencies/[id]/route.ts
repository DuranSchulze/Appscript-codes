import { NextRequest, NextResponse } from "next/server";
import { auth, db } from "@/server/auth";

export const runtime = "nodejs";

export async function PUT(
  request: NextRequest,
  ctx: { params: Promise<{ id: string }> },
) {
  try {
    const session = await auth.api.getSession({
      headers: request.headers as any,
    });

    if (!session?.user) {
      return NextResponse.json({ error: "Not authenticated" }, { status: 401 });
    }

    const { id } = await ctx.params;

    const body = await request.json();
    const agency = body?.agency;

    if (!agency || !String(agency.name || "").trim()) {
      return NextResponse.json({ error: "Agency name is required" }, { status: 400 });
    }

    const existing = await db.agency.findFirst({
      where: { id, userId: session.user.id },
    });

    if (!existing) {
      return NextResponse.json({ error: "Agency not found" }, { status: 404 });
    }

    const updated = await db.agency.update({
      where: { id },
      data: {
        name: String(agency.name).trim(),
        addressLine1: agency.addressLine1 || null,
        addressLine2: agency.addressLine2 || null,
        website: agency.website || null,
        telephoneNumber: agency.telephoneNumber || null,
        contactNumber: agency.contactNumber || null,
        email: agency.email || null,
        logoBase64: agency.logoBase64 || null,
      },
    });

    return NextResponse.json({ agency: updated });
  } catch (error: any) {
    console.error("Update agency error:", error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}

export async function DELETE(
  request: NextRequest,
  ctx: { params: Promise<{ id: string }> },
) {
  try {
    const session = await auth.api.getSession({
      headers: request.headers as any,
    });

    if (!session?.user) {
      return NextResponse.json({ error: "Not authenticated" }, { status: 401 });
    }

    const { id } = await ctx.params;

    const existing = await db.agency.findFirst({
      where: { id, userId: session.user.id },
      select: { id: true },
    });

    if (!existing) {
      return NextResponse.json({ error: "Agency not found" }, { status: 404 });
    }

    await db.agency.delete({ where: { id } });

    return NextResponse.json({ ok: true });
  } catch (error: any) {
    console.error("Delete agency error:", error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}
