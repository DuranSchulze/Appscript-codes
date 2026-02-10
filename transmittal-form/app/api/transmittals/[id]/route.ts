import { NextRequest, NextResponse } from "next/server";
import { auth, db } from "@/server/auth";

export const runtime = "nodejs";

const ensureDbTransmittalPrefix = (value: string) => {
  const trimmed = String(value || "").trim();
  if (!trimmed) return "";
  return trimmed.startsWith("TR-FP-") ? trimmed : `TR-FP-${trimmed}`;
};

const stripTransmittalPrefix = (value: string) =>
  value.startsWith("TR-FP-") ? value.slice("TR-FP-".length) : value;

const mapTransmittalForApi = (transmittal: any) => {
  if (!transmittal) return transmittal;
  const project = (transmittal.project || {}) as any;
  const nextProject = {
    ...project,
    transmittalNumber: stripTransmittalPrefix(
      String(project.transmittalNumber || ""),
    ),
  };
  return {
    ...transmittal,
    project: nextProject,
  };
};

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

    const transmittal = await db.transmittal.findFirst({
      where: { id, userId: session.user.id },
      select: { id: true },
    });

    if (!transmittal) {
      return NextResponse.json({ error: "Transmittal not found" }, { status: 404 });
    }

    await db.transmittal.delete({ where: { id } });

    return NextResponse.json({ ok: true });
  } catch (error: any) {
    console.error("Delete transmittal error:", error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}

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

    const { id: transmittalId } = await ctx.params;

    const body = await request.json();
    const data = body?.data;
    if (!data) {
      return NextResponse.json(
        { error: "Missing transmittal data" },
        { status: 400 },
      );
    }

    const agencyId = data.agencyId ? String(data.agencyId) : null;
    if (agencyId) {
      const agency = await db.agency.findFirst({
        where: { id: agencyId, userId: session.user.id },
        select: { id: true },
      });
      if (!agency) {
        return NextResponse.json(
          { error: "Invalid agency" },
          { status: 400 },
        );
      }
    }

    const rawTransmittalNumber = String(data.project?.transmittalNumber || "").trim();
    const dbTransmittalNumber = ensureDbTransmittalPrefix(rawTransmittalNumber);

    const project = {
      ...(data.project || {}),
      transmittalNumber: dbTransmittalNumber,
    };

    const existing = await db.transmittal.findFirst({
      where: { id: transmittalId, userId: session.user.id },
    });

    if (!existing) {
      return NextResponse.json({ error: "Transmittal not found" }, { status: 404 });
    }

    if (dbTransmittalNumber) {
      const duplicate = await db.transmittal.findFirst({
        where: {
          transmittalNumber: dbTransmittalNumber,
          id: { not: transmittalId },
        },
        select: { id: true },
      });
      if (duplicate) {
        return NextResponse.json(
          { error: `Transmittal number "${rawTransmittalNumber}" is already in use.` },
          { status: 409 },
        );
      }
    }

    const transmittal = await db.transmittal.update({
      where: { id: transmittalId },
      data: {
        notes: data.notes || "",
        transmittalNumber: dbTransmittalNumber || null,
        agencyId,
        handDelivery: Boolean(data.transmissionMethod?.personalDelivery),
        pickUp: Boolean(data.transmissionMethod?.pickUp),
        courier: Boolean(data.transmissionMethod?.grabLalamove),
        registeredMail: Boolean(data.transmissionMethod?.registeredMail),
        projectName: data.project?.projectName || "",
        projectNumber: data.project?.projectNumber || null,
        engagementRefNumber: data.project?.engagementRef || null,
        projectPurpose: data.project?.purpose || null,
        project,
        sender: data.sender || {},
        receivedBy: data.receivedBy || {},
        footerNotes: data.footerNotes || {},
        preparedBy: data.signatories?.preparedBy || "",
        preparedByRole: data.signatories?.preparedByRole || "",
        notedBy: data.signatories?.notedBy || "",
        notedByRole: data.signatories?.notedByRole || "",
        timeReleased: data.signatories?.timeReleased || "",
        recipients: {
          deleteMany: {},
          create: [
            {
              recipientName: data.recipient?.to || "",
              recipientOrganization: data.recipient?.company || null,
              recipientAttention: data.recipient?.attention || null,
              recipientFullAddress: data.recipient?.address || null,
              recipientAgencyContactNumber: data.recipient?.contactNumber || null,
              recipientAgencyEmail: data.recipient?.email || null,
            },
          ],
        },
        items: {
          deleteMany: {},
          create: Array.isArray(data.items)
            ? data.items.map((item: any) => ({
                qty: item.qty || "",
                noOfItems: item.noOfItems || "",
                documentNumber: item.documentNumber || "",
                description: item.description || "",
                remarks: item.remarks || "",
                fileType: item.fileType || null,
                fileSource: item.fileSource || null,
              }))
            : [],
        },
      },
      include: {
        items: true,
        recipients: true,
        agency: true,
      },
    });

    return NextResponse.json({ transmittal: mapTransmittalForApi(transmittal) });
  } catch (error: any) {
    console.error("Update transmittal error:", error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}
