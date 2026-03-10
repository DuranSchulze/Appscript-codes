import { NextResponse } from "next/server";
import { auth, db } from "@/server/auth";

export const runtime = "nodejs";

const stripTransmittalPrefix = (value: string) =>
  value.startsWith("TR-FP-") ? value.slice("TR-FP-".length) : value;

const ensureDbTransmittalPrefix = (value: string) => {
  const trimmed = String(value || "").trim();
  if (!trimmed) return "";
  return trimmed.startsWith("TR-FP-") ? trimmed : `TR-FP-${trimmed}`;
};

const mapTransmittalForApi = (transmittal: any) => {
  if (!transmittal) return transmittal;
  const project = (transmittal.project || {}) as any;
  const nextProject = {
    ...project,
    transmittalNumber: stripTransmittalPrefix(
      String(project.transmittalNumber || ""),
    ),
    department: String(project.department || transmittal.department || ""),
  };
  return {
    ...transmittal,
    project: nextProject,
  };
};

export async function GET(request: Request) {
  try {
    const session = await auth.api.getSession({
      headers: request.headers as any,
    });

    if (!session?.user) {
      return NextResponse.json({ error: "Not authenticated" }, { status: 401 });
    }

    const transmittals = await db.transmittal.findMany({
      where: { userId: session.user.id },
      include: {
        items: true,
        recipients: true,
        agency: true,
      },
    });

    return NextResponse.json({
      transmittals: transmittals.map(mapTransmittalForApi),
    });
  } catch (error: any) {
    console.error("Load transmittals error:", error);
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

    const transmittal = await db.$transaction(async (tx) => {
      const requestedTransmittalNumber = String(
        data.project?.transmittalNumber || "",
      ).trim();
      const now = new Date();
      const yearMonth = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}`;
      const prefix = `TR-FP-${yearMonth}-`;
      let dbTransmittalNumber = ensureDbTransmittalPrefix(requestedTransmittalNumber);

      if (dbTransmittalNumber) {
        const duplicate = await tx.transmittal.findFirst({
          where: {
            userId: session.user.id,
            transmittalNumber: dbTransmittalNumber,
          },
          select: { id: true },
        });
        if (duplicate) {
          throw new Error(
            `Transmittal number "${requestedTransmittalNumber}" is already in use.`,
          );
        }
      } else {
        const existing = await tx.transmittal.findMany({
          where: {
            userId: session.user.id,
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
        dbTransmittalNumber = `${prefix}${String(nextSeq).padStart(4, "0")}`;
      }
      const department = String(data.project?.department || "").trim();

      const project = {
        ...(data.project || {}),
        transmittalNumber: dbTransmittalNumber,
        department,
      };

      return tx.transmittal.create({
        data: {
          userId: session.user.id,
          notes: data.notes || "",
          transmittalNumber: dbTransmittalNumber,
          agencyId,
          handDelivery: Boolean(data.transmissionMethod?.personalDelivery),
          pickUp: Boolean(data.transmissionMethod?.pickUp),
          courier: Boolean(data.transmissionMethod?.grabLalamove),
          registeredMail: Boolean(data.transmissionMethod?.registeredMail),
          projectName: data.project?.projectName || "",
          projectNumber: data.project?.projectNumber || null,
          engagementRefNumber: data.project?.engagementRef || null,
          projectPurpose: data.project?.purpose || null,
          department: department || null,
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
    });

    return NextResponse.json({ transmittal: mapTransmittalForApi(transmittal) });
  } catch (error: any) {
    console.error("Save transmittal error:", error);
    if (
      error?.message?.includes("already in use") ||
      error?.code === "P2002"
    ) {
      return NextResponse.json(
        { error: error.message || "Transmittal number is already in use." },
        { status: 409 },
      );
    }
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}
