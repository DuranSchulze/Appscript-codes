import "dotenv/config";
import express from "express";
import cors from "cors";
import cookieParser from "cookie-parser";
import { toNodeHandler } from "better-auth/node";
import { auth, prisma } from "./auth";

const app = express();
const port = Number(process.env.PORT) || 8000;
const SEND_API_TOKEN = process.env.SEND_API_TOKEN;

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
    transmittalNumber: stripTransmittalPrefix(String(project.transmittalNumber || "")),
  };
  return {
    ...transmittal,
    project: nextProject,
  };
};

app.use(
  cors({
    origin: ["http://localhost:3000"],
    credentials: true,
  })
);

app.use(cookieParser());

// Better Auth handler must be mounted before json parsing.
app.all("/api/auth/*", toNodeHandler(auth));

app.use(express.json());

app.get("/health", (_req, res) => {
  res.json({ ok: true });
});

app.get("/api/export-transmittals", async (req, res) => {
  try {
    if (!SEND_API_TOKEN) {
      return res.status(500).json({ error: "SEND_API_TOKEN not configured" });
    }

    const tokenFromHeader = req.header("x-api-token");
    const tokenFromQuery = typeof req.query.token === "string" ? req.query.token : null;
    const token = tokenFromHeader || tokenFromQuery;

    if (!token || token !== SEND_API_TOKEN) {
      return res.status(401).json({ error: "Invalid token" });
    }

    const transmittals = await prisma.transmittal.findMany({
      include: {
        items: true,
        recipients: true,
        user: true,
      },
    });

    const rows = transmittals.flatMap((t) => {
      const project = (t.project || {}) as any;
      const transmittalNo = stripTransmittalPrefix(
        String(project.transmittalNumber || ""),
      );
      const date = String(project.date || "");

      const primaryRecipient = t.recipients?.[0];
      const client =
        primaryRecipient?.recipientOrganization ||
        primaryRecipient?.recipientName ||
        "";

      const mode = [
        t.handDelivery ? "Hand Delivery" : null,
        t.pickUp ? "Pick Up" : null,
        t.courier ? "Courier" : null,
        t.registeredMail ? "Registered Mail" : null,
      ]
        .filter(Boolean)
        .join(", ");

      const transmittedBy = t.preparedBy || t.user?.name || t.user?.email || "";

      const items = (t.items || []).length ? t.items : [null];
      return items.map((item) => {
        const docs = item
          ? [item.documentNumber, item.description].filter(Boolean).join(" - ")
          : "";

        return {
          client,
          documents: docs,
          transmittalDocumentNo: transmittalNo,
          dateOfTransmittal: date,
          transmittedBy,
          mode,
          gdLink:
            item && item.fileType === "gdrive" ? item.fileSource || "" : "",
          remarks: item?.remarks || "",
        };
      });
    });

    return res.json({
      rows,
      totalRows: rows.length,
      totalTransmittals: transmittals.length,
    });
  } catch (error: any) {
    console.error("Export transmittals error:", error);
    return res.status(500).json({ error: error.message });
  }
});

// Endpoint to get Google access token for Drive operations
app.get("/api/google-token", async (req, res) => {
  try {
    const session = await auth.api.getSession({
      headers: req.headers as any,
    });

    if (!session?.user) {
      return res.status(401).json({ error: "Not authenticated" });
    }

    // Get the account linked to Google for this user
    const account = await prisma.account.findFirst({
      where: {
        userId: session.user.id,
        providerId: "google",
      },
    });

    if (!account) {
      return res.status(404).json({ error: "No Google account linked" });
    }

    // Check if token is expired and refresh if needed
    const expiresAt = account.accessTokenExpiresAt ? new Date(account.accessTokenExpiresAt) : new Date(0);
    if (expiresAt < new Date() && account.refreshToken) {
      // Refresh the token
      const tokenResponse = await fetch("https://oauth2.googleapis.com/token", {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: process.env.GOOGLE_CLIENT_ID!,
          client_secret: process.env.GOOGLE_CLIENT_SECRET!,
          refresh_token: account.refreshToken,
          grant_type: "refresh_token",
        }),
      });

      const tokens = await tokenResponse.json();
      if (tokens.access_token) {
        // Update the database with new token
        const newExpiresAt = new Date(Date.now() + tokens.expires_in * 1000);
        await prisma.account.update({
          where: { id: account.id },
          data: {
            accessToken: tokens.access_token,
            accessTokenExpiresAt: newExpiresAt,
          },
        });

        return res.json({ accessToken: tokens.access_token });
      }
    }

    return res.json({ accessToken: account.accessToken });
  } catch (error: any) {
    console.error("Token fetch error:", error);
    return res.status(500).json({ error: error.message });
  }
});

app.put("/api/transmittals/:id", async (req, res) => {
  try {
    const session = await auth.api.getSession({
      headers: req.headers as any,
    });

    if (!session?.user) {
      return res.status(401).json({ error: "Not authenticated" });
    }

    const transmittalId = req.params.id;
    const data = req.body?.data;
    if (!data) {
      return res.status(400).json({ error: "Missing transmittal data" });
    }

    const project = {
      ...(data.project || {}),
      transmittalNumber: ensureDbTransmittalPrefix(
        data.project?.transmittalNumber,
      ),
    };

    const existing = await prisma.transmittal.findFirst({
      where: { id: transmittalId, userId: session.user.id },
    });

    if (!existing) {
      return res.status(404).json({ error: "Transmittal not found" });
    }

    const transmittal = await prisma.transmittal.update({
      where: { id: transmittalId },
      data: {
        notes: data.notes || "",
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
        agency: {
          upsert: {
            create: {
              name: data.sender?.agencyName || "",
              website: data.sender?.website || null,
              telephoneNumber: data.sender?.telephone || null,
              contactNumber: data.sender?.mobile || null,
              email: data.sender?.email || null,
            },
            update: {
              name: data.sender?.agencyName || "",
              website: data.sender?.website || null,
              telephoneNumber: data.sender?.telephone || null,
              contactNumber: data.sender?.mobile || null,
              email: data.sender?.email || null,
            },
          },
        },
      },
      include: {
        items: true,
        recipients: true,
        agency: true,
      },
    });

    return res.json({ transmittal: mapTransmittalForApi(transmittal) });
  } catch (error: any) {
    console.error("Update transmittal error:", error);
    return res.status(500).json({ error: error.message });
  }
});

app.post("/api/transmittals", async (req, res) => {
  try {
    const session = await auth.api.getSession({
      headers: req.headers as any,
    });

    if (!session?.user) {
      return res.status(401).json({ error: "Not authenticated" });
    }

    const data = req.body?.data;
    if (!data) {
      return res.status(400).json({ error: "Missing transmittal data" });
    }

    const project = {
      ...(data.project || {}),
      transmittalNumber: ensureDbTransmittalPrefix(
        data.project?.transmittalNumber,
      ),
    };

    const transmittal = await prisma.transmittal.create({
      data: {
        userId: session.user.id,
        notes: data.notes || "",
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
        agency: {
          create: {
            name: data.sender?.agencyName || "",
            website: data.sender?.website || null,
            telephoneNumber: data.sender?.telephone || null,
            contactNumber: data.sender?.mobile || null,
            email: data.sender?.email || null,
          },
        },
      },
      include: {
        items: true,
        recipients: true,
        agency: true,
      },
    });

    return res.json({ transmittal: mapTransmittalForApi(transmittal) });
  } catch (error: any) {
    console.error("Save transmittal error:", error);
    return res.status(500).json({ error: error.message });
  }
});

app.get("/api/transmittals", async (req, res) => {
  try {
    const session = await auth.api.getSession({
      headers: req.headers as any,
    });

    if (!session?.user) {
      return res.status(401).json({ error: "Not authenticated" });
    }

    const transmittals = await prisma.transmittal.findMany({
      where: { userId: session.user.id },
      include: {
        items: true,
        recipients: true,
        agency: true,
      },
    });

    return res.json({
      transmittals: transmittals.map(mapTransmittalForApi),
    });
  } catch (error: any) {
    console.error("Load transmittals error:", error);
    return res.status(500).json({ error: error.message });
  }
});

app.listen(port, () => {
  console.log(`Auth server running on http://localhost:${port}`);
});
