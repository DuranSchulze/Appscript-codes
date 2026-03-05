"use client";

import React, { useState, useRef, useEffect } from "react";
import { Pencil, Eye, Loader2 } from "lucide-react";
import { TransmittalTemplate } from "../NewReportTemplate";
import { FloatingAccount } from "../FloatingAccount";
import { BulkAddModal } from "../modals/BulkAddModal";
import { AgencyPresetModal } from "../modals/AgencyPresetModal";
import { DriveFileModal } from "../modals/DriveFileModal";
import { DocxPreviewModal } from "../modals/DocxPreviewModal";
import { ErrorBoundary } from "../ErrorBoundary";
import {
  listFilesInFolder,
  extractFolderIdFromLink,
  extractFileIdFromLink,
  isFolderLink,
  getFileMetadata,
  getFileContentAsBase64,
  listDriveFiles,
  checkDriveAccess,
  clearGoogleToken,
  uploadFileToDrive,
} from "../../services/googleDriveService";
import { useFileProcessing, resizeImage } from "../../hooks/useFileProcessing";
import { generateTransmittalDocx } from "../../services/docxGenerator";
import {
  getLinkedSheetId,
  appendTransmittalRow,
  isSheetUrl,
  readSheetRows,
  extractSheetIdFromUrl,
} from "../../services/googleSheetsService";
import { authClient, signIn, signOut, useSession } from "../../lib/auth-client";
import {
  AppData,
  TransmittalItem,
  Signatories,
  ReceivedBy,
  FooterNotes,
  SenderInfo,
} from "../../types";
import * as mammoth from "mammoth";
import {
  ParseResult,
  resolveDocumentNumberWithFallback,
} from "../../services/geminiService";

type ParsedDocumentResponse = ParseResult;

// Modular UI components
import { LoadingScreen } from "./LoadingScreen";
import { LoginScreen } from "./LoginScreen";
import { SidebarHeader } from "./SidebarHeader";
import { SidebarMenuBar } from "./SidebarFooter";
import { TabBar, type TabKey } from "./TabBar";
import { ContentTab } from "./tabs/ContentTab";
import { SenderTab } from "./tabs/SenderTab";
import { RecipientTab } from "./tabs/RecipientTab";
import { ProjectTab } from "./tabs/ProjectTab";
import { SignatoriesTab } from "./tabs/SignatoriesTab";
import { TransmittalListModal } from "../modals/TransmittalListModal";
import {
  ExportChoiceModal,
  type ExportFormat,
} from "../modals/ExportChoiceModal";
import { FolderPickerModal } from "../modals/FolderPickerModal";
import { FileUploadModal } from "../modals/FileUploadModal";
import { PreviewToolbar, ZOOM_STEPS } from "./PreviewToolbar";

// Add type declaration
declare global {
  interface Window {
    html2pdf: any;
    html2canvas: any;
  }
}

type DbAgency = {
  id: string;
  name: string;
  addressLine1: string | null;
  addressLine2: string | null;
  website: string | null;
  telephoneNumber: string | null;
  contactNumber: string | null;
  email: string | null;
  logoBase64: string | null;
  updatedAt: string;
};

const DEFAULT_COLUMN_WIDTHS = {
  qty: 55,
  noOfItems: 65,
  documentNumber: 130,
  description: 260,
  remarks: 150,
} as const;

type ColumnWidthField = keyof typeof DEFAULT_COLUMN_WIDTHS;

const COLUMN_ORDER: ColumnWidthField[] = [
  "noOfItems",
  "qty",
  "documentNumber",
  "description",
  "remarks",
];

const COLUMN_WIDTH_LIMITS: Record<ColumnWidthField, { min: number; max: number }> =
  {
    noOfItems: { min: 55, max: 140 },
    qty: { min: 55, max: 120 },
    documentNumber: { min: 100, max: 260 },
    description: { min: 160, max: 360 },
    remarks: { min: 120, max: 260 },
  };

const createInitialData = (): AppData => ({
  recipient: {
    to: "",
    email: "",
    company: "",
    attention: "",
    address: "",
    contactNumber: "",
  },
  project: {
    projectName: "",
    projectNumber: "",
    engagementRef: "",
    purpose: "",
    transmittalNumber: "",
    department: "Admin",
    date: new Date().toISOString().split("T")[0],
    timeGenerated: new Date().toLocaleTimeString("en-US", {
      hour: "2-digit",
      minute: "2-digit",
    }),
  },
  sender: {
    agencyName: "FILEPINO",
    addressLine1: "Unit 1212 High Street South Corporate Plaza Tower 2",
    addressLine2: "26th Street Bonifacio Global City, Taguig City 1634",
    website: "www.filepino.com",
    mobile: "0917 892 2337",
    telephone: "(028) 372 5023 • (02) 8478-5826",
    email: "info@filepino.com",
    logoBase64: null,
  },
  signatories: {
    preparedBy: "Admin Staff",
    preparedByRole: "Admin Staff",
    notedBy: "Operations Manager",
    notedByRole: "Operations Manager",
    timeReleased: new Date().toLocaleTimeString("en-US", {
      hour: "2-digit",
      minute: "2-digit",
    }),
  },
  transmissionMethod: {
    personalDelivery: false,
    pickUp: false,
    grabLalamove: false,
    registeredMail: false,
  },
  receivedBy: { name: "", date: "", time: "", remarks: "" },
  footerNotes: {
    acknowledgement:
      "This is to acknowledge and confirm that the items/documents listed above are complete and in good condition.",
    disclaimer:
      "For documentation purposes, please return the signed transmittal form to our office via email or courier at your earliest convenience.",
  },
  notes: "",
  agencyId: null,
  items: [],
});

const createEmptySenderDraft = (): SenderInfo => ({
  agencyName: "",
  addressLine1: "",
  addressLine2: "",
  website: "",
  mobile: "",
  telephone: "",
  email: "",
  logoBase64: null,
});

const stripFileExtension = (fileName: string): string =>
  fileName.replace(/\.[^/.]+$/, "").trim();

const DRIVE_IMPORT_REMARK = "Imported from Google Drive";
const DRIVE_ANALYZABLE_EXTENSION_PATTERN =
  /\.(pdf|png|jpe?g|webp|gif|bmp|tiff?)$/i;

type DriveFileMeta = {
  id: string;
  name: string;
  mimeType: string;
};

const toDriveFileSource = (fileId: string): string =>
  `https://drive.google.com/file/d/${fileId}/view`;

const isDriveFileAnalyzable = (file: DriveFileMeta): boolean => {
  const mimeType = String(file.mimeType || "").toLowerCase();
  if (mimeType === "application/pdf") return true;
  if (mimeType.startsWith("image/")) return true;
  return DRIVE_ANALYZABLE_EXTENSION_PATTERN.test(file.name || "");
};

const normalizeAutoDocumentNumber = (value: string): string =>
  value
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .replace(/-{2,}/g, "-");

const createDriveDocumentNumber = (file: {
  id: string;
  name: string;
}): string => {
  const resolved = normalizeAutoDocumentNumber(
    resolveDocumentNumberWithFallback({
      sourceName: file.name,
      description: stripFileExtension(file.name) || file.name,
      documentType: "File",
    }),
  );
  if (resolved) return resolved;

  const baseName = normalizeAutoDocumentNumber(stripFileExtension(file.name));
  const idSuffix = normalizeAutoDocumentNumber(file.id).slice(-4) || "0000";

  if (baseName) {
    const trimmedBase = baseName.slice(0, 16).replace(/-+$/g, "");
    return `DRV-${trimmedBase}-${idSuffix}`;
  }

  return `DRV-${idSuffix}`;
};

const resolveSessionErrorNotice = (error: unknown): string | undefined => {
  if (!error || typeof error !== "object") {
    return undefined;
  }

  const typedError = error as {
    status?: number;
    statusText?: string;
    message?: string;
    error?: { message?: string; code?: string } | string;
  };

  const nestedError =
    typeof typedError.error === "string"
      ? typedError.error
      : `${typedError.error?.message || ""} ${typedError.error?.code || ""}`.trim();

  const details = [typedError.message, typedError.statusText, nestedError]
    .filter(Boolean)
    .join(" ")
    .toLowerCase();

  const status = Number(typedError.status || 0);
  const likelyConnectivityIssue =
    status >= 500 ||
    /unable to connect|network|dns|accelerate|failed to fetch|timeout|please_restart_the_process|internal_server_error/.test(
      details,
    );

  if (likelyConnectivityIssue) {
    return "Session check failed due to a temporary connection issue (internet, DNS, or server). Please check your connection and sign in again.";
  }

  if (status === 401 || status === 403) {
    return "Your session has expired. Please sign in again.";
  }

  return "We could not restore your session. Please sign in again.";
};

const parseTransmittalViaApi = async (
  apiBaseUrl: string,
  input: {
    content: string;
    mimeType: string;
    isText?: boolean;
    fileName?: string;
  },
): Promise<ParseResult> => {
  const response = await fetch(`${apiBaseUrl}/api/parse-transmittal`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    credentials: "include",
    body: JSON.stringify(input),
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload?.error || "Failed to analyze document.");
  }

  return payload as ParseResult;
};

const AppContent: React.FC = () => {
  const [smartInput, setSmartInput] = useState("");
  const [isAnalyzingText, setIsAnalyzingText] = useState(false);
  const [isGeneratingPdf, setIsGeneratingPdf] = useState(false);
  const [isGeneratingDocx, setIsGeneratingDocx] = useState(false);
  const [statusMsg, setStatusMsg] = useState("");
  const [statusType, setStatusType] = useState<"info" | "error">("info");
  const [isBulkModalOpen, setIsBulkModalOpen] = useState(false);
  const [isPreviewModalOpen, setIsPreviewModalOpen] = useState(false);
  const [docxPreviewHtml, setDocxPreviewHtml] = useState("");
  const [isDriveModalOpen, setIsDriveModalOpen] = useState(false);
  const [driveFiles, setDriveFiles] = useState<DriveFileMeta[]>([]);
  const [driveSearch, setDriveSearch] = useState("");
  const [driveSelected, setDriveSelected] = useState<Record<string, boolean>>(
    {},
  );
  const [driveError, setDriveError] = useState("");
  const [isDriveLoading, setIsDriveLoading] = useState(false);
  const [isBulkImporting, setIsBulkImporting] = useState(false);
  const [isDriveSelectionImporting, setIsDriveSelectionImporting] =
    useState(false);
  const [activeTab, setActiveTab] = useState<TabKey>("content");
  const [isTransmittalListOpen, setIsTransmittalListOpen] = useState(false);
  const [isFileUploadOpen, setIsFileUploadOpen] = useState(false);
  const [exportChoiceOpen, setExportChoiceOpen] = useState(false);
  const [exportFolderPickerOpen, setExportFolderPickerOpen] = useState(false);
  const [pendingExportBlob, setPendingExportBlob] = useState<Blob | null>(null);
  const [pendingExportFormat, setPendingExportFormat] =
    useState<ExportFormat>("pdf");
  const [pendingExportFileName, setPendingExportFileName] = useState("");
  const [isUploadingToDrive, setIsUploadingToDrive] = useState(false);
  const [data, setData] = useState<AppData>(() => createInitialData());
  const [activeTransmittalId, setActiveTransmittalId] = useState<string | null>(
    null,
  );
  const [showPreview, setShowPreview] = useState(false);

  const [columnWidths, setColumnWidths] = useState({ ...DEFAULT_COLUMN_WIDTHS });
  const [isDriveReady, setIsDriveReady] = useState(false);
  const [zoomPercent, setZoomPercent] = useState(100);
  const [autoFitZoom, setAutoFitZoom] = useState(100);
  const [isManualZoom, setIsManualZoom] = useState(false);
  const previewContainerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleResize = () => {
      if (previewContainerRef.current) {
        const containerWidth = previewContainerRef.current.offsetWidth;
        const targetWidth = 850;
        let fitPercent = 100;
        if (containerWidth < targetWidth) {
          fitPercent = Math.round(((containerWidth - 32) / targetWidth) * 100);
          fitPercent = Math.max(25, Math.min(100, fitPercent));
        }
        setAutoFitZoom(fitPercent);
        if (!isManualZoom) {
          setZoomPercent(fitPercent);
        }
      }
    };

    handleResize();
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, [showPreview, isManualZoom]);

  const previewScale = zoomPercent / 100;

  const handleZoomIn = () => {
    setIsManualZoom(true);
    setZoomPercent((prev) => {
      const next = ZOOM_STEPS.find((s) => s > prev);
      return next ?? prev;
    });
  };

  const handleZoomOut = () => {
    setIsManualZoom(true);
    setZoomPercent((prev) => {
      const next = [...ZOOM_STEPS].reverse().find((s) => s < prev);
      return next ?? prev;
    });
  };

  const handleZoomReset = () => {
    setIsManualZoom(false);
    setZoomPercent(autoFitZoom);
  };

  const handleZoomSet = (percent: number) => {
    setIsManualZoom(true);
    setZoomPercent(percent);
  };

  const { data: session, isPending, error: sessionError } = useSession();
  const apiBaseUrl =
    process.env.NEXT_PUBLIC_BETTER_AUTH_URL ||
    (typeof window !== "undefined" ? window.location.origin : "");
  const sessionErrorNotice = resolveSessionErrorNotice(sessionError);

  const handleGoogleSignIn = async () => {
    await authClient.signIn.oauth2({
      providerId: "google",
      callbackURL: window.location.origin,
    });
  };

  const handleDDSSignIn = async () => {
    await authClient.signIn.oauth2({
      providerId: "google-dds",
      callbackURL: window.location.origin,
    });
  };

  const handleSignOut = async () => {
    await clearGoogleToken();
    await signOut();
  };

  const fileInputRef = useRef<HTMLInputElement>(null);
  const printRef = useRef<HTMLDivElement>(null);

  const [agencies, setAgencies] = useState<DbAgency[]>([]);
  const [selectedAgencyId, setSelectedAgencyId] = useState<string>("");
  const [isAgencyModalOpen, setIsAgencyModalOpen] = useState(false);
  const [agencyDraft, setAgencyDraft] = useState<SenderInfo>(
    () => createInitialData().sender,
  );
  const [agencyDriveLogoLink, setAgencyDriveLogoLink] = useState("");
  const [isImportingAgencyDriveLogo, setIsImportingAgencyDriveLogo] =
    useState(false);
  const [agencyModalMode, setAgencyModalMode] = useState<"create" | "update">(
    "create",
  );

  const getFileTimestamp = () =>
    new Date()
      .toISOString()
      .replace(/T/, "_")
      .replace(/\..+/, "")
      .replace(/:/g, "-");

  // Check Drive access when user is logged in
  useEffect(() => {
    if (session?.user) {
      checkDriveAccess().then(setIsDriveReady);
    }
  }, [session]);

  const fetchNextTransmittalNumber = async (force = false) => {
    if (!apiBaseUrl) return;
    try {
      const response = await fetch(
        `${apiBaseUrl}/api/transmittals/next-number`,
        {
          credentials: "include",
        },
      );
      if (!response.ok) return;
      const payload = await response.json().catch(() => ({}));
      if (payload?.transmittalNumber) {
        setData((prev) => {
          if (!force && prev.project.transmittalNumber) return prev;
          return {
            ...prev,
            project: {
              ...prev.project,
              transmittalNumber: payload.transmittalNumber,
            },
          };
        });
      }
    } catch (error) {
      console.error("Failed to fetch next transmittal number", error);
    }
  };

  useEffect(() => {
    if (session?.user) {
      fetchNextTransmittalNumber();
    }
  }, [session?.user?.id]);

  const loadAgenciesFromDb = async () => {
    if (!session?.user || !apiBaseUrl) return;
    try {
      const response = await fetch(`${apiBaseUrl}/api/agencies`, {
        credentials: "include",
      });
      if (!response.ok) return;
      const payload = await response.json().catch(() => ({}));
      const list = Array.isArray(payload.agencies)
        ? (payload.agencies as DbAgency[])
        : [];
      setAgencies(list);
    } catch (error) {
      console.error("Failed to load agencies", error);
    }
  };

  useEffect(() => {
    loadAgenciesFromDb();
  }, [session?.user?.id]);

  useEffect(() => {
    const id = data.agencyId ? String(data.agencyId) : "";
    if (id) setSelectedAgencyId(id);
  }, [data.agencyId]);

  const mapDbAgencyToSender = (agency: DbAgency): SenderInfo => ({
    agencyName: agency.name || "",
    addressLine1: agency.addressLine1 || "",
    addressLine2: agency.addressLine2 || "",
    website: agency.website || "",
    mobile: agency.contactNumber || "",
    telephone: agency.telephoneNumber || "",
    email: agency.email || "",
    logoBase64: agency.logoBase64 || null,
  });

  useEffect(() => {
    if (!selectedAgencyId) return;
    const agency = agencies.find((a) => a.id === selectedAgencyId);
    if (!agency) return;
    setData((prev) => ({
      ...prev,
      agencyId: agency.id,
      sender: mapDbAgencyToSender(agency),
    }));
  }, [selectedAgencyId, agencies]);

  const mapDbTransmittalToAppData = (transmittal: any): AppData => {
    const projectData = transmittal.project || {};
    const senderData = transmittal.sender || {};
    const receivedBy = transmittal.receivedBy || {};
    const footerNotes = transmittal.footerNotes || {};
    const primaryRecipient = transmittal.recipients?.[0];
    const stripPrefix = (value: string) =>
      value.startsWith("TR-FP-") ? value.slice("TR-FP-".length) : value;

    return {
      agencyId: transmittal.agencyId || transmittal.agency?.id || null,
      recipient: {
        to: primaryRecipient?.recipientName || "",
        email: primaryRecipient?.recipientAgencyEmail || "",
        company: primaryRecipient?.recipientOrganization || "",
        attention: primaryRecipient?.recipientAttention || "",
        address: primaryRecipient?.recipientFullAddress || "",
        contactNumber: primaryRecipient?.recipientAgencyContactNumber || "",
      },
      project: {
        projectName: projectData.projectName || transmittal.projectName || "",
        projectNumber:
          projectData.projectNumber || transmittal.projectNumber || "",
        engagementRef:
          projectData.engagementRef || transmittal.engagementRefNumber || "",
        purpose: projectData.purpose || transmittal.projectPurpose || "",
        transmittalNumber: stripPrefix(
          String(projectData.transmittalNumber || ""),
        ),
        department: projectData.department || transmittal.department || "",
        date: projectData.date || "",
        timeGenerated: projectData.timeGenerated || "",
      },
      items: Array.isArray(transmittal.items)
        ? transmittal.items.map((item: any) => ({
            id: item.id,
            qty: item.qty || "",
            noOfItems: item.noOfItems || "",
            documentNumber: resolveDocumentNumberWithFallback({
              currentDocumentNumber: item.documentNumber || "",
              sourceName: item.description || "",
              description: item.description || "",
              documentType: item.documentType || "",
            }),
            description: item.description || "",
            remarks: item.remarks || "",
            fileType: item.fileType || undefined,
            fileSource: item.fileSource || undefined,
          }))
        : [],
      sender: {
        agencyName: senderData.agencyName || "",
        addressLine1: senderData.addressLine1 || "",
        addressLine2: senderData.addressLine2 || "",
        website: senderData.website || "",
        mobile: senderData.mobile || "",
        telephone: senderData.telephone || "",
        email: senderData.email || "",
        logoBase64: senderData.logoBase64 || null,
      },
      signatories: {
        preparedBy: transmittal.preparedBy || "",
        preparedByRole: transmittal.preparedByRole || "",
        notedBy: transmittal.notedBy || "",
        notedByRole: transmittal.notedByRole || "",
        timeReleased: transmittal.timeReleased || "",
      },
      receivedBy: {
        name: receivedBy.name || "",
        date: receivedBy.date || "",
        time: receivedBy.time || "",
        remarks: receivedBy.remarks || "",
      },
      footerNotes: {
        acknowledgement: footerNotes.acknowledgement || "",
        disclaimer: footerNotes.disclaimer || "",
      },
      notes: transmittal.notes || "",
      transmissionMethod: {
        personalDelivery: Boolean(transmittal.handDelivery),
        pickUp: Boolean(transmittal.pickUp),
        grabLalamove: Boolean(transmittal.courier),
        registeredMail: Boolean(transmittal.registeredMail),
      },
    };
  };

  const saveTransmittalToDb = async () => {
    if (!session?.user || !apiBaseUrl) return;
    const isEditing = Boolean(activeTransmittalId);
    const url = isEditing
      ? `${apiBaseUrl}/api/transmittals/${activeTransmittalId}`
      : `${apiBaseUrl}/api/transmittals`;
    const response = await fetch(url, {
      method: isEditing ? "PUT" : "POST",
      headers: { "Content-Type": "application/json" },
      credentials: "include",
      body: JSON.stringify({ data }),
    });

    if (!response.ok) {
      const payload = await response.json().catch(() => ({}));
      throw new Error(payload.error || "Failed to save transmittal");
    }

    const payload = await response.json().catch(() => ({}));
    if (payload?.transmittal) {
      setActiveTransmittalId(payload.transmittal.id);

      // Sync the form with the server-assigned transmittal number
      const serverProject = payload.transmittal.project || {};
      const serverNumber = String(serverProject.transmittalNumber || "");
      if (serverNumber) {
        setData((prev) => ({
          ...prev,
          project: { ...prev.project, transmittalNumber: serverNumber },
        }));
      }

      // Append row to linked Google Sheet (non-blocking, only on create)
      if (!isEditing && getLinkedSheetId()) {
        const methods: string[] = [];
        if (data.transmissionMethod?.personalDelivery)
          methods.push("Hand Delivery");
        if (data.transmissionMethod?.pickUp) methods.push("Pick Up");
        if (data.transmissionMethod?.grabLalamove) methods.push("Courier");
        if (data.transmissionMethod?.registeredMail)
          methods.push("Registered Mail");

        appendTransmittalRow({
          transmittalNumber: serverNumber || data.project.transmittalNumber,
          date: data.project.date,
          projectName: data.project.projectName,
          recipientName: data.recipient.to,
          recipientCompany: data.recipient.company,
          preparedBy: data.signatories.preparedBy,
          notedBy: data.signatories.notedBy,
          itemsCount: data.items.length,
          transmissionMethod: methods.join(", ") || "—",
        }).catch(() => {});
      }
    }
  };

  const handleOpenTransmittal = async (id: string) => {
    if (!apiBaseUrl) return;
    try {
      const response = await fetch(`${apiBaseUrl}/api/transmittals`, {
        credentials: "include",
      });
      if (!response.ok) throw new Error("Failed to fetch transmittals");
      const payload = await response.json();
      const transmittals = Array.isArray(payload.transmittals)
        ? payload.transmittals
        : [];
      const match = transmittals.find((t: any) => t.id === id);
      if (!match) throw new Error("Transmittal not found");

      const appData = mapDbTransmittalToAppData(match);
      setData(appData);
      setActiveTransmittalId(id);
      setActiveTab("content");
      setStatusMsg("Transmittal loaded");
      setStatusType("info");
      setTimeout(() => setStatusMsg(""), 3000);
    } catch (e: any) {
      setStatusMsg(e.message || "Failed to open transmittal");
      setStatusType("error");
      setTimeout(() => setStatusMsg(""), 5000);
    }
  };

  const handleDeleteTransmittal = async (id: string) => {
    if (!apiBaseUrl) return;
    const response = await fetch(`${apiBaseUrl}/api/transmittals/${id}`, {
      method: "DELETE",
      credentials: "include",
    });
    if (!response.ok) {
      const payload = await response.json().catch(() => ({}));
      throw new Error(payload.error || "Failed to delete");
    }
    // If the deleted transmittal is currently active, clear the workspace
    if (activeTransmittalId === id) {
      setData(createInitialData());
      setActiveTransmittalId(null);
      fetchNextTransmittalNumber(true);
    }
  };

  const resetForNewAnalysis = () => {
    setData(createInitialData());
    setActiveTransmittalId(null);
    fetchNextTransmittalNumber(true);
    setStatusMsg("Form Reset");
    setTimeout(() => setStatusMsg(""), 3000);
  };

  const updateField = (section: keyof AppData, field: string, value: any) => {
    setData((prev) => ({
      ...prev,
      [section]: { ...(prev[section] as any), [field]: value },
    }));
  };

  const handleUpdateNotes = (value: string) =>
    setData((prev) => ({ ...prev, notes: value }));
  const handleUpdateSignatory = (field: keyof Signatories, value: string) =>
    setData((prev) => ({
      ...prev,
      signatories: { ...prev.signatories, [field]: value },
    }));
  const handleUpdateReceivedBy = (field: keyof ReceivedBy, value: string) =>
    setData((prev) => ({
      ...prev,
      receivedBy: { ...prev.receivedBy, [field]: value },
    }));
  const handleUpdateFooter = (field: keyof FooterNotes, value: string) =>
    setData((prev) => ({
      ...prev,
      footerNotes: { ...prev.footerNotes, [field]: value },
    }));
  const updateTransmission = (method: string, checked: boolean) =>
    setData((prev) => ({
      ...prev,
      transmissionMethod: { ...prev.transmissionMethod, [method]: checked },
    }));

  const reindexItems = (items: TransmittalItem[]): TransmittalItem[] =>
    items.map((item, index) => ({
      ...item,
      noOfItems: (index + 1).toString(),
    }));
  const toDriveItem = (
    file: Pick<DriveFileMeta, "id" | "name">,
    remarks: string = DRIVE_IMPORT_REMARK,
  ): TransmittalItem => ({
    id: file.id,
    qty: "1",
    noOfItems: "1",
    documentNumber: createDriveDocumentNumber(file),
    description: stripFileExtension(file.name),
    remarks,
    fileType: "gdrive",
    fileSource: toDriveFileSource(file.id),
  });
  const addItems = (newItems: TransmittalItem[]) => {
    setData((prev) => ({
      ...prev,
      items: reindexItems([...prev.items, ...newItems]),
    }));
  };

  const mapAnalyzedDriveItems = (
    file: DriveFileMeta,
    parsedItems: ParsedDocumentResponse["items"],
    forceDefaultRemarks = false,
  ): TransmittalItem[] => {
    const fallbackDescription = stripFileExtension(file.name) || file.name;
    return parsedItems.map((res, index) => {
      const resolvedDocumentNumber = resolveDocumentNumberWithFallback({
        currentDocumentNumber: res.documentNumber || "",
        sourceName: file.name || "",
        description: res.description || fallbackDescription,
        documentType: res.documentType || "File",
      });
      return {
        id: `${file.id}-${Date.now()}-${index}-${Math.random().toString(36).slice(2, 8)}`,
        qty: String(res.qty || "1").trim() || "1",
        noOfItems: "1",
        documentNumber:
          resolvedDocumentNumber || createDriveDocumentNumber(file),
        description: (
          res.description ||
          fallbackDescription ||
          `File ${index + 1}`
        ).trim(),
        remarks: forceDefaultRemarks
          ? DRIVE_IMPORT_REMARK
          : (res.remarks || DRIVE_IMPORT_REMARK).trim() || DRIVE_IMPORT_REMARK,
        fileType: "gdrive",
        fileSource: toDriveFileSource(file.id),
      };
    });
  };

  const analyzeDriveFileForItems = async (
    file: DriveFileMeta,
  ): Promise<{ items: TransmittalItem[]; usedFallback: boolean }> => {
    if (!isDriveFileAnalyzable(file)) {
      return { items: [toDriveItem(file)], usedFallback: true };
    }

    try {
      const { base64, mimeType } = await getFileContentAsBase64(file.id);
      const result = await parseTransmittalViaApi(apiBaseUrl, {
        content: base64,
        mimeType: mimeType || file.mimeType || "application/pdf",
        isText: false,
        fileName: file.name,
      });

      const parsedItems = Array.isArray(result.items) ? result.items : [];
      if (parsedItems.length === 0) {
        return { items: [toDriveItem(file)], usedFallback: true };
      }

      const usedFallback =
        Boolean(result.error) || Number(result.fallbackCount || 0) > 0;
      return {
        items: mapAnalyzedDriveItems(file, parsedItems, usedFallback),
        usedFallback,
      };
    } catch {
      return { items: [toDriveItem(file)], usedFallback: true };
    }
  };

  const importDriveFilesWithAi = async (
    files: DriveFileMeta[],
    progressLabel: string,
  ): Promise<{ addedCount: number; fallbackCount: number }> => {
    if (files.length === 0) {
      return { addedCount: 0, fallbackCount: 0 };
    }

    const importedItems: TransmittalItem[] = [];
    let fallbackCount = 0;

    for (let index = 0; index < files.length; index += 1) {
      const file = files[index];
      setStatusMsg(
        `${progressLabel}: analyzing file ${index + 1}/${files.length} (${file.name})...`,
      );
      setStatusType("info");

      const { items, usedFallback } = await analyzeDriveFileForItems(file);
      importedItems.push(...items);
      if (usedFallback) fallbackCount += 1;
    }

    if (importedItems.length > 0) {
      addItems(importedItems);
    }

    return { addedCount: importedItems.length, fallbackCount };
  };

  const mergeHeaderData = (header: any) => {
    if (!header) return;
    setData((prev) => ({
      ...prev,
      recipient: {
        ...prev.recipient,
        to: header.recipientName || prev.recipient.to,
        email: header.recipientEmail || prev.recipient.email,
        company: header.companyName || prev.recipient.company,
        address: header.address || prev.recipient.address,
      },
      project: {
        ...prev.project,
        projectName: header.projectName || prev.project.projectName,
        projectNumber: header.projectNumber || prev.project.projectNumber,
        purpose: header.purpose || prev.project.purpose,
      },
    }));
  };

  const handleManualAdd = () =>
    addItems([
      {
        id: Date.now().toString() + Math.random(),
        qty: "1",
        noOfItems: "0",
        documentNumber: "",
        description: "",
        remarks: "",
        fileType: "link",
      },
    ]);

  const handleBulkImportDriveLink = async (link: string) => {
    const folderId = extractFolderIdFromLink(link);
    if (!folderId) {
      const msg = "Invalid Google Drive folder link.";
      setStatusMsg(msg);
      setStatusType("error");
      setTimeout(() => setStatusMsg(""), 5000);
      throw new Error(msg);
    }

    if (!isDriveReady) {
      const msg = "Drive access not available. Please sign in again.";
      setStatusMsg(msg);
      setStatusType("error");
      setTimeout(() => setStatusMsg(""), 5000);
      throw new Error(msg);
    }

    try {
      setIsBulkImporting(true);
      setStatusMsg("Listing Drive files...");
      setStatusType("info");
      const files = await listFilesInFolder(folderId);
      if (files.length === 0) {
        setStatusMsg("No files found in the selected Drive folder.");
        setStatusType("info");
        setTimeout(() => setStatusMsg(""), 5000);
        return;
      }

      const { addedCount, fallbackCount } = await importDriveFilesWithAi(
        files,
        "Drive bulk import",
      );

      setStatusMsg(
        fallbackCount > 0
          ? `Imported ${addedCount} item(s) from ${files.length} file(s). Document # fallback was used for ${fallbackCount} file(s).`
          : `Imported ${addedCount} item(s) from ${files.length} file(s) using AI extraction.`,
      );
      setStatusType("info");
      setTimeout(() => setStatusMsg(""), 5000);
    } catch (e: any) {
      setStatusMsg("Error: " + e.message);
      setStatusType("error");
      setTimeout(() => setStatusMsg(""), 5000);
      throw e;
    } finally {
      setIsBulkImporting(false);
    }
  };

  const updateItem = (
    index: number,
    field: keyof TransmittalItem,
    value: string,
  ) => {
    setData((prev) => {
      const newItems = [...prev.items];
      newItems[index] = { ...newItems[index], [field]: value };
      return { ...prev, items: newItems };
    });
  };

  const adjustItemQty = (index: number, delta: number) => {
    setData((prev) => {
      const newItems = [...prev.items];
      const item = newItems[index];
      if (!item) return prev;

      const currentQty = Number.parseInt(String(item.qty || "").trim(), 10);
      const baseQty =
        Number.isFinite(currentQty) && currentQty > 0 ? currentQty : 1;
      const nextQty = Math.max(1, baseQty + delta);

      newItems[index] = {
        ...item,
        qty: String(nextQty),
      };

      return { ...prev, items: newItems };
    });
  };

  const removeItem = (index: number) => {
    setData((prev) => ({
      ...prev,
      items: reindexItems(prev.items.filter((_, i) => i !== index)),
    }));
  };
  const moveItem = (index: number, direction: "up" | "down") => {
    setData((prev) => {
      const newItems = [...prev.items];
      if (direction === "up" && index > 0)
        [newItems[index], newItems[index - 1]] = [
          newItems[index - 1],
          newItems[index],
        ];
      else if (direction === "down" && index < newItems.length - 1)
        [newItems[index], newItems[index + 1]] = [
          newItems[index + 1],
          newItems[index],
        ];
      return { ...prev, items: reindexItems(newItems) };
    });
  };

  const handleReorderItems = (fromIndex: number, toIndex: number) => {
    setData((prev) => {
      const newItems = [...prev.items];
      const [movedItem] = newItems.splice(fromIndex, 1);
      newItems.splice(toIndex, 0, movedItem);
      return { ...prev, items: reindexItems(newItems) };
    });
  };

  const handleResizeDivider = (leftField: ColumnWidthField, deltaX: number) => {
    if (!Number.isFinite(deltaX) || deltaX === 0) return;

    setColumnWidths((prev) => {
      const leftIndex = COLUMN_ORDER.indexOf(leftField);
      if (leftIndex < 0 || leftIndex >= COLUMN_ORDER.length - 1) {
        return prev;
      }

      const rightField = COLUMN_ORDER[leftIndex + 1];

      const leftCurrent = prev[leftField];
      const rightCurrent = prev[rightField];

      const leftLimits = COLUMN_WIDTH_LIMITS[leftField];
      const rightLimits = COLUMN_WIDTH_LIMITS[rightField];

      const maxGrowLeft = Math.min(
        leftLimits.max - leftCurrent,
        rightCurrent - rightLimits.min,
      );
      const maxShrinkLeft = Math.min(
        leftCurrent - leftLimits.min,
        rightLimits.max - rightCurrent,
      );

      let effectiveDelta = deltaX;
      if (deltaX > 0) {
        effectiveDelta = Math.min(deltaX, maxGrowLeft);
      } else {
        effectiveDelta = -Math.min(-deltaX, maxShrinkLeft);
      }

      if (!Number.isFinite(effectiveDelta) || Math.abs(effectiveDelta) < 0.01) {
        return prev;
      }

      const nextLeft = Math.round(leftCurrent + effectiveDelta);
      const nextRight = Math.round(rightCurrent - effectiveDelta);

      if (nextLeft === leftCurrent && nextRight === rightCurrent) {
        return prev;
      }

      return {
        ...prev,
        [leftField]: nextLeft,
        [rightField]: nextRight,
      };
    });
  };

  const handleResetColumnWidths = () =>
    setColumnWidths({ ...DEFAULT_COLUMN_WIDTHS });

  const {
    processFiles: processDocs,
    isProcessing: isParsing,
    progress: parseProgress,
    error: processingError,
  } = useFileProcessing<any>(
    async (base64, mimeType, fileName) => {
      const result = await parseTransmittalViaApi(apiBaseUrl, {
        content: base64,
        mimeType,
        isText: false,
        fileName,
      });
      const hasParsedItems =
        Array.isArray(result.items) && result.items.length > 0;
      const usedFallback =
        Boolean((result as any).error) ||
        Number((result as any).fallbackCount || 0) > 0;

      if ((result as any).error && !hasParsedItems) {
        throw new Error((result as any).error);
      }

      return {
        items: result.items.map((res) => ({
          id: Date.now().toString() + Math.random(),
          qty: res.qty || "1",
          noOfItems: "1",
          documentNumber: resolveDocumentNumberWithFallback({
            currentDocumentNumber: res.documentNumber || "",
            sourceName: fileName || "",
            description:
              res.description ||
              stripFileExtension(fileName || "") ||
              fileName ||
              "",
            documentType: res.documentType || "File",
          }),
          description: res.description,
          remarks: res.remarks || "",
          fileType: "upload",
        })),
        header: result.header,
        usedFallback,
      };
    },
    ["application/pdf", "image/*"],
  );

  useEffect(() => {
    if (processingError) {
      setStatusMsg(processingError);
      setStatusType("error");
      setTimeout(() => setStatusMsg(""), 5000);
    }
  }, [processingError]);

  const handleAgencyLogoUpload = async (
    e: React.ChangeEvent<HTMLInputElement>,
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const base64 = await resizeImage(file, 400);
      setAgencyDraft((prev) => ({
        ...prev,
        logoBase64: `data:image/jpeg;base64,${base64}`,
      }));
    } catch (err: any) {
      setStatusMsg(`Logo Error: ${err.message}`);
      setStatusType("error");
      setTimeout(() => setStatusMsg(""), 3000);
    } finally {
      e.target.value = "";
    }
  };

  const handleAgencyDriveLogoImport = async () => {
    const link = agencyDriveLogoLink.trim();
    if (!link) {
      setStatusMsg("Paste a Google Drive file link first.");
      setStatusType("error");
      setTimeout(() => setStatusMsg(""), 3000);
      return;
    }

    if (!isDriveReady) {
      setStatusMsg("Google Drive is not connected. Please sign in again.");
      setStatusType("error");
      setTimeout(() => setStatusMsg(""), 3000);
      return;
    }

    const fileId = extractFileIdFromLink(link);
    if (!fileId) {
      setStatusMsg("Invalid Google Drive file link.");
      setStatusType("error");
      setTimeout(() => setStatusMsg(""), 3000);
      return;
    }

    setIsImportingAgencyDriveLogo(true);
    try {
      const meta = await getFileMetadata(fileId);
      const metaMimeType = String(meta.mimeType || "");
      if (!metaMimeType.startsWith("image/")) {
        throw new Error("Selected Google Drive file is not an image.");
      }

      const { base64, mimeType } = await getFileContentAsBase64(fileId);
      const effectiveMimeType = mimeType?.startsWith("image/")
        ? mimeType
        : metaMimeType;

      setAgencyDraft((prev) => ({
        ...prev,
        logoBase64: `data:${effectiveMimeType};base64,${base64}`,
      }));
      setAgencyDriveLogoLink("");
      setStatusMsg(`Logo imported from Google Drive (${meta.name}).`);
      setStatusType("info");
      setTimeout(() => setStatusMsg(""), 3000);
    } catch (error: any) {
      setStatusMsg(
        error?.message || "Failed to import logo from Google Drive.",
      );
      setStatusType("error");
      setTimeout(() => setStatusMsg(""), 4000);
    } finally {
      setIsImportingAgencyDriveLogo(false);
    }
  };

  const openAgencyModal = (mode: "create" | "update") => {
    setAgencyDriveLogoLink("");

    if (mode === "update") {
      const agency = agencies.find((item) => item.id === selectedAgencyId);
      if (!agency) {
        setStatusMsg("Select an agency first to update.");
        setStatusType("error");
        setTimeout(() => setStatusMsg(""), 3000);
        return;
      }
      setAgencyDraft(mapDbAgencyToSender(agency));
      setAgencyModalMode("update");
    } else {
      setAgencyDraft(createEmptySenderDraft());
      setAgencyModalMode("create");
    }

    setIsAgencyModalOpen(true);
  };

  const saveAgencyPreset = async () => {
    if (!session?.user || !apiBaseUrl) return;
    const name = agencyDraft.agencyName.trim();
    if (!name) return;
    const isUpdating =
      agencyModalMode === "update" && Boolean(selectedAgencyId);
    const endpoint = isUpdating
      ? `${apiBaseUrl}/api/agencies/${selectedAgencyId}`
      : `${apiBaseUrl}/api/agencies`;

    try {
      const response = await fetch(endpoint, {
        method: isUpdating ? "PUT" : "POST",
        headers: { "Content-Type": "application/json" },
        credentials: "include",
        body: JSON.stringify({
          agency: {
            name,
            addressLine1: agencyDraft.addressLine1 || null,
            addressLine2: agencyDraft.addressLine2 || null,
            website: agencyDraft.website || null,
            telephoneNumber: agencyDraft.telephone || null,
            contactNumber: agencyDraft.mobile || null,
            email: agencyDraft.email || null,
            logoBase64: agencyDraft.logoBase64 || null,
          },
        }),
      });

      if (!response.ok) {
        const payload = await response.json().catch(() => ({}));
        throw new Error(payload.error || "Failed to save agency");
      }

      const payload = await response.json().catch(() => ({}));
      const saved = payload?.agency as DbAgency | undefined;
      if (saved?.id) {
        await loadAgenciesFromDb();
        setSelectedAgencyId(saved.id);
        setData((prev) => ({
          ...prev,
          agencyId: saved.id,
          sender: mapDbAgencyToSender(saved),
        }));
        setStatusMsg(isUpdating ? "Agency updated" : "Agency added");
        setStatusType("info");
        setTimeout(() => setStatusMsg(""), 3000);
      }

      setIsAgencyModalOpen(false);
    } catch (error: any) {
      console.error("Save agency failed:", error);
      setStatusMsg(error.message || "Failed to save agency");
      setStatusType("error");
      setTimeout(() => setStatusMsg(""), 3000);
    }
  };

  const handleDeleteAgency = async () => {
    if (!session?.user || !apiBaseUrl || !selectedAgencyId) return;

    const selectedAgency = agencies.find(
      (item) => item.id === selectedAgencyId,
    );
    const agencyName = selectedAgency?.name || "this agency";
    const shouldDelete = window.confirm(
      `Delete \"${agencyName}\" from saved agencies?`,
    );
    if (!shouldDelete) return;

    try {
      const response = await fetch(
        `${apiBaseUrl}/api/agencies/${selectedAgencyId}`,
        {
          method: "DELETE",
          credentials: "include",
        },
      );

      if (!response.ok) {
        const payload = await response.json().catch(() => ({}));
        throw new Error(payload.error || "Failed to delete agency");
      }

      await loadAgenciesFromDb();
      setSelectedAgencyId("");
      setData((prev) => ({ ...prev, agencyId: null }));
      setStatusMsg("Agency deleted");
      setStatusType("info");
      setTimeout(() => setStatusMsg(""), 3000);
    } catch (error: any) {
      setStatusMsg(error.message || "Failed to delete agency");
      setStatusType("error");
      setTimeout(() => setStatusMsg(""), 3000);
    }
  };

  const processUploadedFiles = async (files: File[]) => {
    if (files.length === 0) return;

    const results = await processDocs(files, (result) => {
      if (result && result.items) {
        addItems(result.items);
        if (result.header) mergeHeaderData(result.header);
      }
    });

    if (!results || results.length === 0) return;

    const addedCount = results.reduce(
      (total, result) =>
        total + (Array.isArray(result?.items) ? result.items.length : 0),
      0,
    );
    const fallbackCount = results.filter(
      (result) => result?.usedFallback,
    ).length;

    if (addedCount <= 0) return;

    setStatusMsg(
      fallbackCount > 0
        ? `Imported ${addedCount} item(s) from ${results.length} uploaded file(s). Document # fallback was used for ${fallbackCount} file(s).`
        : `Imported ${addedCount} item(s) from ${results.length} uploaded file(s) using AI extraction.`,
    );
    setStatusType("info");
    setTimeout(() => setStatusMsg(""), 5000);
  };

  const handleBatchUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = (e.target.files ? Array.from(e.target.files) : []) as File[];
    e.target.value = "";
    if (files.length === 0) return;
    void processUploadedFiles(files);
  };

  const handleUploadFiles = (files: File[]) => {
    if (files.length === 0) return;
    void processUploadedFiles(files);
  };

  const loadDriveFiles = async (query?: string) => {
    if (!isDriveReady) {
      setDriveError("Drive access not available. Please sign in again.");
      return;
    }
    setIsDriveLoading(true);
    setDriveError("");
    try {
      const files = await listDriveFiles(query?.trim() || undefined);
      setDriveFiles(files);
      setDriveSelected({});
    } catch (err: any) {
      setDriveError(err.message || "Failed to load Drive files.");
    } finally {
      setIsDriveLoading(false);
    }
  };

  const handleOpenDriveModal = () => {
    setIsDriveModalOpen(true);
    loadDriveFiles();
  };

  const handleDriveSearch = () => loadDriveFiles(driveSearch);

  const handleDriveToggle = (id: string) => {
    setDriveSelected((prev) => ({ ...prev, [id]: !prev[id] }));
  };

  const handleDriveToggleAll = () => {
    if (driveFiles.length === 0) return;
    const nextValue = !driveFiles.every((file) => driveSelected[file.id]);
    const nextSelected = driveFiles.reduce<Record<string, boolean>>(
      (acc, file) => ({ ...acc, [file.id]: nextValue }),
      {},
    );
    setDriveSelected(nextSelected);
  };

  const handleDriveAddSelected = async () => {
    const selectedFiles = driveFiles.filter((file) => driveSelected[file.id]);
    if (selectedFiles.length === 0) return;

    setIsDriveSelectionImporting(true);
    try {
      const { addedCount, fallbackCount } = await importDriveFilesWithAi(
        selectedFiles,
        "Drive selection",
      );

      setStatusMsg(
        fallbackCount > 0
          ? `Added ${addedCount} item(s) from ${selectedFiles.length} selected file(s). Document # fallback was used for ${fallbackCount} file(s).`
          : `Added ${addedCount} item(s) from ${selectedFiles.length} selected file(s) using AI extraction.`,
      );
      setStatusType("info");
      setIsDriveModalOpen(false);
    } catch (e: any) {
      setStatusMsg("Error: " + (e?.message || "Failed to import Drive files."));
      setStatusType("error");
    } finally {
      setIsDriveSelectionImporting(false);
      setTimeout(() => setStatusMsg(""), 5000);
    }
  };

  const handleSmartAnalysis = async () => {
    const input = smartInput.trim();
    if (!input) return;
    setIsAnalyzingText(true);
    setStatusMsg("Analyzing link...");
    setStatusType("info");
    try {
      if (!isDriveReady) {
        setStatusMsg("Drive access not available. Please sign in again.");
        setStatusType("error");
        setIsAnalyzingText(false);
        return;
      }

      let totalAdded = 0;
      let fallbackCount = 0;

      // Google Sheets link
      if (isSheetUrl(input)) {
        const sheetId = extractSheetIdFromUrl(input);
        if (!sheetId) {
          setStatusMsg("Invalid Google Sheets link.");
          setStatusType("error");
          setIsAnalyzingText(false);
          return;
        }
        setStatusMsg("Reading spreadsheet...");
        const { headers, rows } = await readSheetRows(sheetId);
        if (rows.length === 0) {
          setStatusMsg("Spreadsheet is empty or has no data rows.");
          setStatusType("error");
          setIsAnalyzingText(false);
          return;
        }

        // Map columns by header name (case-insensitive)
        const lowerHeaders = headers.map((h) => h.toLowerCase().trim());
        const col = (name: string) => {
          const idx = lowerHeaders.findIndex((h) => h.includes(name));
          return idx >= 0 ? idx : -1;
        };
        const descIdx =
          col("description") >= 0
            ? col("description")
            : col("name") >= 0
              ? col("name")
              : col("title") >= 0
                ? col("title")
                : 0;
        const qtyIdx = col("qty") >= 0 ? col("qty") : col("quantity");
        const docNumIdx =
          col("document") >= 0
            ? col("document")
            : col("doc no") >= 0
              ? col("doc no")
              : col("number");
        const remarksIdx = col("remark") >= 0 ? col("remark") : col("note");

        const items: TransmittalItem[] = rows
          .filter((row) => row.some((cell) => cell?.trim()))
          .map((row, i) => ({
            id: `sheet-${Date.now()}-${i}`,
            qty: (qtyIdx >= 0 ? row[qtyIdx]?.trim() : "") || "1",
            noOfItems: "1",
            documentNumber: docNumIdx >= 0 ? row[docNumIdx]?.trim() || "" : "",
            description: row[descIdx]?.trim() || `Row ${i + 1}`,
            remarks:
              (remarksIdx >= 0 ? row[remarksIdx]?.trim() : "") ||
              "Via Google Sheet",
            fileType: "link" as const,
          }));

        if (items.length > 0) {
          addItems(items);
          totalAdded += items.length;
        }

        // Drive folder link
      } else if (isFolderLink(input)) {
        const folderId = extractFolderIdFromLink(input);
        if (!folderId) {
          setStatusMsg("Invalid Drive folder link.");
          setStatusType("error");
          setIsAnalyzingText(false);
          return;
        }
        setStatusMsg("Listing files from folder...");
        const files = await listFilesInFolder(folderId);
        const driveImport = await importDriveFilesWithAi(
          files,
          "Drive folder import",
        );
        totalAdded += driveImport.addedCount;
        fallbackCount += driveImport.fallbackCount;

        // Individual Drive file link
      } else if (extractFileIdFromLink(input)) {
        setStatusMsg("Fetching file info...");
        const fileId = extractFileIdFromLink(input)!;
        const meta = await getFileMetadata(fileId);
        const driveImport = await importDriveFilesWithAi(
          [meta],
          "Drive file import",
        );
        totalAdded += driveImport.addedCount;
        fallbackCount += driveImport.fallbackCount;
      } else {
        setStatusMsg("Please paste a valid Google Drive or Sheets link.");
        setStatusType("error");
        setIsAnalyzingText(false);
        return;
      }

      if (totalAdded > 0) {
        setSmartInput("");
        setStatusMsg(
          fallbackCount > 0
            ? `Imported ${totalAdded} item(s). Document # fallback was used for ${fallbackCount} file(s).`
            : `Imported ${totalAdded} item(s) using AI extraction.`,
        );
        setStatusType("info");
      }
    } catch (e: any) {
      setStatusMsg("Error: " + e.message);
      setStatusType("error");
    } finally {
      setIsAnalyzingText(false);
      setTimeout(() => setStatusMsg(""), 5000);
    }
  };

  const downloadBlobLocally = (blob: Blob, fileName: string) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  };

  const showExportChoice = (
    blob: Blob,
    format: ExportFormat,
    fileName: string,
  ) => {
    setPendingExportBlob(blob);
    setPendingExportFormat(format);
    setPendingExportFileName(fileName);
    setExportChoiceOpen(true);
  };

  const handleExportLocalDownload = () => {
    if (pendingExportBlob) {
      downloadBlobLocally(pendingExportBlob, pendingExportFileName);
    }
    setExportChoiceOpen(false);
    setPendingExportBlob(null);
  };

  const handleExportUploadToDrive = () => {
    setExportChoiceOpen(false);
    setExportFolderPickerOpen(true);
  };

  const handleFolderSelected = async (folderId: string, folderName: string) => {
    if (!pendingExportBlob) return;
    setExportFolderPickerOpen(false);
    setIsUploadingToDrive(true);
    setStatusMsg("Uploading to Google Drive...");
    setStatusType("info");
    try {
      const result = await uploadFileToDrive(
        pendingExportBlob,
        pendingExportFileName,
        folderId,
      );
      setStatusMsg(`Uploaded to Drive: ${folderName}`);
      setStatusType("info");
      if (result.webViewLink) {
        window.open(result.webViewLink, "_blank");
      }
    } catch (err: any) {
      setStatusMsg(`Upload failed: ${err.message}`);
      setStatusType("error");
    } finally {
      setIsUploadingToDrive(false);
      setPendingExportBlob(null);
      setTimeout(() => setStatusMsg(""), 5000);
    }
  };

  const handlePrint = async () => {
    setIsGeneratingPdf(true);
    setTimeout(async () => {
      const element = document.getElementById("print-container");
      if (!element || !window.html2pdf) {
        setIsGeneratingPdf(false);
        return;
      }

      const timestamp = getFileTimestamp();
      const fileName = `${data.project.transmittalNumber}_${timestamp}.pdf`;
      const opt = {
        margin: 0.5,
        filename: fileName,
        image: { type: "jpeg", quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true, letterRendering: true },
        jsPDF: { unit: "in", format: "letter", orientation: "portrait" },
      };

      const pdfBlob: Blob = await window
        .html2pdf()
        .from(element)
        .set(opt)
        .toPdf()
        .get("pdf")
        .then((pdf: any) => {
          const totalPages = pdf.internal.getNumberOfPages();
          for (let i = 1; i <= totalPages; i++) {
            pdf.setPage(i);
            pdf.setFontSize(8);
            pdf.setTextColor(150);
            const text = `Page ${i} of ${totalPages}`;
            const pageWidth = pdf.internal.pageSize.getWidth();
            const pageHeight = pdf.internal.pageSize.getHeight();
            pdf.text(text, pageWidth - 1.25, pageHeight - 0.35);
          }
        })
        .outputPdf("blob");

      setIsGeneratingPdf(false);
      showExportChoice(pdfBlob, "pdf", fileName);
    }, 500);
  };

  const handleDownloadDocx = async () => {
    setIsGeneratingDocx(true);
    try {
      const blob = await generateTransmittalDocx(data);
      const timestamp = getFileTimestamp();
      const fileName = `${data.project.transmittalNumber}_${timestamp}.docx`;
      setIsGeneratingDocx(false);
      showExportChoice(blob, "docx", fileName);
    } catch {
      setIsGeneratingDocx(false);
    }
  };

  const handlePreviewDocx = async () => {
    setDocxPreviewHtml("");
    setIsPreviewModalOpen(true);
    try {
      const blob = await generateTransmittalDocx(data);
      const result = await mammoth.convertToHtml({
        arrayBuffer: await blob.arrayBuffer(),
      });
      setDocxPreviewHtml(result.value);
    } catch (e) {
      setIsPreviewModalOpen(false);
    }
  };

  const handleExportCSV = () => {
    if (data.items.length === 0) return;
    const headers = ["No.", "QTY", "Document Number", "Description", "Remarks"];
    const rows = data.items.map((item) =>
      [
        item.noOfItems,
        item.qty,
        item.documentNumber,
        item.description,
        item.remarks,
      ].join(","),
    );
    const csvContent = [headers.join(","), ...rows].join("\n");
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8" });
    const fileName = `${data.project.transmittalNumber}_items.csv`;
    showExportChoice(blob, "csv", fileName);
  };

  const handleSaveTransmittal = async () => {
    const isEditing = Boolean(activeTransmittalId);
    try {
      await saveTransmittalToDb();
      setStatusMsg(isEditing ? "Transmittal updated" : "Transmittal saved");
      setStatusType("info");
    } catch (e: any) {
      setStatusMsg(e.message || "Failed to save transmittal");
      setStatusType("error");
    }
    setTimeout(() => setStatusMsg(""), 5000);
  };

  useEffect(() => {
    const onSaveShortcut = (event: KeyboardEvent) => {
      const isSaveShortcut =
        (event.metaKey || event.ctrlKey) &&
        event.key.toLowerCase() === "s" &&
        !event.repeat;

      if (!isSaveShortcut) return;

      event.preventDefault();
      void handleSaveTransmittal();
    };

    window.addEventListener("keydown", onSaveShortcut);
    return () => window.removeEventListener("keydown", onSaveShortcut);
  }, [handleSaveTransmittal]);

  const handleSendEmail = () => {
    if (!data.recipient.email) {
      setStatusMsg("Add Recipient Email first.");
      setStatusType("error");
      setTimeout(() => setStatusMsg(""), 3000);
      return;
    }
    const subject = encodeURIComponent(
      `Transmittal: ${data.project.projectName} - ${data.project.transmittalNumber}`,
    );
    const body = encodeURIComponent(
      `Dear ${data.recipient.to},\n\nPlease find the attached transmittal for your documents.\n\nBest regards,\n${data.sender.agencyName}`,
    );
    window.open(
      `mailto:${data.recipient.email}?subject=${subject}&body=${body}`,
      "_blank",
    );
  };

  const isDocumentProcessing =
    isParsing || isBulkImporting || isDriveSelectionImporting;

  const hasFormData =
    data.items.length > 0 ||
    [
      data.recipient.to,
      data.recipient.email,
      data.recipient.company,
      data.recipient.attention,
      data.recipient.address,
      data.recipient.contactNumber,
      data.project.projectName,
      data.project.projectNumber,
      data.project.engagementRef,
      data.project.purpose,
      data.project.transmittalNumber,
      data.notes,
      data.receivedBy.name,
      data.receivedBy.date,
      data.receivedBy.time,
      data.receivedBy.remarks,
    ].some((value) => String(value || "").trim().length > 0);

  const documentProcessingStatus = isParsing
    ? parseProgress.total > 0
      ? `Parsing ${parseProgress.current}/${parseProgress.total} file(s)...`
      : "Processing uploaded files..."
    : statusMsg ||
      (isBulkImporting
        ? "Bulk importing files from Google Drive..."
        : "Importing selected files from Google Drive...");

  if (isPending) {
    return <LoadingScreen />;
  }

  if (!session?.user) {
    return (
      <LoginScreen
        onGoogleSignIn={handleGoogleSignIn}
        onDDSSignIn={handleDDSSignIn}
        authNotice={sessionErrorNotice}
      />
    );
  }

  return (
    <div className="flex flex-col lg:flex-row h-screen bg-slate-100 overflow-hidden font-sans selection:bg-brand-500/20 relative">
      {isDocumentProcessing ? (
        <div className="pointer-events-none fixed left-1/2 top-5 z-[130] w-[min(92vw,560px)] -translate-x-1/2">
          <div className="rounded-2xl border border-brand-200 bg-white/95 px-4 py-3 shadow-2xl backdrop-blur">
            <div className="flex items-start gap-3">
              <Loader2 className="mt-0.5 h-4 w-4 animate-spin text-brand-600" />
              <div className="min-w-0">
                <p className="text-[10px] font-black uppercase tracking-widest text-brand-700">
                  Processing files
                </p>
                <p className="mt-1 text-xs text-slate-600 break-words">
                  {documentProcessingStatus}
                </p>
              </div>
            </div>
          </div>
        </div>
      ) : null}

      {statusMsg && !isDocumentProcessing ? (
        <div className="pointer-events-none fixed right-4 top-5 z-[140] w-[min(92vw,360px)] animate-in slide-in-from-right-5 fade-in duration-300">
          <div
            className={`rounded-2xl border px-4 py-3 shadow-2xl backdrop-blur ${
              statusType === "error"
                ? "border-red-200 bg-red-50/95 text-red-700"
                : "border-brand-200 bg-white/95 text-slate-700"
            }`}
          >
            <p
              className={`text-[10px] font-black uppercase tracking-widest ${
                statusType === "error" ? "text-red-600" : "text-brand-700"
              }`}
            >
              {statusType === "error" ? "Action failed" : "Action complete"}
            </p>
            <p className="mt-1 text-xs break-words">{statusMsg}</p>
          </div>
        </div>
      ) : null}

      {/* Mobile Toggle Button */}
      <div className="lg:hidden absolute bottom-6 right-6 z-50 flex gap-2">
        <button
          onClick={() => setShowPreview(!showPreview)}
          className="w-14 h-14 bg-slate-900 text-white rounded-full shadow-2xl flex items-center justify-center hover:bg-slate-800 transition-all active:scale-95"
        >
          {showPreview ? (
            <Pencil className="w-6 h-6" />
          ) : (
            <Eye className="w-6 h-6" />
          )}
        </button>
      </div>

      {/* ─── Sidebar ─── */}
      <div
        className={`${showPreview ? "hidden" : "flex"} w-full lg:flex lg:w-[420px] bg-white/80 backdrop-blur-xl border-r border-slate-200/60 flex-col h-full shadow-2xl z-20 overflow-hidden absolute inset-0 lg:static`}
      >
        <SidebarHeader
          isDriveReady={isDriveReady}
          showPreview={showPreview}
          onTogglePreview={setShowPreview}
        />

        <SidebarMenuBar
          onNewTransmittal={resetForNewAnalysis}
          onOpenTransmittal={() => setIsTransmittalListOpen(true)}
          onSaveTransmittal={handleSaveTransmittal}
          isEditingTransmittal={Boolean(activeTransmittalId)}
          transmittalNumber={data.project.transmittalNumber}
          onExportPdf={handlePrint}
          onExportDocx={handleDownloadDocx}
          onExportCsv={handleExportCSV}
          onSendEmail={handleSendEmail}
          onPreviewDocx={handlePreviewDocx}
          onOpenAiSettings={() => window.dispatchEvent(new Event("open-ai-settings"))}
          onSignOut={handleSignOut}
          onResetWorkspace={resetForNewAnalysis}
          isGeneratingPdf={isGeneratingPdf}
          isGeneratingDocx={isGeneratingDocx}
        />

        <TabBar activeTab={activeTab} onTabChange={setActiveTab} />

        <div className="flex-1 overflow-y-auto p-8 space-y-10 scrollbar-hide">
          {activeTab === "content" && (
            <ContentTab
              smartInput={smartInput}
              onSmartInputChange={setSmartInput}
              isAnalyzingText={isAnalyzingText}
              onSmartAnalysis={handleSmartAnalysis}
              isParsing={isParsing}
              parseProgress={parseProgress}
              isDocumentProcessing={isDocumentProcessing}
              onOpenUploadModal={() => setIsFileUploadOpen(true)}
              isDriveReady={isDriveReady}
              onOpenDriveModal={handleOpenDriveModal}
              transmissionMethod={data.transmissionMethod}
              onUpdateTransmission={updateTransmission}
              notes={data.notes}
              onUpdateNotes={handleUpdateNotes}
            />
          )}

          {activeTab === "sender" && (
            <SenderTab
              agencies={agencies}
              selectedAgencyId={selectedAgencyId}
              onSelectAgency={setSelectedAgencyId}
              onOpenAgencyModal={openAgencyModal}
              onDeleteAgency={handleDeleteAgency}
              sender={data.sender}
            />
          )}

          {activeTab === "recipient" && (
            <RecipientTab
              recipient={data.recipient}
              onUpdateField={updateField}
            />
          )}

          {activeTab === "project" && (
            <ProjectTab project={data.project} onUpdateField={updateField} />
          )}

          {activeTab === "signatories" && (
            <SignatoriesTab
              signatories={data.signatories}
              onUpdateSignatory={handleUpdateSignatory}
            />
          )}
        </div>
      </div>

      {/* ─── Preview Panel ─── */}
      <div
        ref={previewContainerRef}
        className={`${showPreview ? "flex" : "hidden"} lg:flex flex-1 h-full overflow-y-auto overflow-x-hidden bg-slate-200 custom-scrollbar w-full absolute inset-0 lg:static z-10 flex-col items-center`}
      >
        <PreviewToolbar
          zoomPercent={zoomPercent}
          onZoomIn={handleZoomIn}
          onZoomOut={handleZoomOut}
          onZoomReset={handleZoomReset}
          onResetColumnWidths={handleResetColumnWidths}
          onZoomSet={handleZoomSet}
          transmittalNumber={data.project.transmittalNumber}
          showSaveNotice={hasFormData}
        />
        <div className="flex-1 overflow-y-auto overflow-x-hidden w-full flex flex-col items-center p-4 lg:p-8 pt-0">
          <div
            className={`transition-all duration-300 ease-out origin-top shadow-[0_40px_100px_rgba(0,0,0,0.15)] rounded-sm shrink-0 ${
              isDocumentProcessing ? "opacity-60" : "opacity-100"
            }`}
            style={{
              transform: `scale(${previewScale})`,
              width: "816px",
              marginBottom: "200px",
            }}
          >
            <div
              id="print-container"
              className="bg-white min-h-[1056px]"
              aria-busy={isDocumentProcessing}
            >
              <TransmittalTemplate
                data={data}
                onUpdateItem={updateItem}
                onAdjustItemQty={adjustItemQty}
                onRemoveItem={removeItem}
                onMoveItem={moveItem}
                onReorderItems={handleReorderItems}
                onAddItem={handleManualAdd}
                onBulkAdd={() => setIsBulkModalOpen(true)}
                onUpdateSignatory={handleUpdateSignatory}
                onUpdateReceivedBy={handleUpdateReceivedBy}
                onUpdateFooter={handleUpdateFooter}
                onUpdateNotes={handleUpdateNotes}
                isGeneratingPdf={isGeneratingPdf}
                columnWidths={
                  isGeneratingPdf ? DEFAULT_COLUMN_WIDTHS : columnWidths
                }
                onResizeDivider={handleResizeDivider}
              />
            </div>
          </div>
        </div>
      </div>

      {/* ─── Modals ─── */}
      <FileUploadModal
        isOpen={isFileUploadOpen}
        onClose={() => setIsFileUploadOpen(false)}
        onUploadFiles={handleUploadFiles}
        isDriveReady={isDriveReady}
        isParsing={isParsing}
        parseProgress={parseProgress}
      />
      <TransmittalListModal
        isOpen={isTransmittalListOpen}
        onClose={() => setIsTransmittalListOpen(false)}
        onOpenTransmittal={handleOpenTransmittal}
        onDeleteTransmittal={handleDeleteTransmittal}
        apiBaseUrl={apiBaseUrl}
      />
      <BulkAddModal
        isOpen={isBulkModalOpen}
        onClose={() => setIsBulkModalOpen(false)}
        onImportDriveLink={handleBulkImportDriveLink}
      />
      <DocxPreviewModal
        isOpen={isPreviewModalOpen}
        onClose={() => setIsPreviewModalOpen(false)}
        html={docxPreviewHtml}
      />
      <DriveFileModal
        isOpen={isDriveModalOpen}
        onClose={() => setIsDriveModalOpen(false)}
        files={driveFiles}
        searchValue={driveSearch}
        isLoading={isDriveLoading}
        error={driveError}
        selectedIds={driveSelected}
        isAllSelected={
          driveFiles.length > 0 &&
          driveFiles.every((file) => driveSelected[file.id])
        }
        onSearchChange={setDriveSearch}
        onSearch={handleDriveSearch}
        onToggle={handleDriveToggle}
        onToggleAll={handleDriveToggleAll}
        onAddSelected={handleDriveAddSelected}
        isImporting={isDriveSelectionImporting}
        processingMessage={statusMsg}
      />
      <ExportChoiceModal
        isOpen={exportChoiceOpen}
        format={pendingExportFormat}
        fileName={pendingExportFileName}
        isUploading={isUploadingToDrive}
        onDownloadLocal={handleExportLocalDownload}
        onUploadToDrive={handleExportUploadToDrive}
        onClose={() => {
          setExportChoiceOpen(false);
          setPendingExportBlob(null);
        }}
      />
      <FolderPickerModal
        isOpen={exportFolderPickerOpen}
        onClose={() => {
          setExportFolderPickerOpen(false);
          setPendingExportBlob(null);
        }}
        onSelect={handleFolderSelected}
      />
      <AgencyPresetModal
        isOpen={isAgencyModalOpen}
        onClose={() => setIsAgencyModalOpen(false)}
        mode={agencyModalMode}
        isDriveReady={isDriveReady}
        draft={agencyDraft}
        onChange={setAgencyDraft}
        onLogoUpload={handleAgencyLogoUpload}
        driveLogoLink={agencyDriveLogoLink}
        onDriveLogoLinkChange={setAgencyDriveLogoLink}
        onImportDriveLogo={handleAgencyDriveLogoImport}
        isImportingDriveLogo={isImportingAgencyDriveLogo}
        onSave={saveAgencyPreset}
      />

      <FloatingAccount
        user={{
          name: session?.user?.name,
          email: session?.user?.email,
          image: session?.user?.image,
        }}
        apiBaseUrl={apiBaseUrl}
        isDriveReady={isDriveReady}
        onSignOut={handleSignOut}
      />

      <style>{`
                .scrollbar-hide::-webkit-scrollbar { display: none; }
                .ease-out-expo { transition-timing-function: cubic-bezier(0.19, 1, 0.22, 1); }
            `}</style>
    </div>
  );
};

export const App = () => (
  <ErrorBoundary>
    <AppContent />
  </ErrorBoundary>
);
