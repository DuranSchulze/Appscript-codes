"use client";

import React, { useState, useRef, useEffect } from "react";
import { Pencil, Eye } from "lucide-react";
import { TransmittalTemplate } from "../NewReportTemplate";
import { FloatingAccount } from "../FloatingAccount";
import { BulkAddModal } from "../modals/BulkAddModal";
import { AgencyPresetModal } from "../modals/AgencyPresetModal";
import { DriveFileModal } from "../modals/DriveFileModal";
import { DocxPreviewModal } from "../modals/DocxPreviewModal";
import { ErrorBoundary } from "../ErrorBoundary";
import { parseTransmittalDocument } from "../../services/geminiService";
import {
  listFilesInFolder,
  extractFolderIdFromLink,
  extractFileIdFromLink,
  isFolderLink,
  getFileMetadata,
  listDriveFiles,
  checkDriveAccess,
  clearGoogleToken,
} from "../../services/googleDriveService";
import { useFileProcessing, resizeImage } from "../../hooks/useFileProcessing";
import { generateTransmittalDocx } from "../../services/docxGenerator";
import {
  getLinkedSheetId,
  setLinkedSheetId,
  clearLinkedSheetId,
  extractSheetIdFromUrl,
  getSpreadsheetTitle,
  appendTransmittalRow,
  isSheetUrl,
  readSheetRows,
} from "../../services/googleSheetsService";
import { signIn, signOut, useSession } from "../../lib/auth-client";
import {
  AppData,
  TransmittalItem,
  Signatories,
  ReceivedBy,
  FooterNotes,
  HistoryItem,
  SenderInfo,
} from "../../types";
import * as mammoth from "mammoth";

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
import { HistoryTab } from "./tabs/HistoryTab";
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
  const [driveFiles, setDriveFiles] = useState<
    Array<{ id: string; name: string; mimeType: string }>
  >([]);
  const [driveSearch, setDriveSearch] = useState("");
  const [driveSelected, setDriveSelected] = useState<Record<string, boolean>>(
    {},
  );
  const [driveError, setDriveError] = useState("");
  const [isDriveLoading, setIsDriveLoading] = useState(false);
  const [activeTab, setActiveTab] = useState<TabKey>("content");
  const [history, setHistory] = useState<HistoryItem[]>(() => {
    if (typeof window === "undefined") return [];
    try {
      const saved = window.localStorage.getItem("transmittal_history");
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });
  const [data, setData] = useState<AppData>(() => createInitialData());
  const [activeTransmittalId, setActiveTransmittalId] = useState<string | null>(
    null,
  );
  const [showPreview, setShowPreview] = useState(false);

  const [columnWidths, setColumnWidths] = useState({
    qty: 55,
    noOfItems: 65,
    documentNumber: 130,
    remarks: 150, // Increased from 100 to 150
  });
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

  const { data: session, isPending } = useSession();
  const apiBaseUrl =
    process.env.NEXT_PUBLIC_BETTER_AUTH_URL ||
    (typeof window !== "undefined" ? window.location.origin : "");

  const handleGoogleSignIn = async () => {
    await signIn.social({
      provider: "google",
      callbackURL: window.location.origin,
    });
  };

  const handleSignOut = async () => {
    clearGoogleToken();
    await signOut();
  };

  const fileInputRef = useRef<HTMLInputElement>(null);
  const logoInputRef = useRef<HTMLInputElement>(null);
  const printRef = useRef<HTMLDivElement>(null);
  const [logoInputKey, setLogoInputKey] = useState(0);

  const [agencies, setAgencies] = useState<DbAgency[]>([]);
  const [selectedAgencyId, setSelectedAgencyId] = useState<string>("");
  const [isAgencyModalOpen, setIsAgencyModalOpen] = useState(false);
  const [agencyDraft, setAgencyDraft] = useState<SenderInfo>(
    () => createInitialData().sender,
  );

  // Google Sheets linking state
  const [linkedSheetTitle, setLinkedSheetTitle] = useState<string | null>(null);
  const [sheetUrlInput, setSheetUrlInput] = useState("");
  const [sheetLinkError, setSheetLinkError] = useState("");
  const [isLinkingSheet, setIsLinkingSheet] = useState(false);

  // Restore linked sheet title on mount
  useEffect(() => {
    const id = getLinkedSheetId();
    if (id) {
      getSpreadsheetTitle(id)
        .then((title) => setLinkedSheetTitle(title))
        .catch(() => {
          clearLinkedSheetId();
          setLinkedSheetTitle(null);
        });
    }
  }, []);

  const handleLinkSheet = async () => {
    setSheetLinkError("");
    const url = sheetUrlInput.trim();
    if (!url) return;
    const id = extractSheetIdFromUrl(url);
    if (!id) {
      setSheetLinkError("Invalid Google Sheets URL");
      return;
    }
    setIsLinkingSheet(true);
    try {
      const title = await getSpreadsheetTitle(id);
      setLinkedSheetId(id);
      setLinkedSheetTitle(title);
      setSheetUrlInput("");
    } catch {
      setSheetLinkError(
        "Cannot access this spreadsheet. Make sure it exists and you have edit access.",
      );
    } finally {
      setIsLinkingSheet(false);
    }
  };

  const handleUnlinkSheet = () => {
    clearLinkedSheetId();
    setLinkedSheetTitle(null);
  };

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
    loadHistoryFromDb();
  }, [session?.user?.id]);

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
    if (typeof window === "undefined") return;
    try {
      window.localStorage.setItem(
        "transmittal_history",
        JSON.stringify(history),
      );
    } catch {
      return;
    }
  }, [history]);

  useEffect(() => {
    const id = data.agencyId ? String(data.agencyId) : "";
    if (id) setSelectedAgencyId(id);
  }, [data.agencyId]);

  useEffect(() => {
    if (!selectedAgencyId) return;
    const agency = agencies.find((a) => a.id === selectedAgencyId);
    if (!agency) return;
    setData((prev) => ({
      ...prev,
      agencyId: agency.id,
      sender: {
        agencyName: agency.name || "",
        addressLine1: agency.addressLine1 || "",
        addressLine2: agency.addressLine2 || "",
        website: agency.website || "",
        mobile: agency.contactNumber || "",
        telephone: agency.telephoneNumber || "",
        email: agency.email || "",
        logoBase64: agency.logoBase64 || null,
      },
    }));
    setLogoInputKey((prev) => prev + 1);
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
        department: projectData.department || "",
        date: projectData.date || "",
        timeGenerated: projectData.timeGenerated || "",
      },
      items: Array.isArray(transmittal.items)
        ? transmittal.items.map((item: any) => ({
            id: item.id,
            qty: item.qty || "",
            noOfItems: item.noOfItems || "",
            documentNumber: item.documentNumber || "",
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

  const mapDbTransmittalToHistoryItem = (transmittal: any): HistoryItem => {
    const data = mapDbTransmittalToAppData(transmittal);
    const dateTime = [data.project.date, data.project.timeGenerated]
      .filter(Boolean)
      .join(" ")
      .trim();

    return {
      id: transmittal.id,
      timestamp: dateTime || new Date().toLocaleString(),
      transmittalNumber: data.project.transmittalNumber,
      projectName: data.project.projectName || "Untitled Project",
      recipientName: data.recipient.to || "Unknown Recipient",
      createdBy: session?.user?.name || session?.user?.email || "Unknown",
      preparedBy: data.signatories.preparedBy || "",
      notedBy: data.signatories.notedBy || "",
      data,
    };
  };

  const loadHistoryFromDb = async () => {
    if (!session?.user || !apiBaseUrl) return;
    try {
      const response = await fetch(`${apiBaseUrl}/api/transmittals`, {
        credentials: "include",
      });
      if (!response.ok) return;
      const payload = await response.json();
      const mapped = Array.isArray(payload.transmittals)
        ? payload.transmittals.map(mapDbTransmittalToHistoryItem)
        : [];
      setHistory(mapped);
    } catch (error) {
      console.error("Failed to load transmittals", error);
    }
  };

  const saveTransmittalToDb = async () => {
    if (!session?.user || !apiBaseUrl) return;
    try {
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
        const historyItem = mapDbTransmittalToHistoryItem(payload.transmittal);
        setHistory((prev) => {
          const existingIndex = prev.findIndex(
            (item) => item.id === historyItem.id,
          );
          if (existingIndex === -1) {
            return [historyItem, ...prev].slice(0, 20);
          }
          const next = [...prev];
          next[existingIndex] = historyItem;
          return next;
        });
        setActiveTransmittalId(payload.transmittal.id);

        // Sync the form with the server-assigned transmittal number
        // (the server generates the number atomically to prevent duplicates)
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
    } catch (error: any) {
      console.error("Save transmittal failed:", error);
      setStatusMsg(error.message || "Failed to save transmittal");
      setStatusType("error");
      setTimeout(() => setStatusMsg(""), 3000);
    }
  };

  const addToHistory = () => {
    const newItem: HistoryItem = {
      id: Date.now().toString(),
      timestamp: new Date().toLocaleString(),
      transmittalNumber: data.project.transmittalNumber,
      projectName: data.project.projectName || "Untitled Project",
      recipientName: data.recipient.to || "Unknown Recipient",
      createdBy: session?.user?.name || session?.user?.email || "Unknown",
      preparedBy: data.signatories.preparedBy || "",
      notedBy: data.signatories.notedBy || "",
      data: { ...data },
    };
    setHistory((prev) => [newItem, ...prev.slice(0, 19)]); // Keep last 20
  };

  const loadFromHistory = (item: HistoryItem) => {
    setData(item.data);
    setActiveTransmittalId(item.id);
    setActiveTab("content");
    setStatusMsg("Snapshot Restored");
    setTimeout(() => setStatusMsg(""), 3000);
  };

  const deleteHistoryItem = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    setHistory((prev) => prev.filter((item) => item.id !== id));
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
  const addItems = (newItems: TransmittalItem[]) => {
    setData((prev) => ({
      ...prev,
      items: reindexItems([...prev.items, ...newItems]),
    }));
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
      setStatusMsg("Listing Drive files...");
      setStatusType("info");
      const files = await listFilesInFolder(folderId);
      addItems(
        files.map((f) => ({
          id: f.id,
          qty: "1",
          noOfItems: "1",
          documentNumber: "DRIVE",
          description: f.name.replace(/\.[^/.]+$/, ""),
          remarks: "Via Drive Folder",
          fileType: "gdrive",
          fileSource: `https://drive.google.com/file/d/${f.id}/view`,
        })),
      );
      setStatusMsg(`Fetched ${files.length} files from Drive.`);
      setStatusType("info");
      setTimeout(() => setStatusMsg(""), 3000);
    } catch (e: any) {
      setStatusMsg("Error: " + e.message);
      setStatusType("error");
      setTimeout(() => setStatusMsg(""), 5000);
      throw e;
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

  const handleColumnResize = (
    field: keyof typeof columnWidths,
    newWidth: number,
  ) => setColumnWidths((prev) => ({ ...prev, [field]: newWidth }));

  const {
    processFiles: processDocs,
    isProcessing: isParsing,
    progress: parseProgress,
    error: processingError,
  } = useFileProcessing<any>(
    async (base64, mimeType, fileName) => {
      const result = await parseTransmittalDocument(
        base64,
        mimeType,
        false,
        fileName,
      );
      if ((result as any).error) throw new Error((result as any).error);
      return {
        items: result.items.map((res) => ({
          id: Date.now().toString() + Math.random(),
          qty: res.qty || "1",
          noOfItems: "1",
          documentNumber: res.documentNumber || "",
          description: res.description,
          remarks: res.remarks || "",
          fileType: "upload",
        })),
        header: result.header,
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

  const handleLogoUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const base64 = await resizeImage(file, 400);
      updateField("sender", "logoBase64", `data:image/jpeg;base64,${base64}`);
    } catch (err: any) {
      setStatusMsg(`Logo Error: ${err.message}`);
      setStatusType("error");
      setTimeout(() => setStatusMsg(""), 3000);
    } finally {
      setLogoInputKey((prev) => prev + 1);
    }
  };

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

  const openAgencyModal = () => {
    setAgencyDraft({ ...data.sender });
    setIsAgencyModalOpen(true);
  };

  const saveAgencyPreset = async () => {
    if (!session?.user || !apiBaseUrl) return;
    const name = agencyDraft.agencyName.trim();
    if (!name) return;
    try {
      const response = await fetch(`${apiBaseUrl}/api/agencies`, {
        method: "POST",
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
          sender: { ...agencyDraft, agencyName: name },
        }));
      }

      setIsAgencyModalOpen(false);
    } catch (error: any) {
      console.error("Save agency failed:", error);
      setStatusMsg(error.message || "Failed to save agency");
      setStatusType("error");
      setTimeout(() => setStatusMsg(""), 3000);
    }
  };

  const handleBatchUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = (e.target.files ? Array.from(e.target.files) : []) as File[];
    if (files.length === 0) return;
    processDocs(files, (result) => {
      if (result && result.items) {
        addItems(result.items);
        if (result.header) mergeHeaderData(result.header);
      }
    });
    e.target.value = "";
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

  const handleDriveAddSelected = () => {
    const selectedFiles = driveFiles.filter((file) => driveSelected[file.id]);
    if (selectedFiles.length === 0) return;
    addItems(
      selectedFiles.map((file) => ({
        id: file.id,
        qty: "1",
        noOfItems: "1",
        documentNumber: "DRIVE",
        description: file.name.replace(/\.[^/.]+$/, ""),
        remarks: "From Google Drive",
        fileType: "gdrive",
        fileSource: `https://drive.google.com/file/d/${file.id}/view`,
      })),
    );
    setStatusMsg(`Added ${selectedFiles.length} Drive files.`);
    setStatusType("info");
    setTimeout(() => setStatusMsg(""), 3000);
    setIsDriveModalOpen(false);
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
        addItems(
          files.map((f) => ({
            id: f.id,
            qty: "1",
            noOfItems: "1",
            documentNumber: "SCAN",
            description: f.name.replace(/\.[^/.]+$/, ""),
            remarks: "Via Drive Folder",
            fileType: "gdrive" as const,
            fileSource: `https://drive.google.com/file/d/${f.id}/view`,
          })),
        );
        totalAdded += files.length;

        // Individual Drive file link
      } else if (extractFileIdFromLink(input)) {
        setStatusMsg("Fetching file info...");
        const fileId = extractFileIdFromLink(input)!;
        const meta = await getFileMetadata(fileId);
        addItems([
          {
            id: meta.id,
            qty: "1",
            noOfItems: "1",
            documentNumber: "DRIVE",
            description: meta.name.replace(/\.[^/.]+$/, ""),
            remarks: "Via Drive Link",
            fileType: "gdrive",
            fileSource: `https://drive.google.com/file/d/${meta.id}/view`,
          },
        ]);
        totalAdded += 1;
      } else {
        setStatusMsg("Please paste a valid Google Drive or Sheets link.");
        setStatusType("error");
        setIsAnalyzingText(false);
        return;
      }

      if (totalAdded > 0) {
        setSmartInput("");
        setStatusMsg(`Imported ${totalAdded} item(s).`);
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

  const handlePrint = async () => {
    setIsGeneratingPdf(true);
    setTimeout(async () => {
      const element = document.getElementById("print-container");
      if (!element || !window.html2pdf) {
        setIsGeneratingPdf(false);
        return;
      }

      const timestamp = getFileTimestamp();
      const opt = {
        margin: 0.5,
        filename: `${data.project.transmittalNumber}_${timestamp}.pdf`,
        image: { type: "jpeg", quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true, letterRendering: true },
        jsPDF: { unit: "in", format: "letter", orientation: "portrait" },
      };

      await window
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
        .save();

      setIsGeneratingPdf(false);
    }, 500);
  };

  const handleDownloadDocx = async () => {
    setIsGeneratingDocx(true);
    try {
      const blob = await generateTransmittalDocx(data);
      const timestamp = getFileTimestamp();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${data.project.transmittalNumber}_${timestamp}.docx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } finally {
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
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `${data.project.transmittalNumber}_items.csv`;
    link.click();
  };

  const handleSaveTransmittal = async () => {
    try {
      await saveTransmittalToDb();
      setStatusMsg("Saved to history");
      setStatusType("info");
    } catch (e: any) {
      setStatusMsg(e.message || "Failed to save transmittal");
      setStatusType("error");
    }
    setTimeout(() => setStatusMsg(""), 5000);
  };

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

  if (isPending) {
    return <LoadingScreen />;
  }

  if (!session?.user) {
    return <LoginScreen onGoogleSignIn={handleGoogleSignIn} />;
  }

  return (
    <div className="flex flex-col lg:flex-row h-screen bg-slate-100 overflow-hidden font-sans selection:bg-brand-500/20 relative">
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
          onSaveTransmittal={handleSaveTransmittal}
          onExportPdf={handlePrint}
          onExportDocx={handleDownloadDocx}
          onExportCsv={handleExportCSV}
          onSendEmail={handleSendEmail}
          onPreviewDocx={handlePreviewDocx}
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
              fileInputRef={fileInputRef}
              onBatchUpload={handleBatchUpload}
              isDriveReady={isDriveReady}
              onOpenDriveModal={handleOpenDriveModal}
              statusMsg={statusMsg}
              statusType={statusType}
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
              sender={data.sender}
              onUpdateField={updateField}
              logoInputRef={logoInputRef}
              logoInputKey={logoInputKey}
              onLogoUpload={handleLogoUpload}
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

          {activeTab === "history" && (
            <HistoryTab
              history={history}
              onLoadFromHistory={loadFromHistory}
              onDeleteHistoryItem={deleteHistoryItem}
              linkedSheetTitle={linkedSheetTitle}
              sheetUrlInput={sheetUrlInput}
              onSheetUrlInputChange={(v) => {
                setSheetUrlInput(v);
                setSheetLinkError("");
              }}
              sheetLinkError={sheetLinkError}
              isLinkingSheet={isLinkingSheet}
              onLinkSheet={handleLinkSheet}
              onUnlinkSheet={handleUnlinkSheet}
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
          onZoomSet={handleZoomSet}
        />
        <div className="flex-1 overflow-y-auto overflow-x-hidden w-full flex flex-col items-center p-4 lg:p-8 pt-0">
          <div
            className="transition-all duration-300 ease-out origin-top shadow-[0_40px_100px_rgba(0,0,0,0.15)] rounded-sm shrink-0"
            style={{
              transform: `scale(${previewScale})`,
              width: "816px",
              marginBottom: "200px",
            }}
          >
            <div id="print-container" className="bg-white min-h-[1056px]">
              <TransmittalTemplate
                data={data}
                onUpdateItem={updateItem}
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
                columnWidths={columnWidths}
                onColumnResize={handleColumnResize}
              />
            </div>
          </div>
        </div>
      </div>

      {/* ─── Modals ─── */}
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
      />
      <AgencyPresetModal
        isOpen={isAgencyModalOpen}
        onClose={() => setIsAgencyModalOpen(false)}
        draft={agencyDraft}
        onChange={setAgencyDraft}
        onLogoUpload={handleAgencyLogoUpload}
        onSave={saveAgencyPreset}
      />

      <FloatingAccount
        user={{
          name: session?.user?.name,
          email: session?.user?.email,
          image: session?.user?.image,
        }}
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
