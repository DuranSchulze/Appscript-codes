"use client";

import React, { useState, useRef, useEffect } from "react";
import Link from "next/link";
import { TransmittalTemplate } from "../NewReportTemplate";
import { FloatingAccount } from "../FloatingAccount";
import { BulkAddModal } from "../modals/BulkAddModal";
import { AgencyPresetModal } from "../modals/AgencyPresetModal";
import { DriveFileModal } from "../modals/DriveFileModal";
import { DocxPreviewModal } from "../modals/DocxPreviewModal";
import { ErrorBoundary } from "../ErrorBoundary";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { Card, CardContent } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Separator } from "@/components/ui/separator";
import { parseTransmittalDocument } from "../../services/geminiService";
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
} from "../../services/googleDriveService";
import { useFileProcessing, resizeImage } from "../../hooks/useFileProcessing";
import { generateTransmittalDocx } from "../../services/docxGenerator";
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

// Add type declaration
declare global {
  interface Window {
    html2pdf: any;
    html2canvas: any;
  }
}

const formatTime24hTo12h = (time24h: string): string => {
  const match = time24h.match(/^\s*(\d{1,2}):(\d{2})\s*$/);
  if (!match) return "";
  const hours = Number(match[1]);
  const minutes = match[2];
  const minutesNum = Number(minutes);
  if (
    !Number.isFinite(hours) ||
    hours < 0 ||
    hours > 23 ||
    !Number.isFinite(minutesNum) ||
    minutesNum < 0 ||
    minutesNum > 59
  )
    return "";
  const hour12 = hours % 12 === 0 ? 12 : hours % 12;
  const ampm = hours >= 12 ? "PM" : "AM";
  return `${hour12}:${minutes} ${ampm}`;
};

type AgencyPreset = {
  id: string;
  label: string;
  sender: SenderInfo;
  updatedAt: string;
};

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

const parseTimeTo24h = (value: string): string => {
  const trimmed = value.trim();
  if (!trimmed) return "";

  const direct24h = trimmed.match(/^(\d{1,2}):(\d{2})$/);
  if (direct24h) {
    const h = Number(direct24h[1]);
    const m = Number(direct24h[2]);
    if (
      Number.isFinite(h) &&
      Number.isFinite(m) &&
      h >= 0 &&
      h <= 23 &&
      m >= 0 &&
      m <= 59
    ) {
      return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
    }
    return "";
  }

  const match12h = trimmed.match(/^(\d{1,2}):(\d{2})\s*([AaPp][Mm])$/);
  if (!match12h) return "";

  let hours = Number(match12h[1]);
  const minutes = Number(match12h[2]);
  const ampm = match12h[3].toUpperCase();

  if (
    !Number.isFinite(hours) ||
    !Number.isFinite(minutes) ||
    hours < 1 ||
    hours > 12 ||
    minutes < 0 ||
    minutes > 59
  ) {
    return "";
  }

  if (ampm === "AM") {
    hours = hours === 12 ? 0 : hours;
  } else {
    hours = hours === 12 ? 12 : hours + 12;
  }

  return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
};

const ExpandingTextarea = ({
  value,
  onChange,
  placeholder,
  className = "",
}: {
  value: string;
  onChange: (v: string) => void;
  placeholder?: string;
  className?: string;
}) => {
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  useEffect(() => {
    if (textareaRef.current) {
      textareaRef.current.style.height = "auto";
      textareaRef.current.style.height = `${textareaRef.current.scrollHeight}px`;
    }
  }, [value]);
  return (
    <Textarea
      ref={textareaRef}
      className={`overflow-hidden resize-none min-h-[44px] ${className}`}
      value={value}
      onChange={(e) => onChange(e.target.value)}
      placeholder={placeholder}
      rows={1}
    />
  );
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
  const [activeTab, setActiveTab] = useState<
    "content" | "sender" | "recipient" | "project" | "signatories" | "history"
  >("content");
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
  const [previewScale, setPreviewScale] = useState(1);
  const previewContainerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleResize = () => {
      if (previewContainerRef.current) {
        const containerWidth = previewContainerRef.current.offsetWidth;
        // 8.5in is roughly 816px. We add some padding buffer.
        const targetWidth = 850;
        if (containerWidth < targetWidth) {
          const newScale = (containerWidth - 32) / targetWidth; // 32px padding
          setPreviewScale(Math.max(0.3, Math.min(1, newScale)));
        } else {
          setPreviewScale(1);
        }
      }
    };

    handleResize();
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, [showPreview]); // Re-calculate when preview is shown

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

  const fetchNextTransmittalNumber = async () => {
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
          if (prev.project.transmittalNumber) return prev;
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
    if (confirm("Replace current workspace with this history snapshot?")) {
      setData(item.data);
      setActiveTransmittalId(item.id);
      setActiveTab("content");
      setStatusMsg("Snapshot Restored");
      setTimeout(() => setStatusMsg(""), 3000);
    }
  };

  const deleteHistoryItem = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    setHistory((prev) => prev.filter((item) => item.id !== id));
  };

  const resetForNewAnalysis = () => {
    if (
      confirm(
        "Reset current form? This will not affect your History snapshots.",
      )
    ) {
      setData(createInitialData());
      setActiveTransmittalId(null);
      fetchNextTransmittalNumber();
      setStatusMsg("Form Reset");
      setTimeout(() => setStatusMsg(""), 3000);
    }
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
      alert(`Logo Error: ${err.message}`);
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
      alert(`Logo Error: ${err.message}`);
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
    if (!smartInput.trim()) return;
    setIsAnalyzingText(true);
    setStatusMsg("Analyzing Input...");
    try {
      const lines = smartInput
        .split("\n")
        .map((l) => l.trim())
        .filter(Boolean);
      const folderLines: string[] = [];
      const fileLines: string[] = [];
      const textLines: string[] = [];

      for (const line of lines) {
        if (isFolderLink(line)) {
          folderLines.push(line);
        } else if (extractFileIdFromLink(line)) {
          fileLines.push(line);
        } else {
          textLines.push(line);
        }
      }

      const hasDriveLinks = folderLines.length > 0 || fileLines.length > 0;
      if (hasDriveLinks && !isDriveReady) {
        setStatusMsg("Drive access not available. Please sign in again.");
        setStatusType("error");
        setIsAnalyzingText(false);
        return;
      }

      let totalAdded = 0;

      // Process folder links
      for (const link of folderLines) {
        const folderId = extractFolderIdFromLink(link);
        if (!folderId) continue;
        setStatusMsg(`Listing files from folder...`);
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
      }

      // Process individual file links in parallel
      if (fileLines.length > 0) {
        setStatusMsg(`Fetching ${fileLines.length} file(s)...`);
        const metadataResults = await Promise.allSettled(
          fileLines.map((link) => {
            const fileId = extractFileIdFromLink(link)!;
            return getFileMetadata(fileId).then((meta) => ({ ...meta, link }));
          }),
        );
        const fileItems: TransmittalItem[] = [];
        for (const result of metadataResults) {
          if (result.status === "fulfilled") {
            const f = result.value;
            fileItems.push({
              id: f.id,
              qty: "1",
              noOfItems: "1",
              documentNumber: "DRIVE",
              description: f.name.replace(/\.[^/.]+$/, ""),
              remarks: "Via Drive Link",
              fileType: "gdrive",
              fileSource: `https://drive.google.com/file/d/${f.id}/view`,
            });
          }
        }
        if (fileItems.length > 0) {
          addItems(fileItems);
          totalAdded += fileItems.length;
        }
      }

      // Process remaining plain text through AI
      if (textLines.length > 0) {
        setStatusMsg("Analyzing text...");
        const textBlock = textLines.join("\n");
        const results = await parseTransmittalDocument(
          textBlock,
          "text/plain",
          true,
        );
        if (results.items.length > 0) {
          addItems(
            results.items.map((res) => ({
              id: Date.now().toString() + Math.random(),
              qty: res.qty || "1",
              noOfItems: "1",
              documentNumber: res.documentNumber || "",
              description: res.description,
              remarks: res.remarks || "",
              fileType: "link",
            })),
          );
          if (results.header) mergeHeaderData(results.header);
          totalAdded += results.items.length;
        }
      }

      if (totalAdded > 0) {
        setSmartInput("");
        setStatusMsg(`Imported ${totalAdded} item(s).`);
        setStatusType("info");
      } else if (textLines.length === 0 && !hasDriveLinks) {
        setStatusMsg("No recognizable content found.");
        setStatusType("error");
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
    await saveTransmittalToDb();
    setStatusMsg("Saved to history");
    setStatusType("info");
    setTimeout(() => setStatusMsg(""), 3000);
  };

  const handleSendEmail = () => {
    if (!data.recipient.email) {
      alert("Add Recipient Email first.");
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
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-100 font-sans">
        <div className="bg-white border border-slate-200 rounded-3xl shadow-2xl p-10 text-center">
          <div className="w-12 h-12 border-4 border-brand-500 border-t-transparent rounded-full animate-spin mx-auto mb-4"></div>
          <p className="text-xs font-black uppercase tracking-widest text-slate-500">
            Loading session
          </p>
        </div>
      </div>
    );
  }

  if (!session?.user) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-100 font-sans">
        <div className="bg-white border border-slate-200 rounded-[36px] shadow-2xl w-full max-w-lg p-10 text-center">
          <h1 className="text-2xl font-black text-slate-900 font-display">
            Smart Transmittal
          </h1>
          <p className="text-sm text-slate-500 mt-2">
            This is a private system created for the Internal Document
            Transmittal System. Please sign in
          </p>
          <button
            onClick={handleGoogleSignIn}
            className="mt-8 w-full py-4 rounded-2xl bg-slate-900 text-white font-bold uppercase tracking-widest text-xs shadow-xl hover:bg-slate-800 transition-all"
          >
            Sign in with Google
          </button>
          <div className="flex justify-center gap-6 mt-6">
            <Link
              href="/legal/privacy-policy"
              target="_blank"
              className="text-[10px] font-bold uppercase tracking-widest text-slate-400 hover:text-slate-600 transition-colors"
            >
              Privacy Policy
            </Link>
            <Link
              href="/legal/terms-of-service"
              target="_blank"
              className="text-[10px] font-bold uppercase tracking-widest text-slate-400 hover:text-slate-600 transition-colors"
            >
              Terms of Service
            </Link>
          </div>
        </div>
      </div>
    );
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
            <svg
              className="w-6 h-6"
              fill="none"
              stroke="currentColor"
              viewBox="0 0 24 24"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth="2"
                d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"
              ></path>
            </svg>
          ) : (
            <svg
              className="w-6 h-6"
              fill="none"
              stroke="currentColor"
              viewBox="0 0 24 24"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth="2"
                d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"
              ></path>
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth="2"
                d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"
              ></path>
            </svg>
          )}
        </button>
      </div>

      <div
        className={`${showPreview ? "hidden" : "flex"} w-full lg:flex lg:w-[420px] bg-white/80 backdrop-blur-xl border-r border-slate-200/60 flex-col h-full shadow-2xl z-20 overflow-hidden absolute inset-0 lg:static`}
      >
        <div className="p-8 border-b border-slate-100 bg-white/40">
          <div className="flex justify-between items-center mb-1">
            <h1 className="font-display font-black text-2xl text-slate-900 tracking-tighter">
              Smart Transmittal
            </h1>
            <div className="flex items-center gap-3">
              <button
                onClick={() => setShowPreview(true)}
                className="lg:hidden text-slate-400 hover:text-brand-600 transition-colors"
                title="Minimize Sidebar"
              >
                <svg
                  className="w-5 h-5"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth="2.5"
                    d="M19 9l-7 7-7-7"
                  ></path>
                </svg>
              </button>
              <button
                onClick={handleSignOut}
                className="text-[9px] font-black text-slate-400 hover:text-slate-700 uppercase tracking-widest"
              >
                Sign out
              </button>
            </div>
          </div>
          <div className="flex justify-between items-center mt-1">
            <p className="text-[10px] text-slate-400 font-black uppercase tracking-widest">
              {isDriveReady
                ? "✓ Drive Connected"
                : "Enterprise Controller v2.1"}
            </p>
            <button
              onClick={resetForNewAnalysis}
              className="text-[9px] font-black text-slate-400 hover:text-red-500 uppercase tracking-widest transition-colors"
            >
              Clear Workspace
            </button>
          </div>
        </div>

        <div className="flex flex-wrap p-2 bg-slate-100/50 m-5 rounded-2xl gap-2 shrink-0">
          <button
            onClick={() => setActiveTab("content")}
            className={`flex-1 py-2 px-4 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all whitespace-nowrap ${activeTab === "content" ? "bg-white shadow-sm text-slate-800" : "text-slate-400 hover:text-slate-600"}`}
          >
            Content
          </button>
          <button
            onClick={() => setActiveTab("sender")}
            className={`flex-1 py-2 px-4 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all whitespace-nowrap ${activeTab === "sender" ? "bg-white shadow-sm text-slate-800" : "text-slate-400 hover:text-slate-600"}`}
          >
            Brand
          </button>
          <button
            onClick={() => setActiveTab("recipient")}
            className={`flex-1 py-2 px-4 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all whitespace-nowrap ${activeTab === "recipient" ? "bg-white shadow-sm text-slate-800" : "text-slate-400 hover:text-slate-600"}`}
          >
            Recipient
          </button>
          <button
            onClick={() => setActiveTab("project")}
            className={`flex-1 py-2 px-4 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all whitespace-nowrap ${activeTab === "project" ? "bg-white shadow-sm text-slate-800" : "text-slate-400 hover:text-slate-600"}`}
          >
            Project
          </button>
          <button
            onClick={() => setActiveTab("signatories")}
            className={`flex-1 py-2 px-4 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all whitespace-nowrap ${activeTab === "signatories" ? "bg-white shadow-sm text-slate-800" : "text-slate-400 hover:text-slate-600"}`}
          >
            Sign-off
          </button>
          <button
            onClick={() => setActiveTab("history")}
            className={`flex-1 py-2 px-4 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all whitespace-nowrap ${activeTab === "history" ? "bg-white shadow-sm text-slate-800" : "text-slate-400 hover:text-slate-600"}`}
          >
            History
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-8 space-y-10 scrollbar-hide">
          {activeTab === "content" && (
            <div className="space-y-10 animate-in fade-in slide-in-from-bottom-4 duration-500">
              <section>
                <div className="flex justify-between items-end mb-4">
                  <h2 className="text-xs font-black text-slate-800 uppercase tracking-[0.2em]">
                    Intelligent Import
                  </h2>
                  {isParsing && (
                    <span className="text-[10px] font-bold text-brand-600 animate-pulse uppercase">
                      Parsing {parseProgress.current}/{parseProgress.total}...
                    </span>
                  )}
                </div>
                <div className="space-y-4">
                  <div className="relative group">
                    <textarea
                      className="w-full p-5 bg-slate-100/50 border border-slate-200 rounded-3xl text-xs font-medium focus:bg-white focus:ring-4 focus:ring-brand-500/5 focus:border-brand-500 outline-none transition-all placeholder-slate-300 min-h-[120px]"
                      placeholder="Paste text, file names, Google Drive file links, or folder links here..."
                      value={smartInput}
                      onChange={(e) => setSmartInput(e.target.value)}
                    />
                    <div className="absolute right-4 bottom-4 flex gap-2">
                      <button
                        onClick={handleSmartAnalysis}
                        disabled={!smartInput.trim() || isAnalyzingText}
                        className="p-2 bg-slate-900 text-white rounded-xl shadow-lg hover:shadow-slate-200 transition-all disabled:opacity-50 active:scale-90"
                      >
                        <svg
                          className={`w-4 h-4 ${isAnalyzingText ? "animate-spin" : ""}`}
                          fill="none"
                          stroke="currentColor"
                          viewBox="0 0 24 24"
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            strokeWidth="2.5"
                            d="M13 10V3L4 14h7v7l9-11h-7z"
                          ></path>
                        </svg>
                      </button>
                    </div>
                  </div>
                  <div className="grid grid-cols-2 gap-3">
                    <button
                      onClick={() => fileInputRef.current?.click()}
                      className="flex items-center justify-center gap-2 bg-white border border-slate-200 text-slate-700 py-3 rounded-2xl text-[10px] font-black uppercase tracking-widest hover:bg-slate-50 transition-all shadow-sm"
                    >
                      <svg
                        className="w-4 h-4 text-slate-400"
                        fill="none"
                        stroke="currentColor"
                        viewBox="0 0 24 24"
                      >
                        <path
                          strokeLinecap="round"
                          strokeLinejoin="round"
                          strokeWidth="2.5"
                          d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12"
                        ></path>
                      </svg>
                      Upload Files
                    </button>
                    <button
                      onClick={handleOpenDriveModal}
                      disabled={!isDriveReady}
                      className="flex items-center justify-center gap-2 bg-white border border-slate-200 text-slate-700 py-3 rounded-2xl text-[10px] font-black uppercase tracking-widest hover:bg-slate-50 transition-all shadow-sm"
                    >
                      <svg
                        className="w-4 h-4 text-slate-400"
                        fill="currentColor"
                        viewBox="0 0 24 24"
                      >
                        <path d="M7.71 3.5L1.15 15l2.86 5L10.57 8.5 7.71 3.5zm1.43 0l6.86 11.5h5.71l-6.86-11.5H9.14zm7.14 12.5H3.43l-2.28 4h13.71l2.28-4z" />
                      </svg>
                      Browse Drive
                    </button>
                  </div>
                  <input
                    type="file"
                    multiple
                    ref={fileInputRef}
                    className="hidden"
                    accept=".pdf,image/*"
                    onChange={handleBatchUpload}
                  />
                  {statusMsg && (
                    <div
                      className={`p-4 rounded-2xl text-[10px] font-bold border animate-in slide-in-from-right duration-300 ${statusType === "error" ? "bg-red-50 border-red-100 text-red-600" : "bg-brand-50 border-brand-100 text-brand-600"}`}
                    >
                      {statusMsg}
                    </div>
                  )}
                </div>
              </section>

              <section className="space-y-4">
                <div className="flex justify-between items-center">
                  <h2 className="text-xs font-black text-slate-800 uppercase tracking-[0.2em]">
                    Transmission
                  </h2>
                </div>
                <div className="p-6 bg-slate-50/50 rounded-[32px] border border-slate-100/80 space-y-4">
                  <div className="grid grid-cols-2 gap-4">
                    <label className="flex items-center gap-3 p-3 bg-white rounded-2xl border border-slate-100 cursor-pointer hover:border-brand-500 transition-all">
                      <input
                        type="checkbox"
                        checked={data.transmissionMethod.personalDelivery}
                        onChange={(e) =>
                          updateTransmission(
                            "personalDelivery",
                            e.target.checked,
                          )
                        }
                        className="w-4 h-4 rounded-md text-brand-600 border-slate-300"
                      />
                      <span className="text-[10px] font-black text-slate-600 uppercase tracking-widest">
                        Hand Delivered
                      </span>
                    </label>
                    <label className="flex items-center gap-3 p-3 bg-white rounded-2xl border border-slate-100 cursor-pointer hover:border-brand-500 transition-all">
                      <input
                        type="checkbox"
                        checked={data.transmissionMethod.pickUp}
                        onChange={(e) =>
                          updateTransmission("pickUp", e.target.checked)
                        }
                        className="w-4 h-4 rounded-md text-brand-600 border-slate-300"
                      />
                      <span className="text-[10px] font-black text-slate-600 uppercase tracking-widest">
                        Pick-up
                      </span>
                    </label>
                    <label className="flex items-center gap-3 p-3 bg-white rounded-2xl border border-slate-100 cursor-pointer hover:border-brand-500 transition-all">
                      <input
                        type="checkbox"
                        checked={data.transmissionMethod.grabLalamove}
                        onChange={(e) =>
                          updateTransmission("grabLalamove", e.target.checked)
                        }
                        className="w-4 h-4 rounded-md text-brand-600 border-slate-300"
                      />
                      <span className="text-[10px] font-black text-slate-600 uppercase tracking-widest">
                        Courier
                      </span>
                    </label>
                    <label className="flex items-center gap-3 p-3 bg-white rounded-2xl border border-slate-100 cursor-pointer hover:border-brand-500 transition-all">
                      <input
                        type="checkbox"
                        checked={data.transmissionMethod.registeredMail}
                        onChange={(e) =>
                          updateTransmission("registeredMail", e.target.checked)
                        }
                        className="w-4 h-4 rounded-md text-brand-600 border-slate-300"
                      />
                      <span className="text-[10px] font-black text-slate-600 uppercase tracking-widest">
                        Registered Mail
                      </span>
                    </label>
                  </div>
                  <div className="space-y-1">
                    <Label className="text-[10px] font-black text-muted-foreground uppercase tracking-widest ml-1">
                      Internal Notes
                    </Label>
                    <Textarea
                      rows={2}
                      value={data.notes}
                      onChange={(e) => handleUpdateNotes(e.target.value)}
                    />
                  </div>
                </div>
              </section>
            </div>
          )}

          {activeTab === "sender" && (
            <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
              <h2 className="text-xs font-black text-slate-800 uppercase tracking-[0.2em]">
                Sender Branding
              </h2>

              <div className="flex flex-col sm:flex-row gap-3 items-end">
                <div className="flex-1 space-y-1">
                  <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                    Saved Agencies
                  </Label>
                  <select
                    className="w-full h-9 rounded-md border border-input bg-transparent px-2.5 py-1 text-sm shadow-xs transition-[color,box-shadow] focus-visible:border-ring focus-visible:ring-ring/50 focus-visible:ring-3 outline-none"
                    value={selectedAgencyId}
                    onChange={(e) => setSelectedAgencyId(e.target.value)}
                  >
                    <option value="">Select agency...</option>
                    {agencies.map((agency) => (
                      <option key={agency.id} value={agency.id}>
                        {agency.name}
                      </option>
                    ))}
                  </select>
                </div>
                <Button
                  onClick={openAgencyModal}
                  className="h-[52px] px-5 rounded-2xl text-[10px] font-black uppercase tracking-[0.2em] shadow-lg"
                >
                  Add
                </Button>
              </div>

              <div className="flex flex-col items-center gap-6 p-6 bg-slate-50 rounded-[40px] border border-slate-200/60">
                <div
                  className="relative group cursor-pointer"
                  onClick={() => logoInputRef.current?.click()}
                >
                  <div className="w-32 h-32 rounded-full border-4 border-white shadow-xl flex items-center justify-center bg-white overflow-hidden transition-transform group-hover:scale-105">
                    {data.sender.logoBase64 ? (
                      <img
                        src={data.sender.logoBase64}
                        alt="Company Logo"
                        className="w-full h-full object-contain"
                      />
                    ) : (
                      <div className="flex flex-col items-center text-slate-300">
                        <svg
                          className="w-8 h-8 mb-1"
                          fill="none"
                          stroke="currentColor"
                          viewBox="0 0 24 24"
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            strokeWidth="2"
                            d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z"
                          ></path>
                        </svg>
                        <span className="text-[8px] font-black uppercase tracking-widest">
                          Add Logo
                        </span>
                      </div>
                    )}
                  </div>
                  <input
                    key={logoInputKey}
                    type="file"
                    ref={logoInputRef}
                    className="hidden"
                    accept="image/*"
                    onChange={handleLogoUpload}
                  />
                </div>
                <div className="text-center">
                  <p className="text-[10px] font-black text-slate-800 uppercase tracking-widest">
                    {data.sender.agencyName}
                  </p>
                </div>
              </div>
              <div className="space-y-5">
                <div className="space-y-1">
                  <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                    Agency Name
                  </Label>
                  <Input
                    value={data.sender.agencyName}
                    onChange={(e) =>
                      updateField("sender", "agencyName", e.target.value)
                    }
                  />
                </div>
                <div className="space-y-1">
                  <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                    Website
                  </Label>
                  <Input
                    value={data.sender.website}
                    onChange={(e) =>
                      updateField("sender", "website", e.target.value)
                    }
                  />
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                      Telephone
                    </Label>
                    <Input
                      value={data.sender.telephone}
                      onChange={(e) =>
                        updateField("sender", "telephone", e.target.value)
                      }
                    />
                  </div>
                  <div className="space-y-1">
                    <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                      Mobile
                    </Label>
                    <Input
                      value={data.sender.mobile}
                      onChange={(e) =>
                        updateField("sender", "mobile", e.target.value)
                      }
                    />
                  </div>
                </div>
                <div className="space-y-1">
                  <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                    Email
                  </Label>
                  <Input
                    value={data.sender.email}
                    onChange={(e) =>
                      updateField("sender", "email", e.target.value)
                    }
                  />
                </div>
              </div>
            </div>
          )}

          {activeTab === "recipient" && (
            <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
              <h2 className="text-xs font-black text-slate-800 uppercase tracking-[0.2em] mb-4">
                Recipient Dossier
              </h2>
              <div className="space-y-4">
                <div className="space-y-1">
                  <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                    To
                  </Label>
                  <Input
                    className="opacity-50 cursor-not-allowed bg-slate-50"
                    value={data.recipient.to}
                    disabled
                  />
                </div>
                <div className="space-y-1">
                  <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                    Organization
                  </Label>
                  <Input
                    value={data.recipient.company}
                    onChange={(e) =>
                      updateField("recipient", "company", e.target.value)
                    }
                  />
                </div>
                <div className="space-y-1">
                  <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                    Attention
                  </Label>
                  <Input
                    value={data.recipient.attention}
                    onChange={(e) =>
                      updateField("recipient", "attention", e.target.value)
                    }
                  />
                </div>
                <div className="space-y-1">
                  <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                    Full Address
                  </Label>
                  <ExpandingTextarea
                    value={data.recipient.address}
                    onChange={(v) => updateField("recipient", "address", v)}
                    placeholder="Full address..."
                  />
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                      Contact No.
                    </Label>
                    <Input
                      value={data.recipient.contactNumber}
                      onChange={(e) =>
                        updateField(
                          "recipient",
                          "contactNumber",
                          e.target.value,
                        )
                      }
                    />
                  </div>
                  <div className="space-y-1">
                    <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                      Email
                    </Label>
                    <Input
                      value={data.recipient.email}
                      onChange={(e) =>
                        updateField("recipient", "email", e.target.value)
                      }
                    />
                  </div>
                </div>
              </div>
            </div>
          )}

          {activeTab === "project" && (
            <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
              <h2 className="text-xs font-black text-slate-800 uppercase tracking-[0.2em] mb-4">
                Project Parameters
              </h2>
              <div className="space-y-4">
                <div className="space-y-1">
                  <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                    Contract/Project Title
                  </Label>
                  <Input
                    value={data.project.projectName}
                    onChange={(e) =>
                      updateField("project", "projectName", e.target.value)
                    }
                  />
                </div>
                <div className="space-y-1">
                  <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                    Transmittal ID
                  </Label>
                  <Input
                    className="font-mono text-[10px]"
                    value={data.project.transmittalNumber}
                    onChange={(e) =>
                      updateField(
                        "project",
                        "transmittalNumber",
                        e.target.value,
                      )
                    }
                  />
                </div>
                <div className="space-y-1">
                  <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                    Purpose
                  </Label>
                  <Input
                    value={data.project.purpose}
                    onChange={(e) =>
                      updateField("project", "purpose", e.target.value)
                    }
                  />
                </div>
              </div>
            </div>
          )}

          {activeTab === "signatories" && (
            <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
              <h2 className="text-xs font-black text-slate-800 uppercase tracking-[0.2em] mb-4">
                Signatories & Auth
              </h2>
              <div className="space-y-6">
                <div className="p-6 bg-slate-50 rounded-[32px] border border-slate-100 space-y-4">
                  <div className="space-y-1">
                    <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                      Prepared By Name
                    </Label>
                    <Input
                      value={data.signatories.preparedBy}
                      onChange={(e) =>
                        handleUpdateSignatory("preparedBy", e.target.value)
                      }
                    />
                  </div>
                  <div className="space-y-1">
                    <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                      Prepared By Role
                    </Label>
                    <Input
                      value={data.signatories.preparedByRole}
                      onChange={(e) =>
                        handleUpdateSignatory("preparedByRole", e.target.value)
                      }
                    />
                  </div>
                </div>
                <div className="p-6 bg-slate-50 rounded-[32px] border border-slate-100 space-y-4">
                  <div className="space-y-1">
                    <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                      Noted By Name
                    </Label>
                    <Input
                      value={data.signatories.notedBy}
                      onChange={(e) =>
                        handleUpdateSignatory("notedBy", e.target.value)
                      }
                    />
                  </div>
                  <div className="space-y-1">
                    <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                      Noted By Role
                    </Label>
                    <Input
                      value={data.signatories.notedByRole}
                      onChange={(e) =>
                        handleUpdateSignatory("notedByRole", e.target.value)
                      }
                    />
                  </div>
                </div>
                <div className="space-y-1">
                  <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
                    Time Released
                  </Label>
                  <Input
                    type="time"
                    step={60}
                    value={parseTimeTo24h(data.signatories.timeReleased)}
                    onChange={(e) =>
                      handleUpdateSignatory(
                        "timeReleased",
                        formatTime24hTo12h(e.target.value),
                      )
                    }
                  />
                </div>
              </div>
            </div>
          )}

          {activeTab === "history" && (
            <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
              <div className="flex justify-between items-center mb-4">
                <h2 className="text-xs font-black text-slate-800 uppercase tracking-[0.2em]">
                  Workspace Snapshots
                </h2>
                <span className="text-[9px] font-black text-slate-400 uppercase tracking-widest">
                  {history.length}/20
                </span>
              </div>
              {history.length === 0 ? (
                <div className="text-center py-20 bg-slate-50/50 rounded-[40px] border border-dashed border-slate-200">
                  <p className="text-[10px] font-black text-slate-300 uppercase tracking-widest">
                    No Recent Snapshots
                  </p>
                </div>
              ) : (
                <div className="space-y-3">
                  {history.map((item) => (
                    <div
                      key={item.id}
                      onClick={() => loadFromHistory(item)}
                      className="group p-5 bg-white border border-slate-200 rounded-[32px] cursor-pointer hover:border-brand-500 hover:shadow-xl transition-all relative overflow-hidden"
                    >
                      <div className="flex justify-between items-start mb-2">
                        <div className="shrink-0 w-8 h-8 rounded-full bg-slate-50 flex items-center justify-center text-slate-400 group-hover:bg-brand-50 group-hover:text-brand-500 transition-all">
                          <svg
                            className="w-4 h-4"
                            fill="none"
                            stroke="currentColor"
                            viewBox="0 0 24 24"
                          >
                            <path
                              strokeLinecap="round"
                              strokeLinejoin="round"
                              strokeWidth="2.5"
                              d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z"
                            ></path>
                          </svg>
                        </div>
                        <button
                          onClick={(e) => deleteHistoryItem(item.id, e)}
                          className="p-2 text-slate-300 hover:text-red-500 opacity-0 group-hover:opacity-100 transition-all"
                        >
                          ✕
                        </button>
                      </div>
                      <div className="space-y-1">
                        <p className="text-[10px] font-black text-slate-800 line-clamp-1">
                          {item.projectName}
                        </p>
                        <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">
                          {item.transmittalNumber}
                        </p>
                        <div className="flex flex-wrap gap-x-4 gap-y-1 text-[9px] text-slate-500">
                          <span className="font-semibold text-slate-600">
                            {item.createdBy}
                          </span>
                          <span>Prepared: {item.preparedBy || "—"}</span>
                          <span>Noted: {item.notedBy || "—"}</span>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
        </div>

        <div className="p-8 border-t border-slate-200 bg-white shadow-[0_-10px_30px_rgba(0,0,0,0.03)]">
          <div className="mb-3">
            <Button
              onClick={handleSaveTransmittal}
              className="w-full py-3 rounded-[24px] bg-brand-600 text-white text-[10px] font-black uppercase tracking-[0.2em] shadow-lg hover:bg-brand-500 active:scale-95"
            >
              Save to History
            </Button>
          </div>
          <div className="grid grid-cols-2 gap-3 mb-3">
            <Button
              onClick={handlePrint}
              disabled={isGeneratingPdf}
              className="flex flex-col items-center justify-center py-4 h-auto bg-slate-900 text-white rounded-[24px] shadow-2xl hover:bg-slate-800 active:scale-95 group"
            >
              <svg
                className="w-5 h-5 mb-1 group-hover:scale-110 transition-transform"
                fill="none"
                stroke="currentColor"
                viewBox="0 0 24 24"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth="2"
                  d="M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2z"
                ></path>
              </svg>
              <span className="text-[9px] font-black uppercase tracking-[0.2em]">
                Export PDF
              </span>
            </Button>
            <Button
              onClick={handleDownloadDocx}
              disabled={isGeneratingDocx}
              variant="outline"
              className="flex flex-col items-center justify-center py-4 h-auto rounded-[24px] hover:border-brand-500 hover:text-brand-600 active:scale-95 group"
            >
              <svg
                className="w-5 h-5 mb-1 group-hover:scale-110 transition-transform"
                fill="none"
                stroke="currentColor"
                viewBox="0 0 24 24"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth="2"
                  d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                ></path>
              </svg>
              <span className="text-[9px] font-black uppercase tracking-[0.2em]">
                Export Word
              </span>
            </Button>
          </div>
        </div>
      </div>

      <div
        ref={previewContainerRef}
        className={`${showPreview ? "flex" : "hidden"} lg:flex flex-1 h-full overflow-y-auto overflow-x-hidden bg-slate-200 p-4 lg:p-8 justify-start custom-scrollbar w-full absolute inset-0 lg:static z-10 flex-col items-center`}
      >
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
