import { GoogleGenAI, GenerateContentResponse, Type } from "@google/genai";

// Schema Definition based on Structured Output Generation technique
const RESPONSE_SCHEMA = {
  type: Type.OBJECT,
  properties: {
    extractedHeader: {
        type: Type.OBJECT,
        description: "Global document header information.",
        properties: {
            recipientName: { type: Type.STRING, description: "Name of the primary recipient or subject." },
            recipientEmail: { type: Type.STRING, description: "Email address associated with the recipient." },
            companyName: { type: Type.STRING, description: "Company or organization name of the recipient." },
            address: { type: Type.STRING, description: "Full mailing or physical address found in the header." },
            projectName: { type: Type.STRING, description: "Title of the project or engagement." },
            projectNumber: { type: Type.STRING, description: "Reference number, ID, or ticket number for the project." },
            purpose: { type: Type.STRING, description: "The stated purpose of the document or transmittal." }
        },
        nullable: true
    },
    items: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          originalFilename: { type: Type.STRING, description: "The source file name." },
          qty: { type: Type.STRING, description: "Quantity of the item (usually '1')." },
          documentType: { type: Type.STRING, description: "Category of the document (e.g., Report, Invoice, Certificate)." },
          documentNumber: { type: Type.STRING, description: "Unique identifier: Reference No, Control No, or ID." },
          description: { type: Type.STRING, description: "Clear title or description of the document." },
          remarks: { type: Type.STRING, description: "Additional notes, status, or remarks." }
        },
        required: ["qty", "documentType", "documentNumber", "description", "remarks"]
      }
    }
  },
  required: ["items"]
};

const DOCUMENT_CONTROLLER_PROMPT = `
  Role: Professional Document Controller.
  Task: Analyze the provided file and extract metadata for ONE single entry in a formal Transmittal Form.

  CRITICAL RULE — ONE FILE = ONE ITEM:
  - This file represents a SINGLE physical document being transmitted.
  - Your output MUST contain exactly ONE item in the items array, no more.
  - Do NOT enumerate pages, attachments, or documents referenced inside the file.
  - Focus ONLY on the first/cover page to identify what this document IS.

  EXTRACTION HIERARCHY:
  1. **Header Identification**: Locate the primary recipient or entity the document is addressed to. Identify the requesting organization or client company.
  2. **Project Context**: Look for project titles, engagement references, or case numbers.
  3. **Document Identity**: From the FIRST PAGE ONLY, determine the single best title/description and unique identifier for this document as a whole.
  4. **Document # / Ref # Priority**: Prefer explicit identifiers in this order: Document No, Reference No, Control No, OR No (Official Receipt No), Invoice No, Receipt No, Title No, eCAR No, Tax Declaration No, Certificate No, and serial/control identifiers.
  5. **Remarks Priority**: If the document indicates copy status, include it in remarks (e.g., Original, Certified True Copy, Photocopy, Scanned Copy).

  FORMATTING RULES:
  - Output exactly ONE item representing the whole file — never split into multiple items.
  - If a specific field is missing, leave it as null or an empty string.
  - For 'Qty', always use '1'.
  - Be precise with names and ID numbers.
  - 'documentNumber' must contain the strongest explicit identifier found on the cover page, not a generic label.
  - Keep punctuation and alphanumeric format of identifiers (e.g., OR-12345, eCAR-2024-0012).
  - Put copy-status notes in 'remarks' when present.
  
  Output the result strictly in the provided JSON schema format.
`;

export interface ParseResult {
    items: Array<{ description: string, documentNumber: string, qty: string, remarks: string, documentType?: string }>;
    header?: {
        recipientName?: string;
        recipientEmail?: string;
        companyName?: string;
        address?: string;
        projectName?: string;
        projectNumber?: string;
        purpose?: string;
    };
    error?: string;
    fallbackCount?: number;
}

export interface DocumentNumberResolutionInput {
    currentDocumentNumber?: string | null;
    sourceName?: string | null;
    description?: string | null;
    documentType?: string | null;
}

// --- Helper: Generate Smart Placeholder for Missing IDs ---
const generateSmartPlaceholder = (type: string, description: string, filename: string): string => {
    const commonIdPattern = /(?:INV|PO|SOA|REF|CERT|CTR|NUM|NO)[-._\s#]+([A-Z0-9-]{3,})/i;
    const match = filename.match(commonIdPattern);
    if (match && match[1]) return match[1].toUpperCase();

    const digitMatch = filename.match(/(\d{4,})/);
    if (digitMatch) return digitMatch[1];

    const text = (type + " " + description).toUpperCase();
    let prefix = "DOC";
    
    if (text.includes("INVOICE") || text.includes("BILL")) prefix = "INV";
    else if (text.includes("CERTIFICATE") || text.includes("COMPLETION")) prefix = "CERT";
    else if (text.includes("STATEMENT") || text.includes("SOA")) prefix = "SOA";
    else if (text.includes("PURCHASE") || text.includes("ORDER")) prefix = "PO";
    else if (text.includes("CONTRACT") || text.includes("AGREEMENT")) prefix = "CTR";
    else if (text.includes("RECEIPT")) prefix = "OR";
    else if (text.includes("VOUCHER")) prefix = "CV";
    else if (text.includes("LETTER")) prefix = "LTR";
    else if (text.includes("REPORT")) prefix = "REP";

    return `${prefix}-PENDING`;
};

const stripFileExtension = (value: string): string =>
    value.replace(/\.[^/.]+$/, "").trim();

const GENERIC_DOCUMENT_NUMBER_VALUES = new Set([
    "scan",
    "scanned",
    "file",
    "document",
    "documents",
    "google drive",
    "drive",
    "bulk import",
    "browse drive",
    "from google drive",
    "via drive folder",
    "via drive link",
    "doc",
    "doc no",
    "document no",
    "document number",
    "reference",
    "reference no",
    "reference number",
    "ref",
    "ref no",
    "control",
    "control no",
    "control number",
    "number",
    "id",
    "identifier",
    "official receipt",
    "invoice",
    "receipt",
    "certificate",
    "transmittal",
    "n a",
    "na",
    "none",
    "null",
    "nil",
    "unknown",
    "not available",
    "not found",
    "tbd",
    "to be determined",
    "pending",
]);

type WeakCandidateContext = {
    sourceName?: string | null;
    description?: string | null;
};

const normalizeComparableText = (value: string): string =>
    value
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, " ")
        .trim();

const sanitizeDocumentNumberCandidate = (value: string): string => {
    const raw = String(value || "").trim();
    if (!raw) return "";

    let cleaned = raw
        .replace(/[“”]/g, '"')
        .replace(/[‘’]/g, "'")
        .replace(/[–—]/g, "-")
        .replace(/\s+/g, " ")
        .trim();

    cleaned = cleaned.replace(/^([\[\(\{\"'])+|([\]\)\}\"'])+$/g, "").trim();
    cleaned = cleaned.replace(
        /^(?:doc(?:ument)?|ref(?:erence)?|control|transmittal|invoice|inv|official\s*receipt|receipt|voucher|cv|po|soa|cert(?:ificate)?|title|tax\s*declaration|td|e\s*car|ecar|no\.?|number)\s*(?:no\.?|number|#)?\s*[:\-]\s*/i,
        "",
    );
    cleaned = cleaned.replace(/[.,;:]+$/g, "").trim();

    return cleaned.toUpperCase();
};

const isLikelyDateToken = (value: string): boolean => {
    const normalized = value.replace(/[._]/g, "-").trim();
    if (!normalized) return false;

    if (/^\d{4}-\d{2}-\d{2}$/.test(normalized)) return true;
    if (/^\d{2}-\d{2}-\d{4}$/.test(normalized)) return true;
    if (/^\d{8}$/.test(normalized)) return true;

    if (/^\d{4}$/.test(normalized)) {
        const year = Number(normalized);
        return year >= 1900 && year <= 2100;
    }

    return false;
};

const isWeakDocumentNumberCandidate = (
    value: string,
    context: WeakCandidateContext = {},
): boolean => {
    const candidate = sanitizeDocumentNumberCandidate(value);
    if (!candidate) return true;

    const comparable = normalizeComparableText(candidate);
    if (!comparable) return true;
    if (GENERIC_DOCUMENT_NUMBER_VALUES.has(comparable)) return true;

    if (/\b(?:unable|cannot|could not|not\s+(?:provided|available|found|indicated|specified))\b/.test(comparable)) {
        return true;
    }

    if (/^(?:https?:\/\/|www\.)/i.test(candidate)) return true;
    if (candidate.length > 80) return true;

    const hasDigit = /\d/.test(candidate);
    const hasStructuredSeparator = /[-_/]/.test(candidate);
    const words = comparable.split(" ").filter(Boolean).length;

    if (!hasDigit && !hasStructuredSeparator && words >= 4) return true;
    if (!hasDigit && !hasStructuredSeparator && candidate.length <= 2) return true;

    const comparableSource = normalizeComparableText(
        stripFileExtension(String(context.sourceName || "")),
    );
    if (
        comparableSource &&
        comparable === comparableSource &&
        !hasDigit &&
        !hasStructuredSeparator &&
        words >= 3
    ) {
        return true;
    }

    return false;
};

const extractLabeledIdentifier = (text: string): string => {
    const source = String(text || "");
    if (!source) return "";

    const pattern =
        /\b(?:doc(?:ument)?|ref(?:erence)?|control|transmittal|invoice|inv|official\s*receipt|or|receipt|voucher|cv|po|soa|cert(?:ificate)?|title|tax\s*declaration|td|e\s*car|ecar)\s*(?:no\.?|number|#)?\s*[:\-]?\s*([A-Za-z0-9][A-Za-z0-9\-_/]{1,})\b/gi;

    let match: RegExpExecArray | null;
    while ((match = pattern.exec(source))) {
        const candidate = sanitizeDocumentNumberCandidate(match[1] || "");
        if (!candidate || isLikelyDateToken(candidate)) continue;
        if (!isWeakDocumentNumberCandidate(candidate, { sourceName: source })) {
            return candidate;
        }
    }

    return "";
};

const extractTokenIdentifier = (text: string): string => {
    const source = String(text || "");
    if (!source) return "";

    const tokens =
        source.match(/\b[A-Za-z0-9]+(?:[-_/][A-Za-z0-9]+)*\b/g) || [];

    for (const token of tokens) {
        const cleanToken = token.replace(/^[-_/]+|[-_/]+$/g, "");
        if (cleanToken.length < 4) continue;

        const hasLetters = /[A-Za-z]/.test(cleanToken);
        const hasDigits = /\d/.test(cleanToken);
        const hasSeparator = /[-_/]/.test(cleanToken);

        if (!(hasDigits && (hasLetters || hasSeparator))) continue;
        if (isLikelyDateToken(cleanToken)) continue;

        const candidate = sanitizeDocumentNumberCandidate(cleanToken);
        if (!isWeakDocumentNumberCandidate(candidate, { sourceName: source })) {
            return candidate;
        }
    }

    for (const token of tokens) {
        const cleanToken = token.replace(/^[-_/]+|[-_/]+$/g, "");
        if (!/^\d{5,}$/.test(cleanToken)) continue;
        if (isLikelyDateToken(cleanToken)) continue;

        const candidate = sanitizeDocumentNumberCandidate(cleanToken);
        if (!isWeakDocumentNumberCandidate(candidate, { sourceName: source })) {
            return candidate;
        }
    }

    return "";
};

const extractIdentifierFromText = (text: string): string => {
    const fromLabel = extractLabeledIdentifier(text);
    if (fromLabel) return fromLabel;

    return extractTokenIdentifier(text);
};

export const resolveDocumentNumberWithFallback = ({
    currentDocumentNumber,
    sourceName,
    description,
    documentType,
}: DocumentNumberResolutionInput): string => {
    const safeSourceName = String(sourceName || "").trim();
    const safeDescription = String(description || "").trim();
    const safeDocumentType = String(documentType || "").trim() || "File";

    const current = sanitizeDocumentNumberCandidate(
        String(currentDocumentNumber || ""),
    );
    if (
        current &&
        !isWeakDocumentNumberCandidate(current, {
            sourceName: safeSourceName,
            description: safeDescription,
        })
    ) {
        return current;
    }

    const sourceBase = stripFileExtension(safeSourceName);
    const textCandidates = [safeSourceName, sourceBase, safeDescription].filter(
        Boolean,
    );

    for (const text of textCandidates) {
        const extracted = extractIdentifierFromText(text);
        if (extracted) return extracted;
    }

    const filenameDerived = sanitizeDocumentNumberCandidate(
        generateSmartPlaceholder(
            safeDocumentType,
            safeDescription || sourceBase || safeSourceName || "Document",
            safeSourceName || safeDescription || "Uploaded Document",
        ),
    );
    if (
        filenameDerived &&
        !isWeakDocumentNumberCandidate(filenameDerived, {
            sourceName: safeSourceName,
            description: safeDescription,
        })
    ) {
        return filenameDerived;
    }

    const deterministic = sanitizeDocumentNumberCandidate(
        generateSmartPlaceholder(safeDocumentType, "Document", "Uploaded Document"),
    );
    return deterministic || "DOC-PENDING";
};

const resolveGeminiApiKey = (): string => {
    return (
        process.env.NEXT_PUBLIC_GEMINI_API_KEY ||
        process.env.NEXT_PUBLIC_GOOGLE_API_KEY ||
        process.env.GEMINI_API_KEY ||
        process.env.GOOGLE_API_KEY ||
        ""
    );
};

const resolveGeminiModel = (): string => {
    return process.env.GEMINI_MODEL || "gemini-2.5-flash";
};

const buildFallbackParseResult = (fileName?: string, remarks: string = ""): ParseResult => {
    const safeFileName = (fileName || "Uploaded Document").trim();
    const description = stripFileExtension(safeFileName) || safeFileName;
    const documentNumber = resolveDocumentNumberWithFallback({
        sourceName: safeFileName,
        description,
        documentType: "File",
    });

    return {
        items: [{
            description,
            documentNumber,
            qty: "1",
            remarks,
            documentType: "File",
        }],
        fallbackCount: 1,
    };
};

// --- ROBUST LOCAL PARSER (Fallback) ---
const parseWithRegex = (content: string): ParseResult => {
    const lines = content.split(/\r?\n/);
    const cleanedItems: ParseResult["items"] = [];
    const seenDescriptions = new Set<string>();
    let fallbackCount = 0;
    
    for (let line of lines) {
        line = line.trim();
        if (line.length < 3 || line.startsWith('http')) continue;
        
        if (/^(me|owner|file size|name|folders)$/i.test(line)) continue;

        if (!seenDescriptions.has(line)) {
            const documentNumber = resolveDocumentNumberWithFallback({
                sourceName: line,
                description: line,
                documentType: "File",
            });
            cleanedItems.push({
                description: line,
                documentNumber,
                qty: "1",
                remarks: "",
                documentType: "File"
            });
            fallbackCount += 1;
            seenDescriptions.add(line);
        }
    }
    return { items: cleanedItems, fallbackCount };
};

/**
 * Uses Gemini to analyze documents and extract metadata.
 */
export const parseTransmittalDocument = async (
    content: string, 
    mimeType: string, 
    isText: boolean = false,
    fileName?: string,
    keyOverride?: string
): Promise<ParseResult> => {

  try {
    const apiKey = String(keyOverride || "").trim() || resolveGeminiApiKey();
    if (!apiKey) {
        const userMessage = "AI parser is not configured (missing Gemini API key).";

        if (isText) {
            return parseWithRegex(content);
        }

        return {
            ...buildFallbackParseResult(fileName, userMessage),
            error: userMessage,
        };
    }

    const ai = new GoogleGenAI({ apiKey });
    
    let parts: any[] = [];

    if (isText) {
        parts.push({ text: `File/Text List to Analyze:\n${content}` });
    } else {
        parts.push({
            inlineData: {
                mimeType: mimeType,
                data: content
            }
        });
        parts.push({ text: `Analyze this document file. Filename: ${fileName || 'Unknown'}` });
    }
    
    parts.push({ text: DOCUMENT_CONTROLLER_PROMPT });

    const response: GenerateContentResponse = await ai.models.generateContent({
        model: resolveGeminiModel(),
        contents: { parts },
        config: {
            responseMimeType: "application/json",
            responseSchema: RESPONSE_SCHEMA,
            temperature: 0.0,
        }
    });

    const text = response.text || "{}";
    let data;
    try {
        data = JSON.parse(text);
    } catch (e) {
        console.error("Failed to parse JSON", text);
        const cleanText = text.replace(/```json/g, "").replace(/```/g, "").trim();
        data = JSON.parse(cleanText);
    }

    if (data && (data.items || data.extractedHeader)) {
        const items = Array.isArray(data.items) ? data.items : [];
        let fallbackCount = 0;

        const processedItems = items.map((item: any) => {
            const sourceName = fileName || item.originalFilename || "";
            const description = item.description || "";
            const currentDocumentNumber = String(item.documentNumber || "");
            const shouldFallback = isWeakDocumentNumberCandidate(currentDocumentNumber, {
                sourceName,
                description,
            });
            const documentNumber = resolveDocumentNumberWithFallback({
                currentDocumentNumber,
                sourceName,
                description,
                documentType: item.documentType || "",
            });

            if (shouldFallback) {
                fallbackCount += 1;
            }

            return { ...item, documentNumber };
        });

        return {
            items: processedItems,
            header: data.extractedHeader || undefined,
            fallbackCount: fallbackCount > 0 ? fallbackCount : undefined,
        };
    }
    
    throw new Error("AI returned invalid data structure.");

  } catch (error: any) {
    console.warn("Gemini Service Failed:", error);
    
    let userMessage = error.message || "Unknown analysis error.";
    if (error.message?.includes("401") || error.message?.includes("API key")) {
        userMessage = "Invalid API Key. Please check your credentials.";
    } else if (error.message?.includes("429") || error.message?.includes("quota")) {
        userMessage = "AI Quota Exceeded. Try again later.";
    } else if (error.message?.includes("500") || error.message?.includes("Internal")) {
        userMessage = "AI Service Error. Please retry.";
    }

    if (isText) {
        return parseWithRegex(content);
    }
    
    const fallbackResult = buildFallbackParseResult(fileName, userMessage);
    fallbackResult.error = userMessage;
    return fallbackResult;
  };
};
