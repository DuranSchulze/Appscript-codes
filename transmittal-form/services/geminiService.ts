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

    return {
        items: [{
            description,
            documentNumber: generateSmartPlaceholder("File", description, safeFileName),
            qty: "1",
            remarks,
            documentType: "File",
        }],
    };
};

// --- ROBUST LOCAL PARSER (Fallback) ---
const parseWithRegex = (content: string): ParseResult => {
    const lines = content.split(/\r?\n/);
    const cleanedItems: Array<any> = [];
    const seenDescriptions = new Set<string>();
    
    for (let line of lines) {
        line = line.trim();
        if (line.length < 3 || line.startsWith('http')) continue;
        
        if (/^(me|owner|file size|name|folders)$/i.test(line)) continue;

        if (!seenDescriptions.has(line)) {
            cleanedItems.push({
                description: line,
                documentNumber: generateSmartPlaceholder("File", line, line),
                qty: "1",
                remarks: "",
                documentType: "File"
            });
            seenDescriptions.add(line);
        }
    }
    return { items: cleanedItems };
};

/**
 * Uses Gemini to analyze documents and extract metadata.
 */
export const parseTransmittalDocument = async (
    content: string, 
    mimeType: string, 
    isText: boolean = false,
    fileName?: string
): Promise<ParseResult> => {

  try {
    const apiKey = resolveGeminiApiKey();
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

        const processedItems = items.map((item: any) => {
            if (!item.documentNumber || item.documentNumber.trim() === "") {
                const smartId = generateSmartPlaceholder(
                    item.documentType || "", 
                    item.description || "", 
                    fileName || item.originalFilename || ""
                );
                return { ...item, documentNumber: smartId };
            }
            return item;
        });

        return {
            items: processedItems,
            header: data.extractedHeader || undefined
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