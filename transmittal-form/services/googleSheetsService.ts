// Google Sheets Service - Appends transmittal rows to a linked spreadsheet

import { getGoogleAccessToken } from "./googleDriveService";

const SHEET_ID_KEY = "transmittal_linked_sheet_id";

/** Persist the linked spreadsheet ID in localStorage */
export const setLinkedSheetId = (sheetId: string) => {
  if (typeof window !== "undefined") {
    window.localStorage.setItem(SHEET_ID_KEY, sheetId);
  }
};

/** Read the linked spreadsheet ID from localStorage */
export const getLinkedSheetId = (): string | null => {
  if (typeof window === "undefined") return null;
  return window.localStorage.getItem(SHEET_ID_KEY) || null;
};

/** Clear the linked spreadsheet ID */
export const clearLinkedSheetId = () => {
  if (typeof window !== "undefined") {
    window.localStorage.removeItem(SHEET_ID_KEY);
  }
};

/** Extract spreadsheet ID from a Google Sheets URL */
export const extractSheetIdFromUrl = (url: string): string | null => {
  // Matches: /spreadsheets/d/SPREADSHEET_ID
  const regex = /\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/;
  const match = url.match(regex);
  return match ? match[1] : null;
};

/** Get the spreadsheet title to confirm it's accessible */
export const getSpreadsheetTitle = async (
  spreadsheetId: string,
): Promise<string> => {
  const accessToken = await getGoogleAccessToken();
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=properties.title`;
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!response.ok) {
    throw new Error(`Cannot access spreadsheet: ${response.status}`);
  }
  const data = await response.json();
  return data.properties?.title || "Untitled";
};

/**
 * Ensure the header row exists on Sheet1. If the sheet is empty, write headers first.
 */
const ensureHeaders = async (
  accessToken: string,
  spreadsheetId: string,
): Promise<void> => {
  const rangeUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Sheet1!A1:J1`;
  const res = await fetch(rangeUrl, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) return; // skip – append will still work

  const data = await res.json();
  const firstRow = data.values?.[0];

  if (!firstRow || firstRow.length === 0) {
    // Write header row
    const headers = [
      "Transmittal No.",
      "Date",
      "Project Name",
      "Recipient",
      "Company",
      "Prepared By",
      "Noted By",
      "Items Count",
      "Transmission Method",
      "Created At",
    ];

    await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Sheet1!A1:J1?valueInputOption=RAW`,
      {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ values: [headers] }),
      },
    );
  }
};

/** Check if a URL is a Google Sheets link */
export const isSheetUrl = (url: string): boolean => {
  return /docs\.google\.com\/spreadsheets\/d\//.test(url);
};

/**
 * Read all data rows from the first sheet of a Google Spreadsheet.
 * Returns { headers, rows } where rows are string arrays.
 */
export const readSheetRows = async (
  spreadsheetId: string,
): Promise<{ headers: string[]; rows: string[][] }> => {
  const accessToken = await getGoogleAccessToken();
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Sheet1?majorDimension=ROWS`;
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!response.ok) {
    throw new Error(`Cannot read spreadsheet: ${response.status}`);
  }
  const data = await response.json();
  const allRows: string[][] = data.values || [];
  if (allRows.length === 0) return { headers: [], rows: [] };
  const [headers, ...rows] = allRows;
  return { headers, rows };
};

export interface TransmittalRowData {
  transmittalNumber: string;
  date: string;
  projectName: string;
  recipientName: string;
  recipientCompany: string;
  preparedBy: string;
  notedBy: string;
  itemsCount: number;
  transmissionMethod: string;
}

/**
 * Append a transmittal summary row to the linked Google Sheet.
 * Returns true on success, false if no sheet is linked or on error.
 */
export const appendTransmittalRow = async (
  row: TransmittalRowData,
): Promise<boolean> => {
  const spreadsheetId = getLinkedSheetId();
  if (!spreadsheetId) return false;

  try {
    const accessToken = await getGoogleAccessToken();

    // Ensure headers exist
    await ensureHeaders(accessToken, spreadsheetId);

    const values = [
      [
        row.transmittalNumber,
        row.date,
        row.projectName,
        row.recipientName,
        row.recipientCompany,
        row.preparedBy,
        row.notedBy,
        String(row.itemsCount),
        row.transmissionMethod,
        new Date().toISOString(),
      ],
    ];

    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Sheet1!A:J:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`;
    const response = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ values }),
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      console.error("Sheets append error:", err);
      return false;
    }

    return true;
  } catch (error) {
    console.error("Failed to append row to Google Sheet:", error);
    return false;
  }
};
