// Google Drive Service - Uses access token from Better Auth session

const AUTH_SERVER =
  process.env.NEXT_PUBLIC_BETTER_AUTH_URL ||
  (typeof window !== "undefined" ? window.location.origin : "");

// Cached access token
let cachedAccessToken: string | null = null;

// Fetch the Google access token from the server (uses session cookie)
export const getGoogleAccessToken = async (): Promise<string> => {
    if (cachedAccessToken) return cachedAccessToken;
    
    const response = await fetch(`${AUTH_SERVER}/api/google-token`, {
        credentials: 'include',
    });
    
    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error || 'Failed to get Google token');
    }
    
    const data = await response.json();
    cachedAccessToken = data.accessToken;
    return data.accessToken;
};

// Clear the cached token (call on logout)
export const clearGoogleToken = () => {
    cachedAccessToken = null;
};

// Check if user has Google Drive access
export const checkDriveAccess = async (): Promise<boolean> => {
    try {
        await getGoogleAccessToken();
        return true;
    } catch {
        return false;
    }
};

export const getFileContentAsBase64 = async (fileId: string): Promise<{base64: string, mimeType: string}> => {
    const accessToken = await getGoogleAccessToken();

    const response = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
    });

    if (!response.ok) throw new Error(`Download failed: ${response.statusText}`);

    const blob = await response.blob();
    const mimeType = blob.type;

    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
            const result = reader.result as string;
            resolve({ base64: result.split(',')[1], mimeType });
        };
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
};

export const extractFolderIdFromLink = (link: string): string | null => {
    const regex = /(?:\/folders\/|\/d\/|[?&]id=)([a-zA-Z0-9_-]{15,})/;
    const match = link.match(regex);
    return match ? match[1] : null;
};

export const isFolderLink = (link: string): boolean => {
    return /\/folders\//.test(link);
};

export const extractFileIdFromLink = (link: string): string | null => {
    // Matches: /file/d/ID, /document/d/ID, /spreadsheets/d/ID, /presentation/d/ID, ?id=ID
    const regex = /(?:\/(?:file|document|spreadsheets|presentation)\/d\/|[?&]id=)([a-zA-Z0-9_-]{15,})/;
    const match = link.match(regex);
    return match ? match[1] : null;
};

export const getFileMetadata = async (fileId: string): Promise<{ id: string; name: string; mimeType: string }> => {
    const accessToken = await getGoogleAccessToken();
    const url = `https://www.googleapis.com/drive/v3/files/${fileId}?fields=id,name,mimeType&supportsAllDrives=true`;
    const response = await fetch(url, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
    });
    if (!response.ok) {
        throw new Error(`Failed to get file metadata: ${response.status} ${response.statusText}`);
    }
    return response.json();
};

export const listFilesInFolder = async (folderId: string): Promise<Array<{name: string, id: string, mimeType: string}>> => {
    const accessToken = await getGoogleAccessToken();

    try {
        const query = `'${folderId}' in parents and trashed = false`;
        const allFiles: Array<{ name: string; id: string; mimeType: string }> = [];
        let pageToken: string | undefined;

        do {
            const url = new URL('https://www.googleapis.com/drive/v3/files');
            url.searchParams.append('pageSize', '1000');
            url.searchParams.append('fields', 'nextPageToken, files(id, name, mimeType)');
            url.searchParams.append('q', query);
            url.searchParams.append('supportsAllDrives', 'true');
            url.searchParams.append('includeItemsFromAllDrives', 'true');
            if (pageToken) url.searchParams.append('pageToken', pageToken);

            const response = await fetch(url.toString(), {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            });

            if (!response.ok) {
                let errorDetails = '';
                try {
                    const errorJson = await response.json();
                    errorDetails = errorJson?.error?.message ? `: ${errorJson.error.message}` : '';
                } catch {
                    // ignore JSON parsing errors
                }
                throw new Error(`Drive API Error: ${response.status} ${response.statusText}${errorDetails}`);
            }

            const data = await response.json();
            allFiles.push(...(data.files || []));
            pageToken = data.nextPageToken;
        } while (pageToken);

        return allFiles;

    } catch (err: any) {
        console.error("Error listing files", err);
        throw new Error(err.message || "Failed to list files");
    }
};

// List files in user's Drive root or search
export const listDriveFiles = async (query?: string): Promise<Array<{name: string, id: string, mimeType: string}>> => {
    const accessToken = await getGoogleAccessToken();

    const url = new URL('https://www.googleapis.com/drive/v3/files');
    url.searchParams.append('pageSize', '50');
    url.searchParams.append('fields', 'files(id, name, mimeType, modifiedTime)');
    url.searchParams.append('orderBy', 'modifiedTime desc');
    
    if (query) {
        url.searchParams.append('q', `name contains '${query}' and trashed = false`);
    } else {
        url.searchParams.append('q', 'trashed = false');
    }

    const response = await fetch(url.toString(), {
        headers: { 'Authorization': `Bearer ${accessToken}` }
    });

    if (!response.ok) {
        throw new Error(`Drive API Error: ${response.status}`);
    }

    const data = await response.json();
    return data.files || [];
};
