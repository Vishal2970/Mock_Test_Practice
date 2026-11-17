const DRIVE_SCOPE = 'https://www.googleapis.com/auth/drive.readonly';
const GIS_SCRIPT = 'https://accounts.google.com/gsi/client';
const DRIVE_API = 'https://www.googleapis.com/drive/v3/files';

type GoogleTokenResponse = {
  access_token?: string;
  expires_in?: number;
  scope?: string;
  error?: string;
  error_description?: string;
};

type GooglePopupError = {
  type?: 'popup_closed' | 'popup_failed_to_open' | 'unknown' | string;
  message?: string;
};

type GoogleTokenClient = {
  requestAccessToken: (options?: { prompt?: string }) => void;
};

type GoogleOAuth2 = {
  initTokenClient: (options: {
    client_id: string;
    scope: string;
    callback: (response: GoogleTokenResponse) => void;
    error_callback: (error: GooglePopupError) => void;
  }) => GoogleTokenClient;
  revoke: (token: string, callback: () => void) => void;
};

declare global {
  interface Window {
    google?: {
      accounts?: {
        oauth2?: GoogleOAuth2;
      };
    };
  }
}

export type DriveItem = {
  id: string;
  name: string;
  mimeType: string;
  modifiedTime?: string;
};

export class DriveApiError extends Error {
  constructor(message: string, readonly status: number) {
    super(message);
    this.name = 'DriveApiError';
  }
}

export class GoogleAuthCancelledError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'GoogleAuthCancelledError';
  }
}

let gisReady: Promise<void> | null = null;

function loadGoogleIdentityServices(): Promise<void> {
  if (window.google?.accounts?.oauth2) return Promise.resolve();
  if (gisReady) return gisReady;

  const ready = new Promise<void>((resolve, reject) => {
    const script = document.createElement('script');
    script.src = GIS_SCRIPT;
    script.async = true;
    script.defer = true;
    script.onload = () => window.google?.accounts?.oauth2 ? resolve() : reject(new Error('Google Identity Services did not initialize.'));
    script.onerror = () => reject(new Error('Could not load Google Identity Services.'));
    document.head.appendChild(script);
  }).catch(error => {
    gisReady = null;
    throw error;
  });
  gisReady = ready;
  return ready;
}

export function prepareGoogleIdentityServices(): Promise<void> {
  return loadGoogleIdentityServices();
}

export function requestDriveAccessToken(clientId: string): Promise<{ token: string; expiresIn: number }> {
  const oauth2 = window.google?.accounts?.oauth2;
  if (oauth2) return startTokenRequest(oauth2, clientId);
  return loadGoogleIdentityServices().then(() => {
    const loadedOAuth2 = window.google?.accounts?.oauth2;
    if (!loadedOAuth2) throw new Error('Google Identity Services is unavailable.');
    return startTokenRequest(loadedOAuth2, clientId);
  });
}

function startTokenRequest(oauth2: GoogleOAuth2, clientId: string): Promise<{ token: string; expiresIn: number }> {
  return new Promise((resolve, reject) => {
    let settled = false;
    let popupCloseTimer: number | undefined;
    const rejectOnce = (error: Error) => {
      if (settled) return;
      settled = true;
      if (popupCloseTimer !== undefined) window.clearTimeout(popupCloseTimer);
      reject(error);
    };
    const resolveOnce = (auth: { token: string; expiresIn: number }) => {
      if (settled) return;
      settled = true;
      if (popupCloseTimer !== undefined) window.clearTimeout(popupCloseTimer);
      resolve(auth);
    };
    const client = oauth2.initTokenClient({
      client_id: clientId,
      scope: DRIVE_SCOPE,
      callback: response => {
        if (settled) {
          console.warn('[MockPractice Drive OAuth] A token callback arrived after the OAuth request ended.', {
            hasAccessToken: Boolean(response.access_token),
            error: response.error || '',
          });
          return;
        }
        if (response.error || !response.access_token) {
          rejectOnce(new Error(response.error_description || response.error || 'Google did not return an access token.'));
          return;
        }
        if (!response.scope?.split(/\s+/).includes(DRIVE_SCOPE)) {
          rejectOnce(new Error('Google did not grant the required read-only Drive permission.'));
          return;
        }
        resolveOnce({ token: response.access_token, expiresIn: Number(response.expires_in) || 3600 });
      },
      error_callback: error => {
        if (settled) return;
        if (error.type === 'popup_closed') {
          if (popupCloseTimer === undefined) {
            popupCloseTimer = window.setTimeout(() => {
              console.warn('[MockPractice Drive OAuth] Popup closed before a token response was received.', {
                type: error.type,
                message: error.message || '',
              });
              rejectOnce(new GoogleAuthCancelledError('The Google authorization popup closed before a token response was received.'));
            }, 5000);
          }
          return;
        }
        console.warn('[MockPractice Drive OAuth] Popup error callback received.', {
          type: error.type || 'unknown',
          message: error.message || '',
        });
        if (error.type === 'popup_failed_to_open') {
          rejectOnce(new Error('Google could not open the sign-in popup. Allow popups for this site and try again.'));
          return;
        }
        rejectOnce(new Error(error.message || 'Google sign-in failed before authorization completed.'));
      },
    });
    try {
      client.requestAccessToken({ prompt: 'select_account' });
    } catch (error) {
      rejectOnce(error instanceof Error ? error : new Error('Google could not start the sign-in flow.'));
    }
  });
}

async function driveFetch(url: string, token: string): Promise<Response> {
  const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!response.ok) {
    let message = `Google Drive request failed (${response.status}).`;
    try {
      const body = await response.json();
      message = body.error?.message || message;
    } catch { /* use the HTTP status message */ }
    console.error('[MockPractice Drive] API request failed.', { status: response.status, message });
    throw new DriveApiError(message, response.status);
  }
  return response;
}

export async function getDriveFolder(token: string, folderId: string): Promise<DriveItem> {
  const query = new URLSearchParams({ fields: 'id,name,mimeType', supportsAllDrives: 'true' });
  const response = await driveFetch(`${DRIVE_API}/${encodeURIComponent(folderId)}?${query}`, token);
  const folder = await response.json() as DriveItem;
  if (folder.mimeType !== 'application/vnd.google-apps.folder') {
    throw new DriveApiError('The configured Google Drive root is not a folder.', 400);
  }
  return folder;
}

export async function listDriveItems(token: string, folderId: string): Promise<DriveItem[]> {
  const items: DriveItem[] = [];
  let pageToken: string | undefined;
  do {
    const query = new URLSearchParams({
      q: `'${folderId.replace(/\\/g, '\\\\').replace(/'/g, "\\'")}' in parents and trashed = false`,
      fields: 'nextPageToken,files(id,name,mimeType,modifiedTime)',
      pageSize: '1000',
      orderBy: 'name',
      supportsAllDrives: 'true',
      includeItemsFromAllDrives: 'true',
    });
    if (pageToken) query.set('pageToken', pageToken);
    const response = await driveFetch(`${DRIVE_API}?${query}`, token);
    const page = await response.json() as { files?: DriveItem[]; nextPageToken?: string };
    items.push(...(page.files || []));
    pageToken = page.nextPageToken;
  } while (pageToken);
  return items;
}

export async function getDriveFile(token: string, fileId: string): Promise<DriveItem> {
  const query = new URLSearchParams({ fields: 'id,name,mimeType,modifiedTime', supportsAllDrives: 'true' });
  const response = await driveFetch(`${DRIVE_API}/${encodeURIComponent(fileId)}?${query}`, token);
  return response.json() as Promise<DriveItem>;
}

export async function downloadDriveFile(token: string, fileId: string): Promise<ArrayBuffer> {
  const query = new URLSearchParams({ alt: 'media', supportsAllDrives: 'true' });
  const response = await driveFetch(`${DRIVE_API}/${encodeURIComponent(fileId)}?${query}`, token);
  return response.arrayBuffer();
}

export function revokeDriveAccessToken(token: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const revoke = window.google?.accounts?.oauth2?.revoke;
    if (!revoke) {
      reject(new Error('Google Identity Services is unavailable to revoke the session.'));
      return;
    }
    try {
      revoke(token, resolve);
    } catch (error) {
      reject(error instanceof Error ? error : new Error('Google could not revoke the session.'));
    }
  });
}
