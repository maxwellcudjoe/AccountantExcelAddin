/**
 * onedrive.js — MSAL + Microsoft Graph API integration
 * Handles sign-in and read/write of encrypted blobs to the user's own OneDrive.
 * Your server never touches this data.
 */

import * as msal from "@azure/msal-browser";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const LF_FOLDER = "LedgerFlow";

const SCOPES = ["Files.ReadWrite", "User.Read"];

const msalConfig = {
  auth: {
    clientId: process.env.AZURE_CLIENT_ID,
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "sessionStorage",   // Never localStorage — sessionStorage is cleared on tab close
    storeAuthStateInCookie: false,
  },
};

const pca = new msal.PublicClientApplication(msalConfig);
let _account = null;

/**
 * Sign in to OneDrive via MSAL popup.
 * @returns {Promise<{email: string}>}
 */
export async function signInOneDrive() {
  await pca.initialize();
  const response = await pca.loginPopup({ scopes: SCOPES });
  _account = response.account;
  pca.setActiveAccount(_account);
  await ensureFolder();
  return { email: _account.username };
}

/**
 * Silently acquire a Graph API access token; falls back to popup.
 * @returns {Promise<string>} access token
 */
export async function getGraphToken() {
  await pca.initialize();
  const accounts = pca.getAllAccounts();
  if (accounts.length === 0) throw new Error("Not signed in to OneDrive");

  const account = _account || accounts[0];
  try {
    const result = await pca.acquireTokenSilent({ scopes: SCOPES, account });
    return result.accessToken;
  } catch {
    const result = await pca.acquireTokenPopup({ scopes: SCOPES, account });
    return result.accessToken;
  }
}

/**
 * Sign out of OneDrive and clear MSAL cache.
 */
export async function signOutOneDrive() {
  await pca.initialize();
  if (_account) {
    await pca.logoutPopup({ account: _account });
  }
  _account = null;
}

/**
 * Returns true if the user is currently signed in to OneDrive.
 */
export async function isOneDriveConnected() {
  await pca.initialize();
  return pca.getAllAccounts().length > 0;
}

/**
 * Creates the /LedgerFlow/ folder in the user's OneDrive root if it doesn't exist.
 */
export async function ensureFolder() {
  const token = await getGraphToken();
  const res = await fetch(`${GRAPH_BASE}/me/drive/root:/${LF_FOLDER}`, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (res.status === 404) {
    // Folder doesn't exist — create it
    await fetch(`${GRAPH_BASE}/me/drive/root/children`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        name: LF_FOLDER,
        folder: {},
        "@microsoft.graph.conflictBehavior": "rename",
      }),
    });
  }
}

/**
 * Writes an encrypted payload as a file in /LedgerFlow/{filename}.
 * The payload is { iv, ciphertext } — only unreadable ciphertext reaches OneDrive.
 * @param {string} filename — e.g. "transactions.lf"
 * @param {object} encryptedPayload — { iv: string, ciphertext: string }
 */
export async function writeFile(filename, encryptedPayload) {
  const token = await getGraphToken();
  const content = JSON.stringify(encryptedPayload);

  const res = await fetch(
    `${GRAPH_BASE}/me/drive/root:/${LF_FOLDER}/${filename}:/content`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: content,
    }
  );

  if (!res.ok) {
    throw new Error(`OneDrive write failed: ${res.status} ${res.statusText}`);
  }
}

/**
 * Reads an encrypted payload from /LedgerFlow/{filename}.
 * Returns null if the file does not exist (404).
 * @param {string} filename
 * @returns {Promise<{iv: string, ciphertext: string} | null>}
 */
export async function readFile(filename) {
  const token = await getGraphToken();

  const res = await fetch(
    `${GRAPH_BASE}/me/drive/root:/${LF_FOLDER}/${filename}:/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`OneDrive read failed: ${res.status} ${res.statusText}`);

  return res.json();
}

/**
 * Lists all .lf files in /LedgerFlow/.
 * @returns {Promise<string[]>} array of filenames
 */
export async function listFiles() {
  const token = await getGraphToken();

  const res = await fetch(
    `${GRAPH_BASE}/me/drive/root:/${LF_FOLDER}:/children`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  if (!res.ok) return [];
  const data = await res.json();
  return (data.value || [])
    .map((f) => f.name)
    .filter((n) => n.endsWith(".lf"));
}
