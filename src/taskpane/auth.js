/* global msal */
// Auth via MSAL: Nested App Authentication (NAA) when the Outlook host supports
// brokering, falling back to a standard MSAL popup flow otherwise.

const CFG = window.SFNC_CONFIG || {};
const SCOPES = CFG.scopes || ["User.Read", "Files.ReadWrite.All", "Sites.Read.All"];

let pca = null;

export function isConfigured() {
  return !!(CFG.clientId && !CFG.clientId.startsWith("<"));
}

export function getAccount() {
  return pca ? pca.getActiveAccount() : null;
}

export async function initAuth() {
  if (!isConfigured()) {
    throw new Error("src/config.js is missing or has placeholder values. Run setup\\2026_06.10_Setup_EntraApp_v1.0.ps1 first.");
  }
  const config = {
    auth: {
      clientId: CFG.clientId,
      authority: "https://login.microsoftonline.com/" + (CFG.tenantId || "common"),
      // resolve against the current page so it works at a domain root or a /repo/ subpath
      redirectUri: new URL("../auth-redirect.html", location.href).href
    },
    cache: { cacheLocation: "localStorage" }
  };
  try {
    pca = await msal.createNestablePublicClientApplication(config);
  } catch (e) {
    if (msal.createStandardPublicClientApplication) {
      pca = await msal.createStandardPublicClientApplication(config);
    } else {
      pca = new msal.PublicClientApplication(config);
      await pca.initialize();
    }
  }
  const accounts = pca.getAllAccounts();
  if (accounts.length > 0) pca.setActiveAccount(accounts[0]);
  return getAccount();
}

// Returns an access token, trying silent first then interactive popup.
export async function getToken(interactiveOk = true) {
  if (!pca) await initAuth();
  try {
    const result = await pca.acquireTokenSilent({
      scopes: SCOPES,
      account: pca.getActiveAccount() || undefined
    });
    pca.setActiveAccount(result.account);
    return result.accessToken;
  } catch (e) {
    if (!interactiveOk) throw e;
    const result = await pca.acquireTokenPopup({ scopes: SCOPES, prompt: "select_account" });
    pca.setActiveAccount(result.account);
    return result.accessToken;
  }
}

export async function signIn() {
  await getToken(true);
  return getAccount();
}

export async function trySilentSignIn() {
  try {
    await getToken(false);
    return getAccount();
  } catch (e) {
    return null;
  }
}

export async function signOut() {
  if (!pca) return;
  try {
    if (pca.clearCache) await pca.clearCache();
  } catch (e) {
    // best effort
  }
  pca.setActiveAccount(null);
}
