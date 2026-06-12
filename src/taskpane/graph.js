// Thin Microsoft Graph layer for OneDrive/SharePoint sharing.
import { getToken } from "./auth.js";

const BASE = "https://graph.microsoft.com/v1.0";
const SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024; // above this, use an upload session
const CHUNK_SIZE = 5 * 1024 * 1024; // must be a multiple of 320 KiB

async function call(path, { method = "GET", body, headers = {} } = {}) {
  const token = await getToken();
  const isBinary = body instanceof Uint8Array || body instanceof ArrayBuffer;
  const res = await fetch(BASE + path, {
    method,
    headers: {
      Authorization: "Bearer " + token,
      ...(body && !isBinary ? { "Content-Type": "application/json" } : {}),
      ...(isBinary ? { "Content-Type": "application/octet-stream" } : {}),
      ...headers
    },
    body: isBinary ? body : body ? JSON.stringify(body) : undefined
  });
  let data = null;
  const text = await res.text();
  if (text) {
    try { data = JSON.parse(text); } catch (e) { data = text; }
  }
  if (!res.ok) {
    const err = new Error((data && data.error && data.error.message) || res.status + " " + res.statusText);
    err.code = (data && data.error && data.error.code) || String(res.status);
    err.status = res.status;
    throw err;
  }
  return { status: res.status, data };
}

// ---- browsing ----

export async function recentFiles() {
  const { data } = await call("/me/drive/recent?$top=30");
  return data.value.map(normalizeItem).filter((i) => i.kind === "file");
}

export async function myDriveChildren(itemId) {
  const seg = itemId && itemId !== "root" ? "items/" + itemId : "root";
  const { data } = await call("/me/drive/" + seg + "/children?$top=200&$orderby=name");
  return data.value.map(normalizeItem);
}

export async function driveChildren(driveId, itemId) {
  const seg = itemId && itemId !== "root" ? "items/" + itemId : "root";
  const { data } = await call("/drives/" + driveId + "/" + seg + "/children?$top=200&$orderby=name");
  return data.value.map(normalizeItem);
}

export async function searchMyDrive(q) {
  const safe = encodeURIComponent(q.replace(/'/g, "''"));
  const { data } = await call("/me/drive/root/search(q='" + safe + "')?$top=30");
  return data.value.map(normalizeItem);
}

export async function listSites() {
  const { data } = await call("/sites?search=*&$top=100");
  return data.value
    .map((s) => ({ kind: "site", siteId: s.id, name: s.displayName || s.name, webUrl: s.webUrl }))
    .sort((a, b) => (a.name || "").localeCompare(b.name || ""));
}

export async function siteDrives(siteId) {
  const { data } = await call("/sites/" + siteId + "/drives");
  return data.value
    .filter((d) => d.driveType === "documentLibrary")
    .map((d) => ({ kind: "drive", driveId: d.id, name: d.name, webUrl: d.webUrl }));
}

function normalizeItem(raw) {
  // /me/drive/recent returns shared items wrapped in remoteItem
  const it = raw.remoteItem || raw;
  return {
    kind: it.folder ? "folder" : "file",
    name: raw.name || it.name,
    size: it.size || 0,
    itemId: it.id,
    driveId: (it.parentReference && it.parentReference.driveId) || (raw.parentReference && raw.parentReference.driveId) || null,
    webUrl: it.webUrl,
    modified: it.lastModifiedDateTime || raw.lastModifiedDateTime,
    childCount: it.folder ? it.folder.childCount : undefined
  };
}

// ---- upload (attachment conversion) ----

export async function ensureFolder(name) {
  try {
    const { data } = await call("/me/drive/root:/" + encodeURIComponent(name));
    return data;
  } catch (e) {
    if (e.status !== 404) throw e;
  }
  try {
    const { data } = await call("/me/drive/root/children", {
      method: "POST",
      body: { name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }
    });
    return data;
  } catch (e) {
    if (e.code === "nameAlreadyExists") {
      const { data } = await call("/me/drive/root:/" + encodeURIComponent(name));
      return data;
    }
    throw e;
  }
}

export async function uploadFile(folder, name, bytes, onProgress) {
  const path = "/me/drive/root:/" + encodeURIComponent(folder) + "/" + encodeURIComponent(name);
  if (bytes.length <= SIMPLE_UPLOAD_LIMIT) {
    const { data } = await call(path + ":/content?@microsoft.graph.conflictBehavior=rename", {
      method: "PUT",
      body: bytes
    });
    if (onProgress) onProgress(1);
    return data;
  }
  const { data: session } = await call(path + ":/createUploadSession", {
    method: "POST",
    body: { item: { "@microsoft.graph.conflictBehavior": "rename", name } }
  });
  const total = bytes.length;
  let item = null;
  for (let start = 0; start < total; start += CHUNK_SIZE) {
    const end = Math.min(start + CHUNK_SIZE, total);
    const res = await fetch(session.uploadUrl, {
      method: "PUT",
      headers: { "Content-Range": "bytes " + start + "-" + (end - 1) + "/" + total },
      body: bytes.subarray(start, end)
    });
    if (!res.ok) {
      const text = await res.text();
      throw new Error("Upload failed (" + res.status + "): " + text.slice(0, 200));
    }
    if (onProgress) onProgress(end / total);
    if (res.status === 200 || res.status === 201) item = await res.json();
  }
  return item;
}

// ---- sharing links & permissions ----

export function createLink(driveId, itemId, body) {
  return call("/drives/" + driveId + "/items/" + itemId + "/createLink", { method: "POST", body });
}

export async function listPermissions(driveId, itemId) {
  const { data } = await call("/drives/" + driveId + "/items/" + itemId + "/permissions");
  return data.value;
}

export function deletePermission(driveId, itemId, permId) {
  return call("/drives/" + driveId + "/items/" + itemId + "/permissions/" + permId, { method: "DELETE" });
}

export function updatePermission(driveId, itemId, permId, body) {
  return call("/drives/" + driveId + "/items/" + itemId + "/permissions/" + permId, { method: "PATCH", body });
}

// Grant specific people access to a scope:"users" sharing link.
export async function grantOnLink(sharingUrl, emails, role) {
  const bytes = new TextEncoder().encode(sharingUrl);
  let bin = "";
  bytes.forEach((b) => (bin += String.fromCharCode(b)));
  const encoded = "u!" + btoa(bin).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  const { data } = await call("/shares/" + encoded + "/permission/grant", {
    method: "POST",
    body: { recipients: emails.map((email) => ({ email })), roles: [role] }
  });
  return data;
}

// ---- policy & admin ----

export async function getUserGroups() {
  try {
    const { data } = await call("/me/memberOf?$select=id,displayName&$top=200");
    return data.value || [];
  } catch (e) {
    return [];
  }
}

export async function getSharePointPolicy() {
  const CFG = window.SFNC_CONFIG || {};
  const siteId = CFG.adminPolicySite || "root";
  try {
    const { data: listsData } = await call(
      "/sites/" + siteId + "/lists?$select=id,displayName,webUrl&$top=200"
    );
    const list = (listsData.value || []).find((l) => l.displayName === "SFNC_ShareLinks_Policy");
    if (!list) return null;
    const { data: itemsData } = await call(
      "/sites/" + siteId + "/lists/" + list.id +
      "/items?$expand=fields($select=Title,SettingValue)&$select=fields&$top=200"
    );
    const raw = {};
    for (const item of (itemsData.value || [])) {
      const f = item.fields;
      if (f && f.Title) raw[f.Title] = f.SettingValue !== undefined ? f.SettingValue : "";
    }
    return { _listUrl: list.webUrl, _raw: raw };
  } catch (e) {
    return null;
  }
}

export async function getTenantSharingDefaults() {
  try {
    const { data } = await call("/admin/sharepoint/settings");
    return {
      defaultScope: mapSpScope(data.defaultSharingLinkType),
      defaultType: data.defaultLinkPermission === "edit" ? "edit" : "view",
      allowAnonymousLinks: data.sharingCapability !== "disabled",
    };
  } catch (e) {
    return null;
  }
}

function mapSpScope(linkType) {
  if (!linkType) return null;
  const lt = linkType.toLowerCase();
  if (lt === "anonymous") return "anonymous";
  if (lt === "direct") return "users";
  if (lt === "internal") return "organization";
  return null;
}
