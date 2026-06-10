/* global Office */
import { initAuth, isConfigured, getAccount, signIn, trySilentSignIn, signOut } from "./auth.js";
import * as graph from "./graph.js";

// ---------- helpers ----------

const $ = (sel) => document.querySelector(sel);

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function fmtSize(bytes) {
  if (!bytes && bytes !== 0) return "";
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(0) + " KB";
  return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

function fmtDate(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  return isNaN(d) ? "" : d.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" });
}

let toastTimer = null;
function toast(msg, isError = false) {
  const el = $("#toast");
  el.textContent = msg;
  el.className = isError ? "error" : "";
  el.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.hidden = true; }, isError ? 9000 : 4000);
}

async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text);
  } catch (e) {
    const ta = document.createElement("textarea");
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    ta.remove();
  }
}

// ---------- state ----------

const DEFAULT_SETTINGS = {
  scope: "organization",
  type: "view",
  expDays: 30,
  folder: "Email attachments",
  thresholdMB: 5
};

let settings = loadSettings();
let inOutlook = false;
let composeMode = false;
const state = {
  source: "recent",
  crumbs: [],
  searching: null,
  file: null // selected file in detail view
};

function loadSettings() {
  try {
    return { ...DEFAULT_SETTINGS, ...(JSON.parse(localStorage.getItem("sfnc_settings")) || {}) };
  } catch (e) {
    return { ...DEFAULT_SETTINGS };
  }
}

function saveSettings() {
  localStorage.setItem("sfnc_settings", JSON.stringify(settings));
  if (inOutlook && Office.context.roamingSettings) {
    try {
      Office.context.roamingSettings.set("sfnc_thresholdMB", settings.thresholdMB);
      Office.context.roamingSettings.saveAsync(() => {});
    } catch (e) { /* read item may not allow saving; non-fatal */ }
  }
}

// ---------- boot ----------

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady((info) => {
    inOutlook = !!(info.host && Office.context.mailbox && Office.context.mailbox.item);
    if (inOutlook) {
      const item = Office.context.mailbox.item;
      composeMode = !!(item.subject && typeof item.subject.getAsync === "function");
    }
    boot();
  });
} else {
  boot();
}

async function boot() {
  wireUi();
  if (!composeMode) {
    $("#tab-btn-attachments").hidden = true;
    $("#btn-create-insert").hidden = true;
  }
  if (!isConfigured()) {
    $("#view-signin").hidden = false;
    $("#btn-signin").hidden = true;
    $("#signin-blurb").textContent = "Setup required: src/config.js is missing or has placeholder values.";
    showSigninError("Run setup\\2026_06.10_Setup_EntraApp_v1.0.ps1 from the AddIn folder, then reload this pane.");
    return;
  }
  try {
    await initAuth();
  } catch (e) {
    $("#view-signin").hidden = false;
    showSigninError(e.message);
    return;
  }
  const account = await trySilentSignIn();
  if (account) showApp();
  else $("#view-signin").hidden = false;
}

function showSigninError(msg) {
  const el = $("#signin-error");
  el.textContent = msg;
  el.hidden = false;
}

function showApp() {
  $("#view-signin").hidden = true;
  $("#app").hidden = false;
  const account = getAccount();
  if (account) {
    $("#account-chip").hidden = false;
    $("#account-name").textContent = account.name || account.username;
  }
  loadCurrent();
  if (composeMode) refreshAttachments();
  renderSettingsForm();
}

// ---------- UI wiring ----------

function wireUi() {
  $("#btn-signin").addEventListener("click", async () => {
    try {
      await signIn();
      showApp();
    } catch (e) {
      showSigninError("Sign-in failed: " + e.message);
    }
  });

  $("#btn-signout").addEventListener("click", async () => {
    await signOut();
    $("#app").hidden = true;
    $("#account-chip").hidden = true;
    $("#view-signin").hidden = false;
  });

  document.querySelectorAll(".tab").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach((b) => b.classList.toggle("active", b === btn));
      ["files", "attachments", "settings"].forEach((t) => {
        $("#tab-" + t).hidden = t !== btn.dataset.tab;
      });
    });
  });

  document.querySelectorAll(".seg").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".seg").forEach((b) => b.classList.toggle("active", b === btn));
      state.source = btn.dataset.source;
      state.searching = null;
      state.crumbs =
        state.source === "onedrive" ? [{ label: "OneDrive", kind: "od-root" }]
        : state.source === "sharepoint" ? [{ label: "Sites", kind: "sites" }]
        : [];
      $("#search-row").hidden = state.source !== "onedrive";
      closeDetail();
      loadCurrent();
    });
  });

  $("#search-input").addEventListener("keydown", (e) => {
    if (e.key !== "Enter") return;
    const q = e.target.value.trim();
    state.searching = q || null;
    closeDetail();
    loadCurrent();
  });

  $("#btn-back").addEventListener("click", closeDetail);
  $("#opt-scope").addEventListener("change", onScopeChange);
  $("#btn-create-insert").addEventListener("click", () => createFromDetail(true));
  $("#btn-create-copy").addEventListener("click", () => createFromDetail(false));

  $("#btn-refresh-attachments").addEventListener("click", refreshAttachments);
  $("#btn-convert").addEventListener("click", convertSelectedAttachments);

  $("#btn-save-settings").addEventListener("click", () => {
    settings.scope = $("#set-scope").value;
    settings.type = $("#set-type").value;
    settings.expDays = Math.max(0, Number($("#set-expdays").value) || 0);
    settings.folder = $("#set-folder").value.trim() || DEFAULT_SETTINGS.folder;
    settings.thresholdMB = Math.max(0, Number($("#set-threshold").value) || 0);
    saveSettings();
    toast("Settings saved.");
  });
}

function renderSettingsForm() {
  $("#set-scope").value = settings.scope;
  $("#set-type").value = settings.type;
  $("#set-expdays").value = settings.expDays;
  $("#set-folder").value = settings.folder;
  $("#set-threshold").value = settings.thresholdMB;
}

// ---------- file browser ----------

async function loadCurrent() {
  const list = $("#file-list");
  list.innerHTML = `<div class="loading">Loading&hellip;</div>`;
  renderBreadcrumb();
  try {
    let items;
    if (state.searching) {
      items = await graph.searchMyDrive(state.searching);
    } else {
      const crumb = state.crumbs[state.crumbs.length - 1];
      if (!crumb) items = await graph.recentFiles();
      else if (crumb.kind === "od-root") items = await graph.myDriveChildren("root");
      else if (crumb.kind === "folder") items = await graph.driveChildren(crumb.driveId, crumb.itemId);
      else if (crumb.kind === "sites") items = await graph.listSites();
      else if (crumb.kind === "site") items = await graph.siteDrives(crumb.siteId);
      else if (crumb.kind === "drive-root") items = await graph.driveChildren(crumb.driveId, "root");
    }
    renderList(items || []);
  } catch (e) {
    list.innerHTML = `<div class="empty">${esc(e.message)}</div>`;
  }
}

function renderBreadcrumb() {
  const bc = $("#breadcrumb");
  bc.innerHTML = "";
  if (state.searching) {
    const span = document.createElement("span");
    span.className = "current";
    span.textContent = `Search: "${state.searching}"`;
    const clear = document.createElement("button");
    clear.type = "button";
    clear.textContent = "clear";
    clear.addEventListener("click", () => {
      state.searching = null;
      $("#search-input").value = "";
      loadCurrent();
    });
    bc.append(span, clear);
    return;
  }
  state.crumbs.forEach((crumb, i) => {
    if (i > 0) {
      const sep = document.createElement("span");
      sep.className = "sep";
      sep.textContent = "›";
      bc.appendChild(sep);
    }
    if (i === state.crumbs.length - 1) {
      const span = document.createElement("span");
      span.className = "current";
      span.textContent = crumb.label;
      bc.appendChild(span);
    } else {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = crumb.label;
      btn.addEventListener("click", () => {
        state.crumbs = state.crumbs.slice(0, i + 1);
        loadCurrent();
      });
      bc.appendChild(btn);
    }
  });
}

const ICONS = { site: "🌐", drive: "📚", folder: "📁", file: "📄" };

function renderList(items) {
  const list = $("#file-list");
  list.innerHTML = "";
  if (!items.length) {
    list.innerHTML = `<div class="empty">Nothing here.</div>`;
    return;
  }
  const sorted = [...items].sort((a, b) => {
    const rank = (x) => ({ site: 0, drive: 0, folder: 1, file: 2 }[x.kind] ?? 3);
    return rank(a) - rank(b);
  });
  for (const item of sorted) {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "row";
    const meta =
      item.kind === "file" ? fmtSize(item.size)
      : item.kind === "folder" && item.childCount !== undefined ? item.childCount + " items"
      : "";
    row.innerHTML = `<span class="ic">${ICONS[item.kind] || "📄"}</span><span class="nm">${esc(item.name)}</span><span class="mt">${esc(meta)}</span>`;
    row.addEventListener("click", () => {
      if (item.kind === "site") {
        state.crumbs.push({ label: item.name, kind: "site", siteId: item.siteId });
        loadCurrent();
      } else if (item.kind === "drive") {
        state.crumbs.push({ label: item.name, kind: "drive-root", driveId: item.driveId });
        loadCurrent();
      } else if (item.kind === "folder") {
        state.searching = null;
        state.crumbs.push({ label: item.name, kind: "folder", driveId: item.driveId, itemId: item.itemId });
        loadCurrent();
      } else {
        openDetail(item);
      }
    });
    list.appendChild(row);
  }
}

// ---------- file detail: create link + manage access ----------

async function openDetail(file) {
  state.file = file;
  $("#files-browser").hidden = true;
  $("#file-detail").hidden = false;
  $("#detail-name").textContent = file.name;
  $("#detail-meta").textContent = [fmtSize(file.size), file.modified ? "modified " + fmtDate(file.modified) : ""].filter(Boolean).join(" · ");

  $("#opt-scope").value = settings.scope;
  $("#opt-type").value = settings.type;
  $("#opt-password").value = "";
  $("#opt-expiry").value = settings.expDays > 0 ? datePlusDays(settings.expDays) : "";
  await onScopeChange();
  loadPermissions();
}

function closeDetail() {
  state.file = null;
  $("#file-detail").hidden = true;
  $("#files-browser").hidden = false;
}

function datePlusDays(days) {
  const d = new Date();
  d.setDate(d.getDate() + days);
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
}

async function onScopeChange() {
  const isUsers = $("#opt-scope").value === "users";
  $("#recipients-row").hidden = !isUsers;
  if (isUsers && !$("#opt-recipients").value) {
    $("#opt-recipients").value = (await getComposeRecipients()).join(", ");
  }
}

function getComposeRecipients() {
  return new Promise((resolve) => {
    if (!composeMode) return resolve([]);
    const item = Office.context.mailbox.item;
    const emails = [];
    item.to.getAsync((r1) => {
      if (r1.status === Office.AsyncResultStatus.Succeeded) emails.push(...r1.value.map((x) => x.emailAddress));
      item.cc.getAsync((r2) => {
        if (r2.status === Office.AsyncResultStatus.Succeeded) emails.push(...r2.value.map((x) => x.emailAddress));
        resolve([...new Set(emails.filter(Boolean))]);
      });
    });
  });
}

function readLinkOptions() {
  const expiryDate = $("#opt-expiry").value;
  return {
    type: $("#opt-type").value,
    scope: $("#opt-scope").value,
    expiry: expiryDate ? expiryDate + "T23:59:59Z" : null,
    password: $("#opt-password").value.trim() || null,
    recipients: $("#opt-recipients").value.split(",").map((s) => s.trim()).filter(Boolean)
  };
}

// Create a sharing link (with expiration/password/recipients as requested).
// Returns { url, notes[] }.
async function createLinkFlow(file, opts) {
  const notes = [];
  const body = { type: opts.type, scope: opts.scope };
  if (opts.expiry) body.expirationDateTime = opts.expiry;
  if (opts.password) body.password = opts.password;
  const { status, data: perm } = await graph.createLink(file.driveId, file.itemId, body);
  let url = perm.link.webUrl;

  // 200 = an existing link of this type/scope was reused; our expiration may not have applied.
  if (status === 200 && opts.expiry && perm.expirationDateTime !== opts.expiry) {
    try {
      await graph.updatePermission(file.driveId, file.itemId, perm.id, { expirationDateTime: opts.expiry });
    } catch (e) {
      notes.push("An existing link was reused and its expiration could not be updated (" + e.message + ").");
    }
  }
  if (opts.scope === "users") {
    if (opts.recipients.length === 0) throw new Error("Enter at least one recipient for a specific-people link.");
    const res = await graph.grantOnLink(url, opts.recipients, opts.type === "edit" ? "write" : "read");
    const updated = (res.value || []).find((p) => p.link && p.link.webUrl);
    if (updated) url = updated.link.webUrl;
  }
  return { url, notes };
}

async function createFromDetail(insert) {
  const file = state.file;
  if (!file) return;
  const opts = readLinkOptions();
  const btn = insert ? $("#btn-create-insert") : $("#btn-create-copy");
  btn.disabled = true;
  try {
    const { url, notes } = await createLinkFlow(file, opts);
    if (insert && composeMode) {
      await insertHtml(linkHtml(file.name, url, opts.expiry));
      toast("Link inserted." + (notes.length ? " " + notes.join(" ") : ""));
    } else {
      await copyText(url);
      toast("Link copied to clipboard." + (notes.length ? " " + notes.join(" ") : ""));
    }
    loadPermissions();
  } catch (e) {
    toast("Couldn't create link: " + e.message, true);
  } finally {
    btn.disabled = false;
  }
}

function linkHtml(name, url, expiryIso) {
  let html = `<a href="${esc(url)}">${esc(name)}</a>`;
  if (expiryIso) html += ` <span style="color:#666;font-size:90%">(link expires ${esc(fmtDate(expiryIso))})</span>`;
  return html;
}

function insertHtml(html) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(r.error ? r.error.message : "insert failed"));
    });
  });
}

// ---------- existing access (manage) ----------

const SCOPE_LABELS = {
  anonymous: "Anyone with the link",
  organization: "People in the organization",
  users: "Specific people",
  existingAccess: "People with existing access"
};

async function loadPermissions() {
  const file = state.file;
  const list = $("#perm-list");
  list.innerHTML = `<div class="loading">Loading&hellip;</div>`;
  try {
    const perms = await graph.listPermissions(file.driveId, file.itemId);
    renderPermissions(perms);
  } catch (e) {
    list.innerHTML = `<div class="empty">${esc(e.message)}</div>`;
  }
}

function renderPermissions(perms) {
  const file = state.file;
  const list = $("#perm-list");
  list.innerHTML = "";
  if (!perms.length) {
    list.innerHTML = `<div class="empty">No sharing links or extra permissions on this file.</div>`;
    return;
  }
  for (const p of perms) {
    const card = document.createElement("div");
    card.className = "perm";
    const isOwner = (p.roles || []).includes("owner");
    if (p.link) {
      const bits = [];
      bits.push((p.link.type === "edit" ? "can edit" : "can view"));
      if (p.expirationDateTime) bits.push("expires " + fmtDate(p.expirationDateTime));
      if (p.hasPassword) bits.push("password");
      const people = (p.grantedToIdentitiesV2 || [])
        .map((g) => g.user && (g.user.displayName || g.user.email))
        .filter(Boolean);
      card.innerHTML =
        `<div class="who">🔗 ${esc(SCOPE_LABELS[p.link.scope] || p.link.scope)} · ${esc(bits.join(" · "))}</div>` +
        (people.length ? `<div class="sub">${esc(people.join(", "))}</div>` : "") +
        `<div class="sub">${esc(p.link.webUrl)}</div>` +
        `<div class="actions"></div><div class="expiry-edit" hidden></div>`;
      const actions = card.querySelector(".actions");

      const copyBtn = document.createElement("button");
      copyBtn.type = "button";
      copyBtn.className = "link-btn";
      copyBtn.textContent = "Copy";
      copyBtn.addEventListener("click", async () => { await copyText(p.link.webUrl); toast("Link copied."); });
      actions.appendChild(copyBtn);

      if (composeMode) {
        const insBtn = document.createElement("button");
        insBtn.type = "button";
        insBtn.className = "link-btn";
        insBtn.textContent = "Insert";
        insBtn.addEventListener("click", async () => {
          try {
            await insertHtml(linkHtml(file.name, p.link.webUrl, p.expirationDateTime));
            toast("Link inserted.");
          } catch (e) { toast(e.message, true); }
        });
        actions.appendChild(insBtn);
      }

      const expBtn = document.createElement("button");
      expBtn.type = "button";
      expBtn.className = "link-btn";
      expBtn.textContent = "Expiration";
      const editRow = card.querySelector(".expiry-edit");
      expBtn.addEventListener("click", () => {
        if (!editRow.hidden) { editRow.hidden = true; return; }
        editRow.hidden = false;
        editRow.innerHTML = "";
        const input = document.createElement("input");
        input.type = "date";
        input.value = p.expirationDateTime ? p.expirationDateTime.slice(0, 10) : "";
        const setBtn = document.createElement("button");
        setBtn.type = "button";
        setBtn.className = "btn";
        setBtn.textContent = "Set";
        setBtn.addEventListener("click", () => patchExpiry(p, input.value ? input.value + "T23:59:59Z" : null));
        const clearBtn = document.createElement("button");
        clearBtn.type = "button";
        clearBtn.className = "link-btn";
        clearBtn.textContent = "Clear";
        clearBtn.addEventListener("click", () => patchExpiry(p, null));
        editRow.append(input, setBtn, clearBtn);
      });
      actions.appendChild(expBtn);

      actions.appendChild(removeBtn(p));
    } else {
      const who =
        (p.grantedToV2 && p.grantedToV2.user && (p.grantedToV2.user.displayName || p.grantedToV2.user.email)) ||
        (p.grantedToIdentitiesV2 || []).map((g) => g.user && (g.user.displayName || g.user.email)).filter(Boolean).join(", ") ||
        "Unknown";
      card.innerHTML =
        `<div class="who">👤 ${esc(who)}</div>` +
        `<div class="sub">${esc((p.roles || []).join(", "))}${p.inheritedFrom ? " · inherited" : ""}</div>` +
        `<div class="actions"></div>`;
      if (!isOwner && !p.inheritedFrom) card.querySelector(".actions").appendChild(removeBtn(p));
    }
    list.appendChild(card);
  }

  function removeBtn(p) {
    // window.confirm is unreliable inside Office webviews; use two-click confirm
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "link-btn danger";
    btn.textContent = "Remove";
    let armed = false;
    btn.addEventListener("click", async () => {
      if (!armed) {
        armed = true;
        btn.textContent = "Really remove?";
        setTimeout(() => { armed = false; btn.textContent = "Remove"; }, 4000);
        return;
      }
      try {
        await graph.deletePermission(file.driveId, file.itemId, p.id);
        toast("Access removed.");
        loadPermissions();
      } catch (e) {
        toast("Couldn't remove: " + e.message, true);
      }
    });
    return btn;
  }

  async function patchExpiry(p, iso) {
    try {
      await graph.updatePermission(file.driveId, file.itemId, p.id, { expirationDateTime: iso });
      toast(iso ? "Expiration updated." : "Expiration cleared.");
      loadPermissions();
    } catch (e) {
      toast("Couldn't update expiration: " + e.message, true);
    }
  }
}

// ---------- attachments ----------

let attachments = [];

function refreshAttachments() {
  if (!composeMode) return;
  const list = $("#attachment-list");
  list.innerHTML = `<div class="loading">Loading&hellip;</div>`;
  $("#convert-log").textContent = "";
  Office.context.mailbox.item.getAttachmentsAsync((r) => {
    if (r.status !== Office.AsyncResultStatus.Succeeded) {
      list.innerHTML = `<div class="empty">Couldn't read attachments: ${esc(r.error ? r.error.message : "")}</div>`;
      return;
    }
    attachments = r.value.filter((a) => !a.isInline);
    list.innerHTML = "";
    if (!attachments.length) {
      list.innerHTML = `<div class="empty">No attachments on this draft.</div>`;
      $("#btn-convert").disabled = true;
      return;
    }
    for (const a of attachments) {
      const isFile = a.attachmentType === "file";
      const row = document.createElement("label");
      row.className = "att-row";
      row.innerHTML =
        `<input type="checkbox" data-id="${esc(a.id)}" ${isFile ? "checked" : "disabled"}>` +
        `<span class="nm">${esc(a.name)}</span>` +
        `<span class="st" data-st="${esc(a.id)}">${isFile ? fmtSize(a.size) : "already a cloud link"}</span>`;
      list.appendChild(row);
    }
    $("#btn-convert").disabled = false;
  });
}

function getAttachmentBytes(id) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(id, (r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) {
        return reject(new Error(r.error ? r.error.message : "couldn't read attachment"));
      }
      if (r.value.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
        return reject(new Error("unsupported attachment format: " + r.value.format));
      }
      const bin = atob(r.value.content);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      resolve(bytes);
    });
  });
}

function removeAttachment(id) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.removeAttachmentAsync(id, (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(r.error ? r.error.message : "couldn't remove attachment"));
    });
  });
}

async function convertSelectedAttachments() {
  const checked = [...document.querySelectorAll('#attachment-list input[type="checkbox"]:checked')].map((c) => c.dataset.id);
  const targets = attachments.filter((a) => checked.includes(a.id));
  if (!targets.length) return toast("Select at least one attachment.", true);

  const btn = $("#btn-convert");
  btn.disabled = true;
  const log = $("#convert-log");
  const setSt = (id, text, cls) => {
    const el = document.querySelector(`[data-st="${CSS.escape(id)}"]`);
    if (el) { el.textContent = text; el.className = "st" + (cls ? " " + cls : ""); }
  };

  // attachment conversions use the default link settings; "specific people"
  // defaults fall back to org-wide links when there are no recipients yet
  let scope = settings.scope;
  let recipients = [];
  const notes = [];
  if (scope === "users") {
    recipients = await getComposeRecipients();
    if (!recipients.length) {
      scope = "organization";
      notes.push("No recipients on the draft yet, so links were created for your whole organization.");
    }
  }
  const expiry = settings.expDays > 0 ? datePlusDays(settings.expDays) + "T23:59:59Z" : null;

  const done = [];
  try {
    await graph.ensureFolder(settings.folder);
  } catch (e) {
    toast("Couldn't create the upload folder: " + e.message, true);
    btn.disabled = false;
    return;
  }
  for (const a of targets) {
    try {
      setSt(a.id, "reading…");
      const bytes = await getAttachmentBytes(a.id);
      setSt(a.id, "uploading…");
      const item = await graph.uploadFile(settings.folder, a.name, bytes, (p) => setSt(a.id, "uploading " + Math.round(p * 100) + "%"));
      setSt(a.id, "creating link…");
      const file = { driveId: item.parentReference.driveId, itemId: item.id, name: item.name };
      const { url, notes: n2 } = await createLinkFlow(file, { type: settings.type, scope, expiry, password: null, recipients });
      notes.push(...n2);
      done.push({ id: a.id, name: a.name, url });
      setSt(a.id, "done", "ok");
    } catch (e) {
      setSt(a.id, e.message, "err");
    }
  }

  if (done.length) {
    const items = done
      .map((d) => `<li>${linkHtml(d.name, d.url, expiry)}</li>`)
      .join("");
    try {
      await insertHtml(`<p>Shared via OneDrive:</p><ul>${items}</ul>`);
    } catch (e) {
      notes.push("Links could not be inserted into the body (" + e.message + ") — they were left as attachments.");
    }
    for (const d of done) {
      try {
        await removeAttachment(d.id);
      } catch (e) {
        notes.push(`Couldn't remove attachment "${d.name}": ${e.message}`);
      }
    }
  }
  log.textContent =
    `${done.length} of ${targets.length} attachment(s) converted.` + (notes.length ? " " + notes.join(" ") : "");
  btn.disabled = false;
  refreshAttachments();
}
