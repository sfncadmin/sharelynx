# SFNC Share Links — Outlook Add-in

OneDrive/SharePoint sharing from inside Outlook, modeled on the ShareFile/Egnyte add-ins: create sharing links with expiration and audience controls, manage/revoke existing access, convert attachments to links, and get warned on send when large attachments could be links instead.

Office.js web add-in — works in classic Outlook for Windows (M365, WebView2), new Outlook, and OWA. Pure static frontend calling Microsoft Graph directly; no backend.

## Features

| Feature | Where |
|-|-|
| Browse/search OneDrive, browse SharePoint sites, Recent files | Files tab |
| Create link: audience (org / specific people / anyone), view/edit, expiration, password | Files tab → pick a file |
| Insert link into the draft or copy to clipboard | Files tab |
| View existing links & permissions, copy/insert, edit expiration, revoke | Files tab → pick a file → Existing access |
| Convert draft attachments → OneDrive uploads + sharing links, removes the attachments | Attachments tab |
| On-send Smart Alert when attachments exceed a size threshold | automatic (OnMessageSend) |
| Defaults: audience, permission, expiration days, upload folder, alert threshold | Settings tab |

## Hosting model

The add-in is a pure static frontend. Two ways to host it:

- **Production (recommended):** the `src/` files are served from **GitHub Pages** at
  `https://sfncadmin.github.io/sfnc-share-links/`. Nothing runs on your PC. `manifest.xml`
  points here and is **admin-deployed** via the M365 admin center (no sideloading, no role policy).
- **Local dev:** `npm start` serves the same files over `https://localhost:3000`.
  `manifest.local.xml` points here (distinct add-in Id + "(Local)" name so it can coexist
  with the deployed one). Use this only when changing the code.

The MSAL library is vendored at `src/vendor/msal-browser.min.js` and `config.js` ships with the
deploy, so the hosted site has no `node_modules` or external-CDN dependency. All in-page paths are
relative, so it works at a domain root or a `/repo/` subpath unchanged.

## Production deploy (GitHub Pages + admin deploy)

Prereqs: an M365 work account that can create Entra app registrations and admin-deploy add-ins.

```powershell
cd C:\zGit\SFNC\AddIn

# 1. Register/refresh the Entra app with BOTH localhost and production redirect URIs,
#    and (re)write src\config.js. Already run once for localhost; re-run with -BaseUrl
#    to add the Pages redirect URIs:
.\setup\2026_06.10_Setup_EntraApp_v1.0.ps1 -BaseUrl https://sfncadmin.github.io/sfnc-share-links
```

2. Push this folder to the public repo `sfncadmin/sfnc-share-links` and enable GitHub Pages
   (main branch, root). The site goes live at `https://sfncadmin.github.io/sfnc-share-links/`.
3. **Admin-deploy the manifest:** [M365 admin center](https://admin.microsoft.com) → Settings →
   **Integrated apps** → **Upload custom apps** → app type **Office Add-in** → **Upload manifest file**
   → choose `manifest.xml` → assign to **Just me** (or a group) → **Deploy**. Propagation can take a
   few hours the first time.
4. Restart classic Outlook. The **Share Links** button appears on the Message ribbon (compose and read).

> What's public in the repo: the add-in HTML/JS/CSS and `config.js` (the Entra **clientId** and
> **tenantId**). These are **not secrets** — this is a public-client SPA with no client secret, and a
> tenant ID is already publicly derivable from any email domain via Microsoft's OIDC discovery
> endpoint. No customer data, keys, or tokens are committed.

## Local development

```powershell
npm install              # dev tooling only (validator, cert helper)
npm run certs            # trust the localhost dev cert (one-time prompt)
npm start                # serves https://localhost:3000
```

Sideload `manifest.local.xml` via <https://aka.ms/olksideload> → My add-ins → Custom Addins →
Add from file, or just open <https://localhost:3000/src/taskpane/taskpane.html> in a browser to
work on the UI (compose features disabled outside Outlook).

### First run

Open a draft → **Share Links** → **Sign in**. With NAA-capable hosts sign-in is silent/SSO; otherwise
an MSAL popup appears. First sign-in asks for consent (or pre-consent for the tenant with the
admin-consent URL the setup script prints).

## Shipping a change (the iteration loop)

There are two hosting layers, and only one of them is slow:

| Layer | Hosts | Update speed |
|-|-|-|
| GitHub Pages | the code (`src/**`) | ~1-2 min after `git push`, automatic |
| M365 Integrated Apps | the manifest (buttons, permissions, URLs) | up to ~24h, **only when the manifest changes** |

So **code changes are fast** — Pages rebuilds on every push to `GitMain` on its own (no GitHub
Action needed; there's no build step). The ~24h M365 propagation only recurs if you edit
`manifest.xml` itself (new ribbon button, changed permissions, renamed, repointed URL) — rare.

```powershell
npm run deploy           # validates manifest.xml, then git add/commit/push -> Pages auto-rebuilds
# ...wait ~1-2 min, then in Outlook close & reopen the task pane to pull new code
```

**Gotcha — stale pane after a push.** Outlook's embedded browser (WebView2) caches the old JS, so
the pane can show old code even though Pages updated. Fixes, easiest first:

1. Close and reopen the task pane.
2. Restart Outlook.
3. Clear the Office web-add-in cache: delete the contents of
   `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` while Outlook is closed.

This is a local cache, not a deploy delay — seconds, not hours. Don't mistake one for the other.

## Known platform limits (tenant/Graph, not bugs)

- **"Anyone" links** fail if anonymous sharing is disabled in the SharePoint admin center — Graph returns the policy error and the pane surfaces it.
- **Expiration** is generally only honored on "Anyone" links (SharePoint Online policy). Setting it on org links may be rejected.
- **Password-protected links** are OneDrive Personal only in Graph v1.0 — expect a clear error on business tenants.
- **Block download** isn't exposed in Graph v1.0 `createLink` — roadmap (beta API / SharePoint REST).
- The on-send alert threshold is read via roamingSettings; changes apply after Outlook reloads the add-in.

## Project layout

```
manifest.xml                     Production manifest (GitHub Pages URLs) — admin-deployed
manifest.local.xml               Local-dev manifest (localhost URLs, distinct Id) — sideloaded
server.js                        Dev static server (HTTPS via office-addin-dev-certs; --http for UI preview)
setup/                           Entra app registration + icon generation scripts
src/config.js                    clientId/tenantId (committed; non-secret) — see config.example.js
src/vendor/msal-browser.min.js   Vendored MSAL (no node_modules/CDN dependency at runtime)
src/taskpane/                    Taskpane UI (taskpane.html/css/js, auth.js = MSAL/NAA, graph.js = Graph calls)
src/launchevent/launchevent.js   OnMessageSend Smart Alert handler (ES5, classic-Outlook JS runtime)
src/commands/commands.html       Event runtime page for new Outlook / OWA
src/assets/                      Icons (generated by setup script)
.nojekyll                        Tells GitHub Pages to serve files as-is
```

## Validate / troubleshoot

```powershell
npm run validate                 # validates manifest.xml against Microsoft's service
```

- Taskpane blank → is `npm start` running? Visit <https://localhost:3000/src/taskpane/taskpane.html> in a browser (the UI runs standalone for testing; compose features disabled).
- Sign-in errors → confirm `src/config.js` exists and consent was granted; check the redirect URIs on the app registration match port 3000.
- Smart Alert not firing → classic Outlook needs an M365 subscription build with event-based activation (Version 2206+); check that the add-in loaded (button visible) and threshold isn't 0.

## Roadmap (v2 candidates)

- Block-download links, real hosting (GitHub Pages/Azure SWA) + multi-tenant app for client rollout via M365 Integrated Apps, centralized deployment, dark mode, drag-drop upload in the pane, "auto-convert on send" instead of warn.
