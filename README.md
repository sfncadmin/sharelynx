# SFNC Share Links -- Outlook Add-in

OneDrive/SharePoint sharing from inside Outlook, modeled on the ShareFile/Egnyte add-ins: browse and share files, create links with expiration and audience controls, manage and revoke existing access, convert attachments to links, and get warned on send when large attachments could be links instead.

Office.js web add-in. Works in classic Outlook for Windows (M365, WebView2), new Outlook, and OWA. Pure static frontend calling Microsoft Graph directly; no backend, no stored data, no secrets. Deployable to any Microsoft 365 tenant.

## Features

### File browsing (Files tab)
- **Recent / OneDrive / SharePoint** sources with search (OneDrive)
- **Explorer-style tree view**: compact rows with inline expand chevrons, file-type icons, and modified-date and size columns; full breadcrumb navigation
- **Pinned shortcuts**: star any site, library, folder, or file to pin it to the top of the Files tab; pins roam with the user across devices (Outlook roaming settings)
- **Home directory**: users can land directly in "their" site/library/folder when opening the SharePoint tab, driven by tenant policy (see below)

### Sharing links
- **Create link** on any file: audience (people in the organization / specific people / anyone), view or edit, expiration date, optional password
- **Insert into the draft** or **copy to clipboard**; recipients prefilled from To/Cc for "specific people" links
- **Existing access**: list current links and permissions on a file, copy or insert them, edit expiration, revoke

### Shared-link review (Manage Links tab)
- Lists everything the user has shared from their OneDrive (delta scan of the `shared` facet, cached for fast incremental refreshes) plus SharePoint items shared through the add-in
- Folder-aware: a shared folder shows as one entry, not thousands -- items inside an already-shared folder are collapsed into it
- Every link shows audience, permission, and expiration, with a red **never expires** badge on links that have none, and a filter to show only never-expiring links
- Prune in place: copy, edit/clear expiration, open the item's folder in OneDrive web, or revoke any link or direct grant without leaving the pane

### Attachments
- **Convert attachments to links** (Attachments tab): uploads selected draft attachments to OneDrive, replaces them with sharing links created with the user's default link settings, and removes the attachments
- **On-send Smart Alert** (OnMessageSend): warns when outgoing attachments exceed a configurable size threshold and offers the conversion instead

### Settings
- Per-user defaults: link audience, permission, expiration days, upload folder, on-send alert threshold (roam with the user)
- **Admin section** (visible to members of the policy admin groups): live policy summary, direct link to edit the policy list in SharePoint, and the full supported key reference

## Tenant policy and home directory (optional)

Admins shape the add-in per tenant without code changes via a SharePoint list named `SFNC_ShareLinks_Policy` on the tenant's root site (two text columns: `Title` = key, `SettingValue` = value). Changes take effect at the next sign-in.

| Key | Controls |
|-|-|
| `allowAnonymousLinks` | allow/deny "Anyone with the link" |
| `adminOnlyAnonymousLinks` | restrict "Anyone" links to policy admins |
| `allowOrganizationLinks` | allow/deny org-wide links |
| `allowSpecificPeopleLinks` | allow/deny specific-people links |
| `defaultScope`, `defaultType`, `defaultExpDays` | tenant defaults for new links |
| `adminGroupIds` | comma-separated group IDs or display names whose members see the admin section |
| `homeAttribute` | Exchange custom attribute slot (1-15) holding each user's home path; default 10 |
| `spHome:<group>` | home path for members of a group, e.g. `VSTH Files/Documents/PROJECTS` |

Policy is enforced at link creation time, not just hidden in the UI. If no policy list exists, all audiences are available and built-in defaults apply.

**Home directory paths** are `Site Name/Library/Folder/...`. Per user: set the Exchange custom attribute, e.g. `Set-Mailbox jdoe -CustomAttribute10 "VSTH Files/Documents"`. Per group: add an `spHome:<group name or ID>` policy row. The user attribute wins over group rules, and users can always navigate back up to all sites.

## Architecture

- Pure static SPA: HTML/CSS/JS served from any static host (currently **GitHub Pages** at `https://sfncadmin.github.io/sfnc-share-links/`)
- Auth via **MSAL.js** (vendored at `src/vendor/msal-browser.min.js`, no CDN or node_modules dependency at runtime) with Nested App Authentication (NAA) where the host supports it, popup fallback otherwise
- All data access is direct delegated calls to **Microsoft Graph**; nothing is proxied or stored server-side
- The Entra app registration is **multi-tenant** (work/school accounts from any organization; personal Microsoft accounts blocked)
- Delegated scopes: `User.Read`, `Files.ReadWrite.All`, `Sites.Read.All`, `GroupMember.Read.All` (admin section), plus optional `SharePointTenantSettings.Read.All` to mirror the tenant's SharePoint sharing defaults when no policy list exists (skipped silently if not consented)

What's public in the repo: the add-in HTML/JS/CSS and `config.js` (the Entra **clientId**). These are not secrets -- this is a public-client SPA with no client secret. No customer data, keys, or tokens are committed.

## Publisher setup (one time)

Already done for the hosted instance; repeat only if standing up a fork under a different app/host.

```powershell
# Register/refresh the multi-tenant Entra app and write src\config.js
.\setup\2026_06.10_Setup_EntraApp_v1.0.ps1 -MultiTenant -BaseUrl https://sfncadmin.github.io/sfnc-share-links
```

Then push to the public repo and enable GitHub Pages (branch root). `manifest.xml` points at the Pages URL.

## Deploying to a customer tenant

Prereq: a Global Admin (or Application + Exchange admin) in the customer tenant.

1. **Grant admin consent** for the app in the customer tenant. Open the admin-consent URL (the setup script prints it; the format is `https://login.microsoftonline.com/organizations/adminconsent?client_id=<clientId>`) signed in as the customer admin, and accept.
2. **Admin-deploy the manifest**: [M365 admin center](https://admin.microsoft.com) > Settings > **Integrated apps** > **Upload custom apps** > app type **Office Add-in** > **Upload manifest file** > choose `manifest.xml` > assign to a pilot group or everyone > **Deploy**. First-time propagation can take a few hours.
3. *(Optional)* Create the `SFNC_ShareLinks_Policy` list on the tenant's SharePoint root site and add policy rows (see table above). Set `adminGroupIds` so the right people get the admin section. The setup script does this in one step:
   ```powershell
   .\setup\2026_06.12_Setup_PolicyList_v1.0.ps1 -AdminGroupIds "IT Team" -Settings @{ defaultExpDays = "30" }
   ```
4. *(Optional)* Set home directories: per-user Exchange custom attributes and/or `spHome:<group>` rows.
5. Users restart Outlook; the **Share Links** button appears on the Message ribbon (compose and read). First use: open a draft > Share Links > Sign in (silent/SSO on NAA-capable hosts, MSAL popup otherwise).

Nothing is installed per machine and there is no per-tenant code or hosting; one hosted instance serves every tenant.

## Shipping a change

Two layers, only one of them slow:

| Layer | Hosts | Update speed |
|-|-|-|
| GitHub Pages | the code (`src/**`) | ~1-2 min after push to `GitMain`, automatic |
| M365 Integrated Apps | the manifest (buttons, permissions, URLs) | up to ~24h, only when `manifest.xml` changes, per tenant |

Code changes are just a push -- Pages rebuilds on its own (no build step). Manifest changes (new ribbon button, changed permissions, renamed, repointed URL) must be re-uploaded in each tenant's Integrated Apps and re-propagate; rare.

**Gotcha -- stale pane after a push.** Outlook's embedded browser (WebView2) caches the old JS, so the pane can show old code even though Pages updated. Fixes, easiest first:

1. Close and reopen the task pane.
2. Restart Outlook.
3. Clear the Office web-add-in cache: delete the contents of `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` while Outlook is closed.

This is a local cache, not a deploy delay -- seconds, not hours.

## Known platform limits (tenant/Graph, not bugs)

- **"Anyone" links** fail if anonymous sharing is disabled in the SharePoint admin center -- Graph returns the policy error and the pane surfaces it.
- **Expiration** is generally only honored on "Anyone" links (SharePoint Online policy). Setting it on org links may be rejected.
- **Password-protected links** are OneDrive Personal only in Graph v1.0 -- expect a clear error on business tenants.
- **Block download** isn't exposed in Graph v1.0 `createLink` -- roadmap (beta API / SharePoint REST).
- The on-send alert threshold is read via roamingSettings; changes apply after Outlook reloads the add-in.

## Project layout

```
manifest.xml                     Production manifest (GitHub Pages URLs) -- admin-deployed per tenant
setup/                           Entra app registration + icon generation scripts
src/config.js                    clientId + scopes (committed; non-secret) -- see config.example.js
src/vendor/msal-browser.min.js   Vendored MSAL (no node_modules/CDN dependency at runtime)
src/taskpane/                    Taskpane UI (taskpane.html/css/js, auth.js = MSAL/NAA, graph.js = Graph calls)
src/launchevent/launchevent.js   OnMessageSend Smart Alert handler (ES5, classic-Outlook JS runtime)
src/commands/commands.html       Event runtime page for new Outlook / OWA
src/assets/                      Icons (generated by setup script)
.nojekyll                        Tells GitHub Pages to serve files as-is
```

## Troubleshooting

- **Sign-in errors**: confirm admin consent was granted in the user's tenant and the redirect URIs on the app registration include the hosting URL.
- **Policy not applying**: the list must be named `SFNC_ShareLinks_Policy` on the tenant root site with `Title`/`SettingValue` text columns; policy loads at sign-in.
- **Smart Alert not firing**: classic Outlook needs an M365 subscription build with event-based activation (Version 2206+); check that the add-in loaded (button visible) and the threshold isn't 0.
- **Add-in button missing**: Integrated Apps propagation can take hours on first deploy; confirm the user is in the assigned group, then restart Outlook.

## Roadmap

- **Rename product to LynxShare** (manifest display name, branding, repo/hosting URL)
- **Streamlined setup wizard**: zero-edit onboarding so deployment doesn't require an M365/PowerShell expert. Running the setup script (or visiting a setup page) launches a Microsoft sign-in directly, prompts for everything interactively, and auto-discovers the tenant ID and other values it needs; the admin never opens or edits the .ps1. Stretch: a web-based wizard that walks through consent, manifest deploy, and policy list creation end to end. Some steps need elevated consent (app registration, Exchange attributes), so the wizard should clearly state required roles up front and degrade gracefully when the signed-in admin lacks one.
- **Snapshot-style links**: optional "send as snapshot" that copies the file at send time and links to the frozen copy, so later edits to the original don't change what the recipient's link serves. Admin policy key to *require* snapshots, especially for SharePoint sources. Not data-room immutability (snapshots are frozen by convention, not enforcement), but an honest middle ground for record documents.
- **Inline education**: "i" info icons beside every sharing choice (audience, expiration, attachment conversion) explaining in plain language how that method works -- in particular, where a converted attachment physically lives and that the emailed link points to it.
- **Protect converted-attachment storage**: today converted attachments land in a OneDrive folder the user can delete during cleanup, silently breaking links in already-sent mail. Candidates: an org-managed SharePoint library with add-but-no-delete permissions (links also survive user offboarding), a warning README file inside the upload folder, and a broken-link health check in the Manage Links tab.
- **View density toggle**: compact rows (current default) vs a larger comfortable view with thumbnails for users who expect modern spacious layouts.
- Block-download links (Graph beta / SharePoint REST)
- Drag-drop upload in the pane
- "Auto-convert on send" instead of warn
- Dark mode
