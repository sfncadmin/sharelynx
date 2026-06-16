# Deployment Guide

How to set up, deploy, update, and troubleshoot ShareLynx.

---

## Publisher setup (one time)

This is already done for the hosted instance at GitHub Pages. Only repeat if standing up a fork under a different Entra app or host.

### 1. Register the Entra app

```powershell
.\setup\ShareLynx_Entra_Setup.ps1 -MultiTenant -BaseUrl https://your-host/your-repo
```

The script:
- Creates (or updates) a public-client SPA app registration with delegated Graph permissions
- Writes `src/config.js` with the resulting `clientId` and scopes
- Prints an admin-consent URL for `GroupMember.Read.All` (needed for the policy/admin feature)

Pass `-MultiTenant` to allow any Microsoft 365 organization to sign in. Without it, the app is single-tenant (your org only).

### 2. Enable GitHub Pages

Push to the repo, then enable Pages (Settings > Pages > branch root). `manifest.xml` references the Pages URL for all source locations.

---

## Deploying to a tenant

Requires a Global Admin, or an account with both Application Admin and Exchange Admin roles.

### 1. Grant admin consent

Open the admin-consent URL in a browser, signed in as the tenant admin:

```
https://login.microsoftonline.com/organizations/adminconsent?client_id=<clientId>
```

The setup script prints this URL. Accept the permissions.

### 2. Deploy the manifest

1. Go to [M365 admin center](https://admin.microsoft.com) > Settings > **Integrated apps**
2. **Upload custom apps** > app type **Office Add-in** > **Upload manifest file**
3. Choose `manifest.xml`
4. Assign to a pilot group or everyone
5. **Deploy**

First-time propagation can take a few hours.

### 3. (Optional) Create the policy list

```powershell
.\setup\ShareLynx_Policy_Setup.ps1 -AdminGroupIds "IT Team" -Settings @{
    defaultExpDays      = "30"
    allowAnonymousLinks = "false"
}
```

This creates (or updates) the `ShareLynx_Policy` SharePoint list on the tenant root site. See [features.md](features.md) for the full policy key reference.

### 4. (Optional) Set home directories

- **Per user**: `Set-Mailbox jdoe -CustomAttribute10 "Site Name/Library/Folder"`
- **Per group**: add an `spHome:<group name or ID>` row to the policy list

### 5. Users open Outlook

Restart Outlook. The **ShareLynx** button appears on the Message ribbon (compose and read). First use: open a draft > ShareLynx > sign in.

Nothing is installed per machine. There is no per-tenant code or hosting -- one hosted instance serves every tenant.

---

## Shipping changes

Two layers, only one of them slow:

| Layer | What it hosts | Update speed |
|-|-|-|
| GitHub Pages | the code (`src/**`) | ~1--2 min after push, automatic |
| M365 Integrated Apps | the manifest (buttons, permissions, URLs) | up to ~24h, only when `manifest.xml` changes |

Code changes are just a push -- Pages rebuilds on its own (no build step). Manifest changes (new ribbon button, changed permissions, repointed URL) must be re-uploaded in each tenant's Integrated Apps and re-propagate. This is rare.

### Bump the version on every manifest redeploy

Before re-uploading a changed `manifest.xml`, increment the `<Version>` element (e.g. `1.1.0.0` -> `1.2.0.0`). The `<Id>` GUID stays the same so M365 recognizes it as the same add-in; the version is how it detects a newer build to replace. Re-uploading at the same version is rejected with:

> Failed. Please update the version number in the manifest file and try again.

Commit and push after bumping so the hosted copy at the Pages URL matches the file you upload. Pure `src/**` code changes do not need a bump -- Pages serves those live.

### Stale pane after a push

Outlook's embedded browser (WebView2) caches old JS, so the pane can show old code even though Pages updated. Fixes, easiest first:

1. Close and reopen the task pane.
2. Restart Outlook.
3. Clear the Office web-add-in cache: delete the contents of `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` while Outlook is closed.

This is a local cache issue, not a deploy delay -- seconds, not hours.

---

## Troubleshooting

**Sign-in errors**
Confirm admin consent was granted in the user's tenant and that the redirect URIs on the app registration include the hosting URL.

**Policy not applying**
The list must be named `ShareLynx_Policy` on the tenant root site with `Title` and `SettingValue` text columns. Policy loads at sign-in -- changes take effect after the next sign-in.

**Smart Alert not firing**
Classic Outlook needs an M365 subscription build with event-based activation (Version 2206+). Check that the add-in loaded (button visible on the ribbon) and the on-send threshold isn't set to 0.

**Add-in button missing after deploy**
Integrated Apps propagation can take hours on first deploy. Confirm the user is in the assigned group, then restart Outlook.

**Stale code after a push**
See "Stale pane after a push" above.
