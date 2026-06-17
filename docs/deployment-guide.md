# Deployment Guide

How to set up, deploy, update, and troubleshoot ShareLynx.

---

## Publisher setup (one time)

### 1. Register the Entra app and generate the manifest

```powershell
.\setup\ShareLynx_Entra_Setup.ps1 -MultiTenant -BaseUrl https://yourorg.github.io/sharelynx
```

If you omit `-BaseUrl`, the script prompts for it. Use whichever HTTPS URL will serve the repo files -- GitHub Pages, Azure Static Web Apps, or any static host.

The script:
- Creates (or updates) a public-client SPA app registration with delegated Graph permissions
- Generates `manifest.xml` from `manifest.template.xml`, stamped with your hosting URL
- Writes `src/config.js` with the resulting `clientId` and scopes
- Prints an admin-consent URL for `GroupMember.Read.All` (needed for the policy/admin feature)

Pass `-MultiTenant` to allow any Microsoft 365 organization to sign in. Without it, the app is single-tenant (your org only).

### 2. Host the files

Push to a public repo and enable GitHub Pages (Settings > Pages > branch root), or deploy the files to your static host of choice. The generated `manifest.xml` already points at your URL.

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

Restart Outlook. The **ShareLynx** button appears on the Message ribbon (compose and read) for Outlook clients that support the Mailbox 1.12 requirement set. First use: open a draft > ShareLynx > sign in.

Nothing is installed per machine. There is no per-tenant code -- one hosted instance serves every tenant.

---

## Shipping changes

Two layers, only one of them slow:

| Layer | What it hosts | Update speed |
|-|-|-|
| Your static host | the code (`src/**`) | depends on host (GitHub Pages: ~1--2 min) |
| M365 Integrated Apps | the manifest (buttons, permissions, URLs) | up to ~24h, only when `manifest.xml` changes |

Code changes are just a push -- no build step. Manifest changes (new ribbon button, changed permissions, repointed URL) must be re-uploaded in each tenant's Integrated Apps and re-propagate. This is rare.

### Regenerate and re-upload the manifest

When `manifest.template.xml` changes, an admin must regenerate `manifest.xml` and upload the new file in the M365 admin center. The setup script does **not** bump the manifest version automatically.

Before regenerating and re-uploading, increment the `<Version>` element in `manifest.template.xml` (e.g. `1.1.0.0` -> `1.2.0.0`). The `<Id>` GUID stays the same so M365 recognizes it as the same add-in; the version is how it detects a newer build to replace. Re-uploading at the same version is rejected with:

> Failed. Please update the version number in the manifest file and try again.

Then regenerate `manifest.xml` with:

```powershell
.\setup\ShareLynx_Entra_Setup.ps1 -MultiTenant -BaseUrl https://yourorg.github.io/sharelynx
```

For this repository's GitHub Pages URL, use:

```powershell
.\setup\ShareLynx_Entra_Setup.ps1 -MultiTenant -BaseUrl https://sfncadmin.github.io/sharelynx
```

The script to run is `setup\ShareLynx_Entra_Setup.ps1`. It updates the Entra app registration, writes `src\config.js`, and generates `manifest.xml` from `manifest.template.xml`.

Commit and push after bumping and regenerating so the hosted copy matches the file you upload. Pure `src/**` code changes do not need a manifest bump -- your host serves those live.

If you change your hosting URL, re-run the setup script with the new `-BaseUrl` to regenerate `manifest.xml`, then re-upload it.

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
The list must be named `ShareLynx_Policy` on the tenant root site with `Title` and `SettingValue` text columns. Policy loads at sign-in -- changes take effect after the next sign-in. If the add-in cannot read the policy list, it limits risky link types for that session and shows a warning.

**Smart Alert not firing**
The Outlook client must support the Mailbox 1.12 requirement set and event-based activation. Check that the add-in loaded (button visible on the ribbon) and the on-send threshold isn't set to 0.

**Add-in button missing after deploy**
Integrated Apps propagation can take hours on first deploy. Confirm the user is in the assigned group, then restart Outlook.

**Stale code after a push**
See "Stale pane after a push" above.
