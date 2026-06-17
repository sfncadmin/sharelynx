# Features

Detailed feature reference for ShareLynx.

---

## File browsing (Files tab)

- **Three sources**: Recent files, OneDrive, and SharePoint -- toggle between them with the segmented control at the top
- **Explorer-style tree view**: compact rows with inline expand chevrons, file-type icons, modified-date and size columns, and full breadcrumb navigation
- **Search**: search OneDrive by keyword (Enter to search, results replace the file list)
- **Pinned shortcuts**: star any site, library, folder, or file to pin it to the top of the Files tab. Pins roam with the user across devices via Outlook roaming settings.
- **Home directory**: users can land directly in "their" site/library/folder when opening the SharePoint tab, driven by tenant policy (see Tenant Policy below)

## Sharing links

- **Create link** on any file with controls for:
  - **Audience**: people in the organization, specific people, or anyone with the link
  - **Permission**: view or edit
  - **Expiration date**
  - **Password** (OneDrive Personal only in Graph v1.0)
- **Insert into the draft** (compose mode) or **copy to clipboard** (compose or read mode)
- **Recipients prefilled** from the email's To/Cc fields when creating "specific people" links
- **Existing access panel**: view all current links and permissions on the selected file, copy or insert existing links, edit expiration, or revoke access

## Shared-link review (Manage Links tab)

- Lists everything the user has shared from their OneDrive, using a delta scan of the `shared` facet with caching for fast incremental refreshes
- Also includes SharePoint items shared through the add-in (tracked in a local registry)
- **Folder-aware grouping**: items inside an already-shared folder are collapsed into the folder entry, so a shared folder shows as one row instead of hundreds
- Each link shows **audience**, **permission**, and **expiration** at a glance
- **Never-expires badge**: red indicator on links with no expiration date
- **Filter**: toggle to show only never-expiring links
- **In-place management**: copy link, edit or clear expiration, open the item's folder in OneDrive web, or revoke any link or direct grant

## Attachment conversion (Attachments tab)

- Lists all file attachments on the current draft
- Select one or more, then **Convert selected to links**:
  1. Uploads each attachment to a configurable OneDrive folder
  2. Creates a sharing link with the user's default settings
  3. Removes the original attachment from the draft
  4. Inserts the link into the email body
- Progress shown per attachment (reading, uploading with percentage, creating link, done)

## On-send Smart Alert

- **OnMessageSend** event handler (runs in classic Outlook's JavaScript-only runtime)
- Warns when outgoing file attachments exceed a configurable size threshold (default 5 MB)
- Lists the oversized attachments by name and size
- Offers a **Convert to links** action that opens the add-in pane
- Threshold of 0 disables the alert; threshold changes apply after Outlook reloads the add-in

## Settings

- **Link defaults**: audience, permission, expiration days
- **Attachment conversion**: upload folder name, on-send alert threshold (MB)
- User settings are stored in browser storage and synced to Outlook roaming settings where available
- **Admin section** (visible only to members of the configured admin groups): live policy summary, direct link to edit the policy list in SharePoint, and the full supported-key reference

---

## Tenant policy

Admins configure the add-in per tenant through a SharePoint list named `ShareLynx_Policy` on the tenant root site. The list has two text columns: `Title` (setting key) and `SettingValue` (value). Changes take effect at the next user sign-in.

### Policy keys

| Key | Type | Controls |
|-|-|-|
| `allowAnonymousLinks` | `true`/`false` | Allow or deny "Anyone with the link" scope |
| `adminOnlyAnonymousLinks` | `true`/`false` | Restrict "Anyone" links to policy admin group members |
| `allowOrganizationLinks` | `true`/`false` | Allow or deny organization-wide links |
| `allowSpecificPeopleLinks` | `true`/`false` | Allow or deny specific-people links |
| `defaultScope` | `organization`, `users`, `anonymous` | Tenant default audience for new links |
| `defaultType` | `view`, `edit` | Tenant default permission level |
| `defaultExpDays` | number | Tenant default expiration (days from creation) |
| `adminGroupIds` | comma-separated | Entra group IDs or display names whose members see the admin section |
| `homeAttribute` | `1`--`15` | Exchange custom attribute slot holding each user's home path (default: 10) |
| `spHome:<group>` | path | Home directory path for members of a group |

### Enforcement

Policy is enforced at link creation time, not just in the UI. If an admin disallows anonymous links via policy, the "Anyone" option is removed from the dropdown and the Graph API call would also reject it. If no policy list exists, all audiences are available and built-in defaults apply.

When the tenant has also consented to `SharePointTenantSettings.Read.All`, the add-in mirrors the SharePoint admin center's sharing defaults automatically as a fallback when no policy list exists.

If the policy list cannot be read because of a permissions, network, or Graph error, the add-in limits risky link types for that session and warns the user instead of silently allowing every option.

### Home directories

Home directory paths follow the format `Site Name/Library/Folder/...` (e.g. `Contoso Files/Documents/PROJECTS`).

- **Per user**: set via Exchange custom attribute -- `Set-Mailbox jdoe -CustomAttribute10 "Contoso Files/Documents"`
- **Per group**: add an `spHome:<group name or ID>` row to the policy list

Per-user attributes take priority over group rules. Users can always navigate back up to the site list from their home directory.

---

## Known platform limits

These are Microsoft Graph and SharePoint Online constraints, not bugs in ShareLynx:

| Limit | Detail |
|-|-|
| "Anyone" links | Fail if anonymous sharing is disabled in the SharePoint admin center. Graph returns the policy error and the add-in surfaces it. |
| Expiration on org links | Generally only honored on "Anyone" links (SharePoint Online policy). Setting expiration on org-wide links may be silently ignored or rejected. |
| Password-protected links | OneDrive Personal only in Graph v1.0. Business tenants get a clear error. |
| Block download | Not exposed in Graph v1.0 `createLink`. On the roadmap via beta API or SharePoint REST. |
| Roaming settings latency | The on-send alert threshold is read from Outlook roaming settings. Changes apply after Outlook reloads the add-in (close/reopen pane or restart Outlook). |
