// Shape of src/config.js, written by setup\ShareLynx_Entra_Setup.ps1.
// config.js IS committed (clientId is not a secret for a public-client SPA).
window.SHARELYNX_CONFIG = {
  clientId: "<entra-app-client-id>",
  // tenantId is optional. Omit (or remove) to allow any Microsoft 365 tenant to sign in.
  // Set it to restrict sign-in to your organization only.
  // tenantId: "<entra-tenant-id>",
  scopes: ["User.Read", "Files.ReadWrite", "Sites.Read.All", "GroupMember.Read.All"],
  // To also mirror the tenant's SharePoint admin sharing defaults when no policy list exists,
  // add "SharePointTenantSettings.Read.All" to scopes and re-run the setup script.
  // Requires SharePoint admin consent. App works fine without it (call is silently skipped).
  // Optional: SharePoint site ID containing the ShareLynx_Policy list.
  // Omit to use the root site. Admins: find a site's ID at
  //   GET https://graph.microsoft.com/v1.0/sites/<hostname>:/<path>
  // adminPolicySite: "contoso.sharepoint.com,abc123-...,abc456-..."
};
