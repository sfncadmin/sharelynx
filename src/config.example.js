// Shape of src/config.js, written by setup\2026_06.10_Setup_EntraApp_v1.0.ps1.
// config.js IS committed (clientId/tenantId are not secrets for a public-client SPA);
// run the setup script to generate it for a new tenant.
window.SFNC_CONFIG = {
  clientId: "<entra-app-client-id>",
  tenantId: "<entra-tenant-id>",
  scopes: ["User.Read", "Files.ReadWrite.All", "Sites.Read.All"]
};
