<#
.SYNOPSIS
    Creates (or updates) the Entra ID app registration for the ShareLynx
    Outlook add-in, generates manifest.xml, and writes src\config.js.

.DESCRIPTION
    Registers a public client (SPA) app with delegated Graph permissions:
    User.Read, Files.ReadWrite, Sites.Read.All, GroupMember.Read.All.

    Generates manifest.xml from manifest.template.xml, stamping in the
    hosting URL you provide via -BaseUrl (or interactively when omitted).

    By default creates a single-tenant app (AzureADMyOrg). Pass -MultiTenant to
    create an app that accepts sign-ins from any Microsoft 365 organization.

    GroupMember.Read.All is required for the optional admin policy feature and
    requires admin consent. The script outputs an admin-consent URL.

    Requires the Microsoft.Graph.Applications module and an account that can
    create app registrations (Application Developer role or higher).

.EXAMPLE
    .\ShareLynx_Entra_Setup.ps1 -MultiTenant -BaseUrl https://yourorg.github.io/sharelynx
#>
[CmdletBinding()]
param(
    [string]$DisplayName = "ShareLynx (Outlook Add-in)",
    [int]$Port = 3000,
    # Production hosting URL (origin+path). Prompted if omitted.
    [string]$BaseUrl = "",
    # Allow any Microsoft 365 tenant to sign in (omit tenantId from config.js)
    [switch]$MultiTenant
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path $PSScriptRoot -Parent

# --- Prompt for BaseUrl if not provided ---
if (-not $BaseUrl) {
    Write-Host ""
    Write-Host "ShareLynx needs a hosting URL for its static files." -ForegroundColor Cyan
    Write-Host "This is usually a GitHub Pages URL for your fork, e.g.:"
    Write-Host "  https://yourorg.github.io/sharelynx" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Other options work too (Azure Static Web Apps, any HTTPS host)."
    Write-Host ""
    $BaseUrl = Read-Host "Enter your hosting URL"
    if (-not $BaseUrl) {
        Write-Error "A hosting URL is required. Pass -BaseUrl or enter one when prompted."
        return
    }
}
$BaseUrl = $BaseUrl.TrimEnd("/")

# --- Generate manifest.xml from template ---
$templatePath = Join-Path $repoRoot "manifest.template.xml"
$manifestPath = Join-Path $repoRoot "manifest.xml"

if (-not (Test-Path $templatePath)) {
    Write-Error "manifest.template.xml not found at $templatePath"
    return
}

$manifestContent = (Get-Content $templatePath -Raw) -replace '\{\{BASE_URL\}\}', $BaseUrl
$manifestContent | Set-Content -Path $manifestPath -Encoding utf8 -NoNewline
Write-Host "Generated manifest.xml with base URL: $BaseUrl" -ForegroundColor Green

# Add hosting domain to AppDomains if not already present
$hostingHost = ([Uri]$BaseUrl).GetLeftPart([UriPartial]::Authority)
if ($manifestContent -notmatch [regex]::Escape($hostingHost)) {
    $manifestContent = $manifestContent -replace '(  </AppDomains>)', "    <AppDomain>$hostingHost</AppDomain>`n  </AppDomains>"
    $manifestContent | Set-Content -Path $manifestPath -Encoding utf8 -NoNewline
    Write-Host "Added $hostingHost to AppDomains" -ForegroundColor Green
}

if (-not (Get-Module -ListAvailable Microsoft.Graph.Applications)) {
    Write-Host "Installing Microsoft.Graph.Applications module (CurrentUser)..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph.Applications -Scope CurrentUser -Force
}
Import-Module Microsoft.Graph.Applications

Connect-MgGraph -Scopes "Application.ReadWrite.All" -NoWelcome
$ctx = Get-MgContext
$tenantId = $ctx.TenantId
Write-Host "Connected to tenant $tenantId as $($ctx.Account)"

# Resolve delegated permission scope IDs dynamically from the Graph service principal
$graphAppId = "00000003-0000-0000-c000-000000000000"
$graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'"
$scopeNames = @("User.Read", "Files.ReadWrite", "Sites.Read.All", "GroupMember.Read.All")
$resourceAccess = foreach ($name in $scopeNames) {
    $scope = $graphSp.Oauth2PermissionScopes | Where-Object { $_.Value -eq $name }
    if (-not $scope) { throw "Could not resolve Graph delegated permission '$name'." }
    @{ Id = $scope.Id; Type = "Scope" }
}

$originHost = ([Uri]$BaseUrl).Host
$spaRedirects = @(
    "https://localhost:$Port/src/auth-redirect.html",
    "https://localhost:$Port/src/taskpane/taskpane.html",
    "brk-multihub://localhost:$Port",
    "$BaseUrl/src/auth-redirect.html",
    "$BaseUrl/src/taskpane/taskpane.html",
    "brk-multihub://$originHost"
)
Write-Host "Including production redirect URIs for $BaseUrl"

$appParams = @{
    DisplayName            = $DisplayName
    SignInAudience         = if ($MultiTenant) { "AzureADMultipleOrgs" } else { "AzureADMyOrg" }
    Spa                    = @{ RedirectUris = $spaRedirects }
    RequiredResourceAccess = @(@{ ResourceAppId = $graphAppId; ResourceAccess = $resourceAccess })
}

$existing = Get-MgApplication -Filter "displayName eq '$DisplayName'" | Select-Object -First 1
if ($existing) {
    Write-Host "Updating existing app registration '$DisplayName'..."
    Update-MgApplication -ApplicationId $existing.Id @appParams
    $app = Get-MgApplication -ApplicationId $existing.Id
} else {
    Write-Host "Creating app registration '$DisplayName'..."
    $app = New-MgApplication @appParams
}

$clientId = $app.AppId
Write-Host "App (client) ID: $clientId" -ForegroundColor Green

# Write src\config.js. The clientId is not a secret for this public-client SPA,
# so the file can be committed with the hosted static add-in.
$configPath = Join-Path (Split-Path $PSScriptRoot -Parent) "src\config.js"
if ($MultiTenant) {
    $configContent = @"
// Generated by setup\ShareLynx_Entra_Setup.ps1 on $(Get-Date -Format yyyy-MM-dd)
// Multi-tenant: any Microsoft 365 organization can sign in.
window.SHARELYNX_CONFIG = {
  clientId: "$clientId",
  scopes: ["User.Read", "Files.ReadWrite", "Sites.Read.All", "GroupMember.Read.All"]
};
"@
} else {
    $configContent = @"
// Generated by setup\ShareLynx_Entra_Setup.ps1 on $(Get-Date -Format yyyy-MM-dd)
window.SHARELYNX_CONFIG = {
  clientId: "$clientId",
  tenantId: "$tenantId",
  scopes: ["User.Read", "Files.ReadWrite", "Sites.Read.All", "GroupMember.Read.All"]
};
"@
}
$configContent | Set-Content -Path $configPath -Encoding utf8
Write-Host "Wrote $configPath" -ForegroundColor Green

Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1. Host the repo files at: $BaseUrl" -ForegroundColor Yellow
Write-Host "     (e.g. enable GitHub Pages on your fork, or deploy to any HTTPS host)"
Write-Host "  2. Grant admin consent for GroupMember.Read.All (required for the policy feature):"
Write-Host "     https://login.microsoftonline.com/$tenantId/adminconsent?client_id=$clientId" -ForegroundColor Yellow
Write-Host "  3. Upload manifest.xml in M365 admin center > Integrated Apps (see README.md)"
Write-Host "  4. (Optional) Create a 'ShareLynx_Policy' SharePoint list on your root site"
Write-Host "     with text columns Title and SettingValue to configure per-tenant link policies."
