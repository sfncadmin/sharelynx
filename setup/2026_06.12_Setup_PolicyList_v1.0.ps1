<#
.SYNOPSIS
    Creates the SFNC_ShareLinks_Policy SharePoint list that the Share Links
    add-in reads for tenant policy, and optionally seeds it with settings.

.DESCRIPTION
    Creates a list named SFNC_ShareLinks_Policy on the tenant root site
    (or the site passed via -SiteId) with a SettingValue text column, which
    is the layout the add-in expects (Title = key, SettingValue = value).

    Safe to re-run: an existing list is reused and existing keys are
    updated in place rather than duplicated.

    Requires an account that can create lists on the target site
    (site owner or SharePoint admin). Sign-in is interactive; nothing is
    stored. Users pick up policy changes at their next sign-in to the add-in.

.EXAMPLE
    .\2026_06.12_Setup_PolicyList_v1.0.ps1 -AdminGroupIds "IT Team"

.EXAMPLE
    .\2026_06.12_Setup_PolicyList_v1.0.ps1 -AdminGroupIds "IT Team" -Settings @{
        defaultExpDays      = "30"
        allowAnonymousLinks = "false"
        "spHome:LKV Staff"  = "LKV Files/Documents"
    }
#>
[CmdletBinding()]
param(
    # Comma-separated Entra group display names or IDs whose members see the admin section
    [string]$AdminGroupIds = "",
    # Additional policy keys to seed, e.g. @{ defaultExpDays = "30" }
    [hashtable]$Settings = @{},
    # Site that hosts the list; default is the tenant root site
    [string]$SiteId = "root"
)

$ErrorActionPreference = "Stop"
$ListName = "SFNC_ShareLinks_Policy"

if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
    Write-Host "Installing Microsoft.Graph.Authentication module (CurrentUser)..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
}
Import-Module Microsoft.Graph.Authentication

Connect-MgGraph -Scopes "Sites.Manage.All" -NoWelcome

$site = Invoke-MgGraphRequest -Method GET -Uri "v1.0/sites/$SiteId"
Write-Host "Site: $($site.displayName) ($($site.webUrl))"

$lists = Invoke-MgGraphRequest -Method GET -Uri "v1.0/sites/$($site.id)/lists?`$select=id,displayName,webUrl&`$top=200"
$list = $lists.value | Where-Object { $_.displayName -eq $ListName } | Select-Object -First 1
if ($list) {
    Write-Host "List '$ListName' already exists; reusing it."
} else {
    $body = @{
        displayName = $ListName
        description = "Tenant policy for the Share Links Outlook add-in. Title = setting key, SettingValue = value."
        list        = @{ template = "genericList" }
        columns     = @(@{ name = "SettingValue"; text = @{} })
    } | ConvertTo-Json -Depth 5
    $list = Invoke-MgGraphRequest -Method POST -Uri "v1.0/sites/$($site.id)/lists" -Body $body -ContentType "application/json"
    Write-Host "Created list '$ListName'." -ForegroundColor Green
}

$seed = @{}
foreach ($k in $Settings.Keys) { $seed[$k] = [string]$Settings[$k] }
if ($AdminGroupIds) { $seed["adminGroupIds"] = $AdminGroupIds }

if ($seed.Count -eq 0) {
    Write-Host "No settings passed; list left as-is. Add rows in SharePoint (Title = key, SettingValue = value)."
} else {
    $existing = Invoke-MgGraphRequest -Method GET -Uri "v1.0/sites/$($site.id)/lists/$($list.id)/items?`$expand=fields(`$select=Title)&`$select=id,fields&`$top=200"
    foreach ($key in $seed.Keys) {
        $row = $existing.value | Where-Object { $_.fields.Title -eq $key } | Select-Object -First 1
        if ($row) {
            $body = @{ SettingValue = $seed[$key] } | ConvertTo-Json
            Invoke-MgGraphRequest -Method PATCH -Uri "v1.0/sites/$($site.id)/lists/$($list.id)/items/$($row.id)/fields" -Body $body -ContentType "application/json" | Out-Null
            Write-Host "Updated $key = $($seed[$key])"
        } else {
            $body = @{ fields = @{ Title = $key; SettingValue = $seed[$key] } } | ConvertTo-Json -Depth 4
            Invoke-MgGraphRequest -Method POST -Uri "v1.0/sites/$($site.id)/lists/$($list.id)/items" -Body $body -ContentType "application/json" | Out-Null
            Write-Host "Added $key = $($seed[$key])"
        }
    }
}

$listUrl = if ($list.webUrl) { $list.webUrl } else { "$($site.webUrl)/Lists/$ListName" }
Write-Host ""
Write-Host "Done. List: $listUrl" -ForegroundColor Green
Write-Host "Supported keys: allowAnonymousLinks, adminOnlyAnonymousLinks, allowOrganizationLinks,"
Write-Host "  allowSpecificPeopleLinks, defaultScope, defaultType, defaultExpDays, adminGroupIds,"
Write-Host "  homeAttribute, spHome:<group>"
Write-Host "Users pick up policy at their next sign-in to the add-in."
