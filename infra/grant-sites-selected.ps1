<#
.SYNOPSIS
    Berechtigt die Managed Identity der ManagePermissions-Function auf GENAU EINE
    SharePoint-Site (Sites.Selected + FullControl — least privilege).

.DESCRIPTION
    Um Berechtigungen auf Listenelementen zu verwalten, braucht die Managed Identity
    pro Ziel-Site die Rolle "fullcontrol" (laut Microsoft-Doku reicht "write" NICHT,
    siehe https://learn.microsoft.com/graph/permissions-selected-overview).

    Das Skript führt nacheinander aus:

      1. Graph-App-Rolle "Sites.Selected" auf der MI sicherstellen
         (Resource: Microsoft Graph, appId 00000003-0000-0000-c000-000000000000).

      2. SharePoint-App-Rolle "Sites.Selected" auf der MI sicherstellen
         (Resource: Office 365 SharePoint Online, appId 00000003-0000-0ff1-ce00-000000000000).
         Das PnP Core SDK nutzt SharePoint-REST (CSOM) — dafür ist ein SharePoint-Token
         nötig; die Graph-Rolle allein genügt für CSOM NICHT.

      3. Site-ID der Ziel-Site über Microsoft Graph ermitteln.

      4. Eintrag /sites/{site-id}/permissions mit Rolle "fullcontrol" anlegen
         (bzw. bestehende Rolle aktualisieren).

    Schritt 1+2 erfordern einmalig Tenant-Admin-Rechte.

.PARAMETER FunctionAppName
    Name der Function App, deren System-Assigned MI berechtigt wird.

.PARAMETER ResourceGroup
    Resource Group der Function App.

.PARAMETER SiteUrl
    Vollständige URL der Ziel-Site, z. B. https://contoso.sharepoint.com/sites/team.

.PARAMETER Role
    Site-Rolle der MI. Default: fullcontrol (für Listenelement-Berechtigungen erforderlich).

.EXAMPLE
    pwsh -NoProfile -File ./grant-sites-selected.ps1 `
        -FunctionAppName func-wsperms -ResourceGroup rg-workshop `
        -SiteUrl https://contoso.sharepoint.com/sites/team
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $FunctionAppName,
    [Parameter(Mandatory = $true)] [string] $ResourceGroup,
    [Parameter(Mandatory = $true)] [string] $SiteUrl,
    [Parameter()] [ValidateSet('read', 'write', 'owner', 'fullcontrol')] [string] $Role = 'fullcontrol'
)

$ErrorActionPreference = 'Stop'

function Write-Step { param($Msg) Write-Host "==> $Msg" -ForegroundColor Cyan }
function Write-Ok { param($Msg) Write-Host "    [OK] $Msg" -ForegroundColor Green }

# Bekannte App-IDs der Ressourcen
$GraphAppId = '00000003-0000-0000-c000-000000000000'   # Microsoft Graph
$SpoAppId   = '00000003-0000-0ff1-ce00-000000000000'   # Office 365 SharePoint Online

# --- Prerequisites -----------------------------------------------------------
Write-Step "Prüfe Module und Anmeldung ..."

# PnP.PowerShell und Microsoft.Graph bringen inkompatible Microsoft.Identity.Client.dll
# mit. Ist PnP geladen, scheitert Connect-MgGraph → in frischer Session ausführen.
if (Get-Module -Name 'PnP.PowerShell') {
    throw "PnP.PowerShell ist geladen (Assembly-Konflikt mit Microsoft.Graph). Bitte in frischer Session ausführen: pwsh -NoProfile -File '$PSCommandPath' ..."
}
foreach ($mod in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications', 'Microsoft.Graph.Sites')) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        throw "Modul '$mod' nicht installiert. Installiere mit: Install-Module Microsoft.Graph -Scope CurrentUser"
    }
}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Applications   -ErrorAction Stop
Import-Module Microsoft.Graph.Sites          -ErrorAction Stop

$null = az account show 2>$null
if ($LASTEXITCODE -ne 0) { throw "Bitte 'az login' ausführen." }

# --- 1. MI Principal-ID ermitteln --------------------------------------------
Write-Step "Hole Managed Identity der Function App '$FunctionAppName' ..."
$mi = az functionapp identity show `
    --name $FunctionAppName `
    --resource-group $ResourceGroup `
    | ConvertFrom-Json
if (-not $mi -or -not $mi.principalId) {
    throw "Function App '$FunctionAppName' hat keine System-Assigned Managed Identity."
}
$miPrincipalId = $mi.principalId
Write-Ok "MI Principal-ID: $miPrincipalId"

# --- Connect Graph -----------------------------------------------------------
Write-Step "Verbinde mit Microsoft Graph (Tenant-Admin nötig) ..."
Connect-MgGraph -Scopes 'AppRoleAssignment.ReadWrite.All', 'Sites.FullControl.All', 'Application.Read.All' -NoWelcome | Out-Null
$ctx = Get-MgContext
Write-Ok "Tenant: $($ctx.TenantId)"

# --- Hilfsfunktion: App-Rolle 'Sites.Selected' zuweisen ----------------------
function Grant-AppRole {
    param([string] $ResourceAppId, [string] $Label)

    $resourceSp = Get-MgServicePrincipal -Filter "appId eq '$ResourceAppId'"
    if (-not $resourceSp) { throw "Service Principal '$Label' ($ResourceAppId) nicht gefunden." }

    $role = $resourceSp.AppRoles | Where-Object {
        $_.Value -eq 'Sites.Selected' -and $_.AllowedMemberTypes -contains 'Application'
    }
    if (-not $role) { throw "AppRole 'Sites.Selected' auf '$Label' nicht gefunden." }

    $existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $miPrincipalId `
        | Where-Object { $_.AppRoleId -eq $role.Id -and $_.ResourceId -eq $resourceSp.Id }
    if ($existing) {
        Write-Ok "'Sites.Selected' ($Label) war bereits zugewiesen"
    }
    else {
        $null = New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $miPrincipalId `
            -PrincipalId $miPrincipalId `
            -ResourceId  $resourceSp.Id `
            -AppRoleId   $role.Id
        Write-Ok "'Sites.Selected' ($Label) zugewiesen"
    }
}

Write-Step "Stelle 'Sites.Selected' (Microsoft Graph) sicher ..."
Grant-AppRole -ResourceAppId $GraphAppId -Label 'Microsoft Graph'

Write-Step "Stelle 'Sites.Selected' (SharePoint Online) sicher ..."
Grant-AppRole -ResourceAppId $SpoAppId -Label 'SharePoint Online'

# --- 3. Site-ID ermitteln ----------------------------------------------------
Write-Step "Ermittle Site-ID für '$SiteUrl' ..."
$uri = [Uri]$SiteUrl
$hostName = $uri.Host
$serverRel = $uri.AbsolutePath.TrimEnd('/')
$site = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/${hostName}:${serverRel}"
$siteId = $site.id  # Format: hostname,siteCollectionId,webId
if (-not $siteId) { throw "Konnte Site-ID nicht ermitteln." }
Write-Ok "Site-ID: $siteId"

# --- 4. Site-Permission anlegen/aktualisieren --------------------------------
Write-Step "Lege /sites/{id}/permissions mit Rolle '$Role' an ..."
$existingPerms = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/permissions"
$alreadyGranted = $existingPerms.value | Where-Object {
    $_.grantedToIdentitiesV2 | Where-Object {
        $_.application -and $_.application.displayName -eq $FunctionAppName
    }
}

if ($alreadyGranted) {
    Write-Ok "Permission war bereits gesetzt — aktualisiere Rolle auf '$Role'."
    foreach ($perm in $alreadyGranted) {
        $null = Invoke-MgGraphRequest -Method PATCH `
            -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/permissions/$($perm.id)" `
            -Body (@{ roles = @($Role) } | ConvertTo-Json) -ContentType 'application/json'
    }
}
else {
    $miSp = Get-MgServicePrincipal -ServicePrincipalId $miPrincipalId
    $body = @{
        roles               = @($Role)
        # Im POST-Body heißt das Feld "grantedToIdentities" (nicht ...V2).
        grantedToIdentities = @(
            @{
                application = @{
                    id          = $miSp.AppId
                    displayName = $FunctionAppName
                }
            }
        )
    } | ConvertTo-Json -Depth 10

    $null = Invoke-MgGraphRequest -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/permissions" `
        -Body $body -ContentType 'application/json'
    Write-Ok "Permission gesetzt"
}

# --- Verification ------------------------------------------------------------
Write-Step "Verifiziere ..."
$perms = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/permissions"
$ourPerm = $perms.value | Where-Object {
    $_.grantedToIdentitiesV2 | Where-Object {
        $_.application -and $_.application.displayName -eq $FunctionAppName
    }
}
if (-not $ourPerm) { throw "Verifikation fehlgeschlagen — keine Permission für '$FunctionAppName' gefunden." }
Write-Ok ("Aktive Rolle(n) für '{0}' auf Site: {1}" -f $FunctionAppName, ($ourPerm.roles -join ', '))

Write-Host ""
Write-Host "Sites.Selected-Setup abgeschlossen." -ForegroundColor Green
Write-Host ""
Write-Host "Hinweis: Token-Caches der Managed Identity können bis zu ~24 h alt sein." -ForegroundColor Yellow
Write-Host "Bei 401-Fehlern nach Rollenänderung: 'az functionapp stop/start' oder Ablauf abwarten." -ForegroundColor Yellow

Disconnect-MgGraph | Out-Null
