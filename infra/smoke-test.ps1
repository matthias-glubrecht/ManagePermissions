#requires -Version 7.0
<#
.SYNOPSIS
    Rauchtest für die ManagePermissions-Function: prüft nebenwirkungsfrei, ob der
    Erfolgs-Pfad (Aufrufer-Token -> Function -> Managed Identity -> SharePoint) grün ist.

.DESCRIPTION
    Sendet einen 'reset'-Aufruf mit einer ECHTEN Listen-GUID, aber einer (bewusst)
    nicht existierenden Element-ID. Der Aufruf durchläuft die komplette Kette bis
    SharePoint, ändert aber garantiert nichts.

    Interpretation des HTTP-Status:
      200 / 404  -> GRÜN: Function erreicht SharePoint via Managed Identity.
                          (404 "Element nicht gefunden" ist hier das Erfolgssignal.)
      502        -> Site-Grant fehlt ODER MI-Token-Cache noch alt.
                    Abhilfe: grant-sites-selected.ps1 ausführen, dann mit
                    -RestartFunction erneut testen.
      401 / 403  -> Aufrufer-TOKEN/Gruppe: das az-Konto hat den Scope
                    'access_as_user' nicht oder ist nicht Mitglied der AllowedGroup
                    (NICHT das MI-Problem). Alternativ über das SPFx-Web-Part testen.
      400        -> Eingabe (SiteUrl/ListId) ungültig.

    Optional leert -RestartFunction vorab den MI-Token-Cache (stop + start), damit ein
    frisch erteilter Site-Grant sofort greift, statt bis zu ~24 h zu warten.

    Nutzt ausschließlich Azure CLI + Invoke-RestMethod — keine PnP/Graph-Module,
    daher kein Assembly-Konflikt und keine frische Session nötig.

.PARAMETER SiteUrl
    Vollständige URL der Ziel-Site, z. B. https://contoso.sharepoint.com/sites/team.

.PARAMETER ListId
    GUID einer ECHTEN Liste auf der Ziel-Site (wird real gelesen, aber nicht verändert).

.PARAMETER ClientId
    Client-ID der API-App-Registrierung (Token-Audience api://<client-id>).
    Default: aus infra/.deployment-output.json (von deploy.ps1), falls vorhanden.

.PARAMETER FunctionAppName
    Name der Function App. Default: aus .deployment-output.json, sonst 'func-wsperms'.

.PARAMETER ResourceGroup
    Resource Group der Function App. Default: aus .deployment-output.json, sonst 'rg-workshop'.

.PARAMETER ItemId
    Bewusst nicht existierende Element-ID. Default: 999999.

.PARAMETER RestartFunction
    Leert vor dem Test den MI-Token-Cache (az functionapp stop + start).

.EXAMPLE
    pwsh -NoProfile -File ./infra/smoke-test.ps1 `
        -SiteUrl https://contoso.sharepoint.com/sites/team `
        -ListId  00000000-0000-0000-0000-000000000000

.EXAMPLE
    # Nach frisch erteiltem Site-Grant: Cache leeren und prüfen
    pwsh -NoProfile -File ./infra/smoke-test.ps1 `
        -SiteUrl https://contoso.sharepoint.com/sites/team `
        -ListId  00000000-0000-0000-0000-000000000000 `
        -RestartFunction
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $SiteUrl,
    [Parameter(Mandatory = $true)] [string] $ListId,
    [Parameter()] [string] $ClientId,
    [Parameter()] [string] $FunctionAppName,
    [Parameter()] [string] $ResourceGroup,
    [Parameter()] [int]    $ItemId = 999999,
    [Parameter()] [switch] $RestartFunction
)

$ErrorActionPreference = 'Stop'

function Write-Step { param($Msg) Write-Host "==> $Msg" -ForegroundColor Cyan }
function Write-Ok   { param($Msg) Write-Host "    [OK] $Msg" -ForegroundColor Green }
function Write-Warn { param($Msg) Write-Host "    [!]  $Msg" -ForegroundColor Yellow }
function Write-Fail { param($Msg) Write-Host "    [X]  $Msg" -ForegroundColor Red }

# --- Defaults aus dem deploy-Output laden ------------------------------------
$outputFile = Join-Path $PSScriptRoot '.deployment-output.json'
if (Test-Path $outputFile) {
    $deployed = Get-Content $outputFile -Raw | ConvertFrom-Json
    if (-not $ClientId)        { $ClientId        = $deployed.clientId }
    if (-not $FunctionAppName) { $FunctionAppName = $deployed.functionApp }
    if (-not $ResourceGroup)   { $ResourceGroup   = $deployed.resourceGroup }
}
if (-not $FunctionAppName) { $FunctionAppName = 'func-wsperms' }
if (-not $ResourceGroup)   { $ResourceGroup   = 'rg-workshop' }
if (-not $ClientId) {
    throw "ClientId fehlt. Per -ClientId angeben oder infra/.deployment-output.json bereitstellen (deploy.ps1)."
}

# --- Eingabe validieren ------------------------------------------------------
$parsedGuid = [Guid]::Empty
if (-not [Guid]::TryParse($ListId, [ref] $parsedGuid)) {
    throw "ListId '$ListId' ist keine gültige GUID."
}
$siteUri = $null
if (-not [Uri]::TryCreate($SiteUrl, [UriKind]::Absolute, [ref] $siteUri) -or $siteUri.Scheme -ne 'https') {
    throw "SiteUrl '$SiteUrl' ist keine gültige https-URL."
}

# --- Anmeldung prüfen --------------------------------------------------------
Write-Step 'Prüfe Azure-CLI-Anmeldung ...'
$null = az account show 2>$null
if ($LASTEXITCODE -ne 0) { throw "Nicht angemeldet. Bitte 'az login' ausführen." }
Write-Ok 'Angemeldet'

# --- Optional: MI-Token-Cache leeren -----------------------------------------
if ($RestartFunction) {
    Write-Step "Leere MI-Token-Cache (stop + start von '$FunctionAppName') ..."
    az functionapp stop  --name $FunctionAppName --resource-group $ResourceGroup --output none
    az functionapp start --name $FunctionAppName --resource-group $ResourceGroup --output none
    Write-Ok 'Function neu gestartet — erster Aufruf holt ein frisches MI-Token'
}

# --- Endpunkt ermitteln ------------------------------------------------------
Write-Step "Ermittle Endpunkt von '$FunctionAppName' ..."
$appHost = az functionapp show --name $FunctionAppName --resource-group $ResourceGroup --query defaultHostName -o tsv
if ($LASTEXITCODE -ne 0 -or -not $appHost) { throw "Function App '$FunctionAppName' nicht gefunden." }
$endpoint = "https://$appHost/api/ManagePermissions"
Write-Ok $endpoint

# --- Aufrufer-Token holen ----------------------------------------------------
Write-Step "Hole Aufrufer-Token (api://$ClientId) ..."
$token = az account get-access-token --resource "api://$ClientId" --query accessToken -o tsv
if ($LASTEXITCODE -ne 0 -or -not $token) {
    throw "Token konnte nicht beschafft werden (Tenant-Richtlinie?). Ggf. über das SPFx-Web-Part testen."
}
Write-Ok 'Token erhalten'

# --- Nebenwirkungsfreier reset-Aufruf ----------------------------------------
Write-Step "Sende nebenwirkungsfreien 'reset' (itemId=$ItemId, wird nichts verändert) ..."
$body = @{
    action = 'reset'
    webUrl = $SiteUrl
    listId = $ListId
    itemId = $ItemId
} | ConvertTo-Json

$status = 0
$response = Invoke-RestMethod -Method Post -Uri $endpoint `
    -Headers @{ Authorization = "Bearer $token" } `
    -ContentType 'application/json' -Body $body `
    -SkipHttpErrorCheck -StatusCodeVariable status

$serverMessage =
    if     ($response.message) { $response.message }
    elseif ($response.error)   { $response.error }
    else                       { '(keine Meldung)' }

# --- Auswertung --------------------------------------------------------------
Write-Host ''
Write-Host "HTTP $status — $serverMessage"
Write-Host ''

switch ($status) {
    { $_ -in 200, 404 } {
        Write-Ok 'GRÜN: Function erreicht SharePoint über die Managed Identity.'
        Write-Ok 'Der Erfolgs-Pfad ist vorführbereit.'
        exit 0
    }
    502 {
        Write-Fail 'ROT (502): Zugriff auf SharePoint verweigert.'
        Write-Warn 'Site-Grant fehlt oder MI-Token-Cache ist noch alt.'
        Write-Warn 'Abhilfe: grant-sites-selected.ps1 ausführen, dann erneut mit -RestartFunction testen.'
        exit 1
    }
    { $_ -in 401, 403 } {
        Write-Warn "GELB ($status): Aufrufer-Token/Gruppe — NICHT das MI-Problem."
        Write-Warn 'Das az-Konto braucht den Scope access_as_user UND Mitgliedschaft in der AllowedGroup.'
        Write-Warn 'Alternativ den Erfolgs-Pfad über das SPFx-Web-Part verifizieren.'
        exit 2
    }
    400 {
        Write-Fail 'ROT (400): Eingabe ungültig — prüfe SiteUrl/ListId.'
        exit 3
    }
    default {
        Write-Fail "Unerwarteter Status $status."
        exit 9
    }
}
