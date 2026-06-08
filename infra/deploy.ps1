<#
.SYNOPSIS
    Deployt die Azure-Infrastruktur für die ManagePermissions-Function (C# / .NET 8 isolated).

.DESCRIPTION
    Erzeugt idempotent (mehrfaches Ausführen aktualisiert, statt zu fehlschlagen):
      - Resource Group
      - Storage Account             (für Function-App-Runtime)
      - Application Insights
      - Function App                (.NET 8 isolated, Linux, Flex Consumption)
      - System-Assigned Managed Identity an der Function App
      - App Settings                (AzureAd__*, SharePoint__AllowedHosts)
      - CORS                        (erlaubt den SharePoint-Origin für SPFx-Aufrufe)

    Nutzt ausschließlich Azure CLI (kein Bicep) — bewusst imperativ gehalten,
    damit jeder Schritt einzeln nachvollziehbar ist.

.PARAMETER ResourceGroup
    Name der Resource Group. Wird angelegt, falls nicht vorhanden. Default: rg-workshop.

.PARAMETER Location
    Azure-Region (z. B. westeurope). Default: westeurope.

.PARAMETER FunctionAppName
    Global eindeutiger Name der Function App. Default: func-wsperms.

.PARAMETER BaseName
    Basisname zur Ableitung von Storage-Account- und Insights-Namen. Default: wsperms.

.PARAMETER ClientId
    Client-ID der API-App-Registrierung dieser Function (= erwartete Token-Audience).

.PARAMETER AllowedGroupId
    Objekt-ID der Sicherheitsgruppe, deren Mitglieder die Function aufrufen dürfen.

.PARAMETER SharePointHost
    SharePoint-Hostname, z. B. contoso.sharepoint.com. Dient als webUrl-Allowlist
    UND als CORS-Origin (https://<host>).

.PARAMETER TenantId
    Verzeichnis-(Mandanten-)ID. Default: aus der aktuellen az-Anmeldung.

.EXAMPLE
    ./deploy.ps1 -ClientId 1111-... -AllowedGroupId 2222-... -SharePointHost contoso.sharepoint.com
#>
[CmdletBinding()]
param(
    [Parameter()] [string] $ResourceGroup = 'rg-workshop',
    [Parameter()] [string] $Location = 'westeurope',
    [Parameter()] [string] $FunctionAppName = 'func-wsperms',
    [Parameter()] [ValidateLength(3, 12)] [string] $BaseName = 'wsperms',
    [Parameter(Mandatory = $true)] [string] $ClientId,
    [Parameter(Mandatory = $true)] [string] $AllowedGroupId,
    [Parameter(Mandatory = $true)] [string] $SharePointHost,
    [Parameter()] [string] $TenantId
)

$ErrorActionPreference = 'Stop'

function Write-Step { param($Msg) Write-Host "==> $Msg" -ForegroundColor Cyan }
function Write-Ok { param($Msg) Write-Host "    [OK] $Msg" -ForegroundColor Green }

# --- Azure-CLI-Check ---------------------------------------------------------
Write-Step "Prüfe Azure CLI ..."
$null = az account show 2>$null
if ($LASTEXITCODE -ne 0) {
    throw "Nicht angemeldet. Bitte 'az login' ausführen."
}
$account = az account show | ConvertFrom-Json
if (-not $TenantId) { $TenantId = $account.tenantId }
Write-Ok ("Subscription: {0} ({1})" -f $account.name, $account.id)
Write-Ok ("Tenant: {0}" -f $TenantId)

# --- Namen ableiten (Storage muss global eindeutig sein) ---------------------
$hashSource = "$($account.id)-$BaseName"
$shortHash = ([System.BitConverter]::ToString(
        [System.Security.Cryptography.SHA1]::Create().ComputeHash(
            [System.Text.Encoding]::UTF8.GetBytes($hashSource))
    ) -replace '-', '').Substring(0, 6).ToLower()

$storageAccount = ("st{0}{1}" -f $BaseName.ToLower(), $shortHash)
if ($storageAccount.Length -gt 24) { $storageAccount = $storageAccount.Substring(0, 24) }
$appInsights = ("appi-{0}-{1}" -f $BaseName.ToLower(), $shortHash)

Write-Host ""
Write-Host "Geplante Ressourcen:" -ForegroundColor Yellow
Write-Host "  Resource Group      : $ResourceGroup"
Write-Host "  Location            : $Location"
Write-Host "  Storage Account     : $storageAccount"
Write-Host "  Function App        : $FunctionAppName"
Write-Host "  Application Insights : $appInsights"
Write-Host ""

# --- Resource Group ----------------------------------------------------------
Write-Step "Resource Group '$ResourceGroup' ..."
az group create --name $ResourceGroup --location $Location --output none
Write-Ok "Resource Group bereit"

# --- Storage Account ---------------------------------------------------------
Write-Step "Storage Account '$storageAccount' ..."
az storage account create `
    --name $storageAccount `
    --resource-group $ResourceGroup `
    --location $Location `
    --sku Standard_LRS `
    --kind StorageV2 `
    --allow-blob-public-access false `
    --min-tls-version TLS1_2 `
    --output none
Write-Ok "Storage Account bereit"

# --- Application Insights ----------------------------------------------------
Write-Step "Application Insights '$appInsights' ..."
az extension add --name application-insights --only-show-errors --output none 2>$null
az monitor app-insights component create `
    --app $appInsights `
    --location $Location `
    --resource-group $ResourceGroup `
    --kind web `
    --application-type web `
    --output none
$aiConnString = az monitor app-insights component show `
    --app $appInsights `
    --resource-group $ResourceGroup `
    --query connectionString -o tsv
Write-Ok "Application Insights bereit"

# --- Function App (.NET 8 isolated / Flex Consumption) -----------------------
Write-Step "Function App '$FunctionAppName' (.NET 8 isolated / Flex Consumption) ..."
az functionapp create `
    --name $FunctionAppName `
    --resource-group $ResourceGroup `
    --flexconsumption-location $Location `
    --runtime dotnet-isolated `
    --runtime-version 8.0 `
    --instance-memory 2048 `
    --maximum-instance-count 40 `
    --storage-account $storageAccount `
    --output none
Write-Ok "Function App bereit"

az functionapp config appsettings set `
    --name $FunctionAppName `
    --resource-group $ResourceGroup `
    --settings "APPLICATIONINSIGHTS_CONNECTION_STRING=$aiConnString" `
    --output none
Write-Ok "Application Insights an Function App verbunden"

# --- Managed Identity --------------------------------------------------------
Write-Step "System-Assigned Managed Identity aktivieren ..."
$miInfo = az functionapp identity assign `
    --name $FunctionAppName `
    --resource-group $ResourceGroup `
    | ConvertFrom-Json
Write-Ok ("MI Principal-ID: {0}" -f $miInfo.principalId)

# --- App Settings ------------------------------------------------------------
Write-Step "App Settings setzen ..."
az functionapp config appsettings set `
    --name $FunctionAppName `
    --resource-group $ResourceGroup `
    --settings `
    "AzureAd__TenantId=$TenantId" `
    "AzureAd__ClientId=$ClientId" `
    "AzureAd__RequiredScope=access_as_user" `
    "AzureAd__AllowedGroupId=$AllowedGroupId" `
    "SharePoint__AllowedHosts=$SharePointHost" `
    --output none
Write-Ok "App Settings gesetzt"

# --- CORS --------------------------------------------------------------------
# Erlaubt dem SPFx-Web-Part (läuft auf der SharePoint-Seite) den Browser-Aufruf.
Write-Step "CORS-Origin 'https://$SharePointHost' erlauben ..."
az functionapp cors add `
    --name $FunctionAppName `
    --resource-group $ResourceGroup `
    --allowed-origins "https://$SharePointHost" `
    --output none 2>$null
Write-Ok "CORS gesetzt"

# --- Zusammenfassung ---------------------------------------------------------
Write-Host ""
Write-Host "Deployment abgeschlossen." -ForegroundColor Green
Write-Host ""
Write-Host "Wichtige Werte (bitte notieren):" -ForegroundColor Yellow
Write-Host "  Function App     : $FunctionAppName"
Write-Host "  MI Principal-ID  : $($miInfo.principalId)"
Write-Host ""
Write-Host "Nächste Schritte:" -ForegroundColor Yellow
Write-Host "  1. MI auf der Ziel-Site berechtigen (Sites.Selected + FullControl):"
Write-Host "       pwsh -NoProfile -File ./infra/grant-sites-selected.ps1 ``"
Write-Host "            -FunctionAppName $FunctionAppName -ResourceGroup $ResourceGroup ``"
Write-Host "            -SiteUrl https://$SharePointHost/sites/<sitename>"
Write-Host "  2. Function-Code deployen (aus src/ManagePermissions):"
Write-Host "       func azure functionapp publish $FunctionAppName"

$outputFile = Join-Path $PSScriptRoot ".deployment-output.json"
@{
    resourceGroup   = $ResourceGroup
    location        = $Location
    storageAccount  = $storageAccount
    functionApp     = $FunctionAppName
    appInsights     = $appInsights
    miPrincipalId   = $miInfo.principalId
    tenantId        = $TenantId
    clientId        = $ClientId
    allowedGroupId  = $AllowedGroupId
    sharePointHost  = $SharePointHost
} | ConvertTo-Json | Set-Content -Path $outputFile -Encoding UTF8
Write-Host ""
Write-Host "Werte gespeichert in: $outputFile" -ForegroundColor DarkGray
