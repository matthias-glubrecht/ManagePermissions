# Setup – ManagePermissions

Schritt-für-Schritt-Anleitung, um die Function `ManagePermissions` einzurichten,
abzusichern und aus einem SPFx-Web-Part aufzurufen.

Reihenfolge:

1. [Entra-App-Registrierung (API)](#1-entra-app-registrierung-api)
2. [Sicherheitsgruppe der berechtigten Aufrufer](#2-sicherheitsgruppe)
3. [Azure-Infrastruktur deployen](#3-azure-infrastruktur-deployen)
4. [Managed Identity auf der Ziel-Site berechtigen](#4-managed-identity-berechtigen)
5. [Function-Code veröffentlichen](#5-function-code-veröffentlichen)
6. [SPFx-Web-Part berechtigen und konfigurieren](#6-spfx-web-part)
7. [Lokale Entwicklung](#7-lokale-entwicklung)

---

## Sicherheitsmodell auf einen Blick

| Frage | Antwort |
|---|---|
| **Wer ruft auf?** | Ein angemeldeter Benutzer über ein SPFx-Web-Part. Das Web-Part holt per `AadHttpClient` ein Entra-ID-Token für die API (`access_as_user`). |
| **Wie wird der Aufrufer geprüft?** | Die Function validiert das JWT **im Code** (Signatur, Issuer, Audience, Lifetime), prüft den Scope `access_as_user` und die Mitgliedschaft in einer **Sicherheitsgruppe**. |
| **Wer darf aufrufen?** | Nur Mitglieder der konfigurierten Gruppe (`AzureAd__AllowedGroupId`). |
| **Wie greift die Function auf SharePoint zu?** | **App-only** über die **Managed Identity** der Function App (kontrollierte Rechte-Erhöhung), begrenzt auf einzelne Sites via **Sites.Selected + FullControl**. |

---

## 1. Entra-App-Registrierung (API)

Diese App-Registrierung repräsentiert die Function als geschützte API.

1. **Entra ID → App registrations → New registration**
   - Name: `ManagePermissions API`
   - Supported account types: **Single tenant**
   - Registrieren. **Application (client) ID** notieren → das ist später `AzureAd__ClientId`.

2. **Expose an API**
   - **Application ID URI**: `api://<client-id>` setzen (Standardvorschlag übernehmen).
   - **Add a scope**:
     - Scope name: `access_as_user`
     - Who can consent: **Admins and users**
     - Consent-Anzeigetexte ausfüllen, State **Enabled**.

3. **Token-Version auf 2 setzen** (damit SPFx v2-Tokens erhält):
   - **Manifest** öffnen → `requestedAccessTokenVersion` auf `2` setzen → speichern.

4. **Gruppen-Claim aktivieren** (für die Autorisierung):
   - **Token configuration → Add groups claim**
   - **Groups assigned to the application** auswählen (vermeidet „Group-Overage").
   - Für **Access**-Tokens aktivieren → speichern.

> Hinweis: „Groups assigned to the application" sorgt dafür, dass nur Gruppen im Token
> erscheinen, die der zugehörigen Enterprise-App zugewiesen sind (Schritt 2 unten).

---

## 2. Sicherheitsgruppe

1. **Entra ID → Groups → New group**
   - Typ: **Security**
   - Name: z. B. `ManagePermissions-Caller`
   - Mitglieder (die berechtigten Benutzer) hinzufügen.
   - **Object Id** der Gruppe notieren → das ist `AzureAd__AllowedGroupId`.

2. Damit der Gruppen-Claim im Token erscheint (wegen „Groups assigned to the application"):
   - **Entra ID → Enterprise applications → `ManagePermissions API` → Users and groups**
   - Die Sicherheitsgruppe der App zuweisen.

---

## 3. Azure-Infrastruktur deployen

Voraussetzung: `az login`.

```powershell
./infra/deploy.ps1 `
    -ClientId       <AzureAd__ClientId> `
    -AllowedGroupId <AzureAd__AllowedGroupId> `
    -SharePointHost <tenant>.sharepoint.com
```

Das Skript legt Resource Group, Storage, Application Insights und die Function App
(.NET 8 isolated, Flex Consumption) an, aktiviert die Managed Identity, setzt die
App Settings (`AzureAd__*`, `SharePoint__AllowedHosts`) und erlaubt den SharePoint-Origin
per CORS.

Optionale Parameter: `-ResourceGroup` (Default `rg-workshop`), `-Location`
(Default `westeurope`), `-FunctionAppName` (Default `func-wsperms`), `-BaseName`,
`-TenantId`.

---

## 4. Managed Identity berechtigen

In einer **frischen** PowerShell-Session (Assembly-Konflikt mit PnP.PowerShell vermeiden):

```powershell
pwsh -NoProfile -File ./infra/grant-sites-selected.ps1 `
    -FunctionAppName func-wsperms `
    -ResourceGroup   rg-workshop `
    -SiteUrl         https://<tenant>.sharepoint.com/sites/<sitename>
```

Das Skript weist der MI die App-Rolle **Sites.Selected** auf **Graph UND SharePoint Online**
zu und legt für die Ziel-Site die Rolle **fullcontrol** an. `fullcontrol` ist für das
Verwalten von Listenelement-Berechtigungen erforderlich (`write` reicht laut Microsoft-Doku
nicht).

> Nach Rollenänderungen kann der MI-Token-Cache bis zu ~24 h alt sein. Bei 401-Fehlern
> hilft `az functionapp stop` + `az functionapp start` oder Abwarten.

---

## 5. Function-Code veröffentlichen

```powershell
Set-Location src/ManagePermissions
func azure functionapp publish func-wsperms
```

Endpunkt danach: `https://<func-wsperms-host>/api/ManagePermissions`

---

## 6. SPFx-Web-Part

Das Beispiel-Web-Part liegt unter [`spfx-sample/`](../spfx-sample/). Es fordert den
API-Scope per `webApiPermissionRequests` an:

```json
"webApiPermissionRequests": [
  { "resource": "ManagePermissions API", "scope": "access_as_user" }
]
```

Nach `gulp bundle --ship` + `gulp package-solution --ship` und Upload in den App-Katalog:

1. **SharePoint Admin Center → Advanced → API access**
2. Die ausstehende Anfrage **`ManagePermissions API / access_as_user`** **genehmigen**.

Das Web-Part holt das Token via `AadHttpClientFactory` für `api://<client-id>` und ruft
`POST /api/ManagePermissions` auf. Die Function-URL wird in den Web-Part-Eigenschaften
hinterlegt.

---

## 7. Lokale Entwicklung

1. `src/ManagePermissions/local.settings.json` aus `local.settings.json.sample` erstellen
   und Werte einsetzen (`AzureAd__*`, `SharePoint__AllowedHosts`).
2. `az login` mit einem Konto, das auf der Ziel-Site Zugriff hat — lokal nutzt
   `DefaultAzureCredential` die Azure-CLI-Anmeldung statt der Managed Identity.
3. Storage-Emulator starten (`azurite`) und `func start` ausführen.
4. Test-Token holen und aufrufen:

```powershell
$token = az account get-access-token --resource api://<client-id> --query accessToken -o tsv
$body = @{
    action            = 'grant'
    webUrl            = 'https://<tenant>.sharepoint.com/sites/<sitename>'
    listId            = '<list-guid>'
    itemId            = 1
    userPrincipalName = 'user@domain.com'
    permissionLevel   = 'Contribute'
} | ConvertTo-Json

Invoke-RestMethod -Method Post -Uri 'http://localhost:7071/api/ManagePermissions' `
    -Headers @{ Authorization = "Bearer $token" } `
    -ContentType 'application/json' -Body $body
```

> Das `az`-CLI-Token enthält den Scope `access_as_user` nur, wenn die Tenant-Richtlinien
> das zulassen; für realistische Tests den Aufruf über das SPFx-Web-Part durchführen.
