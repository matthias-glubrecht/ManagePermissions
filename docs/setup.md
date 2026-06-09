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
     - State: **Enabled**
     - Consent-Anzeigetexte (zum Kopieren):

       | Feld | Wert |
       |---|---|
       | Admin consent display name | ManagePermissions im Namen des Benutzers aufrufen |
       | Admin consent description | Ermöglicht der Anwendung, die ManagePermissions-API im Namen des angemeldeten Benutzers aufzurufen, um Berechtigungen auf SharePoint-Listenelementen zu verwalten. |
       | User consent display name | ManagePermissions in Ihrem Namen aufrufen |
       | User consent description | Ermöglicht der Anwendung, die ManagePermissions-API in Ihrem Namen aufzurufen, um Berechtigungen auf SharePoint-Listenelementen zu verwalten. |

3. **Token-Version auf 2 setzen** (damit die Function die Tokens akzeptiert):

   Die Function validiert eingehende Tokens gegen den **v2.0**-Issuer
   (`https://login.microsoftonline.com/<tenant>/v2.0`). Eine `api://`-App stellt aber
   standardmäßig **v1.0**-Tokens aus (Issuer `https://sts.windows.net/<tenant>/`). Ohne
   diese Umstellung passt der Issuer nicht und jeder Aufruf scheitert mit **401**.

   - In der App-Registrierung links **Manage → Manifest** öffnen.
   - Das Feld `requestedAccessTokenVersion` von `null` auf `2` setzen. Je nach Editor-Format
     liegt es an unterschiedlicher Stelle:
     - **Microsoft Graph App Manifest** (neuer Editor): verschachtelt unter
       `"api": { "requestedAccessTokenVersion": 2 }`.
     - **AAD Graph App Manifest** (älteres/„deprecated"-Format): als Top-Level-Feld
       `"requestedAccessTokenVersion": 2`.
   - **Save** klicken.

   > Kontrolle: Im ausgestellten Token (z. B. über <https://jwt.ms> sichtbar gemacht) muss
   > der Claim `ver` den Wert `2.0` haben.

   > Nachlesen:
   > - [v1.0- vs. v2.0-Tokens und Issuer-Validierung](https://learn.microsoft.com/en-us/entra/identity-platform/access-tokens#v10-and-v20-tokens)
   >   – warum der Issuer je nach Token-Version unterschiedlich ist.
   > - [`requestedAccessTokenVersion` im App-Manifest](https://learn.microsoft.com/en-us/entra/identity-platform/reference-app-manifest#requestedaccesstokenversion-attribute)
   >   – die zulässigen Werte (`null`/`1` = v1.0, `2` = v2.0).

4. **Gruppen-Claim aktivieren** (für die Autorisierung):

   Ein Entra-Token enthält **standardmäßig nicht**, in welchen Gruppen der Benutzer ist.
   Die Function trägt ihre Zugangsentscheidung aber genau über diesen `groups`-Claim aus
   (Vergleich gegen `AzureAd__AllowedGroupId`). Ohne aktivierten Claim fehlt `groups` im
   Token → die Prüfung scheitert für **jeden** mit **403**.

   - In der App-Registrierung links **Manage → Token configuration** öffnen.
   - **+ Add groups claim** wählen.
   - **Groups assigned to the application** ankreuzen (vermeidet „Group-Overage", s. u.).
   - Token-Typ **Access** aktivieren (ID/SAML sind für diese API irrelevant) → **Save**.

   > **Warum „Groups assigned to the application"?** Bei **All groups** packt Entra *alle*
   > Gruppen des Benutzers ins Token. Ist er in zu vielen Mitglied (> 200 im JWT), lässt
   > Entra die Liste weg und setzt nur einen Verweis (`_claim_names`) – „Group-Overage", der
   > `groups`-Claim fehlt dann. „Groups assigned to the application" nimmt nur die der
   > Enterprise-App **zugewiesenen** Gruppen (Schritt 2.2) auf – Overage tritt praktisch nie auf.
   >
   > **Voraussetzungen/Fallen:** Diese Option berücksichtigt nur **direkte** Mitglieder der
   > zugewiesenen Gruppe (keine verschachtelten Gruppen) und erfordert mindestens eine
   > **Entra ID P1**-Lizenz (in einem Free-Tenant lassen sich keine Gruppen einer App zuweisen).
   >
   > Kontrolle: Im ausgestellten Token (z. B. über <https://jwt.ms>) muss der `groups`-Claim
   > die Object-ID der Sicherheitsgruppe enthalten.

   > Nachlesen:
   > - [Gruppen-Claims & Group-Overage konfigurieren](https://learn.microsoft.com/en-us/security/zero-trust/develop/configure-tokens-group-claims-app-roles#group-overages)
   >   – warum „Groups assigned to the application" das Overage-Problem vermeidet.
   > - [Configure groups optional claims](https://learn.microsoft.com/en-us/entra/identity-platform/optional-claims#configure-groups-optional-claims)
   >   – die genauen Portal-Schritte und Optionen.

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

Die beiden GUIDs stammen aus den vorherigen Schritten:

| Parameter | Quelle |
|---|---|
| `-ClientId` | **Application (client) ID** aus [Schritt 1](#1-entra-app-registrierung-api) (App-Registrierung `ManagePermissions API`). |
| `-AllowedGroupId` | **Object Id** aus [Schritt 2](#2-sicherheitsgruppe) (Sicherheitsgruppe `ManagePermissions-Caller`). |

Falls nicht notiert, per CLI nachschlagen (`-o tsv` liefert die nackte GUID):

```powershell
# ClientId der App-Registrierung (appId, NICHT die Object Id der App!)
az ad app list --display-name "ManagePermissions API" --query "[0].appId" -o tsv

# Object Id der Sicherheitsgruppe
az ad group show --group "ManagePermissions-Caller" --query id -o tsv
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
