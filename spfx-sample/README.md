# manage-permissions-sample

Minimales SPFx-Beispiel-Web-Part (No-Framework, reines DOM), das eine bestehende Azure Function aufruft.
Die Function (`POST /api/ManagePermissions`) setzt bzw. entfernt Berechtigungen auf SharePoint-Listenelementen
und ist durch eine Entra-ID-App-Registrierung mit dem delegierten Scope `access_as_user` geschützt.
Das Token wird über `AadHttpClient` / `AadHttpClientFactory` bezogen.

## Verwendete SharePoint-Framework-Version

![version](https://img.shields.io/badge/version-1.21.1-green.svg)

## Voraussetzungen

- Node.js **22.17.0** (LTS), npm 10.x
- Globale Tools: `yo`, `gulp-cli`, `@microsoft/generator-sharepoint@1.21.1`
- Eine bereitgestellte **ManagePermissions**-Azure-Function (HTTP-Endpunkt `POST /api/ManagePermissions`)
- Eine Entra-ID-App-Registrierung für die API mit:
  - Anzeigename **`ManagePermissions API`**
  - exponiertem delegiertem Scope **`access_as_user`**
  - einem Anwendungs-ID-URI (z. B. `api://<client-id>`)
- Ein SharePoint-Online-Tenant mit App-Katalog und ein Konto mit Administratorrechten zum Freigeben der API-Berechtigung

## Funktionsweise

Das Web-Part rendert ein kleines Formular (Beschriftungen auf Deutsch):

- **Aktion** – `grant` oder `reset`
- **Web-URL** – z. B. `https://contoso.sharepoint.com/sites/team`
- **Listen-ID** – GUID der Liste
- **Element-ID** – numerische ID des Listenelements
- **Benutzer (UPN)** – nur bei `grant`
- **Berechtigungsstufe** – `Read`, `Contribute`, `Edit`, `Design`, `FullControl` (nur bei `grant`)
- Schaltfläche **Ausführen**

Beim Absenden wird ein JSON-Body gebaut und per `AadHttpClient` an die Function gesendet:

```http
POST {functionBaseUrl}/api/ManagePermissions
Authorization: Bearer <Token für die API-Ressource>
Content-Type: application/json

{ "action": "grant", "webUrl": "...", "listId": "<guid>", "itemId": 42,
  "userPrincipalName": "user@domain.com", "permissionLevel": "Contribute" }
```

Antwort: HTTP 200 `{ "ok": true, "message": "..." }` bei Erfolg, sonst ein nicht-2xx-Status mit
`{ "error": "..." }`. Im Statusbereich werden HTTP-Status und `message` bzw. `error` angezeigt.

## Web-Part-Eigenschaften

Im Eigenschaftenbereich des Web-Parts (Bearbeiten-Modus) konfigurieren:

| Eigenschaft        | Beispiel                                  | Bedeutung                                                              |
| ------------------ | ----------------------------------------- | --------------------------------------------------------------------- |
| `functionBaseUrl`  | `https://func-wsperms.azurewebsites.net`  | Basis-URL der Function-App                                            |
| `apiResourceUri`   | `api://<client-id>`                        | Ressourcen-/App-ID-URI der API – wird an `getClient(...)` übergeben    |

Sind die Eigenschaften nicht gesetzt, erscheint im Statusbereich ein freundlicher deutscher Hinweis.

## Lokale Entwicklung

```powershell
npm install
gulp serve
```

- Lokale Workbench: `https://localhost:4321/workbench`
- SPO-Workbench: `https://<tenant>.sharepoint.com/_layouts/15/workbench.aspx`

> Hinweis: `AadHttpClient` liefert in der lokalen Workbench ggf. kein gültiges Token für die API.
> Für einen echten End-to-End-Test das Paket in den App-Katalog hochladen und auf einer SharePoint-Seite testen.

## Paketieren und Bereitstellen

```powershell
gulp bundle --ship
gulp package-solution --ship
```

Das erzeugte Paket liegt unter `sharepoint/solution/manage-permissions-sample.sppkg`.

1. `.sppkg` in den **App-Katalog** des Tenants (oder einer Site Collection) hochladen und bereitstellen.
2. **API-Berechtigung freigeben**: SharePoint Admin Center → **Erweitert** → **API-Zugriff**
   (SharePoint Admin Center → *Advanced* → *API access*) und die ausstehende Anfrage
   **`ManagePermissions API` / `access_as_user`** genehmigen.
3. Web-Part auf einer Seite hinzufügen und im Eigenschaftenbereich `functionBaseUrl` sowie `apiResourceUri` setzen.

Die API-Berechtigungsanfrage ist in [config/package-solution.json](config/package-solution.json) hinterlegt:

```json
"webApiPermissionRequests": [
  { "resource": "ManagePermissions API", "scope": "access_as_user" }
]
```

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

## Weiterführende Links

- [SharePoint Framework](https://aka.ms/spfx)
- [Connect to Azure AD-secured APIs (AadHttpClient)](https://learn.microsoft.com/sharepoint/dev/spfx/use-aadhttpclient)
- [API access page (SharePoint Admin Center)](https://learn.microsoft.com/sharepoint/dev/spfx/use-aadhttpclient#configure-the-api-permissions-in-the-sharepoint-admin-center)
