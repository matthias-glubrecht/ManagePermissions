import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';

import styles from './ManagePermissionsWebPart.module.scss';
import * as strings from 'ManagePermissionsWebPartStrings';

export interface IManagePermissionsWebPartProps {
  functionBaseUrl: string;
  apiResourceUri: string;
}

const PERMISSION_LEVELS: string[] = ['Read', 'Contribute', 'Edit', 'Design', 'FullControl'];

export default class ManagePermissionsWebPart extends BaseClientSideWebPart<IManagePermissionsWebPartProps> {

  public render(): void {
    const levelOptions: string = PERMISSION_LEVELS
      .map((level: string) => `<option value="${level}"${level === 'Contribute' ? ' selected' : ''}>${level}</option>`)
      .join('');

    this.domElement.innerHTML = `
    <section class="${styles.managePermissions}">
      <h2 class="${styles.title}">SharePoint-Berechtigungen verwalten</h2>

      <div class="${styles.field}">
        <label class="${styles.label}" for="mp-action">Aktion</label>
        <select id="mp-action" class="${styles.control}">
          <option value="grant" selected>grant – Berechtigung erteilen</option>
          <option value="reset">reset – Vererbung wiederherstellen</option>
        </select>
      </div>

      <div class="${styles.field}">
        <label class="${styles.label}" for="mp-webUrl">Web-URL</label>
        <input id="mp-webUrl" class="${styles.control}" type="text" placeholder="https://contoso.sharepoint.com/sites/team" />
      </div>

      <div class="${styles.field}">
        <label class="${styles.label}" for="mp-listId">Listen-ID (GUID)</label>
        <input id="mp-listId" class="${styles.control}" type="text" placeholder="00000000-0000-0000-0000-000000000000" />
      </div>

      <div class="${styles.field}">
        <label class="${styles.label}" for="mp-itemId">Element-ID</label>
        <input id="mp-itemId" class="${styles.control}" type="number" min="1" step="1" placeholder="42" />
      </div>

      <div class="${styles.field}" id="mp-row-upn">
        <label class="${styles.label}" for="mp-upn">Benutzer (UPN)</label>
        <input id="mp-upn" class="${styles.control}" type="text" placeholder="user@domain.com" />
      </div>

      <div class="${styles.field}" id="mp-row-level">
        <label class="${styles.label}" for="mp-level">Berechtigungsstufe</label>
        <select id="mp-level" class="${styles.control}">${levelOptions}</select>
      </div>

      <button id="mp-submit" class="${styles.button}" type="button">Ausführen</button>

      <div id="mp-status" class="${styles.status} ${styles.hidden}" role="status" aria-live="polite"></div>
    </section>`;

    this._bindEvents();
    this._toggleGrantFields();
  }

  private _bindEvents(): void {
    const actionEl: HTMLSelectElement | null = this.domElement.querySelector('#mp-action');
    const submitEl: HTMLButtonElement | null = this.domElement.querySelector('#mp-submit');

    if (actionEl) {
      actionEl.addEventListener('change', () => this._toggleGrantFields());
    }
    if (submitEl) {
      submitEl.addEventListener('click', () => {
        this._onSubmit().catch((error: unknown) => {
          const statusEl: HTMLElement | null = this.domElement.querySelector('#mp-status');
          if (statusEl) {
            this._setStatus(statusEl, 'error', error instanceof Error ? error.message : String(error));
          }
        });
      });
    }
  }

  private _toggleGrantFields(): void {
    const actionEl: HTMLSelectElement | null = this.domElement.querySelector('#mp-action');
    const isGrant: boolean = (actionEl ? actionEl.value : 'grant') === 'grant';
    const rowUpn: HTMLElement | null = this.domElement.querySelector('#mp-row-upn');
    const rowLevel: HTMLElement | null = this.domElement.querySelector('#mp-row-level');

    if (rowUpn) { rowUpn.classList.toggle(styles.hidden, !isGrant); }
    if (rowLevel) { rowLevel.classList.toggle(styles.hidden, !isGrant); }
  }

  private async _onSubmit(): Promise<void> {
    const statusEl: HTMLElement | null = this.domElement.querySelector('#mp-status');
    if (!statusEl) { return; }

    const functionBaseUrl: string = (this.properties.functionBaseUrl ?? '').trim().replace(/\/+$/, '');
    const apiResourceUri: string = (this.properties.apiResourceUri ?? '').trim();

    if (!functionBaseUrl || !apiResourceUri) {
      this._setStatus(statusEl, 'error', 'Bitte zuerst die Eigenschaften „Function-Basis-URL“ und „API-Ressourcen-URI“ im Eigenschaftenbereich konfigurieren.');
      return;
    }

    const action: string = (this.domElement.querySelector('#mp-action') as HTMLSelectElement).value;
    const webUrl: string = (this.domElement.querySelector('#mp-webUrl') as HTMLInputElement).value.trim();
    const listId: string = (this.domElement.querySelector('#mp-listId') as HTMLInputElement).value.trim();
    const itemIdRaw: string = (this.domElement.querySelector('#mp-itemId') as HTMLInputElement).value.trim();
    const itemId: number = Number(itemIdRaw);

    if (!webUrl || !listId || !itemIdRaw || isNaN(itemId)) {
      this._setStatus(statusEl, 'error', 'Bitte Web-URL, Listen-ID und eine gültige Element-ID angeben.');
      return;
    }

    let payload: Record<string, unknown>;
    if (action === 'grant') {
      const userPrincipalName: string = (this.domElement.querySelector('#mp-upn') as HTMLInputElement).value.trim();
      const permissionLevel: string = (this.domElement.querySelector('#mp-level') as HTMLSelectElement).value;
      if (!userPrincipalName) {
        this._setStatus(statusEl, 'error', 'Für die Aktion „grant“ ist ein Benutzer (UPN) erforderlich.');
        return;
      }
      payload = { action, webUrl, listId, itemId, userPrincipalName, permissionLevel };
    } else {
      payload = { action, webUrl, listId, itemId };
    }

    const submitEl: HTMLButtonElement | null = this.domElement.querySelector('#mp-submit');
    if (submitEl) { submitEl.disabled = true; }
    this._setStatus(statusEl, 'info', 'Anfrage wird gesendet …');

    try {
      const client: AadHttpClient = await this.context.aadHttpClientFactory.getClient(apiResourceUri);
      const response: HttpClientResponse = await client.post(
        `${functionBaseUrl}/api/ManagePermissions`,
        AadHttpClient.configurations.v1,
        {
          headers: {
            'Content-Type': 'application/json',
            'Accept': 'application/json'
          },
          body: JSON.stringify(payload)
        }
      );

      const raw: string = await response.text();
      let data: { ok?: boolean; message?: string; error?: string } = {};
      if (raw) {
        try { data = JSON.parse(raw); } catch { data = {}; }
      }

      if (response.ok) {
        this._setStatus(statusEl, 'success', `HTTP ${response.status} – ${data.message ?? 'Aktion erfolgreich ausgeführt.'}`);
      } else {
        this._setStatus(statusEl, 'error', `HTTP ${response.status} – ${data.error ?? response.statusText ?? 'Unbekannter Fehler.'}`);
      }
    } catch (callError) {
      const message: string = callError instanceof Error ? callError.message : String(callError);
      this._setStatus(statusEl, 'error', `Aufruf fehlgeschlagen: ${message}`);
    } finally {
      if (submitEl) { submitEl.disabled = false; }
    }
  }

  private _setStatus(el: HTMLElement, kind: 'info' | 'success' | 'error', message: string): void {
    const modifier: string = kind === 'success'
      ? styles.statusSuccess
      : kind === 'error'
        ? styles.statusError
        : styles.statusInfo;
    el.className = `${styles.status} ${modifier}`;
    el.textContent = message;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('functionBaseUrl', {
                  label: strings.FunctionBaseUrlFieldLabel,
                  description: strings.FunctionBaseUrlFieldDescription
                }),
                PropertyPaneTextField('apiResourceUri', {
                  label: strings.ApiResourceUriFieldLabel,
                  description: strings.ApiResourceUriFieldDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
