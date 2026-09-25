// Component-test bootstrap: loaded once per browser page by
// @playwright/experimental-ct-react17 before any component mounts.
//
// Real Bootstrap CSS, loaded once here instead of via the production
// `import '.../bootstrap.min.css'` statements (which are aliased away in
// playwright-ct.config.ts, since the SPFx-vendored copy under lib/ doesn't
// exist under src/).
import 'bootstrap/dist/css/bootstrap.min.css';

// Every call a fake custom element makes is recorded here so tests can
// assert on it via `page.evaluate(() => window.__ctCalls)`. Props passed to
// `mount()` are serialized across the Playwright/browser boundary, so this
// is how test code observes what the mounted component actually did to a
// custom element, rather than trying to pass live spies in as props.
declare global {
  interface Window {
    __ctCalls: Array<{ element: string; method: string; args: unknown[] }>;
    // Set by a harness component's render body (before effects run, so
    // there's no race with the fake element's own initAsync call) to
    // simulate a template/field load failure -- e.g. an expired session or a
    // repository permissions error.
    __fieldContainerShouldFailInit?: boolean;
    // Set the same way, to make the fields service's template list take this
    // long to load -- see tests/ct/stubs/lf-ui-components-services.ts.
    __templatesLoadDelayMs?: number;
  }
}
window.__ctCalls = [];

function record(element: string, method: string, args: unknown[]): void {
  window.__ctCalls.push({ element, method, args });
}

// Production code calls `window.open(url)` to open Laserfiche's web client in
// a new tab (RepositoryViewWebPart.tsx). Letting that hit a real (fake) domain
// in a component test either times out or, worse, navigates the CT harness
// page itself away mid-test. Record the call instead of following it.
const realWindowOpen = window.open.bind(window);
window.open = (...args: Parameters<typeof window.open>): ReturnType<typeof window.open> => {
  record('window', 'open', args);
  return null;
};
void realWindowOpen;

// Fakes <lf-repository-browser> from @laserfiche/lf-ui-components, used by
// RepositoryViewWebPart.tsx via a ref + initAsync()/refreshAsync()/
// setColumnsToDisplay(), plus entrySelected/entryDblClicked events.
class FakeLfRepositoryBrowser extends HTMLElement {
  public currentFolder: unknown = undefined;

  async initAsync(...args: unknown[]): Promise<void> {
    record('lf-repository-browser', 'initAsync', args);
  }

  async refreshAsync(...args: unknown[]): Promise<void> {
    record('lf-repository-browser', 'refreshAsync', args);
  }

  setColumnsToDisplay(...args: unknown[]): void {
    record('lf-repository-browser', 'setColumnsToDisplay', args);
  }

  emitEntrySelected(nodes: unknown[]): void {
    this.dispatchEvent(new CustomEvent('entrySelected', { detail: nodes }));
  }

  emitEntryDblClicked(nodes: unknown[]): void {
    this.dispatchEvent(new CustomEvent('entryDblClicked', { detail: nodes }));
  }
}

// Fakes <lf-field-container>, used by RepositoryViewWebPart.tsx's
// ImportFileModal via initAsync()/forceValidation()/getFieldValues()/
// getTemplateValue(), plus dialogOpened/dialogClosed/fieldValuesChanged events.
class FakeLfFieldContainer extends HTMLElement {
  public forceValidationResult = true;
  public fieldValues: Record<string, { values?: Array<{ value: unknown }> }> = {};
  public templateValue: { name: string } | undefined = undefined;
  private fieldsService: { getAvailableTemplatesAsync(): Promise<unknown[]> } | undefined;

  // The real component renders its "Template" section label inside a
  // mat-panel-title, which ImportFileModal portals its templates-loading
  // spinner into -- mirror just that much of its markup.
  connectedCallback(): void {
    if (this.childElementCount > 0) {
      return;
    }
    const title = document.createElement('mat-panel-title');
    const label = document.createElement('p');
    label.textContent = 'Template';
    title.appendChild(label);
    this.appendChild(title);
  }

  async initAsync(...args: unknown[]): Promise<void> {
    record('lf-field-container', 'initAsync', args);
    if (window.__fieldContainerShouldFailInit) {
      throw new Error('Failed to load templates');
    }
    this.fieldsService = args[0] as typeof this.fieldsService;
  }

  // The real component loads the template list from its service the first
  // time the user opens the Template dropdown.
  async openTemplatesDropdown(): Promise<void> {
    record('lf-field-container', 'openTemplatesDropdown', []);
    await this.fieldsService.getAvailableTemplatesAsync();
  }

  forceValidation(): boolean {
    record('lf-field-container', 'forceValidation', []);
    return this.forceValidationResult;
  }

  getFieldValues(): Record<string, { values?: Array<{ value: unknown }> }> {
    record('lf-field-container', 'getFieldValues', []);
    return this.fieldValues;
  }

  getTemplateValue(): { name: string } | undefined {
    record('lf-field-container', 'getTemplateValue', []);
    return this.templateValue;
  }

  emitDialogOpened(): void {
    this.dispatchEvent(new CustomEvent('dialogOpened'));
  }

  emitDialogClosed(): void {
    this.dispatchEvent(new CustomEvent('dialogClosed'));
  }

  emitFieldValuesChanged(isValid: boolean): void {
    this.dispatchEvent(new CustomEvent('fieldValuesChanged', { detail: isValid }));
  }
}

// Fakes <lf-login>, used by SendToLaserficheLoginComponent.tsx,
// SaveToLaserficheDialog.tsx and repository-client.ts.
class FakeLfLogin extends HTMLElement {
  public state = 'LoggedOut';
  public authorization_credentials: { accessToken: string } | undefined = undefined;
  public account_endpoints: { webClientUrl?: string; regionalDomain?: string } = {};
  public account_id: string | undefined = undefined;

  async initLoginFlowAsync(...args: unknown[]): Promise<void> {
    record('lf-login', 'initLoginFlowAsync', args);
  }

  async refreshTokenAsync(...args: unknown[]): Promise<boolean> {
    record('lf-login', 'refreshTokenAsync', args);
    return false;
  }

  emitLoginCompleted(): void {
    this.dispatchEvent(new CustomEvent('loginCompleted'));
  }

  emitLogoutCompleted(detail?: unknown): void {
    this.dispatchEvent(new CustomEvent('logoutCompleted', { detail }));
  }
}

// Fakes <lf-tags>, used by LfTagsPicker.tsx via a `tagsService` property,
// plus the selectedTagsChanged event. Like the real component (whose
// ngAfterViewInit runs synchronously when Angular Elements attaches it), it
// reads tagsService only on its first attach, so a service assigned after
// that is never asked for tags and the list shows "No tags available".
class FakeLfTags extends HTMLElement {
  public tagsService:
    | { getTagDefinitions(): Promise<Array<{ displayName: string }>> }
    | undefined = undefined;
  private initialized = false;

  connectedCallback(): void {
    if (this.initialized) {
      return;
    }
    this.initialized = true;
    if (!this.tagsService) {
      this.renderTags([]);
      return;
    }
    void this.tagsService.getTagDefinitions().then((tags) => this.renderTags(tags));
  }

  private renderTags(tags: Array<{ displayName: string }>): void {
    const list = document.createElement('ul');
    const names = tags.length > 0 ? tags.map((tag) => tag.displayName) : ['No tags available'];
    for (const name of names) {
      const item = document.createElement('li');
      item.textContent = name;
      list.appendChild(item);
    }
    this.replaceChildren(list);
  }

  emitSelectedTagsChanged(tags: Array<{ displayName: string }>): void {
    this.dispatchEvent(
      new CustomEvent('selectedTagsChanged', { detail: tags })
    );
  }
}

if (!customElements.get('lf-repository-browser')) {
  customElements.define('lf-repository-browser', FakeLfRepositoryBrowser);
}
if (!customElements.get('lf-field-container')) {
  customElements.define('lf-field-container', FakeLfFieldContainer);
}
if (!customElements.get('lf-login')) {
  customElements.define('lf-login', FakeLfLogin);
}
if (!customElements.get('lf-tags')) {
  customElements.define('lf-tags', FakeLfTags);
}
