import { override } from '@microsoft/decorators';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IPagesClearProps {
  enabled?: boolean;

  // Defaults ajustados a tu pedido:
  //   - Solo propietarios tienen bypass (true)
  //   - NO bypass para admins/members/visitors (false)
  bypassForAdmins?: boolean;    // default: false
  bypassOwnerGroup?: boolean;   // default: true
  bypassMemberGroup?: boolean;  // default: false
  bypassVisitorGroup?: boolean; // default: false
}

export default class PagesClearApplicationCustomizer
  extends BaseApplicationCustomizer<IPagesClearProps> {

  private styleEl: HTMLStyleElement | null = null;
  private readonly scopeClass = 'hide-chrome';
  private _navHandler: (() => void) | null = null;
  private _observer: MutationObserver | null = null;
  private _bypass: boolean = false;

  private readonly cssText = `
html.hide-chrome #SuiteNavWrapper,
html.hide-chrome #SuiteNavPlaceHolder,
html.hide-chrome .od-SuiteNav,
html.hide-chrome #O365_NavHeader,
html.hide-chrome [data-automationid="SiteHeader"],
html.hide-chrome #spSiteHeader,
html.hide-chrome #spCommandBar,
html.hide-chrome [data-automationid="SiteLeftNav"],
html.hide-chrome .sp-appBar { display:none!important;height:0!important;min-height:0!important;margin:0!important;padding:0!important;overflow:hidden!important; }
html.hide-chrome #spPageCanvasContent,
html.hide-chrome .SPCanvas,
html.hide-chrome .CanvasZone,
html.hide-chrome .Canvas,
html.hide-chrome div[role="main"] { margin:0!important;padding:0!important;max-width:none!important;width:100%!important; }
html.hide-chrome [data-automation-id="socialBar"],
html.hide-chrome [data-automation-id="CommentsWrapper"],
html.hide-chrome [data-automation-id="pageFooter"] { display:none!important;height:0!important;min-height:0!important;margin:0!important;padding:0!important;overflow:hidden!important; }
`.trim();

  @override
  public async onInit(): Promise<void> {
    this._bypass = await this._computeBypass(); // Solo owners por defecto

    this._navHandler = () => this.applyOrCleanup();
    this.context.application.navigatedEvent.add(this, this._navHandler);
    this.applyOrCleanup();

    this.ensureObserver();
    return Promise.resolve();
  }

  /** Bypass SOLO para propietarios del sitio (por defecto) */
  private async _computeBypass(): Promise<boolean> {
    const props = this.properties || {};
    const bypassForAdmins   = props.bypassForAdmins   === true;  // default false
    const bypassOwnerGroup  = props.bypassOwnerGroup  !== false; // default true
    const bypassMemberGroup = props.bypassMemberGroup === true;  // default false
    const bypassVisitorGroup= props.bypassVisitorGroup=== true;  // default false

    // (Opcional) Admin del sitio: por defecto NO bypass
    if (bypassForAdmins) {
      const isSiteAdmin: boolean = (this.context.pageContext as any)?.legacyPageContext?.isSiteAdmin === true;
      if (isSiteAdmin) return true;
    }

    try {
      // Usuario actual
      const me = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/currentUser?$expand=groups`,
        SPHttpClient.configurations.v1
      ).then((r: SPHttpClientResponse) => r.json());

      const myId: number = me?.Id;
      if (!myId) return false;

      // Helper: ¿está en el grupo asociado X?
      const isInAssociatedGroup = async (endpoint: string): Promise<boolean> => {
        const group = await this.context.spHttpClient.get(
          `${this.context.pageContext.web.absoluteUrl}/_api/web/${endpoint}?$select=Id,Users/Id&$expand=Users`,
          SPHttpClient.configurations.v1
        ).then((r: SPHttpClientResponse) => r.json());
        const users: Array<{ Id: number }> = group?.Users || [];
        return users.some(u => u.Id === myId);
      };

      if (bypassOwnerGroup  && await isInAssociatedGroup('associatedownergroup'))   return true;
      if (bypassMemberGroup && await isInAssociatedGroup('associatedmembergroup'))  return true;
      if (bypassVisitorGroup&& await isInAssociatedGroup('associatedvisitorgroup')) return true;
    } catch {
      // Si algo falla, no hacemos bypass (se aplicará ocultar)
    }
    return false;
  }

  private ensureObserver(): void {
    if (this._observer) return;
    this._observer = new MutationObserver(() => {
      if (document.documentElement.classList.contains(this.scopeClass)) {
        this.injectCss();
      }
    });
    this._observer.observe(document.body, { childList: true, subtree: true });
  }

  private applyOrCleanup(): void {
    const enabled = !(this.properties && this.properties.enabled === false);
    if (!enabled || this._bypass) { this.cleanup(); return; }

    const force = new URL(location.href).searchParams.get('chromeless');
    if (force === '1') { this.enableAndInject(); return; }
    if (force === '0') { this.cleanup(); return; }

    const p = (location.pathname || '').toLowerCase();
    const starts = (s: string, sub: string) => s.lastIndexOf(sub, 0) === 0;
    const has    = (s: string, sub: string) => s.indexOf(sub) !== -1;
    const ends   = (s: string, sub: string) => s.length >= sub.length && s.indexOf(sub, s.length - sub.length) !== -1;

    // Solo en Home/SitePages; no en listas/_layouts/etc
    const isLayouts   = starts(p, '/_layouts/15/');
    const isListOrLib = has(p, '/lists/') || has(p, '/forms/') || has(p, '/siteassets/') ||
                        has(p, '/style library/') || has(p, '/shared%20documents') ||
                        has(p, '/documents') || has(p, '/documentos') || has(p, '/_catalogs/');
    const isSitePage  = has(p, '/sitepages/');
    const isHome      = ends(p, '/home.aspx') || ends(p, '/default.aspx');

    const shouldApply = (isSitePage || isHome) && !isLayouts && !isListOrLib;
    if (shouldApply) this.enableAndInject(); else this.cleanup();
  }

  private enableAndInject(): void {
    if (!document.documentElement.classList.contains(this.scopeClass)) {
      document.documentElement.classList.add(this.scopeClass);
    }
    this.injectCss();
    setTimeout(() => this.injectCss(), 0); // por si el shell monta tarde
  }

  private injectCss(): void {
    let s = this.styleEl || document.querySelector('style[data-pagesclear="true"]') as HTMLStyleElement | null;
    if (!s) {
      s = document.createElement('style');
      s.type = 'text/css';
      s.setAttribute('data-pagesclear', 'true');
      document.head.appendChild(s);
      this.styleEl = s;
    }
    if (s.textContent !== this.cssText) s.textContent = this.cssText;
  }

  private cleanup(): void {
    document.documentElement.classList.remove(this.scopeClass);
    const s = document.querySelector('style[data-pagesclear="true"]');
    if (s && s.parentNode) s.parentNode.removeChild(s);
    this.styleEl = null;
  }

  @override
  public onDispose(): void {
    if (this._observer) { this._observer.disconnect(); this._observer = null; }
    if (this._navHandler) {
      this.context.application.navigatedEvent.remove(this, this._navHandler);
      this._navHandler = null;
    }
    this.cleanup();
    super.onDispose();
  }
}
