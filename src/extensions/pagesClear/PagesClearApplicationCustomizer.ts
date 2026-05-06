import { override } from '@microsoft/decorators';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IPagesClearProps {
  enabled?: boolean;

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

  // CSS corregido (SIN comas) + selectores más robustos
private readonly cssText = `
/* Oculta el chrome de SharePoint */
html.hide-chrome #SuiteNavWrapper,
html.hide-chrome #SuiteNavPlaceHolder,
html.hide-chrome .od-SuiteNav,
html.hide-chrome #O365_NavHeader,
html.hide-chrome [data-automationid="SiteHeader"],
html.hide-chrome #spSiteHeader,
html.hide-chrome #spCommandBar,
html.hide-chrome [data-automationid="SiteLeftNav"],
html.hide-chrome .sp-appBar {
  display: none !important;
  height: 0 !important;
  min-height: 0 !important;
  margin: 0 !important;
  padding: 0 !important;
  overflow: hidden !important;
}

/* Oculta social/comments/footer */
html.hide-chrome [data-automation-id="socialBar"],
html.hide-chrome [data-automation-id="CommentsWrapper"],
html.hide-chrome #CommentsWrapper,
html.hide-chrome [data-automation-id="pageFooter"] {
  display: none !important;
  height: 0 !important;
  min-height: 0 !important;
  margin: 0 !important;
  padding: 0 !important;
  overflow: hidden !important;
}

/* Ajuste del layout para que el contenido use todo el ancho */
html.hide-chrome div[role="main"] {
  margin: 0 !important;
  padding: 0 !important;
  max-width: none !important;
  width: 100% !important;
}

/* IMPORTANTÍSIMO: NO ocultar el canvas, solo quitar paddings/márgenes */
html.hide-chrome #spPageCanvasContent,
html.hide-chrome .SPCanvas,
html.hide-chrome .CanvasZone,
html.hide-chrome .Canvas {
  display: block !important;
  margin: 0 !important;
  padding: 0 !important;
  max-width: none !important;
  width: 100% !important;
}
`.trim();


  @override
  public async onInit(): Promise<void> {
    this._bypass = await this._computeBypass(); // por defecto: solo Owners hacen bypass

    this._navHandler = () => this.applyOrCleanup();
    this.context.application.navigatedEvent.add(this, this._navHandler);

    this.ensureObserver();
    this.applyOrCleanup();       // 1ra aplicación
    setTimeout(() => this.applyOrCleanup(), 0); // 2da pasada (shell tarda en montar)

    return Promise.resolve();
  }

  /** Bypass SOLO para propietarios (por defecto) */
  private async _computeBypass(): Promise<boolean> {
    const props = this.properties || {};
    const bypassForAdmins   = props.bypassForAdmins   === true;
    const bypassOwnerGroup  = props.bypassOwnerGroup  !== false; // default true
    const bypassMemberGroup = props.bypassMemberGroup === true;
    const bypassVisitorGroup= props.bypassVisitorGroup=== true;

    // Admin del sitio (opcional)
    if (bypassForAdmins) {
      const isSiteAdmin: boolean = (this.context.pageContext as any)?.legacyPageContext?.isSiteAdmin === true;
      if (isSiteAdmin) return true;
    }

    try {
      const me = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/currentUser?$select=Id`,
        SPHttpClient.configurations.v1
      ).then((r: SPHttpClientResponse) => r.json());

      const myId: number = me?.Id;
      if (!myId) return false;

      // Mucho más liviano: filtra por Id del usuario (no trae todos los users del grupo)
      const isInAssociatedGroup = async (endpoint: string): Promise<boolean> => {
        const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/${endpoint}/users?$select=Id&$filter=Id eq ${myId}`;
        const r = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        const j = await r.json();
        return Array.isArray(j?.value) && j.value.length > 0;
      };

      if (bypassOwnerGroup  && await isInAssociatedGroup('associatedownergroup'))   return true;
      if (bypassMemberGroup && await isInAssociatedGroup('associatedmembergroup'))  return true;
      if (bypassVisitorGroup&& await isInAssociatedGroup('associatedvisitorgroup')) return true;
    } catch {
      // si falla, NO bypass => se oculta
    }

    return false;
  }

  private ensureObserver(): void {
    if (this._observer) return;

    this._observer = new MutationObserver(() => {
      // Si ya está activado, reinyecta por si SharePoint “re-monta” el header
      if (document.documentElement.classList.contains(this.scopeClass)) {
        this.injectCss();
      }
    });

    const target = document.body || document.documentElement;
    this._observer.observe(target, { childList: true, subtree: true });
  }

  private applyOrCleanup(): void {
    const enabled = !(this.properties && this.properties.enabled === false);
    if (!enabled || this._bypass) { this.cleanup(); return; }

    // Forzar por querystring
    const force = new URL(location.href).searchParams.get('chromeless');
    if (force === '1') { this.enableAndInject(); return; }
    if (force === '0') { this.cleanup(); return; }

    const p = (location.pathname || '').toLowerCase();

    const starts = (s: string, sub: string) => s.lastIndexOf(sub, 0) === 0;
    const has    = (s: string, sub: string) => s.indexOf(sub) !== -1;
    const ends   = (s: string, sub: string) => s.length >= sub.length && s.indexOf(sub, s.length - sub.length) !== -1;

    // Evitar layouts y listas/bibliotecas
    const isLayouts   = starts(p, '/_layouts/15/');
    const isListOrLib =
      has(p, '/lists/') || has(p, '/forms/') || has(p, '/_catalogs/') ||
      has(p, '/siteassets/') || has(p, '/style library/') ||
      has(p, '/shared%20documents') || has(p, '/documents') || has(p, '/documentos');

    // Aplicar en páginas modernas y también en raíz (muy común que “Home” sea la raíz)
const isSitePage = has(p, '/sitepages/');
const isHome = ends(p, '/home.aspx') || ends(p, '/default.aspx');

// Detecta correctamente la raíz del sitio aunque no termine en "/"
// Ejemplo: /sites/grp_intra_ec o /sites/grp_intra_ec/
const normalizePath = (value: string): string =>
  (value || '/').toLowerCase().replace(/\/+$/, '') || '/';

const currentPath = normalizePath(p);
const webPath = normalizePath(new URL(this.context.pageContext.web.absoluteUrl).pathname);

const isRoot = currentPath === webPath;



    const shouldApply = (isSitePage || isHome || isRoot) && !isLayouts && !isListOrLib;

    if (shouldApply) this.enableAndInject();
    else this.cleanup();
  }

  private enableAndInject(): void {
    document.documentElement.classList.add(this.scopeClass);
    this.injectCss();
    setTimeout(() => this.injectCss(), 50);
    setTimeout(() => this.injectCss(), 250);
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
    if (s?.parentNode) s.parentNode.removeChild(s);
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
