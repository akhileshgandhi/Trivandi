import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import SearchModal from './components/SearchModal';
import '../../styles/PremiumSearch.module.scss';
import { checkPermissions } from '../Permission/PermissionService';

import { CacheService } from './services/CacheService';

export interface IGlobalSearchCustomizerApplicationCustomizerProperties {
  testMessage: string;
}

const ALLOWED_SITES = [
  'TrivandiHub',
  'PeopleHub',
  'CompanyHub',
  'BrandingMarketing',
  'Projects',
  'TrivandiLondon',
  'TDMCC',
  'TrivandiUSA',
  'TrivandiAustralia',
  'TrivandiKSA',
  'OperationsHub',
  'Freudiger',
  'moreYeahsdepartmentsDMS',
  'PembePortal',
  '/sites/'
];

let g_context: any = null;

export default class GlobalSearchCustomizerApplicationCustomizer
  extends BaseApplicationCustomizer<IGlobalSearchCustomizerApplicationCustomizerProperties> {

  private _modalContainer: HTMLElement | null = null;
  private _observer: MutationObserver | null = null;
  private _attached: Set<Element> = new Set();
  private _isModalOpen: boolean = false;

  public onInit(): Promise<void> {
    console.log('[GlobalSearchCustomizer] App Customizer initialized on:', window.location.href);
    g_context = this.context;
    this._createModalContainer();
    this._interceptSearchBar();

    // Trigger permission check on app start
    checkPermissions(this.context || g_context).catch(err => {
      console.error("[GlobalSearchCustomizer] Error checking permissions on load:", err);
    });

    const wasOpen = sessionStorage.getItem('trivandi_search_modal_open') === 'true';
    if (wasOpen) {
      const currentUrl = window.location.href.toLowerCase();
      const isAllowedSite = ALLOWED_SITES.some(site => currentUrl.includes(site.toLowerCase()));
      const isLibraryOrDocPage = currentUrl.includes('/forms/') || currentUrl.includes('/allitems.aspx');
      if (isAllowedSite && !isLibraryOrDocPage) {
        console.log('[GlobalSearchCustomizer] Re-opening search modal from previous session state.');
        this._openModal();
      }
    }

    CacheService.startCleanup();
    return Promise.resolve();
  }

  private _createModalContainer(): void {
    if (document.getElementById('trivandi-search-root')) {
      this._modalContainer = document.getElementById('trivandi-search-root');
      return;
    }
    this._modalContainer = document.createElement('div');
    this._modalContainer.id = 'trivandi-search-root';
    document.body.appendChild(this._modalContainer);
    console.log('[GlobalSearchCustomizer] Created modal container element #trivandi-search-root');
  }

  private _interceptSearchBar(): void {
    const isSearchBoxTarget = (target: HTMLElement | null): boolean => {
      if (!target) return false;

      // Target must be strictly inside the header search box container (NOT header/nav containers)
      const searchContainer = target.closest(
        '#O365_SearchBoxContainer_container, #sbcId, #O365_SearchBoxField, form[role="search"], .ms-suiteux-searchbox, .ms-suiteux-search-box'
      );

      if (!searchContainer) return false;

      return (
        target.tagName === 'INPUT' ||
        target.getAttribute('role') === 'combobox' ||
        !!target.closest('input, form[role="search"], #sbcId, #O365_SearchBoxField')
      );
    };

    // 1. Direct document level capture listener for maximum responsiveness
    const handleGlobalIntercept = (e: Event): void => {
      const target = e.target as HTMLElement;
      if (!isSearchBoxTarget(target)) return;

      const currentUrl = window.location.href.toLowerCase();
      const isAllowedSite = ALLOWED_SITES.some(site => currentUrl.includes(site.toLowerCase()));
      const isLibraryOrDocPage = currentUrl.includes('/forms/') || currentUrl.includes('/allitems.aspx');

      console.log(`[GlobalSearchCustomizer] Search input event '${e.type}' intercepted. Target:`, target);

      if (!isAllowedSite) return;

      if (isLibraryOrDocPage && (e.type === 'focus' || e.type === 'focusin')) {
        return;
      }

      e.preventDefault();
      e.stopPropagation();
      e.stopImmediatePropagation();

      if (typeof target.blur === 'function') {
        target.blur();
      }

      if (!this._isModalOpen) {
        console.log('[GlobalSearchCustomizer] Opening Search Modal from search input click.');
        this._openModal();
      }
    };

    ['pointerdown', 'mousedown', 'click', 'focusin'].forEach(evtType => {
      document.addEventListener(evtType, handleGlobalIntercept, true);
    });

    // 2. Element level listeners + MutationObserver
    this._attachListeners();
    this._observer = new MutationObserver(() => {
      this._attachListeners();
    });
    this._observer.observe(document.body, { childList: true, subtree: true });
  }

  private _attachListeners(): void {
    const SELECTORS = [
      '#O365_SearchBoxContainer_container input',
      '#sbcId input',
      '#O365_SearchBoxField input',
      '.ms-suiteux-searchbox input',
      '.ms-suiteux-search-box input',
      'form[role="search"] input'
    ];

    const isSearchBoxTarget = (el: Element): boolean => {
      return !!el.closest('#O365_SearchBoxContainer_container, #sbcId, #O365_SearchBoxField, form[role="search"], .ms-suiteux-searchbox, .ms-suiteux-search-box');
    };

    SELECTORS.forEach((sel) => {
      try {
        document.querySelectorAll(sel).forEach((el) => {
          if (!isSearchBoxTarget(el)) return;
          if (this._attached.has(el)) return;
          this._attached.add(el);

          console.log(`[GlobalSearchCustomizer] Attached listener specifically to search input: ${sel}`);

          const handler = (e: Event): void => {
            const currentUrl = window.location.href.toLowerCase();
            const isAllowedSite = ALLOWED_SITES.some(site => currentUrl.includes(site.toLowerCase()));
            const isLibraryOrDocPage = currentUrl.includes('/forms/') || currentUrl.includes('/allitems.aspx');
            
            console.log(`[GlobalSearchCustomizer] Direct element event '${e.type}' triggered on search input target:`, e.target);

            if (!isAllowedSite) {
              return;
            }

            if (isLibraryOrDocPage && (e.type === 'focus' || e.type === 'focusin')) {
              return;
            }

            e.preventDefault();
            e.stopPropagation();
            e.stopImmediatePropagation();

            if (e.target && typeof (e.target as HTMLElement).blur === 'function') {
              (e.target as HTMLElement).blur();
            }
            if (!this._isModalOpen) {
              console.log('[GlobalSearchCustomizer] Opening modal from element handler');
              this._openModal();
            }
          };

          const eventTypes = ['pointerdown', 'mousedown', 'click', 'focus', 'focusin'];
          eventTypes.forEach((evt) => {
            el.addEventListener(evt, handler, true);
          });
        });
      } catch (_err) { /* ignore */ }
    });
  }

  private _openModal(): void {
    if (!this._modalContainer) {
      this._createModalContainer();
    }
    console.log('[GlobalSearchCustomizer] _openModal() executing. Setting _isModalOpen = true');
    this._isModalOpen = true;
    sessionStorage.setItem('trivandi_search_modal_open', 'true');
    const element = React.createElement(SearchModal, {
      context: this.context || g_context,
      isOpen: true,
      onDismiss: () => this._closeModal(),
    });
    // eslint-disable-next-line @microsoft/spfx/pair-react-dom-render-unmount
    ReactDOM.render(element, this._modalContainer!);
    console.log('[GlobalSearchCustomizer] SearchModal component rendered to DOM');
  }

  private _closeModal(): void {
    if (!this._modalContainer) return;
    console.log('[GlobalSearchCustomizer] Closing SearchModal');
    this._isModalOpen = false;
    sessionStorage.removeItem('trivandi_search_modal_open');
    // eslint-disable-next-line @microsoft/spfx/pair-react-dom-render-unmount
    ReactDOM.unmountComponentAtNode(this._modalContainer);
  }

  public onDispose(): void {
    if (this._observer) this._observer.disconnect();
    if (this._modalContainer) {
      // eslint-disable-next-line @microsoft/spfx/pair-react-dom-render-unmount
      ReactDOM.unmountComponentAtNode(this._modalContainer);
      this._modalContainer.remove();
    }
    CacheService.stopCleanup();
  }
}
