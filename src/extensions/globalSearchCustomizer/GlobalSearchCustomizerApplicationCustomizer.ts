import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import SearchModal from './components/SearchModal';
import '../../styles/PremiumSearch.module.scss';

export interface IGlobalSearchCustomizerApplicationCustomizerProperties {
  testMessage: string;
}

export default class GlobalSearchCustomizerApplicationCustomizer
  extends BaseApplicationCustomizer<IGlobalSearchCustomizerApplicationCustomizerProperties> {

  private _modalContainer: HTMLElement | null = null;
  private _observer: MutationObserver | null = null;
  private _attached: Set<Element> = new Set();
  private _isModalOpen: boolean = false;

  public onInit(): Promise<void> {
    this._createModalContainer();
    this._interceptSearchBar();

    const wasOpen = sessionStorage.getItem('trivandi_search_modal_open') === 'true';
    if (wasOpen) {
      this._openModal();
    }

    return Promise.resolve();
  }

  private _createModalContainer(): void {
    this._modalContainer = document.createElement('div');
    this._modalContainer.id = 'trivandi-search-root';
    document.body.appendChild(this._modalContainer);
  }

  private _interceptSearchBar(): void {
    this._attachListeners();
    this._observer = new MutationObserver(() => {
      this._attachListeners();
    });
    this._observer.observe(document.body, { childList: true, subtree: true });
  }

  private _attachListeners(): void {
    const SELECTORS = [
      '#O365_SearchBoxField',
      '[data-automationid="ShyHeader"] input[type="search"]',
      '[data-automationid="ShyHeader"] input',
      'input[placeholder*="Search"]',
      'input[aria-label*="Search"]',
      '[class*="searchBox"] input',
    ];

    SELECTORS.forEach((sel) => {
      try {
        document.querySelectorAll(sel).forEach((el) => {
          if (this._attached.has(el)) return;
          this._attached.add(el);
          const handler = (e: Event): void => {
            e.preventDefault();
            e.stopImmediatePropagation();
            if (!this._isModalOpen) this._openModal();
          };
          el.addEventListener('click', handler, true);
          el.addEventListener('focus', handler, true);
        });
      } catch (_err) { /* ignore */ }
    });
  }

  private _openModal(): void {
    if (!this._modalContainer) return;
    this._isModalOpen = true;
    sessionStorage.setItem('trivandi_search_modal_open', 'true');
    const element = React.createElement(SearchModal, {
      context: this.context,
      isOpen: true,
      onDismiss: () => this._closeModal(),
    });
    ReactDOM.render(element, this._modalContainer);
  }

  private _closeModal(): void {
    if (!this._modalContainer) return;
    this._isModalOpen = false;
    sessionStorage.removeItem('trivandi_search_modal_open');
    ReactDOM.unmountComponentAtNode(this._modalContainer);
  }

  public onDispose(): void {
    if (this._observer) this._observer.disconnect();
    if (this._modalContainer) {
      ReactDOM.unmountComponentAtNode(this._modalContainer);
      this._modalContainer.remove();
    }
  }
}
