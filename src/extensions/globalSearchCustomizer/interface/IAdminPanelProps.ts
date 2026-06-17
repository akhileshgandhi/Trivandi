import { AdminConfigService } from '../services/AdminConfigService';

export interface IAdminConfig {
  sites: string[];
  fileTypes: string[];
  dateFilters: string[];
}

export interface IAdminPanelProps {
  isOpen: boolean;
  onClose: () => void;
  config: IAdminConfig;
  onConfigChange: (config: IAdminConfig) => void;
  adminConfigService?: AdminConfigService | null;
}

export interface IConfigRowProps {
  label: string;
  onDelete: () => void;
}

export interface IAddRowProps {
  value: string;
  placeholder: string;
  onChange: (value: string) => void;
  onAdd: () => void;
}

export interface ISectionHeaderProps {
  icon: string;
  title: string;
}

export const DEFAULT_ADMIN_CONFIG: IAdminConfig = {
  sites: [
    'Intranet',
    'People', 
    'Company',
    'Marketing',
    'Projects',
    'Trivandi London',
    'TDMCC',
    'Trivandi USA',
    'Trivandi Australia',
    'Trivandi KSA',
  ],
  fileTypes: ['PDF', 'DOC', 'XLS', 'PPT'],
  dateFilters: [
    'Today',
    'Yesterday',
    'Last 7 Days',
    'Last 30 Days',
    'This Year',
  ],
};

export const ADMIN_STORAGE_KEY = 'trivandi_admin_config_v1';

export function loadAdminConfig(): IAdminConfig {
  try {
    const stored = localStorage.getItem(ADMIN_STORAGE_KEY);
    if (stored) {
      return { 
        ...DEFAULT_ADMIN_CONFIG, 
        ...JSON.parse(stored) 
      };
    }
  } catch (e) { /* ignore */ }
  return DEFAULT_ADMIN_CONFIG;
}

export function saveAdminConfig(config: IAdminConfig): void {
  try {
    localStorage.setItem(
      ADMIN_STORAGE_KEY, 
      JSON.stringify(config)
    );
  } catch (e) { /* ignore */ }
}
