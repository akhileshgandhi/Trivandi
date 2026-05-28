import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import {
  MoreVertical,
  ExternalLink,
  Monitor,
  Download,
  FolderOpen,
  Copy
} from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';
import { ISearchResult } from '../../../models/ISearchResult';

import { IFileActionMenuProps } from '../interface/IFileActionMenuProps';

export const FileActionMenu: React.FC<IFileActionMenuProps> = ({ file, onOpenChange }) => {
  const [isOpen, setIsOpen] = useState(false);
  const menuRef = useRef<HTMLDivElement>(null);

  // Close menu when clicking outside
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(event.target as Node)) {
        setIsOpen(false);
        onOpenChange?.(false);
      }
    };

    if (isOpen) {
      document.addEventListener('mousedown', handleClickOutside);
    }
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [isOpen, onOpenChange]);

  const toggleMenu = (e: React.MouseEvent) => {
    e.stopPropagation(); // Prevent card selection when clicking menu
    const nextState = !isOpen;
    setIsOpen(nextState);
    onOpenChange?.(nextState);
  };

  // Helper: Get base tenant URL
  const getTenantUrl = (url: string) => {
    if (!url) return 'https://trivandi.sharepoint.com';
    try {
      const parsed = new URL(url);
      return `${parsed.protocol}//${parsed.hostname}`;
    } catch (e) {
      return 'https://trivandi.sharepoint.com';
    }
  };

  // Helper: Parse directory location of the file in SharePoint
  const getFolderUrl = (fileUrl: string) => {
    if (!fileUrl) return '';
    try {
      const parts = fileUrl.split('/');
      parts.pop(); // Remove filename
      return parts.join('/');
    } catch (e) {
      return fileUrl;
    }
  };

  // Helper: Format SharePoint web URL to enforce browser viewing (web=1) or direct download (web=0)
  const getBrowserOpenUrl = (url: string, forceDownload: boolean): string => {
    if (!url) return '';
    try {
      const isSharePoint = url.includes('.sharepoint.com') || url.includes('/sites/') || url.includes('sharepoint.intel');
      if (isSharePoint) {
        if (url.includes('web=')) {
          return url.replace(/web=\d/, `web=${forceDownload ? '0' : '1'}`);
        } else {
          const separator = url.includes('?') ? '&' : '?';
          return `${url}${separator}web=${forceDownload ? '0' : '1'}`;
        }
      }
    } catch (e) {
      // fallback
    }
    return url;
  };

  // Helper: Determine application type
  const getAppName = (fileType: string) => {
    const type = (fileType || '').toLowerCase();
    if (['xls', 'xlsx'].includes(type)) return { name: 'Excel', protocol: 'ms-excel:ofe|u|' };
    if (['doc', 'docx'].includes(type)) return { name: 'Word', protocol: 'ms-word:ofe|u|' };
    if (['ppt', 'pptx'].includes(type)) return { name: 'PowerPoint', protocol: 'ms-powerpoint:ofe|u|' };
    if (['pdf'].includes(type)) return { name: 'Acrobat', protocol: 'pdf:' };
    return { name: 'App', protocol: 'ms-word:ofe|u|' };
  };

  const appInfo = getAppName(file.fileType || '');

  const menuItems = [
    {
      label: 'Open in browser',
      icon: ExternalLink,
      onClick: (e: React.MouseEvent) => {
        e.stopPropagation();
        window.open(getBrowserOpenUrl(file.webUrl, false), '_blank');
        setIsOpen(false);
        onOpenChange?.(false);
      }
    },
    {
      label: 'Open in app',
      icon: Monitor,
      onClick: (e: React.MouseEvent) => {
        e.stopPropagation();
        const launchUrl = `${appInfo.protocol}${file.webUrl}`;
        window.open(launchUrl, '_blank');
        setIsOpen(false);
        onOpenChange?.(false);
      }
    },
    {
      label: 'Download',
      icon: Download,
      onClick: (e: React.MouseEvent) => {
        e.stopPropagation();
        // Dynamic zero-navigation same-page download
        const downloadUrl = getBrowserOpenUrl(file.webUrl, true);
        const iframe = document.createElement('iframe');
        iframe.style.display = 'none';
        iframe.src = downloadUrl;
        document.body.appendChild(iframe);
        setTimeout(() => {
          document.body.removeChild(iframe);
        }, 3000);
        setIsOpen(false);
        onOpenChange?.(false);
      }
    },
    {
      label: 'Open file location',
      icon: FolderOpen,
      onClick: (e: React.MouseEvent) => {
        e.stopPropagation();
        const folderUrl = getFolderUrl(file.webUrl);
        window.open(folderUrl || file.siteUrl || '#', '_blank');
        setIsOpen(false);
        onOpenChange?.(false);
      }
    },
    {
      label: 'Copy link',
      icon: Copy,
      onClick: (e: React.MouseEvent) => {
        e.stopPropagation();
        window.dispatchEvent(new CustomEvent('trivandi-copy-link', { detail: file }));
        setIsOpen(false);
        onOpenChange?.(false);
      }
    },
  ];

  return (
    <div className={styles.actionMenuWrapper} ref={menuRef}>
      <button
        type="button"
        onClick={toggleMenu}
        aria-label="More actions"
        aria-expanded={isOpen}
        aria-haspopup="true"
        className={`${styles.actionTriggerBtn} ${isOpen ? styles.actionTriggerBtnActive : ''}`}
      >
        <MoreVertical size={18} strokeWidth={2.5} />
      </button>

      {isOpen && (
        <div className={styles.actionDropdownList}>
          {menuItems.map((item, index) => (
            <button
              key={index}
              onClick={item.onClick}
              className={styles.actionDropdownItem}
            >
              <item.icon
                size={16}
                className={styles.actionItemIcon}
                strokeWidth={2}
              />
              <span className={styles.actionItemText}>
                {item.label}
              </span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
};
