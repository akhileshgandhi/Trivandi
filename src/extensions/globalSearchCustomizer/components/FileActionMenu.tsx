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
        let finalUrl = url;
        if (forceDownload) {
          // Append SharePoint direct stream download parameter
          if (finalUrl.includes('download=')) {
            finalUrl = finalUrl.replace(/download=\d/, 'download=1');
          } else {
            const separator = finalUrl.includes('?') ? '&' : '?';
            finalUrl = `${finalUrl}${separator}download=1`;
          }
          // Set web=0 to bypass opening in the online editor
          if (finalUrl.includes('web=')) {
            finalUrl = finalUrl.replace(/web=\d/, 'web=0');
          } else {
            finalUrl = `${finalUrl}&web=0`;
          }
        } else {
          // Full preview - open in online reader
          if (finalUrl.includes('web=')) {
            finalUrl = finalUrl.replace(/web=\d/, 'web=1');
          } else {
            const separator = finalUrl.includes('?') ? '&' : '?';
            finalUrl = `${finalUrl}${separator}web=1`;
          }
        }
        return finalUrl;
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
        if (!file.webUrl) return;
        const downloadUrl = getBrowserOpenUrl(file.webUrl, true);
        const link = document.createElement('a');
        link.href = downloadUrl;
        link.setAttribute('download', file.title || 'download');
        link.setAttribute('target', '_self');
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
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
