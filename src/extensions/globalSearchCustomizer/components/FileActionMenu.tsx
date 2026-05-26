import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { 
  MoreVertical, 
  ExternalLink, 
  Monitor, 
  Download, 
  FolderOpen, 
  Copy,
  Check
} from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';
import { ISearchResult } from '../../../models/ISearchResult';

import { IFileActionMenuProps } from '../interface/IFileActionMenuProps';

export const FileActionMenu: React.FC<IFileActionMenuProps> = ({ file, onOpenChange }) => {
  const [isOpen, setIsOpen] = useState(false);
  const [copied, setCopied] = useState(false);
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

  const handleCopyLink = async (e: React.MouseEvent) => {
    e.stopPropagation();
    try {
      const url = file.webUrl || `https://sharepoint.trivandi.com/Shared%20Documents/${encodeURIComponent(file.title)}`;
      await navigator.clipboard.writeText(url);
      setCopied(true);
      setTimeout(() => {
        setCopied(false);
        setIsOpen(false);
        onOpenChange?.(false);
      }, 1500);
    } catch (err) {
      console.error('Failed to copy link!', err);
    }
  };

  const menuItems = [
    { 
      label: 'Open in browser', 
      icon: ExternalLink, 
      onClick: (e: React.MouseEvent) => {
        e.stopPropagation();
        window.open(file.webUrl, '_blank');
        setIsOpen(false);
        onOpenChange?.(false);
      }
    },
    { 
      label: 'Open in app', 
      icon: Monitor, 
      onClick: (e: React.MouseEvent) => {
        e.stopPropagation();
        window.open(`ms-word:ofe|u|${file.webUrl}`, '_blank');
        setIsOpen(false);
        onOpenChange?.(false);
      }
    },
    { 
      label: 'Download', 
      icon: Download, 
      onClick: (e: React.MouseEvent) => {
        e.stopPropagation();
        window.open(file.webUrl + '?web=0', '_blank');
        setIsOpen(false);
        onOpenChange?.(false);
      }
    },
    { 
      label: 'Open file location', 
      icon: FolderOpen, 
      onClick: (e: React.MouseEvent) => {
        e.stopPropagation();
        window.open(file.siteUrl || '#', '_blank');
        setIsOpen(false);
        onOpenChange?.(false);
      }
    },
    { 
      label: copied ? 'Link copied!' : 'Copy link', 
      icon: copied ? Check : Copy, 
      onClick: handleCopyLink,
      className: copied ? styles.actionItemCopied : ''
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
              className={`${styles.actionDropdownItem} ${item.className || ''}`}
            >
              <item.icon 
                size={16} 
                className={`${styles.actionItemIcon} ${copied && item.label === 'Link copied!' ? styles.actionIconCopied : ''}`} 
                strokeWidth={2} 
              />
              <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                {item.label}
              </span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
};
