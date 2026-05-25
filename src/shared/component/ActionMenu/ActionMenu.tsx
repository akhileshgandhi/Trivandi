import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import * as ReactDOM from 'react-dom';
import { MoreVertical } from 'lucide-react';
import styles from './ActionMenu.module.scss';

export interface ActionMenuItem {
  key: string;
  text: string;
  icon?: React.ReactNode;
  onClick?: () => void;
  disabled?: boolean;
  onRender?: () => React.ReactNode;
  className?: string;
}

export interface ActionMenuProps {
  items: ActionMenuItem[];
  icon?: React.ReactNode;
  className?: string;
}

export const ActionMenu: React.FC<ActionMenuProps> = ({ items, icon, className }) => {
  const [isOpen, setIsOpen] = useState(false);
  const [coords, setCoords] = useState({ top: 0, right: 0 });
  const menuRef = useRef<HTMLDivElement>(null);
  const dropdownRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      const target = event.target as Node;
      if (
        menuRef.current && !menuRef.current.contains(target) &&
        (!dropdownRef.current || !dropdownRef.current.contains(target))
      ) {
        setIsOpen(false);
      }
    };

    if (isOpen) {
      document.addEventListener('mousedown', handleClickOutside);
    }
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [isOpen]);

  return (
    <div className={`${styles.actionMenuContainer} ${className || ''}`} ref={menuRef}>
      <button
        type="button"
        className={styles.toggleButton}
        onClick={(e) => {
          e.stopPropagation();
          if (!isOpen && menuRef.current) {
            const rect = menuRef.current.getBoundingClientRect();
            const scrollY = window.pageYOffset || document.documentElement.scrollTop;
            const scrollX = window.pageXOffset || document.documentElement.scrollLeft;
            setCoords({
              top: rect.bottom + scrollY,
              right: (document.documentElement.clientWidth - rect.right) + scrollX,
            });
          }
          setIsOpen(!isOpen);
        }}
        aria-label="More actions"
      >
        {icon || <MoreVertical size={18} />}
      </button>

      {isOpen && ReactDOM.createPortal(
        <div
          ref={dropdownRef}
          className={styles.menuDropdown}
          style={{
            position: 'absolute',
            top: `${coords.top + 4}px`,
            right: `${coords.right}px`,
            left: 'auto',
            zIndex: 999999,
          }}
        >
          {items.map((item) => {
            if (item.onRender) {
              return (
                <div key={item.key} className={`${styles.customItem} ${item.className || ''}`}>
                  {item.onRender()}
                </div>
              );
            }

            if (item.text === '-') {
              return <div key={item.key} className={styles.divider} />;
            }

            return (
              <button
                key={item.key}
                type="button"
                className={`${styles.menuItem} ${item.className || ''}`}
                disabled={item.disabled}
                onClick={(e) => {
                  e.stopPropagation();
                  if (item.onClick) item.onClick();
                  setIsOpen(false);
                }}
              >
                {item.icon && <span className={styles.itemIcon}>{item.icon}</span>}
                <span className={styles.itemText}>{item.text}</span>
              </button>
            );
          })}
        </div>,
        document.body
      )}
    </div>
  );
};
