import * as React from 'react';
import styles from '../../../styles/PremiumSearch.module.scss';
import { ISearchTabsProps } from '../interface/ISearchTabsProps';

const TABS = ['All',  'Folders','Files', 'Images', 'Videos'];

export const SearchTabs: React.FC<ISearchTabsProps> = ({ activeTab, onTabChange, onClearAll }) => {
  return (
    <div className={styles.topTabBar}>
      <div className={styles.tabButtonsWrapper}>
        {TABS.map((tab) => (
          <button
            key={tab}
            onClick={() => onTabChange(tab)}
            className={`${styles.topTabItem} ${activeTab === tab ? styles.topTabItemActive : ''}`}
          >
            {tab}
            {activeTab === tab && (
              <div className={styles.activeTabLine} />
            )}
          </button>
        ))}
      </div>
      
      {onClearAll && (
        <button 
          onClick={onClearAll}
          className={styles.clearAllBtn}
        >
          Clear All
        </button>
      )}
    </div>
  );
};
