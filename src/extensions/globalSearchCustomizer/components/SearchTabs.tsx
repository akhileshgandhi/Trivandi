import * as React from 'react';
import styles from '../../../styles/PremiumSearch.module.scss';
import { ISearchTabsProps } from '../interface/ISearchTabsProps';

const TABS = ['All', 'Files', 'Images', 'Videos'];

export const SearchTabs: React.FC<ISearchTabsProps> = ({ activeTab, onTabChange, onClearAll }) => {
  return (
    <div className={styles.topTabBar} style={{ borderBottom: '1px solid #e2e8f0', display: 'flex', width: '100%', alignItems: 'center', justifyContent: 'space-between', paddingRight: '24px' }}>
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
          style={{ height: '30px' }}
        >
          Clear All
        </button>
      )}
    </div>
  );
};
