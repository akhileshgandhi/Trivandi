import * as React from 'react';
import { ExternalLink, FileText, Image as ImageIcon, FileSpreadsheet, FileBox } from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';
import { ISearchResult } from '../../../models/ISearchResult';
import { ISearchTabsProps } from '../interface/ISearchTabsProps';

// CARD ACCENT COLOR ASSIGNMENTS
const CARD_COLORS = [
  { accent: '#9334e6', bg: '#f3e5f5', shadow: 'rgba(147, 52, 230, 0.2)' },
  { accent: '#00796b', bg: '#e0f2f1', shadow: 'rgba(0, 121, 107, 0.2)' },
  { accent: '#00acc1', bg: '#e0f7fa', shadow: 'rgba(0, 172, 193, 0.2)' },
  { accent: '#1a73e8', bg: '#e8f0fe', shadow: 'rgba(26, 115, 232, 0.2)' },
  { accent: '#e52592', bg: '#fce4ec', shadow: 'rgba(229, 37, 146, 0.2)' },
  { accent: '#fa903e', bg: '#fff4e5', shadow: 'rgba(250, 144, 62, 0.2)' },
  { accent: '#d93025', bg: '#fde7e9', shadow: 'rgba(217, 48, 37, 0.2)' },
  { accent: '#188038', bg: '#e6f4ea', shadow: 'rgba(24, 128, 56, 0.2)' },
];

export const ResultCard: React.FC<{ result: ISearchResult; idx: number; onClick: (id: string) => void }> = ({ result, idx, onClick }) => {
  const color = CARD_COLORS[idx % CARD_COLORS.length];
  
  // Format sizes cleanly
  const formatBytes = (bytes: number): string => {
    if (!bytes) return '4.2 MB';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  };

  const formattedSize = result.size ? formatBytes(result.size) : '4.2 MB';

  return (
    <div
      className={styles.resultCard}
      style={{ borderLeftColor: color.accent } as React.CSSProperties}
      onClick={() => onClick(result.id)}
    >
      <div className={styles.cardLeftBlock}>
        <div 
          className={styles.cardIconBox}
          style={{ backgroundColor: color.bg, color: color.accent }}
        >
          {result.fileType?.toLowerCase() === 'png' ? <ImageIcon size={22} strokeWidth={2} /> : 
           ['xls', 'xlsx'].includes(result.fileType?.toLowerCase()) ? <FileSpreadsheet size={22} strokeWidth={2} /> :
           ['pdf'].includes(result.fileType?.toLowerCase()) ? <FileText size={22} strokeWidth={2} /> :
           ['doc', 'docx'].includes(result.fileType?.toLowerCase()) ? <FileText size={22} strokeWidth={2} /> :
           <FileBox size={22} strokeWidth={2} />}
        </div>
      </div>
      
      <div className={styles.cardBody}>
        <div className={styles.cardHeaderRow}>
          <h4 className={styles.cardTitle}>{result.title}</h4>
          <div className={styles.scoreContainer}>
            <div className={styles.scoreRow}>
              <span className={`${styles.scoreBadge} ${styles.scoreBM25}`}>BM25: 4.80</span>
              <span className={`${styles.scoreBadge} ${styles.scoreTFIDF}`}>TF-IDF: 2.12</span>
            </div>
          </div>
        </div>

        <div className={styles.cardMetadataRow}>
          <span className={styles.metaBoldLabel}>BY: </span>
          <span className={styles.metaValueDark}>{result.author}</span>
          <span className={styles.metaDot} />
          <span>{result.lastModified ? new Date(result.lastModified).toLocaleDateString() : 'Today'}</span>
          <span className={styles.metaDot} />
          <span>{formattedSize}</span>
        </div>

        <p 
          className={styles.cardDescription} 
          dangerouslySetInnerHTML={{ 
            __html: (result.summary || '')
              .replace(/<c0>/g, '<strong>')
              .replace(/<\/c0>/g, '</strong>')
              .replace(/<ddd\/>/g, '&#8230;') 
          }} 
        />
        
        <div className={styles.cardProjectHubRow}>
          <span>PROJECT HUB: </span>
          <span className={styles.projectHubValue} style={{ color: color.accent }}>{result.siteName || 'SharePoint Site'}</span>
        </div>
      </div>
    </div>
  );
};

// SearchTabs Component implementation

const TABS = ['All', 'Files', 'Images'];

export const SearchTabs: React.FC<ISearchTabsProps> = ({ activeTab, onTabChange }) => {
  return (
    <div className={styles.topTabBar} style={{ borderBottom: '1px solid #e2e8f0', display: 'flex', width: '100%' }}>
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
    </div>
  );
};
