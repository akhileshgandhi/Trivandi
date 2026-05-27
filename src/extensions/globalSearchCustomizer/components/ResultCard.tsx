import * as React from 'react';
import { Image as ImageIcon, FileSpreadsheet, FileText, FileBox } from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';
import { ISearchResult } from '../../../models/ISearchResult';

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

import { IResultCardProps } from '../interface/IResultCardProps';

export const ResultCard: React.FC<IResultCardProps> = ({ result, idx, onClick }) => {
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
        </div>

        <div className={styles.cardMetadataRow}>
          <span className={styles.metaBoldLabel}>BY: </span>
          <span className={styles.metaValueDark}>{result.author}</span>
          <span className={styles.metaDot} />
          <span>{result.lastModified ? new Date(result.lastModified).toLocaleDateString() : 'Today'}</span>
          <span className={styles.metaDot} />
          <span>{formattedSize}</span>
        </div>

        <p className={styles.cardDescription}>{result.summary}</p>
        
        <div className={styles.cardProjectHubRow}>
          <span>PROJECT HUB: </span>
          <span className={styles.projectHubValue} style={{ color: color.accent }}>{result.siteName || 'SharePoint Site'}</span>
        </div>
      </div>
    </div>
  );
};
