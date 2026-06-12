import * as React from 'react';
import { BsFiletypeXls, BsFiletypeXlsx, BsFiletypePpt, BsFiletypePptx, BsFiletypeDoc, BsFiletypeDocx, BsFileEarmarkText, BsFilePlay } from 'react-icons/bs';
import { IoMdImages } from 'react-icons/io';
import { FaFolderOpen, FaRegFilePdf } from 'react-icons/fa6';

const IconXls = BsFiletypeXls as any;
const IconXlsx = BsFiletypeXlsx as any;
const IconPpt = BsFiletypePpt as any;
const IconPptx = BsFiletypePptx as any;
const IconPdf = FaRegFilePdf as any;
const IconDoc = BsFiletypeDoc as any;
const IconDocx = BsFiletypeDocx as any;
const IconFolder = FaFolderOpen as any;
const IconPlay = BsFilePlay as any;
const IconText = BsFileEarmarkText as any;
const IconImages = IoMdImages as any;
import styles from '../../../styles/PremiumSearch.module.scss';
import { ISearchResult } from '../../../models/ISearchResult';
import { IResultCardProps } from '../interface/IResultCardProps';
import { renderFormattedSummary } from '../../../utils/SearchHelpers';

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

const getFileColor = (fileType: string) => {
  const ft = fileType ? fileType.toLowerCase() : '';
  if (ft === 'pdf') return { color: '#d93025', bg: '#fde7e9' };
  if (['xls', 'xlsx'].includes(ft)) return { color: '#00796b', bg: '#e0f2f1' };
  if (['ppt', 'pptx'].includes(ft)) return { color: '#e52592', bg: '#fce4ec' };
  if (['png', 'jpg', 'jpeg', 'gif', 'svg'].includes(ft)) return { color: '#9334e6', bg: '#f3e5f5' };
  if (['mp4', 'mov', 'avi', 'wmv', 'mkv', 'flv', 'webm'].includes(ft)) return { color: '#ea4335', bg: '#fce8e6' };
  if (ft === 'folder') return { color: '#f39c12', bg: '#fef5e7' }; // Beautiful yellow-gold folder
  return { color: '#1a73e8', bg: '#e8f0fe' }; // Default clean document blue
};

const getCleanSharePointUrl = (webUrl: string, title: string) => {
  if (!webUrl) return '';
  try {
    let cleaned = webUrl;
    // Replace SharePoint List Item detail page link (DispForm.aspx) with direct file path
    if (cleaned.includes('/Forms/DispForm.aspx')) {
      cleaned = cleaned.replace(/\/Forms\/DispForm\.aspx.*/i, `/${title}`);
    }
    // Remove sharing/redirect tokens like /:i:/r/ or /:f:/g/ or /:x:/r/ etc.
    cleaned = cleaned.replace(/\/:[a-z]:\/[a-z]\//i, '/');
    cleaned = cleaned.replace(/\/:[a-z]:\/r\//i, '/');
    cleaned = cleaned.replace(/\/:[a-z]:\/g\//i, '/');
    cleaned = cleaned.split('?')[0]; // Strip query parameters
    return cleaned;
  } catch (e) {
    return webUrl;
  }
};

export const ResultCard: React.FC<IResultCardProps> = ({ result, idx, onClick }) => {
  const color = CARD_COLORS[idx % CARD_COLORS.length];
  const fileColor = getFileColor(result.fileType);
  const [imageError, setImageError] = React.useState(false);
  
  // Format sizes cleanly
  const formatBytes = (bytes: number): string => {
    if (!bytes) return '4.2 MB';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  };

  const formattedSize = result.size ? formatBytes(result.size) : '4.2 MB';
  const isMedia = ['png', 'jpg', 'jpeg', 'gif', 'svg', 'mp4', 'mov', 'avi'].includes(result.fileType?.toLowerCase() || '');
  const thumbUrl = isMedia && !imageError ? (result.thumbnailUrl || '') : '';

  const getInitials = (name: string) => {
    if (!name) return 'U';
    return name.split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
  };

  return (
    <div
      className={styles.resultCard}
      style={{ borderLeftColor: color.accent } as React.CSSProperties}
      onClick={() => onClick(result.id)}
    >
      <div className={styles.cardLeftBlock}>
        <div 
          className={`${styles.cardIconBox} ${thumbUrl ? styles.cardIconBoxImage : ''}`}
          style={{ backgroundColor: fileColor.bg, color: fileColor.color }}
        >
          {thumbUrl ? (
            <img 
              src={thumbUrl} 
              loading="lazy"
              alt={result.title} 
              onError={() => setImageError(true)}
              className={styles.cardThumbnailImage}
            />
          ) : (
            ['png', 'jpg', 'jpeg', 'gif', 'svg'].includes(result.fileType?.toLowerCase() || '') ? <IconImages size={22} /> : 
            result.fileType?.toLowerCase() === 'xls' ? <IconXls size={22} /> :
            result.fileType?.toLowerCase() === 'xlsx' ? <IconXlsx size={22} /> :
            result.fileType?.toLowerCase() === 'ppt' ? <IconPpt size={22} /> :
            result.fileType?.toLowerCase() === 'pptx' ? <IconPptx size={22} /> :
            result.fileType?.toLowerCase() === 'pdf' ? <IconPdf size={22} /> :
            result.fileType?.toLowerCase() === 'doc' ? <IconDoc size={22} /> :
            result.fileType?.toLowerCase() === 'docx' ? <IconDocx size={22} /> :
            ['folder'].includes(result.fileType?.toLowerCase() || '') ? <IconFolder size={22} /> :
            ['mp4', 'mov', 'avi', 'wmv', 'mkv', 'flv', 'webm'].includes(result.fileType?.toLowerCase() || '') ? <IconPlay size={22} /> :
            <IconText size={22} />
          )}
        </div>
      </div>
      
      <div className={styles.cardBody}>
        <div className={styles.cardHeaderRow}>
          <h4 className={styles.cardTitle}>{result.title}</h4>
        </div>

        <div className={styles.cardMetadataRow}>
          <span className={styles.metaBoldLabel}>BY: </span>
          {result.authorPhotoUrl ? (
            <img 
              src={result.authorPhotoUrl} 
              alt={result.author}
              style={{ width: 18, height: 18, borderRadius: '50%', marginRight: 6, verticalAlign: 'middle', display: 'inline-block' }}
            />
          ) : (
            <div style={{ width: 18, height: 18, borderRadius: '50%', backgroundColor: color.accent, color: '#fff', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontSize: '9px', fontWeight: 'bold', marginRight: 6, verticalAlign: 'middle' }}>
              {getInitials(result.author)}
            </div>
          )}
          <span className={styles.metaValueDark}>{result.author}</span>
          <span className={styles.metaDot} />
          <span>{result.lastModified ? new Date(result.lastModified).toLocaleDateString() : 'Today'}</span>
          <span className={styles.metaDot} />
          <span>{formattedSize}</span>
        </div>

        <p className={styles.cardDescription}>
          {result.fileType === 'folder'
            ? '📁 Folder — Click to explore contents'
            : (result.summary && result.summary !== 'No description preview available.'
               ? renderFormattedSummary(result.summary, '', color.accent)
               : 'No preview available for this file.'
              )
          }
        </p>
        
        <div className={styles.cardProjectHubRow}>
          <span>SITE: </span>
          <span className={styles.projectHubValue} style={{ color: color.accent }}>{result.siteName ? result.siteName.replace(/([a-z])([A-Z])/g, '$1 $2') : 'SharePoint Site'}</span>
        </div>
      </div>
    </div>
  );
};
