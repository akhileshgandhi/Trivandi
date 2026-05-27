import * as React from 'react';
import { FileText, ExternalLink, Download } from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';

export interface IDetailPanelProps {
  selectedFile: {
    id: string;
    title: string;
    author: string;
    date: string;
    size: string;
    description: string;
    url: string;
    type: string;
    starred: boolean;
    color: string;
    badgeColor: string;
    summary: string;
    siteName: string;
    siteUrl: string;
    webUrl: string;
    lastModified: string;
  };
  searchQuery: string;
}

export const DetailPanel: React.FC<IDetailPanelProps> = ({ selectedFile, searchQuery }) => {
  const [imageError, setImageError] = React.useState(false);

  React.useEffect(() => {
    setImageError(false);
  }, [selectedFile.id]);

  const highlightText = (text: string, query: string) => {
    if (!query || !query.trim()) return text;
    const parts = text.split(new RegExp(`(${query})`, 'gi'));
    return (
      <>
        {parts.map((part, i) => 
          part.toLowerCase() === query.toLowerCase() ? (
            <span key={i} style={{ backgroundColor: 'rgba(26, 115, 232, 0.15)', color: '#1a73e8', fontWeight: 'bold', padding: '0 2px', borderRadius: '2px' }}>
              {part}
            </span>
          ) : (
            <span key={i}>{part}</span>
          )
        )}
      </>
    );
  };

  const getActionUrl = (url: string, forceDownload: boolean): string => {
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

  const isRealSharePoint = selectedFile.url && (selectedFile.url.includes('.sharepoint.com') || selectedFile.url.includes('/sites/'));

  const getSharePointPreviewUrl = (url: string): string => {
    if (!url) return '';
    try {
      const urlObj = new URL(url);
      return `${urlObj.origin}/_layouts/15/getpreview.ashx?resolution=3&path=${encodeURIComponent(url)}`;
    } catch (e) {
      return url;
    }
  };

  return (
    <div className={styles.previewContentScroll}>
      {/* Document Visual Preview Card / Thumbnail */}
      <div className={styles.previewDocVisualCard}>
        {isRealSharePoint && !imageError ? (
          <img 
            src={getSharePointPreviewUrl(selectedFile.url)}
            className={styles.previewImage}
            onError={() => setImageError(true)}
            alt={selectedFile.title}
          />
        ) : (
          <div className={styles.mockThumbnailWrapper} style={{ borderTop: `4px solid ${selectedFile.color}` }}>
            <div className={styles.mockThumbnailHeader}>
              <div className={styles.mockThumbnailIcon} style={{ backgroundColor: selectedFile.color + '15', color: selectedFile.color }}>
                <FileText size={20} />
              </div>
              <span className={styles.mockThumbnailBadge} style={{ backgroundColor: selectedFile.color + '20', color: selectedFile.color }}>
                {selectedFile.type}
              </span>
            </div>
            
            <div className={styles.mockThumbnailBody}>
              <h4 className={styles.mockThumbnailTitle}>{selectedFile.title}</h4>
              <div className={styles.mockThumbnailTextLine} style={{ width: '90%' }} />
              <div className={styles.mockThumbnailTextLine} style={{ width: '80%' }} />
              <div className={styles.mockThumbnailTextLine} style={{ width: '95%' }} />
              <div className={styles.mockThumbnailTextLine} style={{ width: '60%' }} />
            </div>

            <div className={styles.mockThumbnailFooter}>
              <span className={styles.mockThumbnailMeta}>BY: {selectedFile.author}</span>
              <span className={styles.mockThumbnailMeta}>{selectedFile.size}</span>
            </div>
          </div>
        )}
      </div>

      {/* Details & Metadata */}
      <div>
        <div className={styles.visualIdentityHeader}>
          <h2 className={styles.visualTitle}>
            {selectedFile.title}
          </h2>
          <div className={styles.visualSubRow}>
            <span className={styles.siteBadge} style={{ backgroundColor: selectedFile.color + '10', color: selectedFile.color, fontSize: '9px', fontWeight: 800, padding: '4px 10px', borderRadius: '6px', whiteSpace: 'nowrap', textOverflow: 'ellipsis', overflow: 'hidden', maxWidth: '300px' }}>
              {selectedFile.url ? selectedFile.url.toUpperCase() : 'SHAREPOINT SITE'}
            </span>
            <span style={{ width: '4px', height: '4px', borderRadius: '50%', backgroundColor: '#cbd5e1', flexShrink: 0 }} />
            <span className={styles.verifiedLabel}>Verified Source</span>
          </div>
        </div>
        
        {/* Bento Grid */}
        <div className={styles.bentoGrid}>
          <div className={styles.bentoItem}>
            <p className={styles.bentoLabel}>Format</p>
            <p className={styles.bentoValue}>{selectedFile.type}</p>
          </div>
          <div className={styles.bentoItem}>
            <p className={styles.bentoLabel}>File Size</p>
            <p className={styles.bentoValue}>{selectedFile.size || '1.2 MB'}</p>
          </div>
          <div className={styles.bentoItem}>
            <p className={styles.bentoLabel}>Last Sync</p>
            <p className={styles.bentoValue}>{selectedFile.date || 'Today'}</p>
          </div>
          <div className={styles.bentoItem}>
            <p className={styles.bentoLabel}>Security</p>
            <p className={styles.bentoValue} style={{ color: '#10b981' }}>Encrypted</p>
          </div>
        </div>

        {/* Author Card */}
        <div className={styles.authorCard}>
          <div className={styles.authorCardLeft}>
            <div className={styles.authorAvatar} style={{ backgroundColor: selectedFile.color }}>
              {selectedFile.author ? selectedFile.author.charAt(0).toUpperCase() : 'U'}
            </div>
            <div>
              <p style={{ fontSize: '9px', fontWeight: 800, textTransform: 'uppercase', letterSpacing: '0.1em', color: '#94a3b8', margin: '0 0 2px 0', lineHeight: 1 }}>Document Owner</p>
              <p style={{ fontSize: '14px', fontWeight: 800, color: '#1e293b', margin: 0 }}>{selectedFile.author || 'SharePoint User'}</p>
            </div>
          </div>
          <div className={styles.authorBadge}>
            Profile
          </div>
        </div>

        {/* Insights Section */}
        <div className={styles.highlightsSection}>
          <div className={styles.highlightsHeader}>
            <h4 className={styles.highlightsTitle}>Content Highlight</h4>
            <span 
              className={styles.copySnippet}
              onClick={() => {
                navigator.clipboard.writeText(selectedFile.description || '');
              }}
            >
              Copy snippet
            </span>
          </div>
          <div className={styles.highlightsContentWrapper}>
            <div className={styles.highlightsLineAccent} style={{ backgroundColor: selectedFile.color + '40' }} />
            <p className={styles.highlightsText}>
              "{highlightText(selectedFile.description, searchQuery)}"
            </p>
          </div>
        </div>

        {/* Action Footer */}
        <div className={styles.actionFooter}>
          <button 
            className={styles.actionPrimaryBtn}
            style={{ backgroundColor: selectedFile.color, boxShadow: `0 10px 20px -5px ${selectedFile.color}40` }}
            onClick={() => window.open(getActionUrl(selectedFile.url, false), '_blank')}
          >
            <ExternalLink size={14} />
            Full Preview
          </button>
          <button 
            className={styles.actionSecondaryBtn} 
            title="Download Asset"
            onClick={() => window.open(getActionUrl(selectedFile.url, true), '_blank')}
          >
            <Download size={18} />
          </button>
        </div>
      </div>
    </div>
  );
};
