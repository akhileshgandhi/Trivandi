import * as React from 'react';
import { FileText, ExternalLink, Download } from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';
import { renderFormattedSummary, getSharePointThumbnailUrl } from '../../../utils/SearchHelpers';

import { IDetailPanelProps } from '../interface/IDetailPanelProps';

export const DetailPanel: React.FC<IDetailPanelProps> = ({ selectedFile, searchQuery }) => {
  const [imageError, setImageError] = React.useState(false);

  React.useEffect(() => {
    setImageError(false);
  }, [selectedFile.id]);

  const getActionUrl = (url: string, forceDownload: boolean): string => {
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

  const isRealSharePoint = selectedFile.url && (selectedFile.url.includes('.sharepoint.com') || selectedFile.url.includes('/sites/'));

  const getSharePointPreviewUrl = (url: string, title: string): string => {
    return getSharePointThumbnailUrl(url, title, 3);
  };

  return (
    <div className={styles.previewContentScroll}>
      {/* Document Visual Preview Card / Thumbnail */}
      <div className={styles.previewDocVisualCard}>
        {isRealSharePoint && !imageError ? (
          <img 
            src={getSharePointPreviewUrl(selectedFile.url, selectedFile.title)}
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
              <div className={`${styles.mockThumbnailTextLine} ${styles.skeleton_w90}`} />
              <div className={`${styles.mockThumbnailTextLine} ${styles.skeleton_w80}`} />
              <div className={`${styles.mockThumbnailTextLine} ${styles.skeleton_w95}`} />
              <div className={`${styles.mockThumbnailTextLine} ${styles.skeleton_w60}`} />
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
            <span 
              className={styles.siteBadge} 
              style={{ 
                backgroundColor: selectedFile.color + '10', 
                color: selectedFile.color 
              }}
            >
              {selectedFile.url ? selectedFile.url.toUpperCase() : 'SHAREPOINT SITE'}
            </span>
            <span className={styles.metaDot} />
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
            <p className={`${styles.bentoValue} ${styles.statusSuccess}`}>Encrypted</p>
          </div>
        </div>

        {/* Author Card */}
        <div className={styles.authorCard}>
          <div className={styles.authorCardLeft}>
            <div className={styles.authorAvatar} style={{ backgroundColor: selectedFile.color }}>
              {selectedFile.author ? selectedFile.author.charAt(0).toUpperCase() : 'U'}
            </div>
            <div>
              <p className={styles.ownerTitle}>Document Owner</p>
              <p className={styles.ownerName}>{selectedFile.author || 'SharePoint User'}</p>
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
              "{renderFormattedSummary(selectedFile.description, searchQuery, selectedFile.color)}"
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
            onClick={() => {
              if (!selectedFile.url) return;
              const downloadUrl = getActionUrl(selectedFile.url, true);
              const link = document.createElement('a');
              link.href = downloadUrl;
              link.setAttribute('download', selectedFile.title || 'download');
              link.setAttribute('target', '_self');
              document.body.appendChild(link);
              link.click();
              document.body.removeChild(link);
            }}
          >
            <Download size={18} />
          </button>
        </div>
      </div>
    </div>
  );
};
