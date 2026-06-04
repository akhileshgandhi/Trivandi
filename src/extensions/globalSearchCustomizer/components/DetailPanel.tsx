import * as React from 'react';
import { FileText, ExternalLink, Download, Folder } from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';
import { renderFormattedSummary, getSharePointThumbnailUrl } from '../../../utils/SearchHelpers';

import { IDetailPanelProps } from '../interface/IDetailPanelProps';

export const DetailPanel: React.FC<IDetailPanelProps> = ({ selectedFile, searchQuery }) => {
  const [imageError, setImageError] = React.useState(false);
  const [isUrlPopupOpen, setIsUrlPopupOpen] = React.useState(false);
  const [copied, setCopied] = React.useState(false);
  const [snippetCopied, setSnippetCopied] = React.useState(false);

  React.useEffect(() => {
    setImageError(false);
  }, [selectedFile.id]);

  const getActionUrl = (url: string, forceDownload: boolean, libraryUrl?: string): string => {
    if (!url) return '';
    try {
      const isSharePoint = url.includes('.sharepoint.com') || url.includes('/sites/') || url.includes('sharepoint.intel');
      if (isSharePoint) {
        if (forceDownload) {
          // Download path — append ?download=1&web=0 for direct stream
          let finalUrl = url;
          if (finalUrl.includes('download=')) {
            finalUrl = finalUrl.replace(/download=\d/, 'download=1');
          } else {
            const separator = finalUrl.includes('?') ? '&' : '?';
            finalUrl = `${finalUrl}${separator}download=1`;
          }
          if (finalUrl.includes('web=')) {
            finalUrl = finalUrl.replace(/web=\d/, 'web=0');
          } else {
            finalUrl = `${finalUrl}&web=0`;
          }
          return finalUrl;
        } else {
          // Full Preview
          const urlWithoutQuery = url.split('?')[0].toLowerCase();
          const mediaExtensions = ['.png', '.jpg', '.jpeg', '.gif', '.svg', '.mp4', '.mov', '.avi', '.wmv', '.mkv', '.flv', '.webm'];
          const isMediaFile = mediaExtensions.some(ext => urlWithoutQuery.endsWith(ext));

          if (isMediaFile) {
            // Force the native SharePoint lightbox viewer: AllItems.aspx?id={encodedFilePath}
            // We omit the 'parent' parameter because SharePoint automatically infers the 
            // context for the id, preventing Ice Cream Cone errors in deep subfolders.
            try {
              const cleanUrl = url.split('?')[0];
              const hostMatch = cleanUrl.match(/^(https?:\/\/[^/]+)/i);
              const host = hostMatch ? hostMatch[1] : '';
              const urlObj = new URL(cleanUrl);
              const filePath = decodeURIComponent(urlObj.pathname);

              if (libraryUrl && host && filePath) {
                const libUrlObj = new URL(libraryUrl);
                const libraryPath = decodeURIComponent(libUrlObj.pathname);
                return `${host}${libraryPath}/Forms/AllItems.aspx?id=${encodeURIComponent(filePath)}`;
              }
            } catch (err) {
              console.warn('Failed to construct AllItems URL for media, falling back to web=1', err);
            }
            // Fallback: If libraryUrl was missing or parsing failed, append ?web=1 
            // to force the SharePoint native viewer instead of downloading the raw file.
            let finalMediaUrl = url;
            if (url.includes('?')) {
              if (!url.includes('web=1')) finalMediaUrl = `${finalMediaUrl}&web=1`;
            } else {
              finalMediaUrl = `${finalMediaUrl}?web=1`;
            }
            return finalMediaUrl;
          } else {
            // For documents (Word, Excel, PPT, PDF), ?web=1 forces the file to open 
            // safely in the Office Online Server browser viewer instead of downloading.
            let finalUrl = url;
            if (url.includes('?')) {
              if (!url.includes('web=1')) finalUrl = `${finalUrl}&web=1`;
            } else {
              finalUrl = `${finalUrl}?web=1`;
            }
            return finalUrl;
          }
        }
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
        {/* Priority: Graph API thumbnailUrlLarge → thumbnailUrl → getpreview.ashx → icon fallback */}
        {!imageError && (selectedFile.thumbnailUrlLarge || selectedFile.thumbnailUrl || (isRealSharePoint && getSharePointPreviewUrl(selectedFile.url, selectedFile.title))) ? (
          <img
            src={selectedFile.thumbnailUrlLarge || selectedFile.thumbnailUrl || getSharePointPreviewUrl(selectedFile.url, selectedFile.title)}
            className={styles.previewImage}
            onError={() => setImageError(true)}
            alt={selectedFile.title}
          />
        ) : (
          <div className={styles.mockThumbnailWrapper} style={{ borderTop: `4px solid ${selectedFile.color}` }}>
            <div className={styles.mockThumbnailHeader}>
              <div className={styles.mockThumbnailIcon} style={{ backgroundColor: selectedFile.color + '15', color: selectedFile.color }}>
                {selectedFile.type === 'FOLDER' ? <Folder size={20} /> : <FileText size={20} />}
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
            <p className={styles.bentoLabel}>Site Name</p>
            <p className={styles.bentoValue} title={selectedFile.siteName || 'SharePoint Site'} style={{ wordBreak: 'break-word' }}>
              {selectedFile.siteName || 'SharePoint Site'}
            </p>
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
                navigator.clipboard.writeText(selectedFile.description || '').catch(() => undefined);
                setSnippetCopied(true);
                setTimeout(() => setSnippetCopied(false), 2000);
              }}
            >
              {snippetCopied ? 'Copied!' : 'Copy snippet'}
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
            onClick={() => window.open(getActionUrl(selectedFile.url, false, selectedFile.libraryUrl), '_blank')}
          >
            <ExternalLink size={14} />
            Full Preview
          </button>
          <button 
            className={styles.actionSecondaryBtn} 
            title="Download Asset"
            onClick={() => {
              if (!selectedFile.url) return;
              const downloadUrl = getActionUrl(selectedFile.url, true, selectedFile.libraryUrl);
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

      {isUrlPopupOpen && (
        <div 
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: 'rgba(15, 23, 42, 0.6)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 99999,
            backdropFilter: 'blur(4px)'
          }}
          onClick={() => setIsUrlPopupOpen(false)}
        >
          <div 
            style={{
              backgroundColor: '#ffffff',
              borderRadius: '16px',
              padding: '24px',
              maxWidth: '500px',
              width: '90%',
              boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04)',
              border: '1px solid #e2e8f0',
              fontFamily: 'Segoe UI, system-ui, sans-serif'
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
              <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700, color: '#0f172a' }}>Full Document URL</h3>
              <button 
                onClick={() => setIsUrlPopupOpen(false)}
                style={{
                  background: 'none',
                  border: 'none',
                  fontSize: '18px',
                  cursor: 'pointer',
                  color: '#64748b',
                  padding: '4px'
                }}
              >
                ✕
              </button>
            </div>
            
            <div 
              style={{
                backgroundColor: '#f8fafc',
                border: '1px solid #e2e8f0',
                borderRadius: '8px',
                padding: '12px 16px',
                wordBreak: 'break-all',
                fontSize: '13px',
                color: '#334155',
                lineHeight: 1.5,
                maxHeight: '200px',
                overflowY: 'auto',
                marginBottom: '20px',
                fontFamily: 'Consolas, Monaco, monospace'
              }}
            >
              {selectedFile.url}
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }}>
              <button
                onClick={async () => {
                  try {
                    await navigator.clipboard.writeText(selectedFile.url);
                    setCopied(true);
                    setTimeout(() => setCopied(false), 2000);
                  } catch (e) {
                    console.error(e);
                  }
                }}
                style={{
                  backgroundColor: copied ? '#10b981' : selectedFile.color,
                  color: '#ffffff',
                  border: 'none',
                  borderRadius: '8px',
                  padding: '8px 16px',
                  fontSize: '13px',
                  fontWeight: 600,
                  cursor: 'pointer',
                  transition: 'background-color 0.2s ease'
                }}
              >
                {copied ? 'Copied!' : 'Copy Link'}
              </button>
              <button
                onClick={() => setIsUrlPopupOpen(false)}
                style={{
                  backgroundColor: '#f1f5f9',
                  color: '#334155',
                  border: 'none',
                  borderRadius: '8px',
                  padding: '8px 16px',
                  fontSize: '13px',
                  fontWeight: 600,
                  cursor: 'pointer'
                }}
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
