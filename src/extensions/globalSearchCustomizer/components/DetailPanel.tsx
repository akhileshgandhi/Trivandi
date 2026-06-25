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

  const [isOpeningPreview, setIsOpeningPreview] = React.useState(false);

  const handleFullPreview = async () => {
    if (!selectedFile.driveId || !selectedFile.itemId) {
      window.open(selectedFile.url, '_blank');
      return;
    }
    try {
      setIsOpeningPreview(true);
      const host = selectedFile.url.match(/^(https?:\/\/[^/]+)/i)?.[1] || '';
      // Fetch the perfect native SharePoint webUrl for any file type dynamically (O(1) API call, only on click)
      const res = await fetch(`${host}/_api/v2.0/drives/${selectedFile.driveId}/items/${selectedFile.itemId}?$select=webUrl`);
      if (res.ok) {
        const data = await res.json();
        if (data && data.webUrl) {
          let finalUrl = data.webUrl;
          
          // Images should NOT get ?web=1 (causes 404). 
          // Videos MUST get ?web=1 to open in the native Stream Web Player instead of downloading.
          const isImage = ['.png', '.jpg', '.jpeg', '.gif', '.svg'].some(ext => finalUrl.toLowerCase().split('?')[0].endsWith(ext));
          
          if (!isImage && !finalUrl.includes('web=1')) {
            finalUrl = finalUrl.includes('?') ? `${finalUrl}&web=1` : `${finalUrl}?web=1`;
          }
          
          window.open(finalUrl, '_blank');
          return;
        }
      }
    } catch (e) {
      // fallback
    } finally {
      setIsOpeningPreview(false);
    }
    // Fallback if API fails
    window.open(selectedFile.url, '_blank');
  };

  const handleDownload = () => {
    if (!selectedFile.driveId || !selectedFile.itemId) {
      window.open(selectedFile.url, '_blank');
      return;
    }
    const host = selectedFile.url.match(/^(https?:\/\/[^/]+)/i)?.[1] || '';
    // Native SharePoint v2.0 content endpoint guarantees a pure file download stream
    window.open(`${host}/_api/v2.0/drives/${selectedFile.driveId}/items/${selectedFile.itemId}/content`, '_blank');
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
              {selectedFile.siteName ? selectedFile.siteName.replace(/([a-z])([A-Z])/g, '$1 $2') : 'SharePoint Site'}
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
              {selectedFile.type === 'FOLDER' || selectedFile.type === 'folder'
                ? '📁 Folder — Click to explore contents'
                : (selectedFile.description && selectedFile.description !== 'No description preview available.'
                   ? <>"{renderFormattedSummary(selectedFile.description, searchQuery, selectedFile.color)}"</>
                   : 'No preview available for this file.'
                  )
              }
            </p>
          </div>
        </div>

        {/* Action Footer */}
        <div className={styles.actionFooter}>
          <button 
            className={styles.actionPrimaryBtn}
            style={{ backgroundColor: selectedFile.color, boxShadow: `0 10px 20px -5px ${selectedFile.color}40` }}
            onClick={handleFullPreview}
            disabled={isOpeningPreview}
          >
            <ExternalLink size={14} />
            {isOpeningPreview ? 'OPENING...' : 'FULL PREVIEW'}
          </button>
          <button 
            className={styles.actionSecondaryBtn} 
            title="Download Asset"
            onClick={handleDownload}
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
