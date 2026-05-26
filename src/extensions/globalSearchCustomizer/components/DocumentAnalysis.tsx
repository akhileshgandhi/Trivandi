import * as React from 'react';
import { X, FileText, ExternalLink, Download } from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';

import { IDocumentAnalysisProps } from '../interface/IDocumentAnalysisProps';

export const DocumentAnalysis: React.FC<IDocumentAnalysisProps> = ({
  selectedFile,
  setSelectedFile,
  previewWidth,
  searchQuery,
  isResizingPreview,
  startResizingPreview
}) => {
  if (!selectedFile) return null;

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

  return (
    <>
      <div 
        onMouseDown={startResizingPreview}
        className={`${styles.resizeHandle} ${isResizingPreview ? styles.resizeHandleActive : ''}`}
      />

      <aside
        className={styles.previewPane}
        style={{ width: previewWidth }}
      >
        {/* Header Bar */}
        <div className={styles.previewHeader}>
          <span className={styles.previewTitleLabel}>Document Analysis</span>
          <button 
            onClick={() => setSelectedFile(null)}
            className={styles.previewCloseBtn}
          >
            <X size={20} />
          </button>
        </div>

        {/* Scrollable Content */}
        <div className={styles.previewContentScroll}>
          
          {/* Skeletons Visual Preview Card */}
          <div className={styles.previewDocVisualCard}>
            <div className={styles.visualCardTop}>
              <div className={styles.visualPulseLine1} />
              <div className={styles.visualPulseLine2} />
              <div className={styles.visualInnerFrame}>
                <div className={styles.frameGripBackground} />
                <FileText size={64} className={styles.frameIcon} />
              </div>
              <div className={styles.visualPulseList}>
                <div className={`${styles.pulseRow} ${styles.pulseRowMid}`} />
                <div className={`${styles.pulseRow} ${styles.pulseRowShort}`} />
              </div>
            </div>
            <div className={styles.visualCardBottom}>
              <span className={styles.visualRestrictedLabel}>Visual Preview Restricted</span>
            </div>
          </div>

          {/* Details & Metadata */}
          <div>
            <div style={{ marginBottom: '24px' }}>
              <h2 style={{ fontSize: '20px', fontWeight: 700, color: '#0f172a', margin: '0 0 4px', lineHeight: 1.2 }}>
                {selectedFile.title}
              </h2>
              <p style={{ fontSize: '12px', fontWeight: 600, color: '#1a73e8', margin: 0, textOverflow: 'ellipsis', overflow: 'hidden', whiteSpace: 'nowrap' }}>
                {selectedFile.url || 'SharePoint Site'}
              </p>
            </div>
            
            <div className={styles.previewMetaGrid}>
              <div className={styles.previewMetaItem}>
                <p className={styles.previewMetaLabel}>Object Type</p>
                <p className={styles.previewMetaValue}>{selectedFile.type || 'DOC'} Asset</p>
              </div>
              <div className={styles.previewMetaItem}>
                <p className={styles.previewMetaLabel}>Storage Size</p>
                <p className={styles.previewMetaValue}>{selectedFile.size || '1.2 MB'}</p>
              </div>
              <div className={styles.previewMetaItem}>
                <p className={styles.previewMetaLabel}>Creation Date</p>
                <p className={styles.previewMetaValue}>{selectedFile.date || 'Today'}</p>
              </div>
              <div className={styles.previewMetaItem}>
                <p className={styles.previewMetaLabel}>Last Activity</p>
                <p className={styles.previewMetaValue}>Today, 2:14 PM</p>
              </div>
              
              <div className={styles.primaryAuthorCard}>
                <p className={styles.previewMetaLabel}>Primary Author</p>
                <div className={styles.authorCardRow}>
                  <div className={styles.authorAvatarSphere} style={{ backgroundColor: selectedFile.color }}>
                    {selectedFile.author ? selectedFile.author.charAt(0).toUpperCase() : 'U'}
                  </div>
                  <p className={styles.previewMetaValue} style={{ margin: 0 }}>{selectedFile.author || 'SharePoint User'}</p>
                </div>
              </div>
            </div>

            {/* Document Content Extracted Section */}
            <div className={styles.extractedSection}>
              <p className={styles.previewMetaLabel}>Document Content</p>
              <div className={styles.extractedContentBlock}>
                "{highlightText(selectedFile.description, searchQuery)}"
              </div>
            </div>

            {/* Action Row */}
            <div className={styles.previewActionsWrapper}>
              <button 
                className={styles.previewFullBtn}
                onClick={() => window.open(selectedFile.url, '_blank')}
              >
                <ExternalLink size={16} strokeWidth={2.5} />
                Full Preview
              </button>
              <button className={styles.previewDownloadBtn} title="Download Asset">
                <Download size={18} strokeWidth={2} />
              </button>
            </div>

          </div>
        </div>
      </aside>
    </>
  );
};
