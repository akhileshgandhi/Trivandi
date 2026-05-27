import * as React from 'react';
import { X, FileText } from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';
import { IDocumentAnalysisProps } from '../interface/IDocumentAnalysisProps';
import { DetailPanel } from './DetailPanel';

export const PreviewPane: React.FC<IDocumentAnalysisProps> = ({
  selectedFile,
  setSelectedFile,
  previewWidth,
  searchQuery,
  isResizingPreview,
  startResizingPreview
}) => {
  if (!selectedFile) return null;

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
        {/* Elegant Header */}
        <div className={styles.previewHeader} style={{ height: '64px', borderBottom: '1px solid #f1f5f9', display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '0 24px', background: '#ffffff' }}>
          <div className={styles.previewHeaderContainer}>
            <div 
              style={{ 
                color: selectedFile.color, 
                backgroundColor: selectedFile.color + '15',
                padding: '8px', 
                borderRadius: '8px',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center'
              }}
            >
              <FileText size={20} />
            </div>
            <div>
              <h3 className={styles.previewHeaderTitle}>Document Preview</h3>
              <p className={styles.previewHeaderSubtitle}>Deep Analysis Mode</p>
            </div>
          </div>
          <button 
            onClick={() => setSelectedFile(null)}
            className={styles.previewCloseBtn}
            style={{ 
              padding: '8px',
              color: '#94a3b8',
              background: 'transparent',
              border: 'none',
              borderRadius: '50%',
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              transition: 'all 0.15s ease'
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.backgroundColor = '#f1f5f9';
              e.currentTarget.style.color = '#0f172a';
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.backgroundColor = 'transparent';
              e.currentTarget.style.color = '#94a3b8';
            }}
          >
            <X size={18} />
          </button>
        </div>

        {/* Embedded Metadata Detail Panel */}
        <DetailPanel selectedFile={selectedFile} searchQuery={searchQuery} />
      </aside>
    </>
  );
};
