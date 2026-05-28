import * as React from 'react';
import { X, FileText, GripVertical } from 'lucide-react';
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
    <aside
      className={styles.previewPane}
      style={{ width: previewWidth }}
    >
      <div 
        onMouseDown={startResizingPreview}
        className={`${styles.resizeHandle} ${styles.previewResizeHandle} ${isResizingPreview ? styles.resizeHandleActive : ''}`}
      >
        <div className={styles.resizeGrip}>
          <GripVertical size={12} strokeWidth={2.5} />
        </div>
      </div>

      {/* Elegant Header */}
      <div className={styles.previewHeader}>
        <div className={styles.previewHeaderContainer}>
          <div 
            className={styles.previewHeaderIconBox}
            style={{ 
              color: selectedFile.color, 
              backgroundColor: selectedFile.color + '15'
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
          title="Close preview"
          aria-label="Close preview"
        >
          <X size={18} />
        </button>
      </div>

      {/* Embedded Metadata Detail Panel */}
      <DetailPanel selectedFile={selectedFile} searchQuery={searchQuery} />
    </aside>
  );
};
