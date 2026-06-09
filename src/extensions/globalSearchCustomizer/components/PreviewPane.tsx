import * as React from 'react';
import { X, GripVertical } from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';
import { IDocumentAnalysisProps } from '../interface/IDocumentAnalysisProps';
import { DetailPanel } from './DetailPanel';
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

const getPreviewIcon = (type: string) => {
  const ft = type ? type.toLowerCase() : '';
  if (['png', 'jpg', 'jpeg', 'gif', 'svg'].includes(ft)) {
    return <IconImages size={20} />;
  }
  if (ft === 'xls') {
    return <IconXls size={20} />;
  }
  if (ft === 'xlsx') {
    return <IconXlsx size={20} />;
  }
  if (ft === 'ppt') {
    return <IconPpt size={20} />;
  }
  if (ft === 'pptx') {
    return <IconPptx size={20} />;
  }
  if (ft === 'pdf') {
    return <IconPdf size={20} />;
  }
  if (ft === 'doc') {
    return <IconDoc size={20} />;
  }
  if (ft === 'docx') {
    return <IconDocx size={20} />;
  }
  if (ft === 'folder') {
    return <IconFolder size={20} />;
  }
  if (['mp4', 'mov', 'avi', 'wmv', 'mkv', 'flv', 'webm'].includes(ft)) {
    return <IconPlay size={20} />;
  }
  return <IconText size={20} />;
};

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
            {getPreviewIcon(selectedFile.type)}
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
