import * as React from 'react';
import { Modal } from '@fluentui/react/lib/Modal';
import { IconButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './FilePreview.module.scss';
import * as pdfjsLib from 'pdfjs-dist';
// eslint-disable-next-line @typescript-eslint/no-var-requires
pdfjsLib.GlobalWorkerOptions.workerSrc = require('pdfjs-dist/build/pdf.worker.entry');

/**
 * Canvas renderer for .ai / .eps files.
 * Adobe Illustrator files embed PDF data — PDF.js can render them directly.
 */
const AiCanvasPreview: React.FC<{ fileUrl: string }> = ({ fileUrl }) => {
  const canvasRef = React.useRef<HTMLCanvasElement | null>(null);
  const [hasError, setHasError] = React.useState(false);

  React.useEffect(() => {
    if (!fileUrl) return;

    const render = async () => {
      try {
        const loadingTask = pdfjsLib.getDocument(fileUrl);
        const pdf = await loadingTask.promise;
        const page = await pdf.getPage(1);

        const viewport = page.getViewport({ scale: 1 });
        const scale = 600 / viewport.width;
        const scaledViewport = page.getViewport({ scale });

        const canvas = canvasRef.current!;
        const context = canvas.getContext('2d')!;
        canvas.width = scaledViewport.width;
        canvas.height = scaledViewport.height;

        await page.render({ canvasContext: context, viewport: scaledViewport }).promise;
      } catch (e) {
        console.error('[AiCanvasPreview] Error:', e);
        setHasError(true);
      }
    };

    render();
  }, [fileUrl]);

  if (hasError) {
    return (
      <div style={{ textAlign: 'center', padding: '48px 24px', color: '#9ca3af' }}>
        <div style={{ fontSize: 48, marginBottom: 12 }}>🎨</div>
        <div style={{ fontWeight: 600, fontSize: 16, color: '#374151', marginBottom: 8 }}>Preview not available</div>
        <div style={{ fontSize: 13 }}>Save the AI file with &quot;Create PDF Compatible File&quot; enabled in Illustrator, then re-upload.</div>
      </div>
    );
  }

  return (
    <div style={{ textAlign: 'center', padding: 8 }}>
      <canvas ref={canvasRef} style={{ maxWidth: '100%', height: 'auto', borderRadius: 4, boxShadow: '0 2px 12px rgba(0,0,0,0.1)' }} />
    </div>
  );
};

export interface IFilePreviewProps {
  isOpen: boolean;
  onDismiss: () => void;
  fileUrl: string;     // Server-relative URL e.g. /sites/Projects/ExternalShareDocument/file.docx
  fileName: string;
  filePath?: string;
  siteUrl?: string;    // Absolute site URL e.g. https://tenant.sharepoint.com/sites/Projects
  canDownload?: boolean; // When false, hides the download and open-in-new-tab buttons (Read permission guests)
}

export interface IFilePreviewState {
  loading: boolean;
  error?: string;
  fileType?: string;
}

export class FilePreview extends React.Component<IFilePreviewProps, IFilePreviewState> {
  
  constructor(props: IFilePreviewProps) {
    super(props);
    
    this.state = {
      loading: false, // Never start with loading
      fileType: this._getFileType(props.fileName)
    };
    
    
  }

  private _getFileIconByExtension = (fileName: string): string => {
    const extension = fileName.split('.').pop()?.toLowerCase() || '';
    if (['doc', 'docx'].includes(extension)) return 'WordDocument';
    if (['xls', 'xlsx'].includes(extension)) return 'ExcelDocument';  
    if (['ppt', 'pptx'].includes(extension)) return 'PowerPointDocument';
    return this._getFileIcon(this._getFileType(fileName));
  }

  private _getFileType = (fileName: string): string => {
    const extension = fileName.split('.').pop()?.toLowerCase() || '';
    
    
    // Image files
    if (['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg', 'webp'].includes(extension)) {
      
      return 'image';
    }
    
    // PDF files
    if (extension === 'pdf') {
      
      return 'pdf';
    }
    
    // Office documents
    if (['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx'].includes(extension)) {
      
      return 'office';
    }
    
    // Text files
    if (['txt', 'csv', 'json', 'xml', 'js', 'ts', 'tsx', 'css', 'scss', 'html'].includes(extension)) {
      
      return 'text';
    }
    
    // Video files
    if (['mp4', 'avi', 'mov', 'wmv', 'flv'].includes(extension)) {
      
      return 'video';
    }
    
    // AI / EPS / DN files — PDF.js canvas rendering
    if (['ai', 'eps', 'dn'].includes(extension)) {
      return 'ai';
    }

    
    return 'unknown';
  }

  private _getFileIcon = (fileType: string): string => {
    switch (fileType) {
      case 'image': return 'FileImage';
      case 'pdf': return 'PDF';
      case 'office': return 'WordDocument';
      case 'text': return 'TextDocument';
      case 'video': return 'Video';
      default: return 'Document';
    }
  }

  private _buildPreviewUrl = (fileUrl: string, fileType: string): string => {
    const { siteUrl } = this.props;
    
    

    if (!siteUrl) return fileUrl;

    // For Office docs + PDF - use SharePoint's built-in WopiFrame viewer
    if (['office', 'pdf'].includes(fileType)) {
      // fileUrl should be server-relative like /sites/.../file.docx
      const serverRelativeUrl = fileUrl.startsWith('http') 
        ? new URL(fileUrl).pathname  // extract path from full URL
        : fileUrl;
      
      const previewUrl = `${siteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(serverRelativeUrl)}&action=embedview`;
      
      return previewUrl;
    }

    return fileUrl;
  }

  private _renderPreviewContent = (): JSX.Element => {
    const { fileUrl, fileName } = this.props;
    const { loading, error } = this.state;
    // Always compute fileType fresh from current fileName (avoids stale state)
    const fileType = this._getFileType(fileName);

    

    if (error) {
      return (
        <div className={styles.errorContainer}>
          <Icon iconName="Error" className={styles.errorIcon} />
          <Text variant="mediumPlus">Failed to load file preview</Text>
          <Text variant="small">{error}</Text>
        </div>
      );
    }

    // Show content immediately for most file types
    switch (fileType) {
      case 'ai': {
        // Use absolute URL for PDF.js fetch
        const absoluteUrl = fileUrl.startsWith('http')
          ? fileUrl
          : `${window.location.origin}${fileUrl}`;
        return (
          <div className={styles.imageContainer}>
            <AiCanvasPreview fileUrl={absoluteUrl} />
          </div>
        );
      }

      case 'image':
        return (
          <div className={styles.imageContainer}>
            <img 
              src={fileUrl} 
              alt={fileName}
              className={styles.previewImage}
              onLoad={() => {
                
              }}
              onError={() => {
                
                this.setState({ error: `Cannot display image: ${fileName}` });
              }}
            />
          </div>
        );

      case 'pdf': {
        const pdfPreviewUrl = this._buildPreviewUrl(fileUrl, 'pdf');
        const readOnly = this.props.canDownload === false;
        
        return (
          <div className={styles.pdfContainer}>
            <iframe
              src={pdfPreviewUrl}
              className={styles.previewFrame}
              title={fileName}
              // Read-only: deny popups, top-navigation and downloads.
              // allow-forms is required — WopiFrame uses form POST to authenticate + load the document.
              // Omitting allow-popups blocks window.open() (expand/zoom button, Download a Copy).
              {...(readOnly ? { sandbox: 'allow-scripts allow-same-origin allow-forms' } : {})}
            />
            {/* Blocker sits over the bottom toolbar (where the … button lives) */}
            {readOnly && (
              <div
                className={styles.iframeBlocker}
                title="Download and print are disabled in read-only view"
              />
            )}
          </div>
        );
      }

      case 'office': {
        const officePreviewUrl = this._buildPreviewUrl(fileUrl, 'office');
        const readOnly = this.props.canDownload === false;
        
        return (
          <div className={styles.officeContainer}>
            <iframe
              src={officePreviewUrl}
              className={styles.previewFrame}
              title={fileName}
              // Read-only: deny popups, top-navigation and downloads.
              // allow-forms is required — WopiFrame uses form POST to authenticate + load the document.
              // Omitting allow-popups blocks window.open() (expand/zoom button, Download a Copy).
              {...(readOnly ? { sandbox: 'allow-scripts allow-same-origin allow-forms' } : {})}
            />
            {/* Blocker sits over the bottom toolbar (where the … button lives) */}
            {readOnly && (
              <div
                className={styles.iframeBlocker}
                title="Download and print are disabled in read-only view"
              />
            )}
          </div>
        );
      }

      case 'text':
        return (
          <div className={styles.textContainer}>
            <iframe
              src={fileUrl}
              className={styles.previewFrame}
              title={fileName}
              onLoad={() => {
                
              }}
            />
          </div>
        );

      default:
        return (
          <div className={styles.unsupportedContainer}>
            <Icon iconName={this._getFileIconByExtension(fileName)} className={styles.fileIcon} />
            <Text variant="mediumPlus">Preview not available</Text>
            <Text variant="small">File: {fileName}</Text>
            <Text variant="small">File type: {this.state.fileType}</Text>
            {this.props.canDownload !== false && (
              <div className={styles.downloadContainer}>
                <IconButton
                  iconProps={{ iconName: 'OpenInNewWindow' }}
                  text="Open in New Tab"
                  href={fileUrl}
                  target="_blank"
                  className={styles.downloadButton}
                />
                <IconButton
                  iconProps={{ iconName: 'Download' }}
                  text="Download File"
                  href={fileUrl}
                  target="_blank"
                  className={styles.downloadButton}
                />
              </div>
            )}
          </div>
        );
    }
  }

  public render(): React.ReactElement<IFilePreviewProps> {
    const { isOpen, onDismiss, fileName, filePath, fileUrl, canDownload = true } = this.props;
    const { error } = this.state;

    

    if (!isOpen) {
      return null;
    }

    const fileIcon = this._getFileIconByExtension(fileName);

    return (
      <Modal
        isOpen={isOpen}
        onDismiss={onDismiss}
        isBlocking={false}
        containerClassName={styles.modalContainer}
      >
        {/* ── Header ── */}
        <div className={styles.header}>
          <Stack horizontal verticalAlign="center" className={styles.headerContent}>
            <Stack.Item grow>
              <div className={styles.fileInfo}>
                <Icon iconName={fileIcon} className={styles.fileTypeIcon} />
                <div className={styles.fileTextGroup}>
                  <Text className={styles.fileName} title={fileName}>{fileName}</Text>
                  {filePath && (
                    <Text className={styles.filePath} title={filePath}>{filePath}</Text>
                  )}
                </div>
              </div>
            </Stack.Item>
            <Stack.Item>
              <IconButton
                iconProps={{ iconName: 'ChromeClose' }}
                ariaLabel="Close"
                onClick={onDismiss}
                className={styles.closeButton}
              />
            </Stack.Item>
          </Stack>
        </div>

        {/* ── Read-only notice bar ── */}
        {!canDownload && (
          <div className={styles.readOnlyBar}>
            <Icon iconName="Lock" />
            Read-only view &mdash; download and print are disabled for your access level.
          </div>
        )}

        {/* ── Preview Content ── */}
        <div className={styles.content}>
          {this._renderPreviewContent()}
        </div>

        {/* ── Footer ── */}
        <div className={styles.footer}>
          {canDownload && (
            <IconButton
              iconProps={{ iconName: 'OpenInNewWindow' }}
              text="Open in New Tab"
              href={fileUrl}
              target="_blank"
              className={styles.actionButton}
            />
          )}
          <Text className={styles.footerLabel}>{fileName}</Text>
          {canDownload && (
            <IconButton
              iconProps={{ iconName: 'Download' }}
              text="Download"
              href={fileUrl}
              target="_blank"
              download={fileName}
              className={styles.actionButton}
            />
          )}
        </div>
      </Modal>
    );
  }
}

export default FilePreview;