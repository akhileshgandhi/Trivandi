import * as React from 'react';
import { Modal } from '@fluentui/react/lib/Modal';
import { IconButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './FilePreview.module.scss';

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
    
    console.log('FilePreview: Constructor - file type:', this.state.fileType, 'for', props.fileName);
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
    console.log('FilePreview: Detecting file type for:', fileName, 'extension:', extension);
    
    // Image files
    if (['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg', 'webp'].includes(extension)) {
      console.log('FilePreview: Detected as image');
      return 'image';
    }
    
    // PDF files
    if (extension === 'pdf') {
      console.log('FilePreview: Detected as PDF');
      return 'pdf';
    }
    
    // Office documents
    if (['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx'].includes(extension)) {
      console.log('FilePreview: Detected as Office document');
      return 'office';
    }
    
    // Text files
    if (['txt', 'csv', 'json', 'xml', 'js', 'ts', 'tsx', 'css', 'scss', 'html'].includes(extension)) {
      console.log('FilePreview: Detected as text');
      return 'text';
    }
    
    // Video files
    if (['mp4', 'avi', 'mov', 'wmv', 'flv'].includes(extension)) {
      console.log('FilePreview: Detected as video');
      return 'video';
    }
    
    console.log('FilePreview: Unknown file type');
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
    
    console.log('FilePreview: Building preview URL', { fileUrl, fileType, siteUrl });

    if (!siteUrl) return fileUrl;

    // For Office docs + PDF - use SharePoint's built-in WopiFrame viewer
    if (['office', 'pdf'].includes(fileType)) {
      // fileUrl should be server-relative like /sites/.../file.docx
      const serverRelativeUrl = fileUrl.startsWith('http') 
        ? new URL(fileUrl).pathname  // extract path from full URL
        : fileUrl;
      
      const previewUrl = `${siteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(serverRelativeUrl)}&action=embedview`;
      console.log('FilePreview: WopiFrame URL:', previewUrl);
      return previewUrl;
    }

    return fileUrl;
  }

  private _renderPreviewContent = (): JSX.Element => {
    const { fileUrl, fileName } = this.props;
    const { loading, error } = this.state;
    // Always compute fileType fresh from current fileName (avoids stale state)
    const fileType = this._getFileType(fileName);

    console.log('FilePreview: Rendering content for', fileName, 'type:', fileType, 'loading:', loading);

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
      case 'image':
        return (
          <div className={styles.imageContainer}>
            <img 
              src={fileUrl} 
              alt={fileName}
              className={styles.previewImage}
              onLoad={() => {
                console.log('Image loaded:', fileName);
              }}
              onError={() => {
                console.error('Image failed to load:', fileName);
                this.setState({ error: `Cannot display image: ${fileName}` });
              }}
            />
          </div>
        );

      case 'pdf': {
        const pdfPreviewUrl = this._buildPreviewUrl(fileUrl, 'pdf');
        const readOnly = this.props.canDownload === false;
        console.log('PDF preview URL:', pdfPreviewUrl);
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
        console.log('Office WopiFrame URL:', officePreviewUrl);
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
                console.log('Text file loaded:', fileName);
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

    console.log('FilePreview: Rendering modal', { 
      isOpen, 
      fileName, 
      fileUrl,
      hasError: !!error
    });

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