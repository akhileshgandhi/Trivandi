import * as React from 'react';
import { Modal } from '@fluentui/react/lib/Modal';
import { IconButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import { FiFile, FiFileText, FiFilePlus, FiVideo, FiMusic, FiImage } from 'react-icons/fi';
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

const OFFICE_EXTENSIONS = ['doc', 'docx', 'xls', 'xlsx', 'xlsm', 'ppt', 'pptx', 'rtf', 'msg', 'eml'];
const TEXT_EXTENSIONS = ['txt', 'csv', 'json', 'xml', 'js', 'ts', 'tsx', 'jsx', 'css', 'scss', 'html', 'cs', 'java', 'py', 'sql', 'yaml', 'yml'];
const VIDEO_EXTENSIONS = ['mp4', 'avi', 'mov', 'wmv', 'flv'];
const AUDIO_EXTENSIONS = ['mp3', 'wav'];

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

  private _renderReactIcon = (IconComponent: React.ComponentType<any>): JSX.Element => {
    return React.createElement(IconComponent, { 'aria-hidden': true });
  }

  constructor(props: IFilePreviewProps) {
    super(props);

    this.state = {
      loading: false, // Never start with loading
      fileType: this._getFileType(props.fileName)
    };


  }

  private _getFileIconByExtension = (fileName: string): JSX.Element => {
    const extension = fileName.split('.').pop()?.toLowerCase() || '';
    if (['doc', 'docx'].includes(extension)) return this._renderReactIcon(FiFileText as React.ComponentType<any>);
    if (['xls', 'xlsx', 'xlsm', 'csv'].includes(extension)) return this._renderReactIcon(FiFilePlus as React.ComponentType<any>);
    if (['ppt', 'pptx'].includes(extension)) return this._renderReactIcon(FiFile as React.ComponentType<any>);
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
    if (OFFICE_EXTENSIONS.includes(extension)) {

      return 'office';
    }

    // Text files
    if (TEXT_EXTENSIONS.includes(extension)) {

      return 'text';
    }

    // Video files
    if (VIDEO_EXTENSIONS.includes(extension)) {

      return 'video';
    }

    // Audio files
    if (AUDIO_EXTENSIONS.includes(extension)) {
      return 'audio';
    }

    // AI / EPS / DN files — PDF.js canvas rendering
    if (['ai', 'eps', 'dn'].includes(extension)) {
      return 'ai';
    }


    return 'unknown';
  }

  private _getFileIcon = (fileType: string): JSX.Element => {
    switch (fileType) {
      case 'image': return this._renderReactIcon(FiImage as React.ComponentType<any>);
      case 'pdf': return this._renderReactIcon(FiFileText as React.ComponentType<any>);
      case 'office': return this._renderReactIcon(FiFileText as React.ComponentType<any>);
      case 'text': return this._renderReactIcon(FiFileText as React.ComponentType<any>);
      case 'video': return this._renderReactIcon(FiVideo as React.ComponentType<any>);
      case 'audio': return this._renderReactIcon(FiMusic as React.ComponentType<any>);
      default: return this._renderReactIcon(FiFile as React.ComponentType<any>);
    }
  }

  private _buildPreviewUrl = (
    fileUrl: string,
    fileType: string,
    action: 'embedview' | 'edit' = 'embedview'
  ): string => {
    const { siteUrl } = this.props;

    if (!siteUrl) return fileUrl;

    // For Office docs - use SharePoint's built-in WopiFrame viewer so they can view directly without auto-downloading
    if (['office'].includes(fileType)) {
      const serverRelativeUrl = fileUrl.startsWith('http')
        ? new URL(fileUrl).pathname  // extract path from full URL
        : fileUrl;

      const trimmedSite = siteUrl.replace(/\/$/, '');
      const previewUrl = action === 'edit'
        ? `${trimmedSite}/_layouts/15/Doc.aspx?sourcedoc=${encodeURIComponent(serverRelativeUrl)}&action=edit`
        : `${trimmedSite}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(serverRelativeUrl)}&action=embedview`;

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

      case 'image': {
        const absoluteUrl = fileUrl.startsWith('http')
          ? fileUrl
          : `${window.location.origin}${fileUrl}`;
        return (
          <div className={styles.imageContainer}>
            <img
              src={absoluteUrl}
              alt={fileName}
              className={styles.previewImage}
              onError={() => {
                this.setState({ error: `Cannot display image: ${fileName}` });
              }}
            />
          </div>
        );
      }

      case 'pdf': {
        const absoluteUrl = fileUrl.startsWith('http')
          ? fileUrl
          : `${window.location.origin}${fileUrl}`;

        return (
          <div className={styles.pdfContainer}>
            <iframe
              src={absoluteUrl}
              className={styles.previewFrame}
              title={fileName}
            />
          </div>
        );
      }

      case 'office': {
        const readOnly = this.props.canDownload === false;
        const officePreviewUrl = this._buildPreviewUrl(
          fileUrl,
          'office',
          readOnly ? 'embedview' : 'edit'
        );

        return (
          <div className={styles.officeContainer}>
            <iframe
              src={officePreviewUrl}
              className={styles.previewFrame}
              title={fileName}
              {...(readOnly ? { sandbox: 'allow-scripts allow-same-origin allow-forms' } : {})}
            />
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

      case 'audio': {
        const absoluteUrl = fileUrl.startsWith('http')
          ? fileUrl
          : `${window.location.origin}${fileUrl}`;

        return (
          <div className={styles.unsupportedContainer}>
            <Icon iconName="MusicInCollection" className={styles.fileIcon} />
            <Text variant="mediumPlus">Audio Preview</Text>
            <audio controls style={{ width: '100%', maxWidth: 720 }}>
              <source src={absoluteUrl} />
              Your browser does not support audio playback.
            </audio>
          </div>
        );
      }

      default:
        return (
          <div className={styles.unsupportedContainer}>
            <span className={styles.fileIcon}>{this._getFileIconByExtension(fileName)}</span>
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
    const fileType = this._getFileType(fileName);
    const openInNewTabUrl = (fileType === 'office')
      ? this._buildPreviewUrl(fileUrl, 'office', canDownload ? 'edit' : 'embedview')
      : fileUrl.startsWith('http') ? fileUrl : `${window.location.origin}${fileUrl}`;

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
                <span className={styles.fileTypeIcon}>{fileIcon}</span>
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
              onClick={() => {
                const url = openInNewTabUrl.startsWith('http')
                  ? openInNewTabUrl
                  : `${window.location.origin}${openInNewTabUrl}`;
                window.open(url, '_blank', 'noopener,noreferrer');
              }}
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