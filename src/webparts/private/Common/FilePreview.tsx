import * as React from 'react';
import { Modal } from '@fluentui/react/lib/Modal';
import { IconButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Icon } from '@fluentui/react/lib/Icon';
import { FiFile } from 'react-icons/fi';
import * as pdfjsLib from 'pdfjs-dist';
// eslint-disable-next-line @typescript-eslint/no-var-requires
pdfjsLib.GlobalWorkerOptions.workerSrc = require('pdfjs-dist/build/pdf.worker.entry');

/**
 * Canvas renderer for .ai / .eps / .dn files.
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
                const scale = 800 / viewport.width;
                const scaledViewport = page.getViewport({ scale });
                const canvas = canvasRef.current!;
                const context = canvas.getContext('2d')!;
                canvas.width = scaledViewport.width;
                canvas.height = scaledViewport.height;
                await page.render({ canvasContext: context, viewport: scaledViewport }).promise;
            } catch (e) {
                setHasError(true);
            }
        };
        render();
    }, [fileUrl]);

    if (hasError) {
        return (
            <div style={{ textAlign: 'center', padding: '40px', color: '#94a3af' }}>
                <Icon iconName="Error" style={{ fontSize: '32px' }} />
                <Text variant="medium" block style={{ marginTop: '12px' }}>AI Preview requires PDF compatibility</Text>
            </div>
        );
    }

    return (
        <div style={{ display: 'flex', justifyContent: 'center', padding: '20px' }}>
            <canvas ref={canvasRef} style={{ maxWidth: '100%', height: 'auto', borderRadius: '4px', boxShadow: '0 4px 12px rgba(0,0,0,0.1)' }} />
        </div>
    );
};

const OFFICE_EXTENSIONS = ['doc', 'docx', 'xls', 'xlsx', 'xlsm', 'ppt', 'pptx', 'rtf', 'msg', 'eml'];
const TEXT_EXTENSIONS = ['txt', 'csv', 'json', 'xml', 'js', 'ts', 'tsx', 'jsx', 'css', 'html', 'cs', 'java', 'py', 'sql', 'yaml', 'yml'];
const VIDEO_EXTENSIONS = ['mp4', 'avi', 'mov', 'wmv', 'flv'];
const AUDIO_EXTENSIONS = ['mp3', 'wav'];

export interface IFilePreviewProps {
    file: {
        name: string;
        url: string;
        serverRelativeUrl: string;
    };
    onClose: () => void;
}

const FilePreview: React.FC<IFilePreviewProps> = ({ file, onClose }) => {
    const getFileType = (fileName: string) => {
        const ext = fileName.split('.').pop()?.toLowerCase();
        if (['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg', 'webp'].includes(ext)) return 'image';
        if (ext === 'pdf') return 'pdf';
        if (OFFICE_EXTENSIONS.includes(ext || '')) return 'office';
        if (TEXT_EXTENSIONS.includes(ext || '')) return 'text';
        if (VIDEO_EXTENSIONS.includes(ext || '')) return 'video';
        if (AUDIO_EXTENSIONS.includes(ext || '')) return 'audio';
        if (['ai', 'eps', 'dn'].includes(ext)) return 'ai';
        return 'unknown';
    };

    const fileType = getFileType(file.name);

    const renderContent = () => {
        switch (fileType) {
            case 'image':
                return (
                    <div style={{ display: 'flex', justifyContent: 'center', padding: '20px' }}>
                        <img src={file.url} alt={file.name} style={{ maxWidth: '100%', maxHeight: '70vh' }} />
                    </div>
                );
            case 'pdf':
            case 'office':
            case 'text':
            case 'video':
                const previewUrl = fileType === 'pdf' 
                    ? file.url 
                    : `${window.location.origin}/sites/Projects/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(file.serverRelativeUrl)}&action=embedview`;
                return (
                    <iframe src={previewUrl} style={{ width: '100%', height: '70vh', border: 'none' }} title={file.name} />
                );
            case 'audio':
                return (
                    <div style={{ padding: '24px' }}>
                        <Text variant="mediumPlus" block style={{ marginBottom: '12px' }}>Audio Preview</Text>
                        <audio controls style={{ width: '100%' }}>
                            <source src={file.url} />
                            Your browser does not support audio playback.
                        </audio>
                    </div>
                );
            case 'ai':
                return <AiCanvasPreview fileUrl={file.url} />;
            default:
                return (
                    <div style={{ padding: '60px', textAlign: 'center' }}>
                        {React.createElement(FiFile as React.ComponentType<any>, {
                            style: { fontSize: '48px', color: '#94a3b8' },
                            'aria-hidden': true,
                        })}
                        <Text variant="large" block style={{ marginTop: '16px' }}>Preview not available for this file type</Text>
                        <a href={file.url} target="_blank" rel="noopener noreferrer" style={{ display: 'inline-block', marginTop: '12px', color: '#3b82f6' }}>Open in new tab</a>
                    </div>
                );
        }
    };

    return (
        <Modal
            isOpen={true}
            onDismiss={onClose}
            isBlocking={false}
            containerClassName="preview-modal-container"
            styles={{ main: { width: '90%', maxWidth: '1000px', borderRadius: '12px' } }}
        >
            <div style={{ padding: '16px 24px', borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <Text variant="large" style={{ fontWeight: 600 }}>{file.name}</Text>
                <IconButton iconProps={{ iconName: 'Cancel' }} onClick={onClose} />
            </div>
            <div style={{ overflowY: 'auto' }}>
                {renderContent()}
            </div>
        </Modal>
    );
};

export default FilePreview;
