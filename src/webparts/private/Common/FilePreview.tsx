import * as React from 'react';
import { Modal } from '@fluentui/react/lib/Modal';
import { IconButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Icon } from '@fluentui/react/lib/Icon';

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
        if (['jpg', 'jpeg', 'png', 'gif', 'svg'].includes(ext)) return 'image';
        if (ext === 'pdf') return 'pdf';
        if (['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx'].includes(ext)) return 'office';
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
                const previewUrl = fileType === 'pdf' 
                    ? file.url 
                    : `${window.location.origin}/sites/Projects/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(file.serverRelativeUrl)}&action=embedview`;
                return (
                    <iframe src={previewUrl} style={{ width: '100%', height: '70vh', border: 'none' }} title={file.name} />
                );
            default:
                return (
                    <div style={{ padding: '60px', textAlign: 'center' }}>
                        <Icon iconName="Document" style={{ fontSize: '48px', color: '#94a3b8' }} />
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
