import * as React from 'react';
import { useState, useEffect } from 'react';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { getDocumentsByServerRelativeUrl } from '../../services/projectService';
import styles from './MoveCopyModal.module.scss';

export interface MoveCopyModalProps {
  isOpen: boolean;
  onClose: () => void;
  onConfirm: (destinationPath: string) => void;
  title: string;
  sourceItemName: string;
  initialPath: string;
  libraryName: string;
}

interface FolderItem {
  Name: string;
  ServerRelativeUrl: string;
}

const MoveCopyModal: React.FC<MoveCopyModalProps> = ({
  isOpen,
  onClose,
  onConfirm,
  title,
  sourceItemName,
  initialPath,
  libraryName
}) => {
  const [currentPath, setCurrentPath] = useState(initialPath);
  const [folders, setFolders] = useState<FolderItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [breadcrumbs, setBreadcrumbs] = useState<{ name: string; path: string }[]>([]);

  useEffect(() => {
    if (isOpen) {
      loadFolders(initialPath);
      // Build simple breadcrumbs - just show folder hierarchy, not full server path
      const parts = initialPath.split('/').filter(Boolean);
      
      // Extract just the folder name from the current path (last part)
      const currentFolderName = parts.length > 0 ? parts[parts.length - 1] : 'Root';
      
      // Create breadcrumbs: Root > CurrentFolder
      const crumbs = [
        { name: 'Root', path: initialPath.split('/').slice(0, -1).join('/') || initialPath },
        { name: currentFolderName, path: initialPath }
      ];
      
      setBreadcrumbs(crumbs);
      setCurrentPath(initialPath);
    }
  }, [isOpen, initialPath]);

  const loadFolders = async (path: string) => {
    setLoading(true);
    try {
      const result = await getDocumentsByServerRelativeUrl(libraryName, path, 1, 100);
      const folderItems = result.items
        .filter(item => item.FSObjType === 1)
        .map(item => ({
          Name: item.FileLeafRef,
          ServerRelativeUrl: item.FileRef || item.ServerRelativeUrl
        }));
      setFolders(folderItems);
    } catch (error) {
      console.error('Error loading folders:', error);
    } finally {
      setLoading(false);
    }
  };

  const handleFolderClick = (folder: FolderItem) => {
    const newPath = folder.ServerRelativeUrl;
    setCurrentPath(newPath);
    setBreadcrumbs([...breadcrumbs, { name: folder.Name, path: newPath }]);
    loadFolders(newPath);
  };

  const handleBreadcrumbClick = (path: string, index: number) => {
    setCurrentPath(path);
    setBreadcrumbs(breadcrumbs.slice(0, index + 1));
    loadFolders(path);
  };

  return (
    <Dialog
      hidden={!isOpen}
      onDismiss={onClose}
      dialogContentProps={{
        type: DialogType.close,
        title: `${title} "${sourceItemName}"`,
      }}
      modalProps={{ isBlocking: true }}
      minWidth={500}
    >
      <div className={styles.modalContent}>
        <div className={styles.breadcrumbBar}>
          {breadcrumbs.map((crumb, index) => (
            <React.Fragment key={crumb.path}>
              <span 
                className={styles.breadcrumbItem} 
                onClick={() => handleBreadcrumbClick(crumb.path, index)}
                title={crumb.path}
              >
                {crumb.name === 'Root' ? (
                  <Icon iconName="Home" className={styles.homeIcon} />
                ) : (
                  crumb.name
                )}
              </span>
              {index < breadcrumbs.length - 1 && <span className={styles.separator}>/</span>}
            </React.Fragment>
          ))}
        </div>

        <div className={styles.folderList}>
          {loading ? (
            <Spinner size={SpinnerSize.large} label="Loading folders..." />
          ) : folders.length === 0 ? (
            <div className={styles.emptyState}>No subfolders found</div>
          ) : (
            folders.map(folder => (
              <div 
                key={folder.ServerRelativeUrl} 
                className={styles.folderRow}
                onClick={() => handleFolderClick(folder)}
              >
                <Icon iconName="Folder" className={styles.folderIcon} />
                <span>{folder.Name}</span>
              </div>
            ))
          )}
        </div>
      </div>
      <DialogFooter>
        <PrimaryButton onClick={() => onConfirm(currentPath)} text={`${title} here`} />
        <DefaultButton onClick={onClose} text="Cancel" />
      </DialogFooter>
    </Dialog>
  );
};

export default MoveCopyModal;
