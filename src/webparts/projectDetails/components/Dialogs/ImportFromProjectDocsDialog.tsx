import * as React from 'react';
import { useState, useEffect } from 'react';
import { toast } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';
import styles from './ImportFromProjectDocsDialog.module.scss';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { getDocumentsByServerRelativeUrl, getProjectDocuments, getProjectById } from '../../../../shared/services/projectService';
import { findProjectLibraryByProject } from '../../../../shared/services/libraryDiscoveryService';

export interface IImportFromProjectDocsDialogProps {
  isOpen: boolean;
  onClose: () => void;
  onSuccess: () => void;
  context: WebPartContext;
  targetPath: string;
  projectId: string | number;
  projectCode?: string;
  projectTitle?: string;
  businessProjectId?: string;
  service?: any;
}

export interface IProjectDocument {
  Id: number | string;
  FileLeafRef: string;
  FileRef: string;
  Modified: string;
  Editor: { Title: string };
  FSObjType: number;
  selected?: boolean;
  isImported?: boolean;
  ItemCount?: number;
  Length?: number;
  EditorTitle?: string;
}

const ImportFromProjectDocsDialog: React.FC<IImportFromProjectDocsDialogProps> = (props) => {
  const [documents, setDocuments] = useState<IProjectDocument[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [importing, setImporting] = useState<boolean>(false);
  const [currentFolderPath, setCurrentFolderPath] = useState<string>('');
  const [activeDocTab, setActiveDocTab] = useState<string>('project');
  const [libraryName, setLibraryName] = useState<string>('');

  const [projectDocumentsUrl, setProjectDocumentsUrl] = useState<string>('');
  const [bidDocumentsUrl, setBidDocumentsUrl] = useState<string>('');
  const [contractsDocumentsUrl, setContractsDocumentsUrl] = useState<string>('');
  const [existingFiles, setExistingFiles] = useState<Set<string>>(new Set());
  const [hasAutoNavigated, setHasAutoNavigated] = useState<boolean>(false);

  const formatFileSize = (bytes?: number): string => {
    if (bytes === undefined || bytes === 0) return '0 KB';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  // 1. Initial Discovery and Metadata Fetch (Mirroring CustomDocumentsList.tsx)
  useEffect(() => {
    const initializeModal = async (): Promise<void> => {
      try {
        const bizProjId = props.businessProjectId || "";
        if (props.projectCode || bizProjId) {
          const found = await findProjectLibraryByProject(bizProjId || props.projectCode, props.projectTitle);
          if (found) setLibraryName(found);
        }

        const id = Number(props.projectId);
        if (!isNaN(id) && id > 0) {
          const projectData = await getProjectById(id);
          if (projectData) {
            const cleanUrl = (url: any) => {
              if (!url) return '';
              const trimmed = String(url).trim();
              const isValidPath = trimmed.toLowerCase().startsWith('/sites') || trimmed.startsWith('http');
              return (trimmed === '-' || trimmed === '' || !isValidPath) ? '' : trimmed;
            };

            const projUrl = cleanUrl(projectData.ProjectDocumentsUrl || projectData.projectDocumentsUrl);
            const bidUrl = cleanUrl(projectData.BidDocumentsUrl || projectData.bidDocumentsUrl);
            const contractUrl = cleanUrl(projectData.ContractsDocumentsUrl || projectData.contractsDocumentsUrl);

            setProjectDocumentsUrl(projUrl);
            setBidDocumentsUrl(bidUrl);
            setContractsDocumentsUrl(contractUrl);

            // Tab selection logic from CustomDocumentsList
            if (projUrl || libraryName) {
              setActiveDocTab('project');
            } else if (bidUrl) {
              setActiveDocTab('bid');
            } else if (contractUrl) {
              setActiveDocTab('contract');
            }
          }
        }

        // Fetch existing files in target path to mark them as "Imported"
        if (props.service && props.targetPath) {
          try {
            const result = await props.service.getSharedDocumentsFromExternalArea(props.targetPath, 1, 500);
            const names = new Set<string>((result.items || []).map((f: any) => String(f.name || '')));
            setExistingFiles(names);
          } catch (e) {

          }
        }

      } catch (err) {

      }
    };

    if (props.isOpen) {
      void initializeModal();
    }
  }, [props.isOpen, props.projectId, props.projectCode, props.projectTitle]);

  // Load documents flow
  useEffect(() => {
    if (props.isOpen) {
      loadDocuments();
    }
  }, [props.isOpen, activeDocTab, currentFolderPath, libraryName, projectDocumentsUrl]);

  const loadDocuments = async (): Promise<void> => {
    try {
      setLoading(true);
      let targetUrl = '';
      let targetLibrary = '';

      switch (activeDocTab) {
        case 'project': targetUrl = projectDocumentsUrl; targetLibrary = libraryName; break;
        case 'bid': targetUrl = bidDocumentsUrl; break;
        case 'contract': targetUrl = contractsDocumentsUrl; break;
      }

      let result: { items: any[]; totalCount: number } = { items: [], totalCount: 0 };
      if (targetUrl) {
        result = await getDocumentsByServerRelativeUrl(targetUrl, currentFolderPath, 1, 100);
      } else if (targetLibrary && activeDocTab === 'project') {
        result = await getProjectDocuments(targetLibrary, currentFolderPath, 1, 100);
      }

      console.log('Doc=====================>>>>>>>>>>>>>>>>>>>>', result);


      setDocuments(
        (result.items || []).map(item => {
          const name = item.FileLeafRef || item.name;
          const isImported =
            existingFiles.has(name) && item.FSObjType !== 1;

          return {
            ...item,
            FileLeafRef: name,
            EditorTitle: item.Editor?.Title || "",
            selected: false,
            isImported,
          };
        })
      );

      // Multi-level Auto-navigation (Double Jump) - Mirroring CustomDocumentsList.tsx
      if (!hasAutoNavigated && (activeDocTab === "project" || activeDocTab === "bid")) {
        const targetFolders = ["Projects", "Bids", "Lost", "Closed"];
        
        // Level 1: At Root, look for organizational folders
        if (currentFolderPath === "") {
          const orgFolder = (result.items || []).find(item => {
            const name = item.FileLeafRef || item.name || "";
            if (item.FSObjType !== 1) return false;
            const folderName = name.toLowerCase();
            return targetFolders.some(target => 
              folderName === target.toLowerCase() || 
              folderName === (target.toLowerCase() + "s") || 
              folderName.includes(target.toLowerCase())
            );
          });
          
          if (orgFolder) {
            const folderName = orgFolder.FileLeafRef || orgFolder.name;
            console.log("IMPORT_MODAL_NAV: Entering Org Folder:", folderName);
            setCurrentFolderPath(folderName);
            return; 
          }
        } 
        
        // Level 2: Inside an organizational folder, look for specific project folder
        else if (targetFolders.some(f => currentFolderPath.toLowerCase().includes(f.toLowerCase()))) {
          const projectFolder = (result.items || []).find(item => {
            const name = item.FileLeafRef || item.name || "";
            return item.FSObjType === 1 && (
              name === props.projectCode || 
              name === props.businessProjectId || 
              (props.projectTitle && name.includes(props.projectTitle))
            );
          });

          if (projectFolder) {
            const folderName = projectFolder.FileLeafRef || projectFolder.name;
            console.log("IMPORT_MODAL_NAV: Entering Project Folder:", folderName);
            setCurrentFolderPath(currentFolderPath + "/" + folderName);
            setHasAutoNavigated(true);
            return;
          } else {
            setHasAutoNavigated(true);
          }
        } else {
          setHasAutoNavigated(true);
        }
      }
    } catch (err) {

    } finally {
      setLoading(false);
    }
  };

  const handleItemClick = (item: any) => {
    if (item.isImported) return; // Prevent selection of already imported files

    // Toggle selection for both files and folders
    setDocuments(docs => docs.map(d => d.Id === item.Id ? { ...d, selected: !d.selected } : d));
  };

  const handleFolderClick = (item: any) => {
    const newPath = currentFolderPath ? `${currentFolderPath}/${item.FileLeafRef}` : item.FileLeafRef;
    setCurrentFolderPath(newPath);
  };

  const handleBreadcrumbClick = (path: string) => {
    setCurrentFolderPath(path);
  };

  const breadcrumbs = React.useMemo(() => {
    const parts = currentFolderPath.split('/').filter(Boolean);
    const crumbs = [{ name: 'Home', path: '' }];
    let currentPath = '';
    parts.forEach(part => {
      currentPath = currentPath ? `${currentPath}/${part}` : part;
      crumbs.push({ name: part, path: currentPath });
    });
    return crumbs;
  }, [currentFolderPath]);

  const handleImport = async () => {
    const selectedDocs = documents.filter(doc => doc.selected);
    if (selectedDocs.length === 0) return;
    setImporting(true);
    try {
      if (props.service) {
        await props.service.importDocuments(selectedDocs, props.targetPath);
        toast.success('Imported successfully');
        props.onSuccess();
        props.onClose();
      }
    } catch (err) {
      toast.error('Import failed');
    } finally {
      setImporting(false);
    }
  };

  const getFileIcon = (item: IProjectDocument) => {
    if (item.FSObjType === 1) return <Icon iconName="FolderHorizontal" style={{ color: '#ffb900' }} />;
    return <Icon iconName="TextDocument" style={{ color: '#605e5c' }} />;
  };

  return (
    <Dialog
      hidden={!props.isOpen}
      onDismiss={props.onClose}
      dialogContentProps={{ type: DialogType.normal, title: `Import from ${props.projectTitle || 'Project'}` }}
      maxWidth={850}
      modalProps={{ isBlocking: true, className: styles.importDialog }}
    >
      <div className={styles.documentsWrapper}>
        <div className={styles.tabRow}>
          <div className={styles.documentTabs}>
            {(projectDocumentsUrl || libraryName) && (
              <button className={`${styles.docTab} ${activeDocTab === 'project' ? styles.activeDocTab : ''}`}
                onClick={() => { setActiveDocTab('project'); setCurrentFolderPath(''); }}>
                Project Documents
              </button>
            )}
            {bidDocumentsUrl && (
              <button className={`${styles.docTab} ${activeDocTab === 'bid' ? styles.activeDocTab : ''}`}
                onClick={() => { setActiveDocTab('bid'); setCurrentFolderPath(''); }}>
                Bid Documents
              </button>
            )}
            {contractsDocumentsUrl && (
              <button className={`${styles.docTab} ${activeDocTab === 'contract' ? styles.activeDocTab : ''}`}
                onClick={() => { setActiveDocTab('contract'); setCurrentFolderPath(''); }}>
                Contract Documents
              </button>
            )}
          </div>
        </div>

        <div className={styles.breadcrumbHeader}>
          <div className={styles.breadcrumbsContainer}>
            {breadcrumbs.map((crumb, index) => (
              <React.Fragment key={index}>
                <button className={`${styles.breadcrumb} ${index === breadcrumbs.length - 1 ? styles.active : ""}`}
                  onClick={() => handleBreadcrumbClick(crumb.path)} disabled={index === breadcrumbs.length - 1}>
                  {crumb.name === 'Home' ? <Icon iconName="Home" style={{ fontSize: '14px' }} /> : crumb.name}
                </button>
                {index < breadcrumbs.length - 1 && <span className={styles.breadcrumbSeparator}>›</span>}
              </React.Fragment>
            ))}
          </div>
        </div>

        <div className={styles.tableContainer}>
          {loading ? <Spinner size={SpinnerSize.large} label="Loading..." /> : (
            <table className={styles.documentsTable}>
              <thead>
                <tr>
                  <th style={{ width: '40px' }}></th>
                  <th>Name</th>
                  <th>Size / Items</th>
                  <th>Modified</th>
                  <th>By</th>
                </tr>
              </thead>
              <tbody>
                {documents.length === 0 ? (
                  <tr><td colSpan={5} style={{ textAlign: 'center', padding: '20px' }}>No documents found</td></tr>
                ) : (
                  documents.map((doc) => (
                    <tr key={doc.Id} className={`${styles.tableRow} ${doc.selected ? styles.selected : ''} ${doc.isImported ? styles.imported : ''}`} onClick={() => handleItemClick(doc)}>
                      <td>
                        {doc.isImported ? (
                          <Icon iconName="CheckMark" style={{ color: '#107c10' }} />
                        ) : (
                          <Icon iconName={doc.selected ? "CheckMark" : "CircleRing"} style={{ color: '#0078d4' }} />
                        )}
                      </td>
                      <td className={styles.nameCol}>
                        <span className={styles.fileIcon}>{getFileIcon(doc)}</span>
                        <span className={doc.FSObjType === 1 ? styles.folderName : styles.fileName}
                          onClick={(e) => {
                            if (doc.FSObjType === 1) {
                              e.stopPropagation(); // Don't toggle selection when navigating
                              handleFolderClick(doc);
                            }
                          }}>
                          {doc.FileLeafRef}
                          {doc.isImported && <span className={styles.importedLabel}>Imported</span>}
                        </span>
                      </td>
                      <td className={styles.sizeCol}>
                        {doc.FSObjType === 1 ? `${doc.ItemCount || 0} items` : formatFileSize(doc.Length)}
                      </td>
                      <td className={styles.modifiedCol}>{doc.Modified ? new Date(doc.Modified).toLocaleDateString() : ''}</td>
                      <td className={styles.byCol}>{doc.Editor?.Title || ""}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          )}
        </div>
      </div>
      <DialogFooter>
        <PrimaryButton onClick={handleImport} disabled={loading || importing || documents.filter(d => d.selected).length === 0} text={importing ? "Importing..." : "Import"} />
        <DefaultButton onClick={props.onClose} text="Cancel" />
      </DialogFooter>
    </Dialog>
  );
};

export default ImportFromProjectDocsDialog;
