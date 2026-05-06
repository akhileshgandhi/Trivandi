import * as React from 'react';
import styles from './PrivateComponent.module.scss';
import { PrivateService } from '../APIServices/PrivateService';
import {
  Folder,
  Search,
  ChevronRight,
  Download,
  Eye,
  FileText,
  File as FileIcon,
  Image as ImageIcon,
  FileSpreadsheet,
  X,
  Upload,
  FolderPlus,
  LayoutGrid,
  List,
  Plus
} from 'lucide-react';
import { BiSolidFilePdf } from "react-icons/bi";
import { BsFiletypeDoc } from "react-icons/bs";
import GlobalLoader from '../../../shared/component/GlobalLoader';
import FilePreview from '../Common/FilePreview';
import { usePermissionStore } from '../../../Permission/PermissionStore';

interface BreadcrumbItem {
  name: string;
  path: string;
}

export interface IPrivateProps {
  context: any;
  libraryName?: string;
}

const PrivateComponent: React.FC<IPrivateProps> = (props) => {
  const { canAdd } = usePermissionStore();
  const libraryName = props.libraryName || "Private";  const service = React.useMemo(() => new PrivateService(props.context, libraryName), [props.context, libraryName]);

  const [content, setContent] = React.useState<{ folders: any[], files: any[] }>({ folders: [], files: [] });
  const [currentPath, setCurrentPath] = React.useState<string>('');
  const [history, setHistory] = React.useState<BreadcrumbItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [searchText, setSearchText] = React.useState<string>('');
  const [selectedFileForPreview, setSelectedFileForPreview] = React.useState<any | null>(null);
  const [showAddModal, setShowAddModal] = React.useState<boolean>(false);
  const [addModalType, setAddModalType] = React.useState<'folder' | 'file'>('folder');
  const [newFolderName, setNewFolderName] = React.useState<string>('');
  const [filesToUpload, setFilesToUpload] = React.useState<File[]>([]);
  const [isSubmitting, setIsSubmitting] = React.useState<boolean>(false);
  const [isDragging, setIsDragging] = React.useState<boolean>(false);
  const [viewMode, setViewMode] = React.useState<'list' | 'grid'>('list');

  // One-click file selection effect
  React.useEffect(() => {
    if (showAddModal && addModalType === 'file' && filesToUpload.length === 0) {
      setTimeout(() => {
        document.getElementById('fileInput')?.click();
      }, 0);
    }
  }, [showAddModal, addModalType]);

  const loadData = React.useCallback(async (path: string = '') => {
    setLoading(true);
    try {
      const result = await service.getLibraryContents(path);
      console.log("🖼️ Private Component Content:", result);
      setContent({ folders: result.folders, files: result.files });
      setCurrentPath(result.currentPath);

      // Update history
      if (!path) {
        setHistory([{ name: libraryName, path: '' }]);
      }
    } catch (error) {
      console.error('PrivateFolders: Error loading data', error);
    } finally {
      setLoading(false);
    }
  }, [service, libraryName]);

  React.useEffect(() => {
    void loadData();
  }, [loadData]);

  const handleFolderClick = (folder: any) => {
    const newPath = folder.ServerRelativeUrl;
    const newHistory = [...history, { name: folder.Name, path: newPath }];
    setHistory(newHistory);
    void loadData(newPath);
  };

  const handleBreadcrumbClick = (item: BreadcrumbItem, index: number) => {
    if (index === history.length - 1) return;
    const newHistory = history.slice(0, index + 1);
    setHistory(newHistory);
    void loadData(item.path);
  };

  const getFileIcon = (fileName: string) => {
    const ext = fileName.split('.').pop()?.toLowerCase();
    const iconSize = 20;

    switch (ext) {
      case 'pdf':
        return React.createElement(BiSolidFilePdf as any, { size: iconSize, color: "#ef4444" });
      case 'doc':
        return React.createElement(BsFiletypeDoc as any, { size: iconSize, color: "#2563eb" });
      case 'docx':
        return <FileText size={iconSize} color="#2563eb" />;
      case 'xls':
      case 'xlsx':
      case 'csv':
        return <FileSpreadsheet size={iconSize} color="#16a34a" />;
      case 'jpg':
      case 'jpeg':
      case 'png':
      case 'gif':
      case 'svg':
        return <ImageIcon size={iconSize} color="#9333ea" />;
      default:
        return <FileIcon size={iconSize} color="#64748b" />;
    }
  };

  const formatDate = (dateStr: string) => {
    if (!dateStr) return "--";
    const date = new Date(dateStr);
    return date.toLocaleDateString('en-US', {
      month: 'short',
      day: 'numeric',
      year: 'numeric'
    });
  };

  const handleDownload = (serverRelativeUrl: string, fileName: string) => {
    void service.downloadFile(serverRelativeUrl, fileName);
  };

  const filteredFolders = content.folders.filter(f => f.Name.toLowerCase().indexOf(searchText.toLowerCase()) > -1);
  const filteredFiles = content.files.filter(f => f.Name.toLowerCase().indexOf(searchText.toLowerCase()) > -1);

  const handleCreateFolder = async () => {
    if (!newFolderName.trim() || !currentPath) return;
    setIsSubmitting(true);
    try {
      await service.createFolder(currentPath, newFolderName);
      setNewFolderName('');
      setShowAddModal(false);
      void loadData(currentPath);
    } catch (error) {
      alert("Failed to create folder");
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleUploadFiles = async () => {
    if (filesToUpload.length === 0 || !currentPath) return;
    setIsSubmitting(true);
    try {
      // Parallel upload for smoother/faster experience
      await Promise.all(filesToUpload.map(file => service.uploadFile(currentPath, file)));

      setFilesToUpload([]);
      setShowAddModal(false);
      await loadData(currentPath);
    } catch (error) {
      console.error("Upload error:", error);
      alert("Failed to upload some files. Please check your connection.");
    } finally {
      setIsSubmitting(false);
    }
  };

  const onFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const selectedFiles = Array.from(e.target.files);
      setFilesToUpload(prev => [...prev, ...selectedFiles]);
      // Reset input value to allow selecting the same file again immediately
      e.target.value = '';
    }
  };

  const totalItems = content.folders.length + content.files.length;

  return (
    <div className={styles.marketingPrivateContainer}>
      {/* Page Title Section */}
      <div className={styles.pageTitleSection}>
        <div className={styles.libraryIcon}>
          <FileIcon size={22} />
        </div>
        <div className={styles.titleInfo}>
          <h1>PrivateFolders</h1>
          <span className={styles.itemCount}>{totalItems} items</span>
        </div>
      </div>

      {/* Header Bar */}
      <div className={styles.headerBar}>
        <div className={styles.leftSection}>
          <div className={styles.breadcrumbWrapper}>
            <span className={styles.breadcrumbItem} onClick={() => void loadData('')}>Root</span>
            {history.length > 1 && (
              <>
                <ChevronRight size={14} className={styles.breadcrumbSeparator} />
                <span className={styles.breadcrumbItem}>{history[history.length - 1].name}</span>
              </>
            )}
          </div>
        </div>

        <div className={styles.rightSection}>
          <div className={styles.searchBox}>
            <div className={styles.searchIcon}><Search size={18} /></div>
            <input
              type="text"
              placeholder="Search..."
              value={searchText}
              onChange={(e) => setSearchText(e.target.value)}
            />
          </div>
          <div className={styles.viewToggles}>
            <button
              className={viewMode === 'grid' ? styles.viewActive : ''}
              onClick={() => setViewMode('grid')}
            >
              <LayoutGrid size={18} />
            </button>
            <button
              className={viewMode === 'list' ? styles.viewActive : ''}
              onClick={() => setViewMode('list')}
            >
              <List size={18} />
            </button>
          </div>
          {canAdd && (
            <button className={styles.addBtn} onClick={() => {
              setAddModalType('folder');
              setShowAddModal(true);
            }}>
              <Plus size={18} />
              Add New
            </button>
          )}
        </div>
      </div>

      {/* Main Content Area */}
      <div className={styles.contentWrapper}>
        {viewMode === 'list' ? (
          <table className={styles.documentTable}>
            <thead>
              <tr>
                <th>NAME</th>
                <th>MODIFIED</th>
                <th>MODIFIED BY</th>
                <th>SIZE/ITEMS</th>
              </tr>
            </thead>
            <tbody>
              {loading ? (
                <tr>
                  <td colSpan={4} className={styles.loaderCell}>
                    <GlobalLoader variant="content" />
                  </td>
                </tr>
              ) : (
                <>
                  {filteredFolders.map((folder, idx) => (
                    <tr key={`folder-${idx}`} className={styles.tableRow} onClick={() => handleFolderClick(folder)}>
                      <td>
                        <div className={styles.itemCell}>
                          <div className={styles.folderIconWrapper}><Folder size={20} /></div>
                          <span className={styles.itemName}>{folder.Name}</span>
                        </div>
                      </td>
                      <td><span className={styles.dateText}>{formatDate(folder.TimeLastModified)}</span></td>
                      <td><span className={styles.dateText}>{folder.Editor?.Title || 'System'}</span></td>
                      <td><span className={styles.sizeText}>{folder.ItemCount || 0} items</span></td>
                    </tr>
                  ))}

                  {filteredFiles.map((file, idx) => (
                    <tr key={`file-${idx}`} className={styles.tableRow} onClick={() => {
                      const fileObj = {
                        name: file.Name,
                        url: window.location.origin + file.ServerRelativeUrl,
                        serverRelativeUrl: file.ServerRelativeUrl
                      };
                      setSelectedFileForPreview(fileObj);
                    }}>
                      <td>
                        <div className={styles.itemCell}>
                          <div className={styles.fileIconWrapper}>{getFileIcon(file.Name)}</div>
                          <span className={styles.itemName}>{file.Name}</span>
                        </div>
                      </td>
                      <td><span className={styles.dateText}>{formatDate(file.TimeLastModified)}</span></td>
                      <td><span className={styles.dateText}>{file.Editor?.Title || 'User'}</span></td>
                      <td><span className={styles.sizeText}>{service.formatBytes(file.Length)}</span></td>
                    </tr>
                  ))}

                  {!loading && filteredFolders.length === 0 && filteredFiles.length === 0 && (
                    <tr>
                      <td colSpan={4} className={styles.emptyCell}>No items found</td>
                    </tr>
                  )}
                </>
              )}
            </tbody>
          </table>
        ) : (
          <div className={styles.documentGrid}>
            {loading ? (
              <div className={styles.loaderCell} style={{ gridColumn: '1 / -1' }}>
                <GlobalLoader variant="content" />
              </div>
            ) : (
              <>
                {filteredFolders.map((folder, idx) => (
                  <div
                    key={`folder-${idx}`}
                    className={`${styles.gridCard} ${styles['card' + (idx % 6)]}`}
                    onClick={() => handleFolderClick(folder)}
                  >
                    <div className={styles.cardIcon}><Folder size={18} /></div>
                    <div className={styles.cardTitle}>{folder.Name}</div>
                    <div className={styles.cardSub}>{folder.ItemCount || 0} items</div>
                    <div className={styles.cardFooter}>
                      <div className={styles.metaRow}>
                        <span className={styles.label}>MODIFIED</span>
                        <span className={styles.value}>{formatDate(folder.TimeLastModified)}</span>
                      </div>
                      <div className={styles.modifiedByInfo}>
                        <div className={styles.avatar}>
                          {folder.Editor?.Title ? folder.Editor.Title.charAt(0).toUpperCase() : 'S'}
                        </div>
                        <span className={styles.userName}>{folder.Editor?.Title || 'System'}</span>
                      </div>
                    </div>
                  </div>
                ))}
                {filteredFiles.map((file, idx) => (
                  <div
                    key={`file-${idx}`}
                    className={`${styles.gridCard} ${styles['card' + ((filteredFolders.length + idx) % 6)]}`}
                    onClick={() => {
                      const fileObj = {
                        name: file.Name,
                        url: window.location.origin + file.ServerRelativeUrl,
                        serverRelativeUrl: file.ServerRelativeUrl
                      };
                      setSelectedFileForPreview(fileObj);
                    }}
                  >
                    <div className={styles.cardIcon}>{getFileIcon(file.Name)}</div>
                    <div className={styles.cardTitle}>{file.Name}</div>
                    <div className={styles.cardSub}>{service.formatBytes(file.Length)}</div>
                    <div className={styles.cardFooter}>
                      <div className={styles.metaRow}>
                        <span className={styles.label}>MODIFIED</span>
                        <span className={styles.value}>{formatDate(file.TimeLastModified)}</span>
                      </div>
                      <div className={styles.modifiedByInfo}>
                        <div className={styles.avatar}>
                          {file.Editor?.Title ? file.Editor.Title.charAt(0).toUpperCase() : 'U'}
                        </div>
                        <span className={styles.userName}>{file.Editor?.Title || 'User'}</span>
                      </div>
                    </div>
                  </div>
                ))}
              </>
            )}
          </div>
        )}
      </div>

      {/* Preview Modal */}
      {selectedFileForPreview && (
        <div className={styles.previewOverlay} onClick={() => setSelectedFileForPreview(null)}>
          <div className={styles.previewModal} onClick={(e) => e.stopPropagation()}>
            <FilePreview file={selectedFileForPreview} onClose={() => setSelectedFileForPreview(null)} />
          </div>
        </div>
      )}

      {/* Add New Modal */}
      {showAddModal && (
        <div className={styles.previewOverlay} onClick={() => !isSubmitting && setShowAddModal(false)}>
          <div className={styles.addModal} onClick={(e) => e.stopPropagation()}>
            <div className={styles.modalHeader}>
              <h3>Add to {history[history.length - 1]?.name || libraryName}</h3>
              <button onClick={() => setShowAddModal(false)}><X size={20} /></button>
            </div>

            <div className={styles.modalBody}>
              <div className={styles.tabContainer}>
                <button
                  className={`${styles.tab} ${addModalType === 'folder' ? styles.tabActive : ''}`}
                  onClick={() => setAddModalType('folder')}
                >
                  <FolderPlus size={18} /> New Folder
                </button>
                <button
                  className={`${styles.tab} ${addModalType === 'file' ? styles.tabActive : ''}`}
                  onClick={() => setAddModalType('file')}
                >
                  <Upload size={18} /> Upload Files
                </button>
              </div>

              <div className={styles.tabContent}>
                {addModalType === 'folder' ? (
                  <div className={styles.formGroup}>
                    <label>Folder Name</label>
                    <input
                      type="text"
                      placeholder="Enter folder name..."
                      value={newFolderName}
                      onChange={(e) => setNewFolderName(e.target.value)}
                      autoFocus
                    />
                  </div>
                ) : (
                  <div className={styles.fileUploadSection}>
                    <label
                      className={`${styles.dropZone} ${isDragging ? styles.dragging : ''}`}
                      htmlFor="fileInput"
                      onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
                      onDragEnter={(e) => { e.preventDefault(); setIsDragging(true); }}
                      onDragLeave={(e) => { e.preventDefault(); setIsDragging(false); }}
                      onDrop={(e) => {
                        e.preventDefault();
                        setIsDragging(false);
                        if (e.dataTransfer.files) setFilesToUpload(prev => [...prev, ...Array.from(e.dataTransfer.files)]);
                      }}
                    >
                      <div className={styles.dropIcon}><Upload size={32} /></div>
                      <p>{filesToUpload.length > 0 ? `${filesToUpload.length} files selected` : "Click to select or drag files here"}</p>
                      <input
                        id="fileInput"
                        type="file"
                        multiple
                        style={{ display: 'none' }}
                        onChange={onFileChange}
                      />
                    </label>
                    {filesToUpload.length > 0 && (
                      <div className={styles.fileList}>
                        {filesToUpload.map((file, idx) => (
                          <div key={idx} className={styles.fileItem}>
                            <div className={styles.fileInfo}>{getFileIcon(file.name)} <span>{file.name}</span></div>
                            <X size={14} className={styles.removeFile} onClick={(e) => {
                              e.stopPropagation();
                              setFilesToUpload(filesToUpload.filter((_, i) => i !== idx));
                            }} />
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                )}
              </div>
            </div>

            <div className={styles.modalFooter}>
              <button className={styles.cancelBtn} onClick={() => setShowAddModal(false)} disabled={isSubmitting}>Cancel</button>
              <button
                className={styles.confirmBtn}
                onClick={addModalType === 'folder' ? handleCreateFolder : handleUploadFiles}
                disabled={isSubmitting || (addModalType === 'folder' ? !newFolderName.trim() : filesToUpload.length === 0)}
              >
                {isSubmitting ? 'Processing...' : (addModalType === 'folder' ? 'Create Folder' : 'Upload Files')}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default PrivateComponent;
