import * as React from 'react';
import styles from './SharedFiles.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ISharedDocument, IBreadcrumb, ISharedFileAccessLog } from '../IProjectExternalPortalState';
import { Icon } from '@fluentui/react/lib/Icon';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { Breadcrumb } from '@fluentui/react/lib/Breadcrumb';
import { ProjectExternalPortalService } from '../../services/ProjectExternalPortalService';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { ArrowDownToLine } from 'lucide-react';
import FilePreview from '../FilePreview/FilePreview';
import { usePermissionStore } from '../../../../Permission/PermissionStore';
import Pagination from '../../../../shared/component/Pagination/Pagination';

export interface ISharedFilesProps {
  context: WebPartContext;
  currentPath: string;
  breadcrumbs: IBreadcrumb[];
  onFolderNavigate: (path: string, breadcrumbs: IBreadcrumb[]) => void;
  onShowNewDocument: () => void;
  onShowShareAccess: (selectedDocs: ISharedDocument[]) => void;
  onShowImportDialog: () => void;
  onRefresh?: () => void;
  service?: ProjectExternalPortalService; // Optional shared service instance
  isUserRestricted?: boolean; // Whether current user is restricted from sharing
  onFolderChange?: (path: string) => void;
}

export interface ISharedFilesState {
  documents: ISharedDocument[];
  loading: boolean;
  error?: string;
  selectedDocuments: number[];
  selectAll: boolean;
  currentPermission: 'Read' | 'Review' | 'Edit' | 'Admin' | null;
  guestRole: 'Viewer' | 'Editor' | null;
  showLogDialog: boolean;
  loadingLogs: boolean;
  logsError?: string;
  selectedLogDocument?: ISharedDocument;
  accessLogs: ISharedFileAccessLog[];
  totalDocumentsFound?: number;
  accessFilteredCount?: number;
  filePreview: {
    isOpen: boolean;
    fileUrl: string;
    fileName: string;
    filePath?: string;
    canDownload?: boolean;
  };
  pagination: {
    currentPage: number;
    pageSize: number;
    totalItems: number;
    totalPages: number;
  };
}

export default class SharedFiles extends React.Component<ISharedFilesProps, ISharedFilesState> {
  private portalService: ProjectExternalPortalService;

  constructor(props: ISharedFilesProps) {
    super(props);

    this.state = {
      documents: [],
      loading: false,
      selectedDocuments: [],
      selectAll: false,
      currentPermission: null,
      guestRole: null,
      showLogDialog: false,
      loadingLogs: false,
      accessLogs: [],
      filePreview: {
        isOpen: false,
        fileUrl: '',
        fileName: '',
        filePath: '',
        canDownload: true
      },
      pagination: {
        currentPage: 1,
        pageSize: 10,
        totalItems: 0,
        totalPages: 0
      }
    };

    this.portalService = props.service || new ProjectExternalPortalService(props.context);
    this._loadDebounceTimer = null;
  }

  private _loadDebounceTimer: number | null;

  public componentDidMount(): void {
    this._loadDocuments();
  }

  public componentDidUpdate(prevProps: ISharedFilesProps): void {
    if (prevProps.currentPath !== this.props.currentPath) {
      // Notify parent of folder change
      if (this.props.onFolderChange) {
        this.props.onFolderChange(this.props.currentPath);
      }
      this._loadDocuments();
    }
  }

  public componentWillUnmount(): void {
    // Clear debounce timer on unmount
    if (this._loadDebounceTimer) {
      clearTimeout(this._loadDebounceTimer as any);
    }
  }

  private _loadDocuments = async (page: number = this.state.pagination.currentPage, pageSize: number = this.state.pagination.pageSize): Promise<void> => {
    // Prevent multiple simultaneous loads
    if (this.state.loading) {
      
      return;
    }

    // Clear any pending debounce timer
    if (this._loadDebounceTimer) {
      clearTimeout(this._loadDebounceTimer as any);
    }

    
    this.setState({
      loading: true,
      selectedDocuments: [],
      selectAll: false,
      totalDocumentsFound: 0,
      accessFilteredCount: 0
    });
    try {
      const [{ items: documents, totalCount }, currentPermission, guestRole] = await Promise.all([
        this.portalService.getSharedDocumentsFromExternalLibrary(this.props.currentPath, page, pageSize),
        this.portalService.getUserPermissionForFolderPath(this.props.currentPath),
        this.portalService.getCurrentGuestRole()
      ]);
      

      this.setState({
        documents,
        currentPermission,
        guestRole,
        loading: false,
        totalDocumentsFound: documents.length,
        accessFilteredCount: 0,
        pagination: {
          currentPage: page,
          pageSize: pageSize,
          totalItems: totalCount,
          totalPages: Math.ceil(totalCount / pageSize)
        }
      });
    } catch (error) {
      const errorMessage = error.message || 'Failed to load documents';
      

      let displayError: string;
      if (errorMessage.includes('404') || errorMessage.includes('does not exist')) {
        displayError = 'ExternalShareDocument library does not exist at this site.';
      } else if (errorMessage.includes('403') || errorMessage.includes('Unauthorized')) {
        displayError = 'You do not have access to the ExternalShareDocument library.';
      } else if (errorMessage.includes('429') || errorMessage.includes('throttle')) {
        displayError = 'Too many requests. Please wait a moment and try again.';
      } else if (errorMessage.includes('500')) {
        displayError = 'Server error occurred. Please try again later.';
      } else {
        displayError = errorMessage;
      }

      this.setState({
        error: displayError,
        loading: false,
        documents: [],
        currentPermission: null
      });
    }
  }

  public refreshDocuments = (): void => {
    
    // Debounce refresh calls to prevent rapid successive requests
    if (this._loadDebounceTimer) {
      clearTimeout(this._loadDebounceTimer as any);
    }

    this._loadDebounceTimer = window.setTimeout(() => {
      
      this._loadDocuments(1); // Reset to page 1 on refresh
    }, 300); // 300ms debounce
  }

  private handlePageChange = (page: number): void => {
    this._loadDocuments(page, this.state.pagination.pageSize);
  };

  private handlePageSizeChange = (pageSize: number): void => {
    this._loadDocuments(1, pageSize);
  };

  private _onSelectAllChange = (ev: React.FormEvent<HTMLInputElement>): void => {
    const checked = (ev.currentTarget as HTMLInputElement).checked;
    const { documents } = this.state;
    if (checked) {
      this.setState({
        selectedDocuments: documents.map((_, index) => index),
        selectAll: true
      });
    } else {
      this.setState({
        selectedDocuments: [],
        selectAll: false
      });
    }
  }

  private _onDocumentSelectChange = (index: number): void => {
    const { selectedDocuments, documents } = this.state;
    const newSelected = [...selectedDocuments];

    const idx = newSelected.indexOf(index);
    if (idx > -1) {
      newSelected.splice(idx, 1);
    } else {
      newSelected.push(index);
    }

    this.setState({
      selectedDocuments: newSelected,
      selectAll: newSelected.length === documents.length
    });
  }

  private _onShareAccessClick = (): void => {
    const { selectedDocuments, documents, currentPermission } = this.state;
    const selectedDocs = selectedDocuments.map(index => documents[index]);

    if (selectedDocs.length === 0) {
      alert('Please select at least one file or folder to share');
      return;
    }

    // Enhanced permission checking with better user feedback
    const readOnlyDocs = selectedDocs.filter(doc => doc.permission === 'Read' || doc.permission === 'Review');
    const hasReadOnlyUser = currentPermission === 'Read' || currentPermission === 'Review';

    if (hasReadOnlyUser) {
      alert(`You only have ${currentPermission} permission for this location. Sharing is not allowed. Please contact your project administrator to request Edit or Admin access.`);
      return;
    }

    if (readOnlyDocs.length > 0) {
      const docNames = readOnlyDocs.map(doc => doc.name).join(', ');
      alert(`You don't have Edit or Admin permission for the following documents: ${docNames}. Sharing requires Edit or Admin permission on all selected items.`);
      return;
    }

    this.props.onShowShareAccess(selectedDocs);
  }

  private _onBreadcrumbClick = (item: any): void => {
    const index = this.props.breadcrumbs.findIndex(b => b.key === item.key);
    const newBreadcrumbs = this.props.breadcrumbs.slice(0, index + 1);
    this.props.onFolderNavigate(item.path, newBreadcrumbs);
  }

  private _handleFolderClick = (folderPath: string): void => {
    // Notify parent of folder change
    if (this.props.onFolderChange) {
      this.props.onFolderChange(folderPath);
    }

    this.setState({
      loading: true
    }, () => {
      this._loadDocuments();
    });
  }

  private _onFolderClick = (folder: ISharedDocument): void => {
    const newPath = `${this.props.currentPath}/${folder.name}`;
    const newBreadcrumbs = [
      ...this.props.breadcrumbs,
      { text: folder.name, key: folder.name, path: newPath }
    ];
    this.props.onFolderNavigate(newPath, newBreadcrumbs);
  }

  private _formatDate = (date?: Date): string => {
    if (!date) return '';
    return new Date(date).toLocaleString('en-US', {
      month: '2-digit',
      day: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: true
    });
  }

  private _formatExpiry = (log: { accessDuration?: number; expiryDate?: Date; actionDate?: Date }): JSX.Element => {
    const now = new Date();

    // Resolve the expiry date: prefer stored ExpiryDate, fall back to ActionDate + Duration
    let resolved: Date | null = null;
    if (log.expiryDate) {
      resolved = log.expiryDate;
    } else if (log.accessDuration && log.accessDuration > 0 && log.actionDate) {
      resolved = new Date(log.actionDate.getTime() + log.accessDuration * 24 * 60 * 60 * 1000);
    }

    // Permanent: duration 0 or -1 with no expiry date
    if (!resolved) {
      return <span style={{ color: '#107c41', fontWeight: 600 }}>Permanent</span>;
    }

    const isExpired = now > resolved;
    const formatted = resolved.toLocaleString('en-US', {
      month: '2-digit', day: '2-digit', year: 'numeric',
      hour: '2-digit', minute: '2-digit', hour12: true
    });

    if (isExpired) {
      return (
        <span>
          <span style={{
            display: 'inline-block', background: '#fde7e9', color: '#a80000',
            borderRadius: 10, padding: '8px 8px', fontSize: 11, fontWeight: 600, marginRight: 6
          }}>Expired</span>
          <br />
          <span style={{ color: '#a80000', fontSize: 12 }}>{formatted}</span>
        </span>
      );
    }

    return <span style={{ color: '#323130', fontSize: 12 }}>{formatted}</span>;
  }

  private _onViewLogClick = async (doc: ISharedDocument): Promise<void> => {
    this.setState({
      showLogDialog: true,
      loadingLogs: true,
      logsError: undefined,
      selectedLogDocument: doc,
      accessLogs: []
    });

    try {
      // External (restricted) guests only see their own log entries
      const guestEmail = this.props.isUserRestricted
        ? (this.props.context.pageContext.user.email || undefined)
        : undefined;
      const accessLogs = await this.portalService.getDocumentAccessLog(doc.fileRef, guestEmail);
      this.setState({ accessLogs, loadingLogs: false });
    } catch (error) {
      this.setState({
        loadingLogs: false,
        logsError: error.message || 'Failed to load access log'
      });
    }
  }

  private _closeLogDialog = (): void => {
    this.setState({
      showLogDialog: false,
      loadingLogs: false,
      logsError: undefined,
      selectedLogDocument: undefined,
      accessLogs: []
    });
  }

  private _openFilePreview = (document: ISharedDocument): void => {
    // All guests with any valid permission (Read, Review, Edit, Admin) can view/preview files.
    // Only block if the user is restricted AND the document has no recognised permission at all.
    if (this.props.isUserRestricted && !document.permission) {
      alert('You do not have access to this file.');
      return;
    }
    // Read-only guests may preview but not download
    const canDownload = !this.props.isUserRestricted ||
      document.permission === 'Edit' || document.permission === 'Admin';
    this.setState({
      filePreview: {
        isOpen: true,
        fileUrl: document.fileRef,
        fileName: document.name,
        filePath: document.fileRef,
        canDownload
      }
    });
  }

  private _closeFilePreview = (): void => {
    this.setState({
      filePreview: {
        isOpen: false,
        fileUrl: '',
        fileName: '',
        filePath: '',
        canDownload: true
      }
    });
  }

  public render(): React.ReactElement<ISharedFilesProps> {
    const { documents, loading, error, selectedDocuments, selectAll, currentPermission, guestRole, showLogDialog, loadingLogs, logsError, selectedLogDocument, accessLogs } = this.state;
    const { breadcrumbs, isUserRestricted } = this.props;
    // Per-document helpers based on the document's own shared permission
    const { canAdd: globalCanAdd, canEdit: globalCanEdit } = usePermissionStore.getState();

    const canPreviewDoc = (doc: ISharedDocument): boolean => {
      if (!isUserRestricted) return true;
      return !!doc.permission; // any assigned permission allows viewing
    };
    const canDownloadDoc = (doc: ISharedDocument): boolean => {
      if (!isUserRestricted) return true;
      // Only Edit/Admin guests may download
      return doc.permission === 'Edit' || doc.permission === 'Admin';
    };
    // Whether the guest can edit/upload/import at the current folder level
    const canEditCurrent = (currentPermission === 'Edit' || currentPermission === 'Admin') && globalCanEdit;
    const canAddCurrent = globalCanAdd;
    const selectedDocs = selectedDocuments.map(index => documents[index]);
    const canShareSelection = selectedDocs.length > 0 && canEditCurrent && !selectedDocs.some(doc => doc.permission === 'Read' || doc.permission === 'Review') && !isUserRestricted;

    const breadcrumbItems = breadcrumbs.map(b => ({
      text: b.text,
      key: b.key,
      onClick: () => this._onBreadcrumbClick(b)
    }));

    return (
      <div className={styles.sharedFiles}>
        <div className={styles.headerRow}>
          <div className={styles.titleArea}>
            <span className={styles.accentBar}></span>
            <h4 className={styles.sectionTitle}>Shared Files</h4>
          </div>

          <div className={styles.headerActions}>
            <DefaultButton
              text="Import"
              iconProps={{ iconName: 'CloudDownload' }}
              className={styles.importButton}
              onClick={() => {
                if (canEditCurrent && !isUserRestricted) {
                  this.props.onShowImportDialog();
                }
              }}
              disabled={!canEditCurrent || isUserRestricted}
              title={
                isUserRestricted ? 'You do not have permission to import documents' :
                  !globalCanEdit ? 'You do not have global edit permission' :
                    !canEditCurrent ? 'Import requires Edit permission for this folder' : ''
              }
            />

            {canAddCurrent && (
              <PrimaryButton
                text="New"
                iconProps={{ iconName: 'Add' }}
                onClick={this.props.onShowNewDocument}
                disabled={!canEditCurrent || isUserRestricted}
                title={
                  isUserRestricted ? 'You do not have permission to add documents' :
                    !globalCanAdd ? 'You do not have global add permission' :
                      !canEditCurrent ? 'Adding requires Edit permission for this folder' : ''
                }
                className={styles.newButton}
              />
            )}
          </div>
        </div>


        {/* Breadcrumb and Actions */}
        <div className={styles.toolbar}>
          <div className={styles.breadcrumbContainer}>
            <Breadcrumb
              items={breadcrumbItems}
              maxDisplayedItems={5}
              className={styles.breadcrumb}
            />
          </div>
          <div className={styles.actions}>
            {selectedDocuments.length > 0 && (
              <DefaultButton
                text={`Share Access (${selectedDocuments.length})`}
                iconProps={{ iconName: 'Share' }}
                onClick={this._onShareAccessClick}
                disabled={!canShareSelection}
                title={
                  isUserRestricted ? 'You do not have permission to share documents' :
                    !canShareSelection ? 'Sharing requires Edit permission' : ''
                }
                className={styles.shareButton}
              />
            )}
          </div>
        </div>

        {/* Permission Info Message */}
        {!this.state.loading && !this.state.error && this.state.currentPermission && (
          <MessageBar
            messageBarType={
              this.state.currentPermission === 'Read' ? MessageBarType.warning :
                this.state.currentPermission === 'Review' ? MessageBarType.info :
                  this.state.currentPermission === 'Admin' ? MessageBarType.success :
                    MessageBarType.info
            }
            isMultiline={false}
            className={styles.permissionInfo}
          >
            {this.state.currentPermission === 'Read' ? (
              <>
                <strong>Read access:</strong> You can view documents in the portal but cannot download, upload, share, or import files.
              </>
            ) : this.state.currentPermission === 'Review' ? (
              <>
                <strong>Review access:</strong> You can view and comment on documents but cannot download, upload, share, or import files.
              </>
            ) : this.state.currentPermission === 'Admin' ? (
              <>
                <strong>Admin access:</strong> You have full control — view, edit, download, upload, share, import and manage guest permissions.
              </>
            ) : (
              <>
                <strong>Edit access:</strong> You can view, edit, download, upload, share, and import documents.
              </>
            )}
          </MessageBar>
        )}

        {/* Documents Table */}
        {this.state.loading ? (
          <div className={styles.loadingContainer}>
            <Spinner size={SpinnerSize.large} label="Loading documents..." />
          </div>
        ) : this.state.error ? (
          <div className={styles.errorContainer}>
            <Icon iconName="Error" className={styles.errorIcon} />
            <p>{this.state.error}</p>
            <DefaultButton text="Retry" onClick={() => this._loadDocuments()} />
          </div>
        ) : (
          <div className={styles.tableContainer}>
            <table className={styles.table}>
              <thead>
                <tr>
                  <th className={styles.checkboxColumn}>
                    <input
                      type="checkbox"
                      checked={this.state.selectAll}
                      onChange={this._onSelectAllChange}
                    />
                  </th>
                  <th className={styles.nameColumn}>Name</th>
                  <th className={styles.modifiedColumn}>Modified</th>
                  <th className={styles.byColumn}>Modified By</th>
                  <th className={styles.actionColumn}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {this.state.documents.length === 0 ? (
                  <tr>
                    <td colSpan={5} className={styles.emptyState}>
                      <Icon iconName="FabricFolder" className={styles.emptyIcon} />
                      {this.state.error && this.state.error.includes('does not exist') ? (
                        <>
                          <p>External Portal not available</p>
                          <p className={styles.emptySubtext}>
                            The External Portal library has not been set up or you don't have access to it.
                            Please contact your site administrator to set up the External Portal for this project.
                          </p>
                        </>
                      ) : isUserRestricted ? (
                        <>
                          <p>No documents have been shared with you</p>
                          <p className={styles.emptySubtext}>
                            As an external guest, you can only see documents that have been specifically shared with you.
                            Please contact your project administrator if you believe you should have access to additional documents.
                          </p>
                        </>
                      ) : this.state.currentPermission === null ? (
                        <>
                          <p>No documents found or you may not have access to this location.</p>
                          <p className={styles.emptySubtext}>
                            If you believe you should have access, please contact your project administrator or try refreshing the page.
                          </p>
                        </>
                      ) : this.state.currentPermission === 'Read' ? (
                        <>
                          <p>No documents found in External Shared Area.</p>
                          <p className={styles.emptySubtext}>
                            You have <strong>Read</strong> permission. Contact your project administrator to upload documents or request higher access.
                          </p>
                        </>
                      ) : this.state.currentPermission === 'Review' ? (
                        <>
                          <p>No documents found in External Shared Area.</p>
                          <p className={styles.emptySubtext}>
                            You have <strong>Review</strong> permission. Contact your project administrator to upload documents or request Edit access.
                          </p>
                        </>
                      ) : (
                        <>
                          <p>No documents found in External Shared Area.</p>
                          <p className={styles.emptySubtext}>
                            Click "New" to upload a document or "Import" to copy files from project documents.
                          </p>
                        </>
                      )}
                    </td>
                  </tr>
                ) : (
                  this.state.documents.map((doc, index) => (
                    <tr
                      key={index}
                      className={`${styles.row} ${this.state.selectedDocuments.includes(index) ? styles.selected : ''}`}
                    >
                      <td className={styles.checkboxColumn}>
                        <input
                          type="checkbox"
                          checked={this.state.selectedDocuments.includes(index)}
                          onChange={() => this._onDocumentSelectChange(index)}
                        />
                      </td>
                      <td className={styles.nameColumn}>
                        <div className={styles.nameCell}>
                          {doc.isFolder ? (
                            <>
                              <Icon iconName="FabricFolder" className={styles.folderIcon} />
                              <span
                                className={styles.folderLink}
                                onClick={() => this._onFolderClick(doc)}
                              >
                                {doc.name}
                              </span>
                            </>
                          ) : (
                            <>
                              <Icon iconName="Page" className={styles.fileIcon} />
                              <span
                                //  href={doc.fileRef}
                                // target="_blank"
                                // rel="noopener noreferrer"
                                className={styles.fileLink}
                              >
                                {doc.name}
                              </span>
                            </>
                          )}
                        </div>
                      </td>
                      <td className={styles.modifiedColumn}>
                        {this._formatDate(doc.modified)}
                      </td>
                      <td className={styles.byColumn}>{doc.modifiedBy}</td>
                      <td className={styles.actionColumn}>
                        <IconButton
                          iconProps={{ iconName: 'MoreVertical' }}
                          title="More actions"
                          className={styles.moreButton}
                          menuProps={{
                            items: [
                              {
                                key: 'preview',
                                text: 'Preview',
                                iconProps: { iconName: 'View' },
                                onClick: () => {
                                  this._openFilePreview(doc);
                                },
                                disabled: doc.isFolder || !canPreviewDoc(doc),
                                title: doc.isFolder
                                  ? 'Cannot preview folders'
                                  : !canPreviewDoc(doc)
                                    ? 'You do not have access to preview this file'
                                    : `Preview ${doc.name}`
                              },
                              {
                                key: 'view',
                                text: 'View Log',
                                iconProps: { iconName: 'History' },
                                onClick: () => {
                                  this._onViewLogClick(doc).catch((error) => {
                                    
                                  });
                                }
                              },
                              {
                                key: 'share',
                                text: 'Share',
                                iconProps: { iconName: 'Share' },
                                onClick: () => {
                                  this.setState({ selectedDocuments: [index] }, () => {
                                    this._onShareAccessClick();
                                  });
                                },
                                disabled: this.state.currentPermission === 'Read' || this.state.currentPermission === 'Review' || doc.permission === 'Read' || doc.permission === 'Review' || isUserRestricted,
                                title: isUserRestricted
                                  ? 'You do not have permission to share documents'
                                  : (this.state.currentPermission === 'Read' || this.state.currentPermission === 'Review')
                                    ? 'You need Edit or Admin permission to share documents. Contact your project administrator.'
                                    : (doc.permission === 'Read' || doc.permission === 'Review') ? 'This document requires Edit or Admin permission to share.' : 'Share this document with external users'
                              },
                              {
                                key: 'download',
                                text: 'Download',
                                iconProps: { iconName: 'Download' },
                                onClick: () => { window.open(doc.fileRef, '_blank'); },
                                // Read guests: view only, no download. Edit/Admin guests: can download.
                                disabled: doc.isFolder || !canDownloadDoc(doc),
                                title: doc.isFolder
                                  ? 'Cannot download folders - navigate into the folder to access individual files'
                                  : !canDownloadDoc(doc)
                                    ? 'Download requires Edit permission'
                                    : 'Download this file'
                              },
                              {
                                key: 'openInNewTab',
                                text: 'Open in New Tab',
                                iconProps: { iconName: 'OpenInNewWindow' },
                                onClick: () => {
                                  // Build SharePoint viewer URL so the file opens in Word/Excel/PowerPoint Online
                                  // instead of downloading. Use action=view for read-only, action=default for edit.
                                  const siteUrl = this.props.context.pageContext.web.absoluteUrl;
                                  const ext = doc.name.split('.').pop()?.toLowerCase() || '';
                                  const officeExts = ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'pdf'];
                                  if (officeExts.includes(ext)) {
                                    const viewerUrl = `${siteUrl}/_layouts/15/Doc.aspx?sourcedoc=${encodeURIComponent(doc.fileRef)}&action=view`;
                                    window.open(viewerUrl, '_blank');
                                  } else {
                                    window.open(doc.fileRef, '_blank');
                                  }
                                },
                                // Read guests: view only, no open-in-new-tab (would allow saving/editing).
                                // Edit/Admin guests: can open in new tab.
                                disabled: doc.isFolder || !canDownloadDoc(doc),
                                title: doc.isFolder
                                  ? 'Cannot open folders in new tab'
                                  : !canDownloadDoc(doc)
                                    ? 'Open in new tab requires Edit permission'
                                    : 'Open in new browser tab'
                              }
                            ]
                          }}
                        />
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        )}

        {this.state.selectedDocuments.length > 0 && (
          <MessageBar messageBarType={MessageBarType.info} className={styles.selectionInfo}>
            {this.state.selectedDocuments.length} item(s) selected
          </MessageBar>
        )}

        {this.state.pagination.totalItems > 0 && (
          <div className={styles.paginationContainer}>
            <Pagination
              currentPage={this.state.pagination.currentPage}
              totalPages={this.state.pagination.totalPages}
              pageSize={this.state.pagination.pageSize}
              totalItems={this.state.pagination.totalItems}
              onPageChange={this.handlePageChange}
              onPageSizeChange={this.handlePageSizeChange}
              label="documents"
            />
          </div>
        )}

        <Dialog
          hidden={!this.state.showLogDialog}
          onDismiss={this._closeLogDialog}
          dialogContentProps={{
            type: DialogType.close,
            title: `Access Log - ${this.state.selectedLogDocument?.name || ''}`,
            showCloseButton: true
          }}
          modalProps={{
            isBlocking: false
          }}
          minWidth={900}
          maxWidth={1100}
          styles={{ main: { background: '#f2f6f9' } }}
        >
          <div className={styles.logDialogContent}>
            {this.state.loadingLogs ? (
              <Spinner size={SpinnerSize.medium} label="Loading access log..." />
            ) : this.state.logsError ? (
              <MessageBar messageBarType={MessageBarType.error}>
                {this.state.logsError}
              </MessageBar>
            ) : this.state.accessLogs.length === 0 ? (
              <MessageBar messageBarType={MessageBarType.info}>
                No log entries found for this item.
              </MessageBar>
            ) : (
              <div className={styles.logTableContainer}>
                <table className={styles.responsiveTable}>
                  <thead className={styles.responsiveTableHead}>
                    <tr className={styles.responsiveTableRow}>
                      {/* <th className={styles.responsiveTableHeadTitle}>Action</th> */}
                      <th className={styles.responsiveTableHeadTitle}>Permission</th>
                      <th className={styles.responsiveTableHeadTitle}>Expiry</th>
                      <th className={styles.responsiveTableHeadTitle}>Guest Emails</th>
                      <th className={styles.responsiveTableHeadTitle}>Shared By</th>
                      <th className={styles.responsiveTableHeadTitle}>Date</th>
                    </tr>
                  </thead>
                  <tbody className={styles.responsiveTableBody}>
                    {this.state.accessLogs.map((log, idx) => (
                      <tr key={`${log.actionDate ? log.actionDate.toISOString() : 'nodate'}-${idx}`} className={styles.responsiveTableRow}>
                        {/* <td>{log.actionType || '-'}</td> */}
                        <td>{log.permission || '-'}</td>
                        <td>{this._formatExpiry(log)}</td>

                        <td>
                          <span className={styles.emailEllipsis} title={log.guestEmail || '-'}>
                            {log.guestEmail || '-'}
                          </span>
                        </td>

                        <td>{log.sharedBy || '-'}</td>
                        <td>{this._formatDate(log.actionDate)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
          <DialogFooter>
            <DefaultButton text="Close" onClick={this._closeLogDialog} />
          </DialogFooter>
        </Dialog>

        <FilePreview
          isOpen={this.state.filePreview.isOpen}
          onDismiss={this._closeFilePreview}
          fileUrl={this.state.filePreview.fileUrl}
          fileName={this.state.filePreview.fileName}
          filePath={this.state.filePreview.filePath}
          siteUrl={this.props.context.pageContext.web.absoluteUrl}
          canDownload={this.state.filePreview.canDownload}
        />
      </div>
    );
  }
}
