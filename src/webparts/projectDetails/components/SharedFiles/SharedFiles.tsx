import * as React from 'react';
import { useState, useEffect, useRef, useCallback } from 'react';
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
import FilePreview from '../FilePreview/FilePreview';
import VersionHistory from '../VersionHistory/VersionHistory';
import { ISelectedDocument } from '../VersionHistory/IVersionHistory';
import { usePermissionStore } from '../../../../Permission/PermissionStore';
import Pagination from '../../../../shared/component/Pagination/Pagination';
import CommonShareComponent from '../../../../shared/Common/CommonShareComponent';
import GlobalLoader from '../../../../shared/component/GlobalLoader';

export interface ISharedFilesProps {
  context: WebPartContext;
  currentPath: string;
  breadcrumbs: IBreadcrumb[];
  onFolderNavigate: (path: string, breadcrumbs: IBreadcrumb[]) => void;
  onShowNewDocument: () => void;
  onShowShareAccess: (selectedDocs: ISharedDocument[]) => void;
  onShowImportDialog: () => void;
  onRefresh?: () => void;
  service?: ProjectExternalPortalService;
  isUserRestricted?: boolean;
  onFolderChange?: (path: string) => void;
}

interface IFilePreviewState {
  isOpen: boolean;
  fileUrl: string;
  fileName: string;
  filePath?: string;
  canDownload?: boolean;
}

interface IPaginationState {
  currentPage: number;
  pageSize: number;
  totalItems: number;
  totalPages: number;
}

const SharedFiles = React.forwardRef<any, ISharedFilesProps>((props, ref) => {
  const {
    context,
    currentPath,
    breadcrumbs,
    onFolderNavigate,
    onShowNewDocument,
    onShowShareAccess,
    onShowImportDialog,
    isUserRestricted,
    onFolderChange,
    service,
  } = props;

  const portalService = useRef<ProjectExternalPortalService>(
    service || new ProjectExternalPortalService(context)
  );
  const loadDebounceTimer = useRef<number | null>(null);
  const isLoadingRef = useRef(false);

  const [documents, setDocuments] = useState<ISharedDocument[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);
  const [currentPermission, setCurrentPermission] = useState<'Read' | 'Review' | 'Edit' | 'Admin' | null>(null);
  const [guestRole, setGuestRole] = useState<'Viewer' | 'Editor' | null>(null);
  const [totalDocumentsFound, setTotalDocumentsFound] = useState<number | undefined>(0);
  const [accessFilteredCount, setAccessFilteredCount] = useState<number | undefined>(0);

  const [showLogDialog, setShowLogDialog] = useState(false);
  const [loadingLogs, setLoadingLogs] = useState(false);
  const [logsError, setLogsError] = useState<string | undefined>(undefined);
  const [selectedLogDocument, setSelectedLogDocument] = useState<ISharedDocument | undefined>(undefined);
  const [accessLogs, setAccessLogs] = useState<ISharedFileAccessLog[]>([]);

  const [filePreview, setFilePreview] = useState<IFilePreviewState>({
    isOpen: false,
    fileUrl: '',
    fileName: '',
    filePath: '',
    canDownload: true,
  });

  const [pagination, setPagination] = useState<IPaginationState>({
    currentPage: 1,
    pageSize: 10,
    totalItems: 0,
    totalPages: 0,
  });

  const [showVersionHistory, setShowVersionHistory] = useState(false);
  const [selectedDocumentForVersions, setSelectedDocumentForVersions] = useState<ISelectedDocument | undefined>(undefined);

  const { canAdd: globalCanAdd, canEdit: globalCanEdit } = usePermissionStore.getState();

  // ─── Data Loading ──────────────────────────────────────────────────────────

  const loadDocuments = useCallback(async (page: number = pagination.currentPage, pageSize: number = pagination.pageSize): Promise<void> => {
    if (isLoadingRef.current) return;

    if (loadDebounceTimer.current) {
      clearTimeout(loadDebounceTimer.current);
    }

    isLoadingRef.current = true;
    setLoading(true);
    setTotalDocumentsFound(0);
    setAccessFilteredCount(0);

    try {
      const [{ items: docs, totalCount }, permission, role] = await Promise.all([
        portalService.current.getSharedDocumentsFromExternalLibrary(currentPath, page, pageSize),
        portalService.current.getUserPermissionForFolderPath(currentPath),
        portalService.current.getCurrentGuestRole(),
      ]);

      setDocuments(docs);
      setCurrentPermission(permission);
      setGuestRole(role);
      setLoading(false);
      setTotalDocumentsFound(docs.length);
      setAccessFilteredCount(0);
      setPagination({
        currentPage: page,
        pageSize,
        totalItems: totalCount,
        totalPages: Math.ceil(totalCount / pageSize),
      });
    } catch (err: any) {
      const errorMessage = err?.message || 'Failed to load documents';
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

      setError(displayError);
      setLoading(false);
      setDocuments([]);
      setCurrentPermission(null);
    } finally {
      isLoadingRef.current = false;
    }
  }, [currentPath, pagination.currentPage, pagination.pageSize]);

  // ─── Lifecycle ─────────────────────────────────────────────────────────────

  useEffect(() => {
    loadDocuments();
    return () => {
      if (loadDebounceTimer.current) {
        clearTimeout(loadDebounceTimer.current);
      }
    };
  }, []); // on mount

  useEffect(() => {
    if (onFolderChange) {
      onFolderChange(currentPath);
    }
    loadDocuments(1, pagination.pageSize);
  }, [currentPath]); // on path change

  // ─── Public refresh (exposed via ref if needed) ────────────────────────────

  const refreshDocuments = useCallback((): void => {
    if (loadDebounceTimer.current) {
      clearTimeout(loadDebounceTimer.current);
    }
    loadDebounceTimer.current = window.setTimeout(() => {
      loadDocuments(1);
    }, 300);
  }, [loadDocuments]);

  // ─── Pagination ────────────────────────────────────────────────────────────

  const handlePageChange = (page: number): void => {
    loadDocuments(page, pagination.pageSize);
  };

  const handlePageSizeChange = (pageSize: number): void => {
    loadDocuments(1, pageSize);
  };

  // ─── Navigation ───────────────────────────────────────────────────────────

  const onBreadcrumbClick = (item: any): void => {
    const index = breadcrumbs.findIndex(b => b.key === item.key);
    const newBreadcrumbs = breadcrumbs.slice(0, index + 1);
    onFolderNavigate(item.path, newBreadcrumbs);
  };

  const onFolderClick = (folder: ISharedDocument): void => {
    const newPath = `${currentPath}/${folder.name}`;
    const newBreadcrumbs = [
      ...breadcrumbs,
      { text: folder.name, key: folder.name, path: newPath },
    ];
    onFolderNavigate(newPath, newBreadcrumbs);
  };

  // ─── Formatting Helpers ────────────────────────────────────────────────────

  const formatDate = (date?: Date): string => {
    if (!date) return '';
    return new Date(date).toLocaleString('en-US', {
      month: '2-digit',
      day: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: true,
    });
  };

  const formatExpiry = (log: { accessDuration?: number; expiryDate?: Date; actionDate?: Date }): JSX.Element => {
    const now = new Date();
    let resolved: Date | null = null;

    if (log.expiryDate) {
      resolved = log.expiryDate;
    } else if (log.accessDuration && log.accessDuration > 0 && log.actionDate) {
      resolved = new Date(log.actionDate.getTime() + log.accessDuration * 24 * 60 * 60 * 1000);
    }

    if (!resolved) {
      return <span style={{ color: '#107c41', fontWeight: 600 }}>Permanent</span>;
    }

    const isExpired = now > resolved;
    const formatted = resolved.toLocaleString('en-US', {
      month: '2-digit', day: '2-digit', year: 'numeric',
      hour: '2-digit', minute: '2-digit', hour12: true,
    });

    if (isExpired) {
      return (
        <span>
          <span style={{
            display: 'inline-block', background: '#fde7e9', color: '#a80000',
            borderRadius: 10, padding: '8px 8px', fontSize: 11, fontWeight: 600, marginRight: 6,
          }}>Expired</span>
          <br />
          <span style={{ color: '#a80000', fontSize: 12 }}>{formatted}</span>
        </span>
      );
    }

    return <span style={{ color: '#323130', fontSize: 12 }}>{formatted}</span>;
  };

  const renderPermissionBadge = (permission?: string): JSX.Element => {
    if (!permission) return <span>-</span>;

    const permissionStyles: { [key: string]: { background: string; color: string } } = {
      Admin: { background: '#fee2e2', color: '#991b1b' },
      Edit: { background: '#dcfce7', color: '#166534' },
      Review: { background: '#fef9c3', color: '#713f12' },
      Read: { background: '#e5e7eb', color: '#374151' },
    };

    const style = permissionStyles[permission] || { background: '#e5e7eb', color: '#374151' };

    return (
      <span style={{
        display: 'inline-block',
        background: style.background,
        color: style.color,
        borderRadius: 4,
        padding: '4px 8px',
        fontSize: 12,
        fontWeight: 600,
      }}>
        {permission}
      </span>
    );
  };

  const renderEmailBadges = (emailString?: string): JSX.Element => {
    if (!emailString) return <span>-</span>;
    const emails = emailString.split(';').map(e => e.trim()).filter(e => e);
    if (emails.length === 0) return <span>-</span>;

    return (
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px' }}>
        {emails.map((email, idx) => (
          <span
            key={idx}
            style={{
              display: 'inline-block',
              background: '#e0f2fe',
              color: '#0369a1',
              borderRadius: 4,
              padding: '4px 8px',
              fontSize: 12,
              wordBreak: 'break-word',
            }}
            title={email}
          >
            {email}
          </span>
        ))}
      </div>
    );
  };

  // ─── Log Dialog ────────────────────────────────────────────────────────────

  const onViewLogClick = async (doc: ISharedDocument): Promise<void> => {
    setShowLogDialog(true);
    setLoadingLogs(true);
    setLogsError(undefined);
    setSelectedLogDocument(doc);
    setAccessLogs([]);

    try {
      const guestEmail = isUserRestricted
        ? (context.pageContext.user.email || undefined)
        : undefined;
      const logs = await portalService.current.getDocumentAccessLog(doc.fileRef, guestEmail);
      setAccessLogs(logs);
      setLoadingLogs(false);
    } catch (err) {
      setLoadingLogs(false);
      setLogsError((err as any)?.message || 'Failed to load access log');
    }
  };

  const closeLogDialog = (): void => {
    setShowLogDialog(false);
    setLoadingLogs(false);
    setLogsError(undefined);
    setSelectedLogDocument(undefined);
    setAccessLogs([]);
  };

  // ─── Version History ──────────────────────────────────────────────────────

  const handleOpenVersionHistory = (doc: ISharedDocument): void => {
    setSelectedDocumentForVersions({
      fileRef: doc.fileRef,
      name: doc.name,
      serverRelativeUrl: doc.fileRef,
      listItemId: doc.id,
    });
    setShowVersionHistory(true);
  };

  const handleCloseVersionHistory = (): void => {
    setShowVersionHistory(false);
    setSelectedDocumentForVersions(undefined);
  };

  const handleVersionRestored = (): void => {
    alert('Version restored successfully');
    loadDocuments();
  };

  const handleVersionDeleted = (): void => {
    alert('Version deleted successfully');
    loadDocuments();
  };

  // ─── File Preview ──────────────────────────────────────────────────────────

  const openFilePreview = (doc: ISharedDocument): void => {
    if (isUserRestricted && !doc.permission) {
      alert('You do not have access to this file.');
      return;
    }
    const canDownload = !isUserRestricted || doc.permission === 'Edit' || doc.permission === 'Admin';
    setFilePreview({
      isOpen: true,
      fileUrl: doc.fileRef,
      fileName: doc.name,
      filePath: doc.fileRef,
      canDownload,
    });
  };

  const closeFilePreview = (): void => {
    setFilePreview({ isOpen: false, fileUrl: '', fileName: '', filePath: '', canDownload: true });
  };

  // ─── Permission Helpers ────────────────────────────────────────────────────

  const canPreviewDoc = (doc: ISharedDocument): boolean => {
    if (!isUserRestricted) return true;
    return !!doc.permission;
  };

  const canDownloadDoc = (doc: ISharedDocument): boolean => {
    if (!isUserRestricted) return true;
    return doc.permission === 'Edit' || doc.permission === 'Admin';
  };

  const canEditCurrent = (currentPermission === 'Edit' || currentPermission === 'Admin') && globalCanEdit;
  const canAddCurrent = globalCanAdd;

  // ─── Breadcrumb Items ──────────────────────────────────────────────────────

  const breadcrumbItems = breadcrumbs.map(b => ({
    text: b.text,
    key: b.key,
    onClick: () => onBreadcrumbClick(b),
  }));

  // ─── Render ────────────────────────────────────────────────────────────────

  return (
    <div className={styles.sharedFiles}>
      {/* Header */}
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
                onShowImportDialog();
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
              onClick={onShowNewDocument}
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

      {/* Breadcrumb */}
      <div className={styles.toolbar}>
        <div className={styles.breadcrumbContainer}>
          <Breadcrumb
            items={breadcrumbItems}
            maxDisplayedItems={5}
            className={styles.breadcrumb}
          />
        </div>
        <div className={styles.actions} />
      </div>

      {/* Permission Info */}
      {!loading && !error && currentPermission && (
        <MessageBar
          messageBarType={
            currentPermission === 'Read' ? MessageBarType.warning :
              currentPermission === 'Review' ? MessageBarType.info :
                currentPermission === 'Admin' ? MessageBarType.success :
                  MessageBarType.info
          }
          isMultiline={false}
          className={styles.permissionInfo}
        >
          {currentPermission === 'Read' ? (
            <><strong>Read access:</strong> You can view documents in the portal but cannot download, upload, share, or import files.</>
          ) : currentPermission === 'Review' ? (
            <><strong>Review access:</strong> You can view and comment on documents but cannot download, upload, share, or import files.</>
          ) : currentPermission === 'Admin' ? (
            <><strong>Admin access:</strong> You have full control — view, edit, download, upload, share, import and manage guest permissions.</>
          ) : (
            <><strong>Edit access:</strong> You can view, edit, download, upload, share, and import documents.</>
          )}
        </MessageBar>
      )}

      {/* Table */}
      {loading && <GlobalLoader variant="bar" />}
      {error ? (
        <div className={styles.errorContainer}>
          <Icon iconName="Error" className={styles.errorIcon} />
          <p>{error}</p>
          <DefaultButton text="Retry" onClick={() => loadDocuments()} />
        </div>
      ) : (
        <div className={styles.tableContainer}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th className={styles.nameColumn}>Name</th>
                <th className={styles.modifiedColumn}>Modified</th>
                <th className={styles.byColumn}>Modified By</th>
                <th className={styles.actionColumn}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {documents.length === 0 ? (
                <tr>
                  <td colSpan={5} className={styles.emptyState}>
                    <Icon iconName="FabricFolder" className={styles.emptyIcon} />
                    {error && error.includes('does not exist') ? (
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
                    ) : currentPermission === null ? (
                      <>
                        <p>No documents found or you may not have access to this location.</p>
                        <p className={styles.emptySubtext}>
                          If you believe you should have access, please contact your project administrator or try refreshing the page.
                        </p>
                      </>
                    ) : currentPermission === 'Read' ? (
                      <>
                        <p>No documents found in External Shared Area.</p>
                        <p className={styles.emptySubtext}>
                          You have <strong>Read</strong> permission. Contact your project administrator to upload documents or request higher access.
                        </p>
                      </>
                    ) : currentPermission === 'Review' ? (
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
                documents.map((doc, index) => (
                  <tr key={index} className={styles.row}>
                    <td className={styles.nameColumn}>
                      <div className={styles.nameCell}>
                        {doc.isFolder ? (
                          <>
                            <Icon iconName="FabricFolder" className={styles.folderIcon} />
                            <span className={styles.folderLink} onClick={() => onFolderClick(doc)}>
                              {doc.name}
                            </span>
                          </>
                        ) : (
                          <>
                            <Icon iconName="Page" className={styles.fileIcon} />
                            <span className={styles.fileLink}>{doc.name}</span>
                          </>
                        )}
                      </div>
                    </td>
                    <td className={styles.modifiedColumn}>{formatDate(doc.modified)}</td>
                    <td className={styles.byColumn}>{doc.modifiedBy}</td>
                    <td className={styles.actionColumn}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
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
                                onClick: () => openFilePreview(doc),
                                disabled: doc.isFolder || !canPreviewDoc(doc),
                                title: doc.isFolder
                                  ? 'Cannot preview folders'
                                  : !canPreviewDoc(doc)
                                    ? 'You do not have access to preview this file'
                                    : `Preview ${doc.name}`,
                              },
                              // {
                              //   key: 'view',
                              //   text: 'View Log',
                              //   iconProps: { iconName: 'History' },
                              //   onClick: () => { onViewLogClick(doc).catch(() => {}); },
                              // },
                              {
                                key: 'version_history',
                                text: 'Version History',
                                iconProps: { iconName: 'History' },
                                onClick: () => handleOpenVersionHistory(doc),
                                disabled: doc.isFolder || isUserRestricted,
                                title: doc.isFolder ? 'Cannot view versions of folders' : 'View file version history'
                              },
                              {
                                key: 'share_native',
                                onRender: () => (
                                  <CommonShareComponent
                                    item={{
                                      ...doc,
                                      ServerRelativeUrl: doc.fileRef,
                                      IsFolder: doc.isFolder,
                                    }}
                                    context={context}
                                    buttonType="menuItem"
                                  />
                                ),
                                disabled:
                                  currentPermission === 'Read' ||
                                  currentPermission === 'Review' ||
                                  doc.permission === 'Read' ||
                                  doc.permission === 'Review' ||
                                  isUserRestricted,
                              },
                              {
                                key: 'manage_access',
                                text: 'Manage Access',
                                iconProps: { iconName: 'Permissions' },
                                onClick: () => onShowShareAccess([doc]),
                                disabled: isUserRestricted || (currentPermission !== 'Admin' && currentPermission !== 'Edit'),
                                title: 'Manage document permissions and access',
                              },
                              {
                                key: 'download',
                                text: 'Download',
                                iconProps: { iconName: 'Download' },
                                onClick: () => { window.open(doc.fileRef, '_blank'); },
                                disabled: doc.isFolder || !canDownloadDoc(doc),
                                title: doc.isFolder
                                  ? 'Cannot download folders - navigate into the folder to access individual files'
                                  : !canDownloadDoc(doc)
                                    ? 'Download requires Edit permission'
                                    : 'Download this file',
                              },
                              {
                                key: 'openInNewTab',
                                text: 'Open in New Tab',
                                iconProps: { iconName: 'OpenInNewWindow' },
                                onClick: () => {
                                  const siteUrl = context.pageContext.web.absoluteUrl;
                                  const ext = doc.name.split('.').pop()?.toLowerCase() || '';
                                  const officeExts = ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'pdf'];
                                  if (officeExts.includes(ext)) {
                                    const viewerUrl = `${siteUrl}/_layouts/15/Doc.aspx?sourcedoc=${encodeURIComponent(doc.fileRef)}&action=view`;
                                    window.open(viewerUrl, '_blank');
                                  } else {
                                    window.open(doc.fileRef, '_blank');
                                  }
                                },
                                disabled: doc.isFolder || !canDownloadDoc(doc),
                                title: doc.isFolder
                                  ? 'Cannot open folders in new tab'
                                  : !canDownloadDoc(doc)
                                    ? 'Open in new tab requires Edit permission'
                                    : 'Open in new browser tab',
                              },
                            ],
                          }}
                        />
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      )}

      {/* Pagination */}
      {pagination.totalItems > 0 && (
        <div className={styles.paginationContainer}>
          <Pagination
            currentPage={pagination.currentPage}
            totalPages={pagination.totalPages}
            pageSize={pagination.pageSize}
            totalItems={pagination.totalItems}
            onPageChange={handlePageChange}
            onPageSizeChange={handlePageSizeChange}
            label="documents"
          />
        </div>
      )}

      {/* Log Dialog */}
      <Dialog
        hidden={!showLogDialog}
        onDismiss={closeLogDialog}
        dialogContentProps={{
          type: DialogType.close,
          title: `View Access Log • ${selectedLogDocument?.name || ''}`,
          showCloseButton: true,
        }}
        modalProps={{ isBlocking: false }}
        minWidth={900}
        maxWidth={1100}
        styles={{ main: { background: '#f2f6f9' } }}
      >
        <div className={styles.logDialogContent}>
          {loadingLogs ? (
            <Spinner size={SpinnerSize.medium} label="Loading access log..." />
          ) : logsError ? (
            <MessageBar messageBarType={MessageBarType.error}>{logsError}</MessageBar>
          ) : accessLogs.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '40px 20px', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center' }}>
              <Icon iconName="People" style={{ fontSize: 48, color: '#c7c6c5', marginBottom: 16 }} />
              <p style={{ margin: '8px 0', fontSize: 16, fontWeight: 600, color: '#323130' }}>No log entries found</p>
              <p style={{ margin: '0', fontSize: 13, color: '#605e5c' }}>This document has not been accessed yet.</p>
            </div>
          ) : (
            <div>
              <div style={{ marginBottom: 20, paddingBottom: 12, borderBottom: '1px solid #e1e1e1' }}>
                <span style={{ fontSize: 14, fontWeight: 600, color: '#323130' }}>
                  👥 {accessLogs.length} Access Log{accessLogs.length !== 1 ? 's' : ''}
                </span>
              </div>
              <div className={styles.logTableContainer}>
                <table className={styles.responsiveTable}>
                  <thead className={styles.responsiveTableHead}>
                    <tr className={styles.responsiveTableRow}>
                      <th className={styles.responsiveTableHeadTitle}>Permission</th>
                      <th className={styles.responsiveTableHeadTitle}>Expiry</th>
                      <th className={styles.responsiveTableHeadTitle}>Guest Emails</th>
                      <th className={styles.responsiveTableHeadTitle}>Shared By</th>
                      <th className={styles.responsiveTableHeadTitle}>Date</th>
                    </tr>
                  </thead>
                  <tbody className={styles.responsiveTableBody}>
                    {accessLogs.map((log, idx) => (
                      <tr
                        key={`${log.actionDate ? log.actionDate.toISOString() : 'nodate'}-${idx}`}
                        className={styles.responsiveTableRow}
                      >
                        <td>{renderPermissionBadge(log.permission)}</td>
                        <td>{formatExpiry(log)}</td>
                        <td>{renderEmailBadges(log.guestEmail)}</td>
                        <td>{log.sharedBy || '-'}</td>
                        <td>{formatDate(log.actionDate)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
        <DialogFooter>
          <DefaultButton text="Close" onClick={closeLogDialog} />
        </DialogFooter>
      </Dialog>

      {/* File Preview */}
      <FilePreview
        isOpen={filePreview.isOpen}
        onDismiss={closeFilePreview}
        fileUrl={filePreview.fileUrl}
        fileName={filePreview.fileName}
        filePath={filePreview.filePath}
        siteUrl={context.pageContext.web.absoluteUrl}
        canDownload={filePreview.canDownload}
      />

      {/* Version History Modal */}
      {selectedDocumentForVersions && (
        <VersionHistory
          context={context}
          isOpen={showVersionHistory}
          document={selectedDocumentForVersions}
          onDismiss={handleCloseVersionHistory}
          onVersionRestored={handleVersionRestored}
          onVersionDeleted={handleVersionDeleted}
        />
      )}
    </div>
  );
});

export default SharedFiles;