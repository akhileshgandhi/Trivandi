import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  Dialog,
  DialogType,
  DialogFooter
} from '@fluentui/react/lib/Dialog';
import {
  PrimaryButton,
  DefaultButton
} from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import styles from './VersionHistory.module.scss';
import {
  IVersionHistoryProps,
  IFileVersion,
  ISelectedDocument
} from './IVersionHistory';
import { VersionHistoryService } from './VersionHistoryService';

const VersionHistory: React.FC<IVersionHistoryProps> = (props) => {
  const {
    context,
    isOpen,
    document,
    onDismiss,
    onVersionRestored,
    onVersionDeleted
  } = props;

  const [versions, setVersions] = useState<IFileVersion[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);
  const [sortBy, setSortBy] = useState<'date' | 'version'>('date');
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('desc');
  const [selectedVersionId, setSelectedVersionId] = useState<string | undefined>();
  const [deleteConfirmOpen, setDeleteConfirmOpen] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState<'single' | 'all'>('single');
  const [deleteVersionId, setDeleteVersionId] = useState<string | undefined>();
  const [deletingAll, setDeletingAll] = useState(false);
  const [restoreLoading, setRestoreLoading] = useState(false);
  const [service, setService] = useState<VersionHistoryService | null>(null);

  // Initialize service
  useEffect(() => {
    if (context) {
      setService(new VersionHistoryService(context));
    }
  }, [context]);

  // Load versions when dialog opens
  useEffect(() => {
    if (isOpen && document && service) {
      loadVersions();
    }
  }, [isOpen, document, service]);

  const loadVersions = async () => {
    if (!document || !service) return;

    setLoading(true);
    setError(undefined);

    try {
      // Try PnPjs first, fallback to REST
      let loadedVersions: IFileVersion[] = [];

      try {
        loadedVersions = await service.getFileVersionsFromPnPjs(
          document.serverRelativeUrl
        );
      } catch (pnpError) {
        console.warn('PnPjs failed, trying REST API:', pnpError);
        loadedVersions = await service.getFileVersionsFromREST(
          document.serverRelativeUrl
        );
      }

      setVersions(loadedVersions);
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Failed to load versions';
      setError(errorMessage);
      console.error('Version loading error:', err);
    } finally {
      setLoading(false);
    }
  };

  const handleSort = (sortType: 'date' | 'version') => {
    if (sortBy === sortType) {
      setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc');
    } else {
      setSortBy(sortType);
      setSortOrder('desc');
    }
  };

  const getSortedVersions = (): IFileVersion[] => {
    const sorted = [...versions];

    sorted.sort((a, b) => {
      let comparison = 0;

      if (sortBy === 'date') {
        comparison = new Date(a.created).getTime() - new Date(b.created).getTime();
      } else {
        comparison = a.versionNumber - b.versionNumber;
      }

      return sortOrder === 'asc' ? comparison : -comparison;
    });

    return sorted;
  };

  const handleDeleteClick = (versionId: string) => {
    setDeleteTarget('single');
    setDeleteVersionId(versionId);
    setDeleteConfirmOpen(true);
  };

  const confirmDelete = async () => {
    if (!document || !service) return;

    setDeleteConfirmOpen(false);

    try {
      if (deleteVersionId) {
        await service.deleteVersion(document.serverRelativeUrl, deleteVersionId);

        // Reload versions
        await loadVersions();

        if (onVersionDeleted) {
          onVersionDeleted();
        }
      }
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Failed to delete version';
      setError(errorMessage);
      console.error('Delete error:', err);
    }
  };

  const handleRestore = async (version: IFileVersion) => {
    if (!document || !service) return;

    setRestoreLoading(true);

    try {
      await service.restoreVersion(document.serverRelativeUrl, version.id);

      // Reload versions
      await loadVersions();

      if (onVersionRestored) {
        onVersionRestored(version);
      }
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Failed to restore version';
      setError(errorMessage);
      console.error('Restore error:', err);
    } finally {
      setRestoreLoading(false);
    }
  };

  const sortedVersions = getSortedVersions();
  console.log(sortedVersions,'sortedVersions');
  
  const currentVersion = versions.find(v => v.isCurrentVersion);

  const canDelete = (version: IFileVersion) => !version.isCurrentVersion;

  return (
    <>
      <Dialog
        hidden={!isOpen}
        onDismiss={onDismiss}
        dialogContentProps={{
          type: DialogType.close,
          title: `Version History • ${document?.name || 'File'}`,
          showCloseButton: true,
          closeButtonAriaLabel: 'Close'
        }}
        modalProps={{
          isBlocking: true,
          isDarkOverlay: false
        }}
        minWidth={1000}
        maxWidth={1300}
      >
        <div className={styles.versionHistoryContainer}>
          {/* Toolbar */}
          <div className={styles.versionHistoryToolbar}>
            <div className={styles.sorting}>
              <label>Sort by:</label>
              <button
                className={`${styles.sortButton} ${sortBy === 'date' ? styles.active : ''}`}
                onClick={() => handleSort('date')}
                title="Sort by modified date"
              >
                <Icon iconName="Calendar" className={styles.sortIcon} />
                Date {sortBy === 'date' && <Icon iconName={sortOrder === 'asc' ? 'SortUp' : 'SortDown'} className={styles.sortArrow} />}
              </button>
              <button
                className={`${styles.sortButton} ${sortBy === 'version' ? styles.active : ''}`}
                onClick={() => handleSort('version')}
                title="Sort by version number"
              >
                <Icon iconName="NumberField" className={styles.sortIcon} />
                Version {sortBy === 'version' && <Icon iconName={sortOrder === 'asc' ? 'SortUp' : 'SortDown'} className={styles.sortArrow} />}
              </button>
            </div>
            {/* <button className={styles.filterButton} title="Filter versions">
              <Icon iconName="Filter" />
            </button> */}
          </div>

          {/* Content */}
          <div className={styles.versionHistoryContent}>
            {loading ? (
              <div className={styles.loadingState}>
                <Spinner
                  size={SpinnerSize.large}
                  label="Loading version history..."
                  className={styles.spinner}
                />
                <p className={styles.loadingText}>
                  Fetching all versions from SharePoint...
                </p>
              </div>
            ) : error ? (
              <div className={styles.errorState}>
                <div>
                  <Icon
                    iconName="ErrorBadge"
                    className={styles.errorIcon}
                    style={{ marginRight: 8 }}
                  />
                  <span className={styles.errorTitle}>Error loading versions</span>
                </div>
                <div className={styles.errorMessage}>{error}</div>
              </div>
            ) : versions.length === 0 ? (
              <div className={styles.emptyState}>
                <Icon
                  iconName="FileCode"
                  className={styles.emptyIcon}
                />
                <div className={styles.emptyTitle}>No version history</div>
                <div className={styles.emptyMessage}>
                  This file hasn't been modified or version tracking is not enabled.
                </div>
              </div>
            ) : (
              <table className={styles.versionTable}>
                <thead>
                  <tr>
                    <th className={styles.versionCol}>Version</th>
                    <th className={styles.dateCol}>Date Modified</th>
                    <th className={styles.sizeCol}>Size</th>
                    <th className={styles.authorCol}>Modified By</th>
                    {/* <th className={styles.commentsCol}>Comments</th> */}
                    {/* <th className={styles.actionsCol}>Actions</th> */}
                  </tr>
                </thead>
                <tbody>
                  {sortedVersions.map((version) => (
                    <tr
                      key={version.id}
                      className={`${styles.row} ${
                        version.isCurrentVersion ? styles.currentVersion : ''
                      } ${selectedVersionId === version.id ? styles.selectedVersion : ''}`}
                      onClick={() => setSelectedVersionId(version.id)}
                    >
                      <td className={styles.versionCol}>
                        <Icon
                          iconName={service?.getFileIcon(document?.name || '') || 'Page'}
                          className={styles.fileIcon}
                        />
                        <span>{version.displayNumber}</span>
                        {version.isCurrentVersion && (
                          <span className={styles.currentBadge}>Current</span>
                        )}
                      </td>
                      <td className={styles.dateCol} title={service?.formatDate(version.created)}>
                        {service?.formatDate(version.created)}
                      </td>
                      <td className={styles.sizeCol}>
                        {service?.formatFileSize(version.size)}
                      </td>
                      <td className={styles.authorCol}>
                        <span title={version.modifiedByEmail}>
                          {version.modifiedBy}
                        </span>
                      </td>
                      {/* <td
                        className={`${styles.commentsCol} ${
                          version.comments ? styles.hasComments : ''
                        }`}
                        title={version.comments || 'No comments'}
                      >
                        {version.comments || '-'}
                      </td> */}
                      {/* <td className={styles.actionsCol}>
                        <div className={styles.actionIconGroup}>
                          <button
                            className={styles.actionIcon}
                            onClick={(e) => {
                              e.stopPropagation();
                              handleRestore(version);
                            }}
                            disabled={restoreLoading || version.isCurrentVersion}
                            title={version.isCurrentVersion ? '' : 'Restore this version'}
                          >
                            <Icon iconName="Undo" />
                          </button>
                          <button
                            className={styles.actionIcon}
                            title="Download version"
                          >
                            <Icon iconName="Download" />
                          </button>
                          <button
                            className={styles.actionIcon}
                            title="View version"
                          >
                            <Icon iconName="View" />
                          </button>
                        </div>
                        {version.isCurrentVersion && (
                          <div className={styles.restoreTooltip}>Restore this version</div>
                        )}
                      </td> */}
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>

          {/* Footer */}
          <DialogFooter>
            <DefaultButton
              text="Close"
              onClick={onDismiss}
            />
          </DialogFooter>
        </div>
      </Dialog>

      {/* Delete Confirmation Dialog */}
      <Dialog
        hidden={!deleteConfirmOpen}
        onDismiss={() => setDeleteConfirmOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Delete Version?',
          showCloseButton: true,
          closeButtonAriaLabel: 'Close'
        }}
        modalProps={{
          isBlocking: true,
          isDarkOverlay: false
        }}
        minWidth={400}
      >
        <div className={styles.deleteConfirmDialog}>
          <div className={styles.dialogContent}>
            <div style={{ display: 'flex', marginBottom: 16 }}>
              <Icon
                iconName="Warning"
                className={styles.warningIcon}
              />
              <div>
                <div className={styles.warningTitle}>
                  Delete this version?
                </div>
                <div className={styles.warningMessage}>
                  This version will be permanently deleted from the version history.
                  <strong> This action cannot be undone.</strong>
                </div>
              </div>
            </div>
          </div>
          <DialogFooter>
            <DefaultButton
              text="Cancel"
              onClick={() => setDeleteConfirmOpen(false)}
            />
            <PrimaryButton
              text="Delete Version"
              onClick={confirmDelete}
              styles={{
                root: {
                  background: '#e81123'
                }
              }}
            />
          </DialogFooter>
        </div>
      </Dialog>
    </>
  );
};

export default VersionHistory;
