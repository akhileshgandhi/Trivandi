import * as React from 'react';
import styles from './ActiveGuests.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IGuestUser } from '../IProjectExternalPortalState';
import { Icon } from '@fluentui/react/lib/Icon';
import { DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { ProjectExternalPortalService } from '../../services/ProjectExternalPortalService';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { ContextualMenu, IContextualMenuItem } from '@fluentui/react/lib/ContextualMenu';
import FilePreview from '../FilePreview/FilePreview';

export interface IActiveGuestsProps {
  context: WebPartContext;
  projectId?: any;
  refreshTrigger?: number; // Timestamp to trigger refresh
  isUserRestricted?: boolean; // Whether current user is restricted from accessing reports
}

export interface IActiveGuestsState {
  guests: IGuestUser[];
  loading: boolean;
  reportGenerating: boolean;
  error?: string;
  filePreview: {
    isOpen: boolean;
    fileUrl: string;
    fileName: string;
    filePath?: string;
  };
  guestFiles: { [guestEmail: string]: Array<{ fileName: string; filePath: string; fileUrl: string; }> };
  contextMenu: {
    isVisible: boolean;
    target?: any;
    fileInfo?: { fileName: string; filePath: string; fileUrl: string; };
  };
}

class ActiveGuestsComponent extends React.Component<IActiveGuestsProps, IActiveGuestsState> {
  private portalService: ProjectExternalPortalService;

  constructor(props: IActiveGuestsProps) {
    super(props);
    
    this.state = {
      guests: [],
      loading: false,
      reportGenerating: false,
      filePreview: {
        isOpen: false,
        fileUrl: '',
        fileName: '',
        filePath: ''
      },
      guestFiles: {},
      contextMenu: {
        isVisible: false
      }
    };

    this.portalService = new ProjectExternalPortalService(props.context);
  }

  public componentDidMount(): void {
    this._loadGuests();
  }

  public componentDidUpdate(prevProps: IActiveGuestsProps): void {
    // Refresh guests if projectId changes or if refreshTrigger changes
    if (prevProps.projectId !== this.props.projectId || 
        prevProps.refreshTrigger !== this.props.refreshTrigger) {
      console.log('ActiveGuests: props changed, refreshing guests with skipCache=true');
      this._loadGuests(true); // Skip cache when props change
    }
  }

  public refreshGuests = (): void => {
    console.log('ActiveGuests: refreshGuests method called');
    this._loadGuests(true); // Skip cache when manually refreshing
  }

  private _loadGuests = async (skipCache: boolean = false): Promise<void> => {
    console.log('ActiveGuests: _loadGuests called, skipCache:', skipCache);
    this.setState({ loading: true });
    try {
      const guests = await this.portalService.getActiveGuests(this.props.projectId, skipCache);
      console.log('ActiveGuests: Loaded guests:', guests);
      
      // Load guest files for each guest
      const guestFiles: { [guestEmail: string]: Array<{ fileName: string; filePath: string; fileUrl: string; }> } = {};
      for (const guest of guests) {
        try {
          const files = await this._loadGuestFiles(guest.email);
          guestFiles[guest.email] = files;
        } catch (error) {
          console.error(`Error loading files for guest ${guest.email}:`, error);
          guestFiles[guest.email] = [];
        }
      }
      
      this.setState({ guests, guestFiles, loading: false, error: undefined });
    } catch (error) {
      console.error('Error loading guests:', error);
      this.setState({ guests: [], guestFiles: {}, loading: false, error: undefined }); // Show empty state instead of error
    }
  }

  private _getInitials = (name: string): string => {
    const parts = name.split(' ');
    if (parts.length >= 2) {
      return parts[0][0] + parts[1][0];
    }
    return name.substring(0, 2);
  }

  private _isExpired = (guest: IGuestUser): boolean => {
    if (guest.status === 'Expired' || guest.status === 'Revoked') return true;
    if (guest.accessExpiryDate) {
      return new Date(guest.accessExpiryDate) < new Date();
    }
    return false;
  }

  private _formatExpiryDate = (date: Date): string => {
    return new Date(date).toLocaleString('en-US', {
      month: 'short',
      day: 'numeric',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: true
    });
  }

  private _getTimeAgo = (date: Date): string => {
    if (!date) return '';
    const now = new Date();
    const diff = now.getTime() - new Date(date).getTime();
    const hours = Math.floor(diff / (1000 * 60 * 60));
    
    if (hours < 1) return 'Just now';
    if (hours < 24) return `${hours} hours ago`;
    const days = Math.floor(hours / 24);
    return `${days} days ago`;
  }

  private _loadGuestFiles = async (guestEmail: string): Promise<Array<{ fileName: string; filePath: string; fileUrl: string; }>> => {
    try {
      const projectId = this.props.projectId ? Number(this.props.projectId) : undefined;
      const reportRows = await this.portalService.getGuestAccessReport(projectId, true);
      
      // Filter files for this specific guest and remove duplicates
      const guestReports = reportRows.filter(row => row.guestEmail === guestEmail);
      const uniqueFiles = new Map();
      
      guestReports.forEach(row => {
        if (row.fileName && row.filePath) {
          const key = `${row.fileName}-${row.filePath}`;
          if (!uniqueFiles.has(key)) {
            uniqueFiles.set(key, {
              fileName: row.fileName,
              filePath: row.filePath,
              fileUrl: this._constructFileUrl(row.filePath) // Construct URL from path
            });
          }
        }
      });
      
      return Array.from(uniqueFiles.values());
    } catch (error) {
      console.error('Error loading guest files:', error);
      return [];
    }
  }

  private _constructFileUrl = (filePath: string): string => {
    if (!filePath) return '';
    
    console.log('ActiveGuests: Original file path:', filePath);
    
    // If it's already a full URL, return as is
    if (filePath.startsWith('http://') || filePath.startsWith('https://')) {
      console.log('ActiveGuests: Path is already full URL:', filePath);
      return filePath;
    }
    
    // Get site URL
    const siteUrl = this.props.context.pageContext.web.absoluteUrl;
    console.log('ActiveGuests: Site URL:', siteUrl);
    
    // Simple construction - just combine site URL with file path
    let finalUrl: string;
    if (filePath.startsWith('/')) {
      // Remove leading slash and construct URL  
      finalUrl = `${siteUrl}${filePath}`;
    } else {
      finalUrl = `${siteUrl}/${filePath}`;
    }
    
    console.log('ActiveGuests: Final constructed URL:', finalUrl);
    return finalUrl;
  }

  private _openFilePreview = (fileUrl: string, fileName: string, filePath?: string): void => {
    console.log('ActiveGuests: Opening file preview:', {
      fileName,
      fileUrl,
      filePath
    });
    
    this.setState({
      filePreview: {
        isOpen: true,
        fileUrl,
        fileName,
        filePath
      }
    });
  }

  private _closeFilePreview = (): void => {
    this.setState({
      filePreview: {
        isOpen: false,
        fileUrl: '',
        fileName: '',
        filePath: ''
      }
    });
  }

  private _showContextMenu = (event: React.MouseEvent<any>, fileInfo: { fileName: string; filePath: string; fileUrl: string; }): void => {
    event.preventDefault();
    this.setState({
      contextMenu: {
        isVisible: true,
        target: event.currentTarget,
        fileInfo
      }
    });
  }

  private _hideContextMenu = (): void => {
    this.setState({
      contextMenu: {
        isVisible: false
      }
    });
  }

  private _getContextMenuItems = (): IContextualMenuItem[] => {
    const { contextMenu } = this.state;
    if (!contextMenu.fileInfo) return [];

    const { fileName, filePath, fileUrl } = contextMenu.fileInfo;

    return [
      {
        key: 'preview',
        text: 'Preview',
        iconProps: { iconName: 'View' },
        onClick: () => {
          this._openFilePreview(fileUrl, fileName, filePath);
          this._hideContextMenu();
        }
      },
      {
        key: 'openInNewTab',
        text: 'Open in New Tab',
        iconProps: { iconName: 'OpenInNewWindow' },
        onClick: () => {
          window.open(fileUrl, '_blank');
          this._hideContextMenu();
        }
      },
      {
        key: 'download',
        text: 'Download',
        iconProps: { iconName: 'Download' },
        onClick: () => {
          const link = document.createElement('a');
          link.href = fileUrl;
          link.download = fileName;
          link.target = '_blank';
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
          this._hideContextMenu();
        }
      },
      {
        key: 'copyPath',
        text: 'Copy Path',
        iconProps: { iconName: 'Copy' },
        onClick: () => {
          navigator.clipboard.writeText(filePath).then(() => {
            // Could show a toast notification here
            console.log('File path copied to clipboard');
          }).catch(() => {
            console.error('Failed to copy file path');
          });
          this._hideContextMenu();
        }
      }
    ];
  }

  private _onManageAccess = (): void => {
    this._downloadAccessReport().catch((error) => {
      console.error('Error triggering access report:', error);
    });
  }

  private _escapeCsvValue = (value: string): string => {
    const safeValue = value || '';
    if (safeValue.includes(',') || safeValue.includes('"') || safeValue.includes('\n')) {
      return `"${safeValue.replace(/"/g, '""')}"`;
    }
    return safeValue;
  }

  private _formatDateForCsv = (value?: Date): string => {
    if (!value) {
      return '';
    }

    return new Date(value).toLocaleString('en-US', {
      month: '2-digit',
      day: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: true
    });
  }

  private _downloadAccessReport = async (): Promise<void> => {
    if (this.state.reportGenerating) {
      return;
    }

    // Check if user is restricted
    if (this.props.isUserRestricted) {
      alert('You do not have permission to export access reports.');
      return;
    }

    this.setState({ reportGenerating: true });

    try {
      const projectId = this.props.projectId ? Number(this.props.projectId) : undefined;
      const reportRows = await this.portalService.getGuestAccessReport(projectId, true);

      if (!reportRows.length) {
        alert('No access data found for active guests.');
        this.setState({ reportGenerating: false });
        return;
      }

      const headers = [
        'Guest Name',
        'Guest Email',
        // 'Role',
        // 'Company',
        'Status',
        'File Name',
        'File Path',
        'Permission',
        'Shared On',
        'Last Access'
      ];

      const csvRows = reportRows.map((row) => ([
        row.guestName,
        row.guestEmail,
        // row.guestRole,
        // row.guestCompany || '',
        row.guestStatus,
        row.fileName || '',
        row.filePath || '',
        row.permission,
        this._formatDateForCsv(row.sharedOn),
        this._formatDateForCsv(row.guestLastAccess)
      ].map(this._escapeCsvValue).join(',')));

      const csvContent = [headers.join(','), ...csvRows].join('\n');
      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      const dateStamp = new Date().toISOString().slice(0, 10);

      link.href = url;
      link.download = `external-portal-access-report-${dateStamp}.csv`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
    } catch (error) {
      console.error('Error generating access report:', error);
      const errorMessage = error.message?.includes('permission') || error.message?.includes('access denied')
        ? 'You do not have permission to export access reports.'
        : `Failed to generate access report: ${error.message || 'Unknown error'}`;
      alert(errorMessage);
    } finally {
      this.setState({ reportGenerating: false });
    }
  }

  public render(): React.ReactElement<IActiveGuestsProps> {
    const { guests, loading, error, reportGenerating, filePreview, guestFiles, contextMenu } = this.state;

    return (
      <div className={styles.activeGuests}>
        <div className={styles.header}>
          <h4 className={styles.title}>Active Guests</h4>
          <span className={styles.badge}>{guests.length} Active</span>
        </div>

        {loading ? (
          <div className={styles.loadingContainer}>
            <Spinner size={SpinnerSize.medium} />
          </div>
        ) : error ? (
          <div className={styles.errorContainer}>
            <Icon iconName="Error" className={styles.errorIcon} />
            <p>{error}</p>
          </div>
        ) : (
          <>
            <div className={styles.guestsContainer}>
              {guests.length === 0 ? (
                <div className={styles.emptyState}>
                  <Icon iconName="People" className={styles.emptyIcon} />
                  <p>No active guests. Click "Invite Guest" to add external users.</p>
                </div>
              ) : (
                guests.map((guest, index) => {
                  const files = guestFiles[guest.email] || [];
                  const isExpired = this._isExpired(guest);
                  return (
                    <div key={index} className={`${styles.guestCard}${isExpired ? ` ${styles.expiredCard}` : ''}`}>
                      <div className={`${styles.avatar}${isExpired ? ` ${styles.expiredAvatar}` : ''}`}>
                        {this._getInitials(guest.title)}
                      </div>
                      <div className={styles.guestInfo}>
                        <div className={styles.guestNameRow}>
                          <h5 className={styles.guestName}>{guest.title}</h5>
                          {isExpired && (
                            <span className={styles.expiredBadge}>Expired</span>
                          )}
                        </div>
                        <p className={styles.guestCompany}>{guest.company} • {guest.email}</p>
                        <p className={styles.guestRole}>
                          <span className={styles.lastAccess}>{this._getTimeAgo(guest.lastAccessDate)}</span>
                          {/* {files.length > 0 && (
                            <>
                              <span className={styles.separator}> • </span>
                              <span className={styles.fileCount}>{files.length} file(s)</span>
                            </>
                          )} */}
                        </p>
                        {guest.accessExpiryDate && (
                          <p className={`${styles.expiryInfo}${isExpired ? ` ${styles.expiryInfoExpired}` : ''}`}>
                            <Icon iconName="Clock" className={styles.expiryIcon} />
                            {isExpired ? 'Expired: ' : 'Expires: '}
                            {this._formatExpiryDate(guest.accessExpiryDate)}
                          </p>
                        )}
                        {/* {files.length > 0 && (
                          <div className={styles.filesSection}>
                            <div className={styles.filesHeader}>
                              <Icon iconName="Document" className={styles.filesIcon} />
                              <span>Accessible Files:</span>
                            </div>
                            <div className={styles.filesList}>
                              {files.slice(0, 3).map((file, fileIndex) => (
                                <div key={fileIndex} className={styles.fileItem}>
                                  <IconButton
                                    iconProps={{ iconName: 'Preview' }}
                                    title={`Preview ${file.fileName}`}
                                    onClick={() => this._openFilePreview(file.fileUrl, file.fileName, file.filePath)}
                                    className={styles.previewButton}
                                  />
                                  <span 
                                    className={styles.fileName} 
                                    title={file.fileName}
                                    onContextMenu={(e) => this._showContextMenu(e, file)}
                                  >
                                    {file.fileName.length > 30 ? `${file.fileName.substring(0, 30)}...` : file.fileName}
                                  </span>
                                  <IconButton
                                    iconProps={{ iconName: 'MoreVertical' }}
                                    title="More options"
                                    onClick={(e) => this._showContextMenu(e, file)}
                                    className={styles.moreButton}
                                  />
                                </div>
                              ))}
                              {files.length > 3 && (
                                <div className={styles.moreFiles}>
                                  <span>+{files.length - 3} more files</span>
                                </div>
                              )}
                            </div>
                          </div>
                        )} */}
                      </div>
                    </div>
                  );
                })
              )}
            </div>

            {guests.length > 0 && !this.props.isUserRestricted && (
              <div className={styles.footer}>
                <DefaultButton 
                  text={reportGenerating ? 'Generating Report...' : 'Export Access Report'}
                  disabled={reportGenerating}
                  onClick={this._onManageAccess}
                  className={styles.manageButton}
                />
              </div>
            )}
          </>
        )}
        
        <FilePreview
          isOpen={filePreview.isOpen}
          onDismiss={this._closeFilePreview}
          fileUrl={filePreview.fileUrl}
          fileName={filePreview.fileName}
          filePath={filePreview.filePath}
          siteUrl={this.props.context.pageContext.web.absoluteUrl}
        />
        
        {contextMenu.isVisible && (
          <ContextualMenu
            items={this._getContextMenuItems()}
            target={contextMenu.target}
            onDismiss={this._hideContextMenu}
            isBeakVisible={true}
            directionalHint={6} // DirectionalHint.bottomLeftEdge
          />
        )}
      </div>
    );
  }
}

// Export the class component directly since we're using class-based refs
export default ActiveGuestsComponent;
