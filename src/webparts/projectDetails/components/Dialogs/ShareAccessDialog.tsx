import * as React from 'react';
import { toast } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';
import styles from './ShareDocumentDialog.module.scss';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { ProjectExternalPortalService } from '../../services/ProjectExternalPortalService';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { ISharedDocument } from '../IProjectExternalPortalState';

export interface IShareAccessDialogProps {
  isOpen: boolean;
  onClose: () => void;
  onSuccess: () => void;
  context: WebPartContext;
  selectedDocuments: ISharedDocument[];
  service?: ProjectExternalPortalService;
}

export interface IShareAccessDialogState {
  // ── Local copy of selected docs so we can deselect conflicting ones ──
  activeDocuments: ISharedDocument[];
  autoRemovedDocs: string[];   // names of docs auto-removed due to conflicts

  selectedGuests: string[];
  searchQuery: string;
  accessDuration: number;
  customExpiryDate: string;
  customExpiryTime: string;
  sharing: boolean;
  checkingConflicts: boolean;
  conflictWarnings: Array<{ guestEmail: string; documentName: string; expiresOn: string }>;
  error?: string;
  success?: string;
  availableGuests: Array<{ email: string; name: string; }>;
  loadingGuests: boolean;
  azureAdUsers: Array<{ email: string; name: string; }>;
  searchingAzureAd: boolean;
  showDurationPresets: boolean;
  selectedPermission: 'Read' | 'Review' | 'Edit' | 'Admin';
}

export default class ShareAccessDialog extends React.Component<IShareAccessDialogProps, IShareAccessDialogState> {
  private portalService: ProjectExternalPortalService;

  constructor(props: IShareAccessDialogProps) {
    super(props);

    this.state = {
      activeDocuments: [...props.selectedDocuments],
      autoRemovedDocs: [],
      selectedGuests: [],
      searchQuery: '',
      accessDuration: 30,
      customExpiryDate: '',
      customExpiryTime: '23:59',
      sharing: false,
      checkingConflicts: false,
      conflictWarnings: [],
      availableGuests: [],
      loadingGuests: false,
      azureAdUsers: [],
      searchingAzureAd: false,
      showDurationPresets: false,
      selectedPermission: 'Read'
    };

    this.portalService = props.service || new ProjectExternalPortalService(props.context);
  }

  public componentDidUpdate(prevProps: IShareAccessDialogProps): void {
    // Reset local docs when dialog reopens with new selection
    if (this.props.isOpen && !prevProps.isOpen) {
      this.setState({
        activeDocuments: [...this.props.selectedDocuments],
        autoRemovedDocs: [],
        conflictWarnings: [],
        selectedGuests: [],
        error: undefined,
        success: undefined
      });
      if (!this.state.loadingGuests) {
        this._loadAvailableGuests();
      }
    }
  }

  private _loadAvailableGuests = async (): Promise<void> => {
    if (this.state.loadingGuests) return;

    this.setState({ loadingGuests: true });
    try {
      const guests = await this.portalService.getActiveGuestsForSharing();
      this.setState({ availableGuests: guests, loadingGuests: false });
    } catch (error) {
      const errorMessage = error.message || 'Failed to load guests';
      this.setState({
        error: errorMessage.includes('429') || errorMessage.includes('throttle')
          ? 'Too many requests. Please wait a moment and try again.'
          : errorMessage,
        loadingGuests: false
      });
    }
  }

  private _onGuestToggle = (email: string): void => {
    const { selectedGuests } = this.state;
    const index = selectedGuests.indexOf(email);
    const updated = index > -1
      ? selectedGuests.filter((_, i) => i !== index)
      : [...selectedGuests, email];
    this.setState({ selectedGuests: updated }, () => this._refreshConflicts(updated));
  }

  private _onSearchQueryChange = (_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ searchQuery: newValue || '' });
  }

  private _onAddUserFromSearch = (): void => {
    const { searchQuery, selectedGuests } = this.state;
    if (searchQuery && !selectedGuests.includes(searchQuery)) {
      const updated = [...selectedGuests, searchQuery];
      this.setState({ selectedGuests: updated, searchQuery: '' }, () => this._refreshConflicts(updated));
    }
  }

  /**
   * Re-check conflicts whenever selected guests change.
   * If conflicts found → automatically remove conflicting documents
   * from activeDocuments and show an info banner listing what was removed.
   */
  private _refreshConflicts = async (guests: string[]): Promise<void> => {
    const { activeDocuments } = this.state;

    if (guests.length === 0 || activeDocuments.length === 0) {
      this.setState({ conflictWarnings: [], autoRemovedDocs: [] });
      return;
    }

    this.setState({ checkingConflicts: true });

    try {
      const docPaths = activeDocuments.map(d => d.fileRef);
      const conflicts = await this.portalService.checkExistingActiveAccess(docPaths, guests);

      if (conflicts.length > 0) {
        // Collect conflicting document names (unique)
        const conflictingDocNames = new Set(conflicts.map(c => c.documentName));
        const conflictingDocPaths = new Set(
          activeDocuments
            .filter(d => conflictingDocNames.has(d.name))
            .map(d => d.fileRef)
        );

        // Remove conflicting docs from the active list
        const cleanedDocs = activeDocuments.filter(d => !conflictingDocPaths.has(d.fileRef));
        const removedNames  = activeDocuments
          .filter(d => conflictingDocPaths.has(d.fileRef))
          .map(d => d.name);

        this.setState({
          activeDocuments: cleanedDocs,
          autoRemovedDocs: removedNames,
          // Re-run conflicts on the cleaned list (should now be empty)
          conflictWarnings: [],
          checkingConflicts: false
        }, async () => {
          // If there are still remaining docs, re-check the cleaned set
          if (cleanedDocs.length > 0) {
            const remaining = await this.portalService.checkExistingActiveAccess(
              cleanedDocs.map(d => d.fileRef),
              guests
            );
            this.setState({ conflictWarnings: remaining });
          }
        });
      } else {
        this.setState({ conflictWarnings: [], autoRemovedDocs: [], checkingConflicts: false });
      }
    } catch {
      this.setState({ checkingConflicts: false });
    }
  }

  /** Manual deselect from the document list inside the dialog */
  private _removeDocument = (fileRef: string): void => {
    const { activeDocuments, selectedGuests } = this.state;
    const updated = activeDocuments.filter(d => d.fileRef !== fileRef);
    this.setState({ activeDocuments: updated }, () => this._refreshConflicts(selectedGuests));
  }

  private _onAccessDurationChange = (_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    this.setState({ accessDuration: option?.key as number || 30 });
  }

  private _onCustomExpiryDateChange = (_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ customExpiryDate: newValue || '' });
  }

  private _onCustomExpiryTimeChange = (_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ customExpiryTime: newValue || '23:59' });
  }

  private _getCustomExpiryDateTime = (): Date | null => {
    const { customExpiryDate, customExpiryTime } = this.state;
    if (!customExpiryDate) return null;
    return new Date(`${customExpiryDate}T${customExpiryTime || '23:59'}:00`);
  }

  private _onPermissionChange = (_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    this.setState({ selectedPermission: (option?.key as 'Read' | 'Review' | 'Edit' | 'Admin') || 'Read' });
  }

  private _onShareAccess = async (): Promise<void> => {
    const { selectedGuests, accessDuration, customExpiryDate, selectedPermission, activeDocuments, conflictWarnings } = this.state;

    if (selectedGuests.length === 0) {
      this.setState({ error: 'Please select at least one person to share with' });
      return;
    }

    if (activeDocuments.length === 0) {
      this.setState({ error: 'No documents left to share. All selected documents have active access for the selected users.' });
      return;
    }

    if (conflictWarnings.length > 0) {
      this.setState({ error: 'Please resolve all conflicts before sharing.' });
      return;
    }

    // Validate expiry
    let expiryDate: Date | null = null;
    if (accessDuration === 0) {
      expiryDate = null;
    } else if (accessDuration === -1) {
      if (!customExpiryDate) {
        this.setState({ error: 'Please select a custom expiry date' });
        return;
      }
      expiryDate = this._getCustomExpiryDateTime();
      if (!expiryDate || isNaN(expiryDate.getTime())) {
        this.setState({ error: 'Invalid custom expiry date/time' });
        return;
      }
      if (expiryDate <= new Date()) {
        this.setState({ error: 'Custom expiry date must be in the future' });
        return;
      }
    } else {
      expiryDate = new Date(Date.now() + (accessDuration * 24 * 60 * 60 * 1000));
    }

    this.setState({ sharing: true, error: undefined, success: undefined });

    try {
      for (const doc of activeDocuments) {
        if (doc.isFolder) {
          await this.portalService.shareFolderAccess(
            doc.fileRef, selectedGuests,
            accessDuration === 0 ? 0 : accessDuration,
            expiryDate || undefined, selectedPermission
          );
        } else {
          await this.portalService.shareDocumentAccess(
            doc.fileRef, selectedGuests,
            accessDuration === 0 ? 0 : accessDuration,
            expiryDate || undefined, selectedPermission
          );
        }
      }

      this.setState({
        success: `Access granted to ${activeDocuments.length} item(s) successfully!`,
        sharing: false
      });

      toast.success(`Access granted to ${activeDocuments.length} item(s) successfully!`, {
        position: 'top-right', autoClose: 3000
      });

      setTimeout(() => {
        this.props.onSuccess();
        this._resetForm();
      }, 1500);
    } catch (error) {
      this.setState({ error: `Failed to share access: ${error.message}`, sharing: false });
      toast.error(`Failed to share access: ${error.message}`, {
        position: 'top-right', autoClose: 5000
      });
    }
  }

  private _resetForm = (): void => {
    this.setState({
      activeDocuments: [...this.props.selectedDocuments],
      autoRemovedDocs: [],
      selectedGuests: [],
      searchQuery: '',
      accessDuration: 30,
      customExpiryDate: '',
      customExpiryTime: '23:59',
      error: undefined,
      success: undefined,
      azureAdUsers: [],
      selectedPermission: 'Read',
      conflictWarnings: [],
      checkingConflicts: false
    });
  }

  private _onClose = (): void => {
    if (!this.state.sharing) {
      this._resetForm();
      this.props.onClose();
    }
  }

  public render(): React.ReactElement<IShareAccessDialogProps> {
    const {
      activeDocuments, autoRemovedDocs,
      selectedGuests, searchQuery,
      accessDuration, customExpiryDate, customExpiryTime,
      sharing, checkingConflicts, conflictWarnings,
      error, success,
      availableGuests, loadingGuests
    } = this.state;

    const hasActiveConflicts  = conflictWarnings.length > 0;
    const todayStr            = new Date().toISOString().split('T')[0];
    const nowTimeStr          = new Date().toTimeString().slice(0, 5);
    const isToday             = customExpiryDate === todayStr;
    const minTime             = isToday ? nowTimeStr : '00:00';
    const customExpiryPreview = customExpiryDate
      ? new Date(`${customExpiryDate}T${customExpiryTime || '23:59'}:00`)
          .toLocaleString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit', hour12: true })
      : '';

    const durationOptions: IDropdownOption[] = [
      { key: 1,   text: '1 Day' },
      { key: 3,   text: '3 Days' },
      { key: 7,   text: '1 Week' },
      { key: 14,  text: '2 Weeks' },
      { key: 30,  text: '1 Month' },
      { key: 60,  text: '2 Months' },
      { key: 90,  text: '3 Months' },
      { key: 180, text: '6 Months' },
      { key: 365, text: '1 Year' },
      { key: 0,   text: 'No Expiration' },
      { key: -1,  text: 'Custom Date' }
    ];

    const permissionOptions: IDropdownOption[] = [
      { key: 'Read', text: 'Read (view only)' },
      { key: 'Edit', text: 'Edit (view, edit & download)' }
    ];

    return (
      <Dialog
        hidden={!this.props.isOpen}
        onDismiss={this._onClose}
        dialogContentProps={{
          type: DialogType.close,
          title: 'Share Document Access',
          showCloseButton: true
        }}
        modalProps={{ isBlocking: sharing, className: styles.dialog }}
        minWidth={550}
      >
        <div className={styles.dialogContent}>

          {success && <MessageBar messageBarType={MessageBarType.success}>{success}</MessageBar>}
          {error   && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}

          {/* ── Count banner ── */}
          <MessageBar messageBarType={MessageBarType.info}>
            {activeDocuments.length} document(s) selected for sharing
          </MessageBar>

          {/* ── Auto-removed banner ── */}
          {autoRemovedDocs.length > 0 && (
            <MessageBar
              messageBarType={MessageBarType.warning}
              isMultiline
              styles={{ root: { marginTop: 8 } }}
            >
              <strong>
                {autoRemovedDocs.length} file{autoRemovedDocs.length > 1 ? 's' : ''} auto-removed
              </strong>{' '}
              — already have active access for the selected user(s):
              <ul style={{ margin: '4px 0 0 16px', padding: 0 }}>
                {autoRemovedDocs.map((name, i) => (
                  <li key={i} style={{ fontSize: 13 }}>
                    <Icon iconName="Page" style={{ marginRight: 4, verticalAlign: 'middle', fontSize: 12 }} />
                    {name}
                  </li>
                ))}
              </ul>
            </MessageBar>
          )}

          {/* ── Selected documents list (with manual X to deselect) ── */}
          <div className={styles.selectedDocuments}>
            <p className={styles.label}>Selected Documents:</p>
            {activeDocuments.length === 0 ? (
              <MessageBar messageBarType={MessageBarType.warning}>
                No documents remaining. All selected files already have active access for the chosen user(s).
              </MessageBar>
            ) : (
              <ul className={styles.documentList}>
                {activeDocuments.map((doc, index) => (
                  <li key={index} className={styles.documentItem}
                    style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}
                  >
                    <span>
                      <Icon
                        iconName={doc.isFolder ? 'FabricFolder' : 'Page'}
                        className={styles.docIcon}
                      />
                      {doc.name}
                    </span>
                    {/* Manual deselect button */}
                    <Icon
                      iconName="Cancel"
                      title={`Remove ${doc.name}`}
                      onClick={() => !sharing && this._removeDocument(doc.fileRef)}
                      style={{
                        cursor: sharing ? 'not-allowed' : 'pointer',
                        color: '#a80000',
                        fontSize: 12,
                        marginLeft: 8,
                        flexShrink: 0
                      }}
                    />
                  </li>
                ))}
              </ul>
            )}
          </div>

          {/* ── People picker ── */}
          <div className={styles.peoplePickerSection}>
            <p className={styles.label}>Search and Add Users</p>
            <div style={{ display: 'flex', gap: '8px', marginBottom: '12px' }}>
              <TextField
                placeholder="Enter email address or search for user"
                value={searchQuery}
                onChange={this._onSearchQueryChange}
                style={{ flex: 1 }}
                disabled={sharing}
              />
              <PrimaryButton
                text="Add"
                onClick={this._onAddUserFromSearch}
                disabled={!searchQuery || sharing}
              />
            </div>
            {selectedGuests.length > 0 && (
              <div className={styles.selectedUsers}>
                <p className={styles.label}>Selected Users:</p>
                <div className={styles.userTags}>
                  {selectedGuests.map((email, index) => (
                    <div key={index} className={styles.userTag}>
                      <span>{email}</span>
                      <Icon
                        iconName="Cancel"
                        onClick={() => {
                          const newGuests = selectedGuests.filter((_, i) => i !== index);
                          this.setState({ selectedGuests: newGuests }, () => this._refreshConflicts(newGuests));
                        }}
                        style={{ cursor: 'pointer', marginLeft: '4px' }}
                      />
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>

          {/* ── Active guests checkboxes ── */}
          <div className={styles.guestsSection}>
            <p className={styles.label}>Or Select from Active Guests</p>
            {loadingGuests ? (
              <Spinner size={SpinnerSize.small} label="Loading guests..." />
            ) : availableGuests.length === 0 ? (
              <MessageBar messageBarType={MessageBarType.info}>
                No active guests found. Use the search above to add new people.
              </MessageBar>
            ) : (
              <div className={styles.guestsList}>
                {availableGuests.map((guest, index) => (
                  <div key={index} className={styles.guestItem}>
                    <Checkbox
                      label={`${guest.name} (${guest.email})`}
                      checked={selectedGuests.includes(guest.email)}
                      onChange={() => this._onGuestToggle(guest.email)}
                      disabled={sharing}
                    />
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* ── Duration ── */}
          <Dropdown
            label="Access Duration"
            placeholder="Select duration"
            options={durationOptions}
            selectedKey={accessDuration}
            onChange={this._onAccessDurationChange}
            className={styles.field}
            disabled={sharing}
          />

          {/* ── Permission ── */}
          <Dropdown
            label="Permission Level"
            placeholder="Select permission"
            options={permissionOptions}
            selectedKey={this.state.selectedPermission}
            onChange={this._onPermissionChange}
            className={styles.field}
            disabled={sharing}
          />

          {this.state.selectedPermission === 'Read' && (
            <MessageBar messageBarType={MessageBarType.info}>
              <strong>Read Only:</strong> Guest can view and download documents. Cannot edit, upload, share, or import.
            </MessageBar>
          )}
          {this.state.selectedPermission === 'Edit' && (
            <MessageBar messageBarType={MessageBarType.warning}>
              <strong>Edit:</strong> Guest can view, download, upload, share, and import documents.
            </MessageBar>
          )}

          {accessDuration > 0 && (
            <MessageBar messageBarType={MessageBarType.info}>
              Access will expire on{' '}
              {new Date(Date.now() + (accessDuration * 24 * 60 * 60 * 1000))
                .toLocaleString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit', hour12: true })}
            </MessageBar>
          )}
          {accessDuration === 0 && (
            <MessageBar messageBarType={MessageBarType.warning}>
              Access will not expire automatically. You'll need to revoke access manually.
            </MessageBar>
          )}

          {accessDuration === -1 && (
            <>
              <div style={{ display: 'flex', gap: '12px', alignItems: 'flex-start', marginBottom: '8px' }}>
                <div style={{ flex: '1 1 55%' }}>
                  <TextField
                    label="Expiry Date"
                    type="date"
                    value={customExpiryDate}
                    onChange={this._onCustomExpiryDateChange}
                    disabled={sharing}
                    min={todayStr}
                  />
                </div>
                <div style={{ flex: '1 1 45%' }}>
                  <TextField
                    label="Expiry Time"
                    type="time"
                    value={customExpiryTime}
                    onChange={this._onCustomExpiryTimeChange}
                    disabled={sharing || !customExpiryDate}
                    min={minTime}
                  />
                </div>
              </div>
              {customExpiryPreview && (
                <MessageBar messageBarType={MessageBarType.info}>
                  Access will expire on <strong>{customExpiryPreview}</strong>
                </MessageBar>
              )}
            </>
          )}

          <MessageBar messageBarType={MessageBarType.warning}>
            Access will be granted at the item level and will expire automatically on the selected date.
          </MessageBar>

          {/* ── Conflict checking spinner ── */}
          {checkingConflicts && (
            <Spinner size={SpinnerSize.small} label="Checking existing access..." style={{ margin: '8px 0' }} />
          )}

          {/* ── Remaining conflicts (after auto-remove) ── */}
          {!checkingConflicts && hasActiveConflicts && (
            <MessageBar messageBarType={MessageBarType.error} isMultiline styles={{ root: { marginTop: 8 } }}>
              <strong>Cannot share — remaining conflicts found:</strong>
              <ul style={{ margin: '4px 0 0 16px', padding: 0 }}>
                {conflictWarnings.map((c, i) => (
                  <li key={i} style={{ fontSize: 13 }}>
                    <strong>{c.guestEmail}</strong> → &quot;{c.documentName}&quot;{' '}
                    <span style={{ color: '#605e5c' }}>(expires: {c.expiresOn})</span>
                  </li>
                ))}
              </ul>
              <div style={{ marginTop: 6, fontSize: 12, color: '#605e5c' }}>
                These have been kept in the list but cannot be shared until access expires.
              </div>
            </MessageBar>
          )}

        </div>

        <DialogFooter>
          <PrimaryButton
            text={sharing ? 'Sharing...' : 'Share Access'}
            onClick={this._onShareAccess}
            disabled={
              selectedGuests.length === 0 ||
              activeDocuments.length === 0 ||
              sharing ||
              hasActiveConflicts ||
              checkingConflicts
            }
            className={styles.primaryButton}
          />
          <DefaultButton text="Cancel" onClick={this._onClose} disabled={sharing} />
        </DialogFooter>
      </Dialog>
    );
  }
}