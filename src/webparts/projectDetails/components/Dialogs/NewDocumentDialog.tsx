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

export interface INewDocumentDialogProps {
  isOpen: boolean;
  onClose: () => void;
  onSuccess: () => void;
  context: WebPartContext;
  currentPath: string;
  service?: ProjectExternalPortalService;
}

export interface ISelectedFile {
  file: File;
  path: string;
  displayName: string;
}

export interface INewDocumentDialogState {
  selectedFiles: ISelectedFile[];
  shareAfterUpload: boolean;
  selectedGuests: string[];
  accessDuration: number;
  customExpiryDate: string;
  uploading: boolean;
  uploadProgress: { [fileName: string]: number };
  error?: string;
  success?: string;
  availableGuests: Array<{ email: string; name: string; }>;
  loadingGuests: boolean;
}

export default class NewDocumentDialog extends React.Component<INewDocumentDialogProps, INewDocumentDialogState> {
  private portalService: ProjectExternalPortalService;
  private fileInputRef: React.RefObject<HTMLInputElement>;
  private folderInputRef: React.RefObject<HTMLInputElement>;

  constructor(props: INewDocumentDialogProps) {
    super(props);
    
    this.state = {
      selectedFiles: [],
      shareAfterUpload: false,
      selectedGuests: [],
      accessDuration: 30,
      customExpiryDate: '',
      uploading: false,
      uploadProgress: {},
      availableGuests: [],
      loadingGuests: false
    };

    this.portalService = props.service || new ProjectExternalPortalService(props.context);
    this.fileInputRef = React.createRef();
    this.folderInputRef = React.createRef();
  }

  public componentDidUpdate(prevProps: INewDocumentDialogProps): void {
    if (this.props.isOpen && !prevProps.isOpen && !this.state.loadingGuests) {
      this._loadAvailableGuests();
    }
  }

  private _loadAvailableGuests = async (): Promise<void> => {
    // Prevent duplicate calls
    if (this.state.loadingGuests) {
      return;
    }
    
    this.setState({ loadingGuests: true });
    try {
      const guests = await this.portalService.getActiveGuestsForSharing();
      this.setState({ 
        availableGuests: guests,
        loadingGuests: false 
      });
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

  private _onFileSelect = (): void => {
    if (this.fileInputRef.current) {
      this.fileInputRef.current.click();
    }
  }

  private _onFolderSelect = (): void => {
    if (this.folderInputRef.current) {
      this.folderInputRef.current.click();
    }
  }

  private _onFileChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const files = event.target.files;
    if (files && files.length > 0) {
      const selectedFiles: ISelectedFile[] = [];
      
      
      
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        selectedFiles.push({
          file,
          path: '', // Individual files go to current path
          displayName: file.name
        });
      }
      
      this.setState({ 
        selectedFiles: [...this.state.selectedFiles, ...selectedFiles]
      });
    }
    // Reset input value to allow re-selection of same files
    event.target.value = '';
  }

  private _onFolderChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const files = event.target.files;
    if (files && files.length > 0) {
      const selectedFiles: ISelectedFile[] = [];
      
      
      
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        // Get relative path from webkitRelativePath
        const relativePath = (file as any).webkitRelativePath || file.name;
        
        
        // Extract folder path from relative path
        const pathParts = relativePath.split('/');
        const fileName = pathParts.pop(); // Remove filename to get folder path
        const folderPath = pathParts.length > 0 ? pathParts.join('/') : '';
        
        selectedFiles.push({
          file,
          path: folderPath, // Store relative folder path (no leading slash)
          displayName: relativePath // Show full relative path to user
        });
      }
      
      
      this.setState({ 
        selectedFiles: [...this.state.selectedFiles, ...selectedFiles]
      });
    }
    // Reset input value to allow re-selection of same folder
    event.target.value = '';
  }

  private _removeFile = (index: number): void => {
    const { selectedFiles } = this.state;
    const updated = [...selectedFiles];
    updated.splice(index, 1);
    this.setState({ selectedFiles: updated });
  }

  private _clearAllFiles = (): void => {
    this.setState({ selectedFiles: [] });
    // Clear input values
    if (this.fileInputRef.current) {
      this.fileInputRef.current.value = '';
    }
    if (this.folderInputRef.current) {
      this.folderInputRef.current.value = '';
    }
  }

  private _onShareAfterUploadChange = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean): void => {
    this.setState({ shareAfterUpload: !!checked });
  }

  private _onGuestToggle = (email: string): void => {
    const { selectedGuests } = this.state;
    const index = selectedGuests.indexOf(email);
    
    if (index > -1) {
      selectedGuests.splice(index, 1);
    } else {
      selectedGuests.push(email);
    }
    
    this.setState({ selectedGuests: [...selectedGuests] });
  }

  private _onAccessDurationChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    this.setState({ accessDuration: option?.key as number || 30 });
  }

  private _onCustomExpiryDateChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ customExpiryDate: newValue || '' });
  }

  private _onUpload = async (): Promise<void> => {
    const { selectedFiles, shareAfterUpload, selectedGuests, accessDuration, customExpiryDate } = this.state;
    
    if (selectedFiles.length === 0) {
      this.setState({ error: 'Please select files or folders to upload' });
      return;
    }

    if (shareAfterUpload && selectedGuests.length === 0) {
      this.setState({ error: 'Please select at least one guest to share with' });
      return;
    }

    this.setState({ uploading: true, error: undefined, success: undefined, uploadProgress: {} });

    try {
      let successCount = 0;
      let errorCount = 0;
      const totalFiles = selectedFiles.length;
      
      for (let i = 0; i < selectedFiles.length; i++) {
        const { file, path } = selectedFiles[i];
        
        try {
          // Update progress
          const progress = Math.round(((i + 1) / totalFiles) * 100);
          this.setState(prevState => ({
            uploadProgress: {
              ...prevState.uploadProgress,
              [file.name]: progress
            }
          }));

          // Upload file to the appropriate path
          let targetPath = this.props.currentPath;
          if (path && path !== '/' && path.trim() !== '') {
            // For folder uploads, append the folder structure
            targetPath = this.props.currentPath.endsWith('/') 
              ? `${this.props.currentPath}${path}`
              : `${this.props.currentPath}/${path}`;
          }
          
          
          const uploadedFileRef = await this.portalService.uploadToExternalShareDocument(
            targetPath,
            file
          );

          // If sharing is enabled, apply permissions
          if (shareAfterUpload && selectedGuests.length > 0) {
            const expiryDate = accessDuration === -1 
              ? null 
              : customExpiryDate 
                ? new Date(customExpiryDate)
                : new Date(Date.now() + (accessDuration * 24 * 60 * 60 * 1000));

            await this.portalService.shareDocumentAccess(
              uploadedFileRef,
              selectedGuests,
              accessDuration,
              expiryDate || undefined
            );
          }
          
          successCount++;
        } catch (error) {
          
          errorCount++;
        }
      }

      const resultMessage = errorCount === 0 
        ? `Successfully uploaded ${successCount} file(s)!`
        : `Upload completed: ${successCount} successful, ${errorCount} failed`;

      this.setState({ 
        success: resultMessage,
        uploading: false 
      });

      toast.success(resultMessage, {
        position: "top-right",
        autoClose: 3000,
      });

      setTimeout(() => {
        this.props.onSuccess();
        this._resetForm();
      }, 1500);
    } catch (error) {
      this.setState({ 
        error: `Failed to upload files: ${error.message}`,
        uploading: false 
      });
      toast.error(`Failed to upload files: ${error.message}`, {
        position: "top-right",
        autoClose: 5000,
      });
    }
  }

  private _resetForm = (): void => {
    this.setState({
      selectedFiles: [],
      shareAfterUpload: false,
      selectedGuests: [],
      accessDuration: 30,
      customExpiryDate: '',
      uploadProgress: {},
      error: undefined,
      success: undefined
    });
  }

  private _onClose = (): void => {
    if (!this.state.uploading) {
      this._resetForm();
      this.props.onClose();
    }
  }

  public render(): React.ReactElement<INewDocumentDialogProps> {
    const { isOpen } = this.props;
    const { 
      selectedFiles,
      shareAfterUpload, 
      selectedGuests, 
      accessDuration,
      customExpiryDate,
      uploading, 
      uploadProgress,
      error,
      success,
      availableGuests,
      loadingGuests
    } = this.state;

    const durationOptions: IDropdownOption[] = [
      { key: 7, text: '7 Days' },
      { key: 30, text: '30 Days' },
      { key: 60, text: '60 Days' },
      { key: 90, text: '90 Days' },
      { key: -1, text: 'Custom Date' }
    ];

    return (
      <Dialog
        hidden={!isOpen}
        onDismiss={this._onClose}
        dialogContentProps={{
          type: DialogType.close,
          title: 'Upload New Document',
          showCloseButton: true
        }}
        modalProps={{
          isBlocking: uploading,
          className: styles.dialog
        }}
        minWidth={500}
      >
        <div className={styles.dialogContent}>
          {success && (
            <MessageBar messageBarType={MessageBarType.success}>
              {success}
            </MessageBar>
          )}

          {error && (
            <MessageBar messageBarType={MessageBarType.error}>
              {error}
            </MessageBar>
          )}

          <div className={styles.uploadSection}>
            <input
              ref={this.fileInputRef}
              type="file"
              multiple
              style={{ display: 'none' }}
              onChange={this._onFileChange}
              disabled={uploading}
            />
            <input
              ref={this.folderInputRef}
              type="file"
              {...({ webkitdirectory: true } as any)}
              multiple
              style={{ display: 'none' }}
              onChange={this._onFolderChange}
              disabled={uploading}
            />
            <div className={styles.uploadButtons}>
              <DefaultButton
                text="Choose Files"
                iconProps={{ iconName: 'Attach' }}
                onClick={this._onFileSelect}
                className={styles.uploadButton}
                disabled={uploading}
              />
              {/* <DefaultButton
                text="Choose Folder"
                iconProps={{ iconName: 'FabricFolderFill' }}
                onClick={this._onFolderSelect}
                className={styles.uploadButton}
                disabled={uploading}
              /> */}
              {selectedFiles.length > 0 && (
                <DefaultButton
                  text="Clear All"
                  iconProps={{ iconName: 'Clear' }}
                  onClick={this._clearAllFiles}
                  className={styles.clearButton}
                  disabled={uploading}
                />
              )}
            </div>
          </div>

          {selectedFiles.length > 0 && (
            <div className={styles.selectedFilesSection}>
              <p className={styles.label}>Selected Files ({selectedFiles.length})</p>
              <div className={styles.filesList}>
                {selectedFiles.map((selectedFile, index) => (
                  <div key={index} className={styles.fileItem}>
                    <div className={styles.fileInfo}>
                      <Icon iconName="Document" className={styles.fileIcon} />
                      <div className={styles.fileDetails}>
                        <span className={styles.fileName}>{selectedFile.displayName}</span>
                        <span className={styles.fileSize}>
                          {(selectedFile.file.size / 1024).toFixed(1)} KB
                        </span>
                      </div>
                    </div>
                    {uploading && uploadProgress[selectedFile.file.name] && (
                      <div className={styles.progressInfo}>
                        <span>{uploadProgress[selectedFile.file.name]}%</span>
                      </div>
                    )}
                    {!uploading && (
                      <DefaultButton
                        iconProps={{ iconName: 'Delete' }}
                        onClick={() => this._removeFile(index)}
                        className={styles.removeButton}
                        disabled={uploading}
                      />
                    )}
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* <Checkbox
            label="Share with guests after upload"
            checked={shareAfterUpload}
            onChange={this._onShareAfterUploadChange}
            className={styles.checkbox}
            disabled={uploading}
          /> */}

          {shareAfterUpload && (
            <>
              <div className={styles.guestsSection}>
                <p className={styles.label}>Select Guests to Share With</p>
                {loadingGuests ? (
                  <Spinner size={SpinnerSize.small} label="Loading guests..." />
                ) : availableGuests.length === 0 ? (
                  <MessageBar messageBarType={MessageBarType.info}>
                    No active guests found. Please invite guests first.
                  </MessageBar>
                ) : (
                  <div className={styles.guestsList}>
                    {availableGuests.map((guest, index) => (
                      <div key={index} className={styles.guestItem}>
                        <Checkbox
                          label={`${guest.name} (${guest.email})`}
                          checked={selectedGuests.includes(guest.email)}
                          onChange={() => this._onGuestToggle(guest.email)}
                          disabled={uploading}
                        />
                      </div>
                    ))}
                  </div>
                )}
              </div>

              <Dropdown
                label="Access Duration"
                placeholder="Select duration"
                options={durationOptions}
                selectedKey={accessDuration}
                onChange={this._onAccessDurationChange}
                className={styles.field}
                disabled={uploading}
              />

              {accessDuration === -1 && (
                <TextField
                  label="Custom Expiry Date"
                  type="date"
                  value={customExpiryDate}
                  onChange={this._onCustomExpiryDateChange}
                  className={styles.field}
                  disabled={uploading}
                />
              )}

              <MessageBar messageBarType={MessageBarType.info}>
                Access to this file will expire automatically on the selected date.
              </MessageBar>
            </>
          )}
        </div>

        <DialogFooter>
          <PrimaryButton
            text={uploading ? `Uploading... (${Object.keys(uploadProgress).length}/${selectedFiles.length})` : "Upload"}
            onClick={this._onUpload}
            disabled={selectedFiles.length === 0 || uploading || (shareAfterUpload && selectedGuests.length === 0)}
            className={styles.primaryButton}
          />
          <DefaultButton
            text="Cancel"
            onClick={this._onClose}
            disabled={uploading}
          />
        </DialogFooter>
      </Dialog>
    );
  }
}
