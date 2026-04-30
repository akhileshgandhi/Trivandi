import * as React from 'react';
import { toast } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';
import styles from './ShareDocumentDialog.module.scss';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Icon } from '@fluentui/react/lib/Icon';
import { ProjectExternalPortalService } from '../../services/ProjectExternalPortalService';

export interface IShareDocumentDialogProps {
  isOpen: boolean;
  onClose: () => void;
  context: WebPartContext;
  currentPath: string;
}

export interface IShareDocumentDialogState {
  externalSharing: string;
  type: string;
  name: string;
  selectedFile: File | null;
  selectedUsers: string[];
  accessDuration: number;
  uploading: boolean;
  error?: string;
}

export default class ShareDocumentDialog extends React.Component<IShareDocumentDialogProps, IShareDocumentDialogState> {
  private portalService: ProjectExternalPortalService;
  private fileInputRef: React.RefObject<HTMLInputElement>;

  constructor(props: IShareDocumentDialogProps) {
    super(props);
    
    this.state = {
      externalSharing: '',
      type: 'ExternalSharing',
      name: '',
      selectedFile: null,
      selectedUsers: [],
      accessDuration: 30,
      uploading: false
    };

    this.portalService = new ProjectExternalPortalService(props.context);
    this.fileInputRef = React.createRef();
  }

  private _onExternalSharingChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ externalSharing: newValue || '' });
  }

  private _onTypeChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    this.setState({ type: option?.key as string || '' });
  }

  private _onNameChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ name: newValue || '' });
  }

  private _onFileSelect = (): void => {
    if (this.fileInputRef.current) {
      this.fileInputRef.current.click();
    }
  }

  private _onFileChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const file = event.target.files?.[0];
    if (file) {
      this.setState({ 
        selectedFile: file,
        name: file.name 
      });
    }
  }

  private _onAccessDurationChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    this.setState({ accessDuration: option?.key as number || 30 });
  }

  private _onUploadAndShare = async (): Promise<void> => {
    const { externalSharing, selectedFile, selectedUsers, accessDuration } = this.state;
    
    if (!externalSharing || !selectedFile) {
      this.setState({ error: 'Please fill in all required fields and select a file' });
      return;
    }

    this.setState({ uploading: true, error: undefined });

    try {
      await this.portalService.uploadAndShareDocument(
        this.props.currentPath,
        selectedFile,
        externalSharing,
        selectedUsers,
        accessDuration
      );
      
      toast.success('Document uploaded and shared successfully!', {
        position: "top-right",
        autoClose: 3000,
      });
      
      this.props.onClose();
      // Refresh the documents list
    } catch (error) {
      this.setState({ error: error.message, uploading: false });
      toast.error(`Failed to upload and share: ${error.message}`, {
        position: "top-right",
        autoClose: 5000,
      });
    }
  }

  private _onShareAccess = async (): Promise<void> => {
    const { externalSharing, selectedUsers, accessDuration } = this.state;
    
    if (!externalSharing || selectedUsers.length === 0) {
      this.setState({ error: 'Please select users to share with' });
      return;
    }

    this.setState({ uploading: true, error: undefined });

    try {
      // Share access logic
      await this.portalService.shareAccess(
        externalSharing,
        selectedUsers,
        accessDuration
      );
      
      toast.success('Access shared successfully!', {
        position: "top-right",
        autoClose: 3000,
      });
      
      this.props.onClose();
    } catch (error) {
      this.setState({ error: error.message, uploading: false });
      toast.error(`Failed to share access: ${error.message}`, {
        position: "top-right",
        autoClose: 5000,
      });
    }
  }

  public render(): React.ReactElement<IShareDocumentDialogProps> {
    const { isOpen, onClose } = this.props;
    const { externalSharing, type, name, selectedFile, accessDuration, uploading, error } = this.state;

    const typeOptions: IDropdownOption[] = [
      { key: 'ExternalSharing', text: 'External Sharing' },
      { key: 'InternalOnly', text: 'Internal Only' }
    ];

    const durationOptions: IDropdownOption[] = [
      { key: 30, text: '30 Days' },
      { key: 60, text: '60 Days' },
      { key: 90, text: '90 Days' },
      { key: 365, text: '1 Year' },
      { key: -1, text: 'No Expiration' }
    ];

    const isUploadMode = !selectedFile;

    return (
      <Dialog
        hidden={!isOpen}
        onDismiss={onClose}
        dialogContentProps={{
          type: DialogType.close,
          title: 'Share New Document',
          showCloseButton: true
        }}
        modalProps={{
          isBlocking: false,
          className: styles.dialog
        }}
      >
        <div className={styles.dialogContent}>
          <TextField
            label="External Sharing"
            placeholder="Files uploaded here will be visible to all 2 portal guests."
            value={externalSharing}
            onChange={this._onExternalSharingChange}
            required
            className={styles.field}
          />

          {isUploadMode ? (
            <>
              <Dropdown
                label="Type"
                placeholder="Select type"
                options={typeOptions}
                selectedKey={type}
                onChange={this._onTypeChange}
                className={styles.field}
              />

              <TextField
                label="Name"
                placeholder="Enter Name"
                value={name}
                onChange={this._onNameChange}
                className={styles.field}
              />

              <div className={styles.uploadSection}>
                <input
                  ref={this.fileInputRef}
                  type="file"
                  style={{ display: 'none' }}
                  onChange={this._onFileChange}
                />
                <DefaultButton
                  text={selectedFile ? selectedFile.name : "Choose File"}
                  iconProps={{ iconName: 'Attach' }}
                  onClick={this._onFileSelect}
                  className={styles.uploadButton}
                />
              </div>
            </>
          ) : (
            <>
              <div className={styles.userSection}>
                <p className={styles.label}>Select Users</p>
                <div className={styles.userList}>
                  <div className={styles.userItem}>
                    <div className={styles.avatar}>J</div>
                    <span className={styles.userName}>Justin</span>
                    <input type="checkbox" className={styles.checkbox} />
                  </div>
                  <div className={styles.userItem}>
                    <div className={styles.avatar}>L</div>
                    <span className={styles.userName}>Liam</span>
                    <input type="checkbox" className={styles.checkbox} />
                  </div>
                </div>
              </div>

              <Dropdown
                label="Access Duration"
                placeholder="Select duration"
                options={durationOptions}
                selectedKey={accessDuration}
                onChange={this._onAccessDurationChange}
                className={styles.field}
              />

              <p className={styles.note}>
                Access to these specific files will expire automatically.
              </p>
            </>
          )}

          {error && (
            <div className={styles.errorMessage}>
              <Icon iconName="Error" className={styles.errorIcon} />
              {error}
            </div>
          )}
        </div>

        <DialogFooter>
          <PrimaryButton
            text={isUploadMode ? "Upload & Share" : "Share Access"}
            onClick={isUploadMode ? this._onUploadAndShare : this._onShareAccess}
            disabled={uploading}
            className={styles.primaryButton}
          />
          <DefaultButton
            text="Cancel"
            onClick={onClose}
            disabled={uploading}
          />
        </DialogFooter>
      </Dialog>
    );
  }
}
