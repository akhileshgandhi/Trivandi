import * as React from 'react';
import { toast } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';
import styles from './ImportFromProjectDocsDialog.module.scss';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { ProjectExternalPortalService } from '../../services/ProjectExternalPortalService';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Breadcrumb } from '@fluentui/react/lib/Breadcrumb';

export interface IImportFromProjectDocsDialogProps {
  isOpen: boolean;
  onClose: () => void;
  onSuccess: () => void;
  context: WebPartContext;
  targetPath: string;
  service?: ProjectExternalPortalService;
  projectCode?: string;
  projectName?: string;
  projectId?: string;
}

export interface IProjectDocument {
  id: number;
  name: string;
  location: string;
  selected: boolean;
  isFolder?: boolean;
  serverRelativeUrl?: string;
  documentType?: string;
}

export interface IImportFromProjectDocsDialogState {
  documents: IProjectDocument[];
  loading: boolean;
  importing: boolean;
  error?: string;
  success?: string;
  currentFolderPath: string;
  breadcrumbs: Array<{ text: string; key: string; path: string }>;
  discoveredLibraryName?: string;
}

export default class ImportFromProjectDocsDialog extends React.Component<IImportFromProjectDocsDialogProps, IImportFromProjectDocsDialogState> {
  private portalService: ProjectExternalPortalService;

  constructor(props: IImportFromProjectDocsDialogProps) {
    super(props);
    
    this.state = {
      documents: [],
      loading: false,
      importing: false,
      currentFolderPath: '',
      breadcrumbs: [{ text: 'Root', key: 'root', path: '' }]
    };

    this.portalService = props.service || new ProjectExternalPortalService(props.context);
  }

  public async componentDidUpdate(prevProps: IImportFromProjectDocsDialogProps): Promise<void> {
    // Only load when dialog is opened (not on every update)
    if (this.props.isOpen && !prevProps.isOpen && !this.state.loading && this.state.documents.length === 0) {
      // Discover library name
      if (this.props.projectCode || this.props.projectName) {
        const libraryName = await this.portalService.findProjectLibrary(this.props.projectCode, this.props.projectName);
        if (libraryName) {
          this.setState({ discoveredLibraryName: libraryName });
        }
      }
      this._loadProjectDocuments();
    }
  }

  private _loadProjectDocuments = async (): Promise<void> => {
    // Prevent duplicate calls
    if (this.state.loading) {
      return;
    }
    
    this.setState({ loading: true });
    try {
      // Use new method to get documents from all 3 types if projectId is available
      let documents: any[];
      
      if (this.props.projectId) {
        // Get documents from all 3 document types (Project, Bid, Contract)
        documents = await this.portalService.getAllProjectDocumentsForImport(this.props.projectId, this.state.currentFolderPath);
      } else {
        // Fallback to original method
        documents = await this.portalService.getProjectDocumentsForImport(this.state.currentFolderPath);
      }
      
      this.setState({ 
        documents: documents.map(doc => ({ ...doc, selected: false })),
        loading: false 
      });
    } catch (error) {
      const errorMessage = error.message || 'Failed to load documents';
      console.error('Error loading project documents:', error);
      this.setState({ 
        error: errorMessage.includes('429') || errorMessage.includes('throttle')
          ? 'Too many requests. Please wait a moment and try again.'
          : errorMessage,
        loading: false 
      });
    }
  }

  private _onToggleDocument = (index: number): void => {
    const documents = [...this.state.documents];
    documents[index].selected = !documents[index].selected;
    this.setState({ documents });
  }

  private _onFolderClick = (folder: IProjectDocument): void => {
    const newPath = this.state.currentFolderPath ? `${this.state.currentFolderPath}/${folder.name}` : folder.name;
    const newBreadcrumbs = [
      ...this.state.breadcrumbs,
      { text: folder.name, key: folder.name, path: newPath }
    ];
    
    this.setState({ 
      currentFolderPath: newPath,
      breadcrumbs: newBreadcrumbs,
      documents: []
    }, () => {
      this._loadProjectDocuments();
    });
  }

  private _onBreadcrumbClick = (item: any): void => {
    const index = this.state.breadcrumbs.findIndex(b => b.key === item.key);
    const newBreadcrumbs = this.state.breadcrumbs.slice(0, index + 1);
    const newPath = newBreadcrumbs[newBreadcrumbs.length - 1].path;
    
    this.setState({
      currentFolderPath: newPath,
      breadcrumbs: newBreadcrumbs,
      documents: []
    }, () => {
      this._loadProjectDocuments();
    });
  }

  private _onImportSelected = async (): Promise<void> => {
    const selectedDocs = this.state.documents.filter(doc => doc.selected);
    
    if (selectedDocs.length === 0) {
      this.setState({ error: 'Please select at least one document to import' });
      return;
    }

    this.setState({ importing: true, error: undefined, success: undefined });

    try {
      // Import documents and cache will be cleared inside this method
      await this.portalService.copyDocumentsToExternalLibrary(
        selectedDocs,
        this.props.targetPath
      );
      
      console.log('✓ Import completed, cache cleared');
      
      this.setState({ 
        success: `Successfully imported ${selectedDocs.length} item(s) to External Shared Area!`,
        importing: false 
      });

      toast.success(`Successfully imported ${selectedDocs.length} item(s) to External Shared Area!`, {
        position: "top-right",
        autoClose: 3000,
      });

      // Call onSuccess to trigger refresh after cache is cleared
      console.log('→ Calling onSuccess to refresh parent list');
      this.props.onSuccess();

      // Close dialog after showing success message
      setTimeout(() => {
        this.props.onClose();
      }, 1500);
    } catch (error) {
      this.setState({ error: error.message, importing: false });
      toast.error(`Failed to import: ${error.message}`, {
        position: "top-right",
        autoClose: 5000,
      });
    }
  }

  public render(): React.ReactElement<IImportFromProjectDocsDialogProps> {
    const { isOpen, onClose } = this.props;
    const { documents, loading, importing, error, success, breadcrumbs } = this.state;

    const selectedCount = documents.filter(doc => doc.selected).length;

    const breadcrumbItems = breadcrumbs.map(b => ({
      text: b.text,
      key: b.key,
      onClick: () => this._onBreadcrumbClick(b)
    }));

    // Use discovered library name if available, otherwise fallback to constructed name
    const dialogTitle = this.state.discoveredLibraryName || 
      (this.props.projectCode && this.props.projectName 
        ? `${this.props.projectCode}-${this.props.projectName}` 
        : 'Import from Project Docs');

    return (
      <Dialog
        hidden={!isOpen}
        onDismiss={onClose}
        dialogContentProps={{
          type: DialogType.close,
          title: dialogTitle,
          showCloseButton: true
        }}
        modalProps={{
          isBlocking: importing,
          className: styles.dialog
        }}
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

          <div className={styles.infoBox}>
            <Icon iconName="Info" className={styles.infoIcon} />
            <p className={styles.infoText}>
              Select files and folders to copy to the External Shared Area. The originals will remain intact; copies will be visible to invited guests.
            </p>
          </div>

          {loading ? (
            <div className={styles.loadingContainer}>
              <Spinner size={SpinnerSize.large} label="Loading project documents..." />
            </div>
          ) : (
            <>
              {/* Breadcrumb Navigation */}
              {breadcrumbs.length > 1 && (
                <div style={{ marginBottom: '12px', paddingBottom: '8px', borderBottom: '1px solid #edebe9' }}>
                  <Breadcrumb
                    items={breadcrumbItems}
                    maxDisplayedItems={5}
                  />
                </div>
              )}

              <div className={styles.tableContainer}>
                <table className={styles.table}>
                  <thead>
                    <tr>
                      <th className={styles.checkboxColumn}></th>
                      <th className={styles.nameColumn}>Name</th>
                      <th className={styles.locationColumn}>Location</th>
                      <th className={styles.typeColumn}>Type</th>
                      <th className={styles.sourceColumn}>Source</th>
                    </tr>
                  </thead>
                  <tbody>
                    {documents.length === 0 ? (
                      <tr>
                        <td colSpan={5} className={styles.emptyState}>
                          <Icon iconName="FabricFolder" className={styles.emptyIcon} />
                          <p>No project documents available for import.</p>
                        </td>
                      </tr>
                    ) : (
                      documents.map((doc, index) => (
                        <tr key={index} className={styles.row}>
                          <td className={styles.checkboxColumn}>
                            <input
                              type="checkbox"
                              checked={doc.selected}
                              onChange={() => this._onToggleDocument(index)}
                              className={styles.checkbox}
                              disabled={importing}
                            />
                          </td>
                          <td 
                            className={styles.nameColumn}
                            onClick={() => doc.isFolder ? this._onFolderClick(doc) : null}
                            style={{ cursor: doc.isFolder ? 'pointer' : 'default' }}
                          >
                            <div className={styles.nameCell}>
                              <Icon 
                                iconName={doc.isFolder ? "FabricFolder" : "Page"} 
                                className={styles.fileIcon} 
                              />
                              <span>{doc.name}</span>
                            </div>
                          </td>
                          <td className={styles.locationColumn}>{doc.location}</td>
                          <td className={styles.typeColumn}>{doc.isFolder ? 'Folder' : 'File'}</td>
                          <td className={styles.sourceColumn}>{doc.documentType || 'Project Library'}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>

              {selectedCount > 0 && (
                <MessageBar messageBarType={MessageBarType.info}>
                  {selectedCount} item{selectedCount > 1 ? 's' : ''} selected for import
                </MessageBar>
              )}
            </>
          )}
        </div>

        <DialogFooter>
          <PrimaryButton
            text={importing ? "Importing..." : "Import Selected"}
            onClick={this._onImportSelected}
            disabled={loading || importing || selectedCount === 0}
            className={styles.primaryButton}
          />
          <DefaultButton
            text="Cancel"
            onClick={onClose}
            disabled={importing}
          />
        </DialogFooter>
      </Dialog>
    );
  }
}
