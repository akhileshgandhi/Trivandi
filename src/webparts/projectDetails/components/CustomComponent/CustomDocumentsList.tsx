import * as React from "react";
import { useState, useEffect } from "react";
import { toast } from "react-toastify";
import "react-toastify/dist/ReactToastify.css";
import styles from "./Documents.module.scss";
import { getProjectDocuments, uploadProjectDocument, deleteProjectDocument, downloadProjectDocument, initializePnP, getDocumentsByServerRelativeUrl, uploadDocumentByServerRelativeUrl, deleteDocumentByServerRelativeUrl, createFolderByServerRelativeUrl, uploadDocumentWithFolderCreation } from "../../../../shared/services/projectService";
import { initLibraryDiscoveryService, getProjectLibraryName, getProjectById, findProjectLibraryByProject } from "../../../../shared/services/libraryDiscoveryService";

import GlobalLoader from "../../../../shared/component/GlobalLoader";
import FileUploadModal from "./FileUploadModal/FileUploadModal";
import { usePermissionStore } from "../../../../Permission/PermissionStore";
import Pagination from "../../../../shared/component/Pagination/Pagination";

interface DocumentItem {
  Id: number;
  FileLeafRef: string;
  Modified: string;
  Editor?: { Title: string };
  File_x0020_Type?: string;
  FileRef?: string;
  FSObjType?: number;
  VersionLabel?: string;
  Status?: string;
  ServerRelativeUrl?: string;
  FileDirRef?: string;
}

interface IDocumentsProps {
  projectId: number;
  projectCode?: string;
  projectTitle?: string;
  hideNewButton?: boolean;
  onNewClick?: () => void;
}

interface BreadcrumbItem {
  name: string;
  path: string;
}

interface PaginationInfo {
  currentPage: number;
  pageSize: number;
  totalItems: number;
  totalPages: number;
}

type DocumentTabType = 'project' | 'contract' | 'bid';

// Helper function to extract library and folder path from server relative URL
const extractLibraryFromUrl = (url: string): { library: string; basePath: string } | null => {
  if (!url) return null;

  // URL format: /sites/sitename/LibraryName/FolderPath
  const urlParts = url.split('/').filter(Boolean);

  // Find the library name (usually after 'sites' and site name)
  const sitesIndex = urlParts.indexOf('sites');
  if (sitesIndex !== -1 && urlParts.length > sitesIndex + 2) {
    const library = urlParts[sitesIndex + 2];
    const basePath = urlParts.slice(sitesIndex + 3).join('/');
    return { library, basePath };
  }

  // Alternative: If no 'sites', assume first part is library
  if (urlParts.length > 0) {
    const library = urlParts[0];
    const basePath = urlParts.slice(1).join('/');
    return { library, basePath };
  }

  return null;
};

const CustomDocumentsList: React.FC<IDocumentsProps> = ({
  projectId,
  projectCode,
  projectTitle,
  hideNewButton = false,
  onNewClick
}) => {
  const { canAdd } = usePermissionStore();
  const [documents, setDocuments] = useState<DocumentItem[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [selectedItems, setSelectedItems] = useState<Set<number>>(new Set());
  const [isUploadModalOpen, setIsUploadModalOpen] = useState<boolean>(false);
  const [currentFolderPath, setCurrentFolderPath] = useState<string>("");
  const [breadcrumbs, setBreadcrumbs] = useState<BreadcrumbItem[]>([]);
  const [activeDocTab, setActiveDocTab] = useState<DocumentTabType>('project');

  // Project URL states
  const [projectDocumentsUrl, setProjectDocumentsUrl] = useState<string>('');
  const [bidDocumentsUrl, setBidDocumentsUrl] = useState<string>('');
  const [contractsDocumentsUrl, setContractsDocumentsUrl] = useState<string>('');
  const [projectStatus, setProjectStatus] = useState<string>('');

  const [pagination, setPagination] = useState<PaginationInfo>({
    currentPage: 1,
    pageSize: 10,
    totalItems: 0,
    totalPages: 0,
  });

  const [libraryName, setLibraryName] = React.useState<string>('');

  // Fetch project details and URLs
  React.useEffect(() => {
    const fetchProjectUrls = async (): Promise<void> => {
      try {
        const projectData = await getProjectById(projectId);
        console.log('📊 Project Data fetched:', projectData);

        if (projectData && typeof projectData === 'object' && 'ProjectDocumentsUrl' in projectData) {
          const cleanUrl = (url: any) => {
            if (!url) return '';
            const trimmed = String(url).trim();
            // Valid path should start with '/sites' or 'http'
            const isValidPath = trimmed.toLowerCase().startsWith('/sites') || trimmed.startsWith('http');
            return (trimmed === '-' || trimmed === '' || !isValidPath) ? '' : trimmed;
          };

          const projUrl = cleanUrl(projectData.ProjectDocumentsUrl);
          const bidUrl = cleanUrl(projectData.BidDocumentsUrl);
          const contractUrl = cleanUrl(projectData.ContractsDocumentsUrl);
          const status = projectData.Status || '';

          let fetchedLibraryName = '';
          if (!projUrl) {
            // Strict check: only show if library actually exists in SharePoint
            const found = await findProjectLibraryByProject(projectCode, projectTitle);
            fetchedLibraryName = found || '';
          }

          setProjectDocumentsUrl(projUrl);
          setBidDocumentsUrl(bidUrl);
          setContractsDocumentsUrl(contractUrl);
          setProjectStatus(status);
          setLibraryName(fetchedLibraryName);

          if (projUrl || fetchedLibraryName) {
            setActiveDocTab('project');
          } else if (bidUrl) {
            setActiveDocTab('bid');
          } else if (contractUrl) {
            setActiveDocTab('contract');
          } else {
            // Default to project, but if hasAnyUrl is false, nothing will show in render
            setActiveDocTab('project');
          }
        } else {
          const found = await findProjectLibraryByProject(projectCode, projectTitle);
          setLibraryName(found || '');
          setActiveDocTab('project');
        }
      } catch (error) {
        console.error('Error fetching project URLs:', error);
        const found = await findProjectLibraryByProject(projectCode, projectTitle);
        setLibraryName(found || '');
      }
    };
    void fetchProjectUrls();
  }, [projectId, projectCode, projectTitle]);

  useEffect(() => {
    const hasUrls = projectDocumentsUrl || bidDocumentsUrl || contractsDocumentsUrl;
    if (libraryName || hasUrls) {
      void loadDocuments(1);
      void updateBreadcrumbs();
    }
  }, [projectId, currentFolderPath, libraryName, activeDocTab, projectDocumentsUrl, bidDocumentsUrl, contractsDocumentsUrl]);

  const updateBreadcrumbs = (): void => {
    const crumbs: BreadcrumbItem[] = [];

    const tabNames = {
      project: projectDocumentsUrl ? 'Project Documents' : libraryName,
      bid: 'Bid Documents',
      contract: 'Contract Documents'
    };

    crumbs.push({
      name: tabNames[activeDocTab],
      path: ""
    });

    if (currentFolderPath) {
      const parts = currentFolderPath.split("/");
      let accumulatedPath = "";

      parts.forEach((part) => {
        if (part) {
          accumulatedPath += (accumulatedPath ? "/" : "") + part;
          crumbs.push({
            name: part,
            path: accumulatedPath
          });
        }
      });
    }

    setBreadcrumbs(crumbs);
  };

  const loadDocuments = async (page: number = pagination.currentPage, pageSize: number = pagination.pageSize): Promise<void> => {
    try {
      setLoading(true);

      let targetUrl = '';
      let targetLibrary = '';

      switch (activeDocTab) {
        case 'project':
          targetUrl = projectDocumentsUrl;
          targetLibrary = libraryName;
          break;
        case 'bid':
          targetUrl = bidDocumentsUrl;
          break;
        case 'contract':
          targetUrl = contractsDocumentsUrl;
          break;
      }

      let result: { items: DocumentItem[]; totalCount: number } = { items: [], totalCount: 0 };

      if (targetUrl) {
        result = await getDocumentsByServerRelativeUrl(targetUrl, currentFolderPath, page, pageSize);
      } else if (targetLibrary) {
        result = await getProjectDocuments(targetLibrary, currentFolderPath, page, pageSize);
      }

      setDocuments(result.items);
      setPagination({
        currentPage: page,
        pageSize: pageSize,
        totalItems: result.totalCount,
        totalPages: Math.ceil(result.totalCount / pageSize),
      });
      setSelectedItems(new Set());
    } catch (error) {
      console.error("Error loading documents:", error);
    } finally {
      setLoading(false);
    }
  };

  // Keep a separate count for files (optional, but since we are paginating, we might want to show the full count if possible)
  // For now, I'll use the pagination.totalItems for the count display
  const documentCount = pagination.totalItems;

  const handlePageChange = (newPage: number): void => {
    if (newPage >= 1 && newPage <= pagination.totalPages) {
      void loadDocuments(newPage);
    }
  };

  const handlePageSizeChange = (newPageSize: number): void => {
    void loadDocuments(1, newPageSize);
  };

  const formatDate = (dateString: string): string => {
    const date = new Date(dateString);
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const year = date.getFullYear();
    const hours = date.getHours().toString().padStart(2, "0");
    const minutes = date.getMinutes().toString().padStart(2, "0");
    return `${month}/${day}/${year} ${hours}:${minutes}`;
  };

  const getFileIcon = (item: DocumentItem): string => {
    if (item.FSObjType === 1) return "📁";
    const ext = item.File_x0020_Type?.toLowerCase();
    switch (ext) {
      case "pdf":
        return "📄";
      case "docx":
      case "doc":
        return "📘";
      case "xlsx":
      case "xls":
        return "📊";
      case "pptx":
      case "ppt":
        return "📙";
      default:
        return "📄";
    }
  };

  const toggleSelectAll = (): void => {
    if (selectedItems.size === documents.length) {
      setSelectedItems(new Set());
    } else {
      setSelectedItems(new Set(documents.map((d) => d.Id)));
    }
  };

  const toggleSelectItem = (id: number): void => {
    const newSelected = new Set(selectedItems);
    if (newSelected.has(id)) {
      newSelected.delete(id);
    } else {
      newSelected.add(id);
    }
    setSelectedItems(newSelected);
  };


  const handleNewDocument = (): void => {
    if (onNewClick) {
      onNewClick();
      return;
    }
    setIsUploadModalOpen(true);
  };

  const handleUpload = async (file: File, relativePath?: string): Promise<void> => {
    try {
      let targetUrl = '';
      let targetLibrary = '';

      switch (activeDocTab) {
        case 'project':
          targetUrl = projectDocumentsUrl;
          targetLibrary = libraryName;
          break;
        case 'bid':
          targetUrl = bidDocumentsUrl;
          break;
        case 'contract':
          targetUrl = contractsDocumentsUrl;
          break;
      }

      if (targetUrl) {
        if (relativePath) {
          await uploadDocumentWithFolderCreation(targetUrl, file, relativePath);
        } else {
          await uploadDocumentByServerRelativeUrl(targetUrl, file, currentFolderPath);
        }
      } else if (targetLibrary) {
        const uploadPath = relativePath ? relativePath : currentFolderPath;
        await uploadProjectDocument(targetLibrary, file, uploadPath);
      }

      await loadDocuments(1);
      toast.success("Document uploaded successfully!", {
        position: "top-right",
        autoClose: 3000,
      });
    } catch (error) {
      console.error("Upload error:", error);
      toast.error("Failed to upload document. Please try again.", {
        position: "top-right",
        autoClose: 5000,
      });
      throw error;
    }
  };

  const handleCreateFolder = async (folderName: string): Promise<void> => {
    try {
      let targetUrl: string = '';

      switch (activeDocTab) {
        case 'project':
          targetUrl = projectDocumentsUrl;
          break;
        case 'bid':
          targetUrl = bidDocumentsUrl;
          break;
        case 'contract':
          targetUrl = contractsDocumentsUrl;
          break;
      }

      if (targetUrl) {
        const folderPath = currentFolderPath ? `${currentFolderPath}` : folderName;
        await createFolderByServerRelativeUrl(targetUrl, folderPath);
      } else {
        throw new Error('No document URL configured for this tab');
      }

      await loadDocuments(1);
      toast.success("Folder created successfully!", {
        position: "top-right",
        autoClose: 3000,
      });
    } catch (error) {
      console.error("Create folder error:", error);
      toast.error("Failed to create folder. Please try again.", {
        position: "top-right",
        autoClose: 5000,
      });
      throw error;
    }
  };

  const handleDelete = async (): Promise<void> => {
    if (selectedItems.size === 0) return;

    const confirmed = window.confirm(
      `Are you sure you want to delete ${selectedItems.size} item(s)?`
    );

    if (!confirmed) return;

    try {
      setLoading(true);
      const selectedDocs = documents.filter(doc => selectedItems.has(doc.Id));

      let targetUrl = '';
      let targetLibrary = '';

      switch (activeDocTab) {
        case 'project':
          targetUrl = projectDocumentsUrl;
          targetLibrary = libraryName;
          break;
        case 'bid':
          targetUrl = bidDocumentsUrl;
          break;
        case 'contract':
          targetUrl = contractsDocumentsUrl;
          break;
      }

      if (targetUrl) {
        await Promise.all(
          selectedDocs.map(doc =>
            deleteDocumentByServerRelativeUrl(doc.FileRef || doc.ServerRelativeUrl || "")
          )
        );
      } else if (targetLibrary) {
        await Promise.all(
          selectedDocs.map(doc =>
            deleteProjectDocument(targetLibrary, doc.FileRef || doc.ServerRelativeUrl || "")
          )
        );
      }

      await loadDocuments(1);
      setSelectedItems(new Set());
      toast.success(`Successfully deleted ${selectedDocs.length} item(s)`, {
        position: "top-right",
        autoClose: 3000,
      });
    } catch (error) {
      console.error("Error deleting documents:", error);
      toast.error("Failed to delete some items. Please try again.", {
        position: "top-right",
        autoClose: 5000,
      });
    } finally {
      setLoading(false);
    }
  };

  const handleDownload = async (): Promise<void> => {
    if (selectedItems.size === 0) return;

    try {
      const selectedDocs = documents.filter(doc => selectedItems.has(doc.Id));
      const fileDocs = selectedDocs.filter(doc => doc.FSObjType !== 1);

      if (fileDocs.length === 0) {
        alert("Cannot download folders. Please select files only.");
        return;
      }

      for (const doc of fileDocs) {
        if (doc.FileRef) {
          await downloadProjectDocument(doc.FileRef, doc.FileLeafRef);
        }
      }

      alert(`Successfully downloaded ${fileDocs.length} file(s)`);
    } catch (error) {
      console.error("Error downloading documents:", error);
      alert("Failed to download some files. Please try again.");
    }
  };

  const handleItemClick = (item: DocumentItem): void => {
    if (item.FSObjType === 1) {
      const newPath = currentFolderPath
        ? `${currentFolderPath}/${item.FileLeafRef}`
        : item.FileLeafRef;
      setCurrentFolderPath(newPath);
      setSelectedItems(new Set());
    } else {
      if (item.FileRef || item.ServerRelativeUrl) {
        const siteUrl = window.location.origin;
        const fileUrl = item.FileRef || item.ServerRelativeUrl || '';
        window.open(`${siteUrl}${fileUrl}`, "_blank");
      }
    }
  };

  const handleBreadcrumbClick = (path: string): void => {
    setCurrentFolderPath(path);
    setSelectedItems(new Set());
  };

  if (loading) return <GlobalLoader variant="content" />;

  const hasAnyUrl = projectDocumentsUrl || bidDocumentsUrl || contractsDocumentsUrl || libraryName;
  const isPipelineProject = projectStatus === 'Potential';

  return (
    <>
      <div className={styles.documentsWrapper}>
        <div className={styles.tabRow}>
          {hasAnyUrl && (
            <div className={styles.documentTabs}>
              {(projectDocumentsUrl || libraryName) && (
                <button
                  className={`${styles.docTab} ${activeDocTab === 'project' ? styles.activeDocTab : ''}`}
                  onClick={() => {
                    setActiveDocTab('project');
                    setCurrentFolderPath('');
                    setSelectedItems(new Set());
                  }}
                >
                  Project Documents
                </button>
              )}

              {bidDocumentsUrl && (
                <button
                  className={`${styles.docTab} ${activeDocTab === 'bid' ? styles.activeDocTab : ''}`}
                  onClick={() => {
                    setActiveDocTab('bid');
                    setCurrentFolderPath('');
                    setSelectedItems(new Set());
                  }}
                >
                  Bid Documents
                </button>
              )}

              {contractsDocumentsUrl && (
                <button
                  className={`${styles.docTab} ${activeDocTab === 'contract' ? styles.activeDocTab : ''}`}
                  onClick={() => {
                    setActiveDocTab('contract');
                    setCurrentFolderPath('');
                    setSelectedItems(new Set());
                  }}
                >
                  Contract Documents
                </button>
              )}
            </div>
          )}

          {!hideNewButton && canAdd && (
            <button className={styles.newButton} onClick={handleNewDocument}>
              <span className={styles.plusIcon}>+</span> New
            </button>
          )}
        </div>

        <div className={styles.h}>
          <div className={styles.breadcrumbsContainer}>
            {breadcrumbs.map((crumb, index) => (
              <React.Fragment key={index}>
                <button
                  className={`${styles.breadcrumb} ${index === breadcrumbs.length - 1 ? styles.active : ""
                    }`}
                  onClick={() => handleBreadcrumbClick(crumb.path)}
                  disabled={index === breadcrumbs.length - 1}
                >
                  {crumb.name}
                </button>
                {index < breadcrumbs.length - 1 && (
                  <span className={styles.breadcrumbSeparator}>›</span>
                )}
              </React.Fragment>
            ))}
          </div>

          {selectedItems.size > 0 && (
            <div className={styles.actionButtons}>
              <button
                className={styles.actionButton}
                onClick={handleDownload}
                title="Download selected items"
              >
                <svg width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M10 3v10m0 0l4-4m-4 4l-4-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                  <path d="M3 14v2a2 2 0 002 2h10a2 2 0 002-2v-2" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                </svg>
              </button>
              <button
                className={`${styles.actionButton} ${styles.deleteButton}`}
                onClick={handleDelete}
                title="Delete selected items"
              >
                <svg width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M3 5h14M8 5V3a1 1 0 011-1h2a1 1 0 011 1v2m3 0v11a2 2 0 01-2 2H7a2 2 0 01-2-2V5h10z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                  <path d="M8 9v6m4-6v6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                </svg>
              </button>
            </div>
          )}

          <div className={styles.toolbar}>
          </div>
        </div>

        <div className={styles.tableContainer}>
          <table className={styles.documentsTable}>
            <thead>
              <tr>
                <th className={styles.checkboxCol}>
                  <input
                    type="checkbox"
                    checked={selectedItems.size === documents.length && documents.length > 0}
                    onChange={toggleSelectAll}
                  />
                </th>
                <th className={styles.nameCol}>Name</th>
                <th className={styles.modifiedCol}>Modified</th>
                <th className={styles.byCol}>By</th>
              </tr>
            </thead>
            <tbody>
              {documents.length === 0 ? (
                <tr>
                  <td colSpan={4} className={styles.noData}>
                    No documents found
                  </td>
                </tr>
              ) : (
                documents.map((doc) => (
                  <tr key={doc.Id} className={styles.tableRow}>
                    <td className={styles.checkboxCol}>
                      <input
                        type="checkbox"
                        checked={selectedItems.has(doc.Id)}
                        onChange={() => toggleSelectItem(doc.Id)}
                      />
                    </td>
                    <td
                      className={styles.nameCol}
                      onClick={() => handleItemClick(doc)}
                      style={{ cursor: "pointer" }}
                    >
                      <span className={styles.fileIcon}>{getFileIcon(doc)}</span>
                      <span className={styles.fileName}>{doc.FileLeafRef}</span>
                    </td>
                    <td className={styles.modifiedCol}>{formatDate(doc.Modified)}</td>
                    <td className={styles.byCol}>{doc.Editor?.Title || ""}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

        {pagination.totalItems > 0 && (
          <Pagination
            currentPage={pagination.currentPage}
            totalPages={pagination.totalPages}
            pageSize={pagination.pageSize}
            totalItems={pagination.totalItems}
            onPageChange={handlePageChange}
            onPageSizeChange={handlePageSizeChange}
            label="documents"
          />
        )}
      </div>

      <FileUploadModal
        isOpen={isUploadModalOpen}
        onClose={() => setIsUploadModalOpen(false)}
        onUpload={handleUpload}
        onCreateFolder={handleCreateFolder}
      />
    </>
  );
};

export default CustomDocumentsList;