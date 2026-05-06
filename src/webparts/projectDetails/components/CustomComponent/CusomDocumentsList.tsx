import * as React from "react";
import { useState, useEffect } from "react";
import { toast } from "react-toastify";
import "react-toastify/dist/ReactToastify.css";
import styles from "./Documents.module.scss";
import { getProjectDocuments, uploadProjectDocument, deleteProjectDocument, downloadProjectDocument, initializePnP, getDocumentsByServerRelativeUrl, uploadDocumentByServerRelativeUrl, deleteDocumentByServerRelativeUrl, createFolderByServerRelativeUrl, uploadDocumentWithFolderCreation } from "../../../../shared/services/projectService";
import { initLibraryDiscoveryService, getProjectLibraryName, getProjectById } from "../../../../shared/services/libraryDiscoveryService";

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

type DocumentTabType = 'project' | 'contract' |'bid' ;

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

const CusomDocumentsList: React.FC<IDocumentsProps> = ({
  projectId,
  projectCode,
  projectTitle,
  hideNewButton = false,
  onNewClick
}) => {
  const { canAdd } = usePermissionStore();
  const [documents, setDocuments] = useState<DocumentItem[]>([]);
  const [allDocuments, setAllDocuments] = useState<DocumentItem[]>([]);
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
    pageSize: 12,
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
        debugger
        // Type guard to check if projectData has the URL properties
        if (projectData && typeof projectData === 'object' && 'ProjectDocumentsUrl' in projectData) {
          const projUrl = projectData.ProjectDocumentsUrl || '';
          const bidUrl = projectData.BidDocumentsUrl || '';
          const contractUrl = projectData.ContractsDocumentsUrl || '';
          const status = projectData.Status || '';

          console.log('📎 URLs found:', {
            ProjectDocumentsUrl: projUrl,
            BidDocumentsUrl: bidUrl,
            ContractsDocumentsUrl: contractUrl,
            Status: status
          });

          setProjectDocumentsUrl(projUrl);
          setBidDocumentsUrl(bidUrl);
          setContractsDocumentsUrl(contractUrl);
          setProjectStatus(status);

          // For Pipeline projects (Potential status), default to Bid Documents
          if (status === 'Potential' && bidUrl) {
            setActiveDocTab('bid');
          }
          // Otherwise use normal priority: project > contract > bid
          else if (projUrl) {
            setActiveDocTab('project');
          } else if (contractUrl) {
            setActiveDocTab('contract');
          } else if (bidUrl) {
            setActiveDocTab('bid');
          }

          // Only discover library if ALL URLs are empty
          if (!projUrl && !bidUrl && !contractUrl) {
            console.log('⚠️ No URLs found, discovering library by code + title');
            const name = await getProjectLibraryName(projectCode, projectTitle);
            setLibraryName(name);
          } else {
            console.log('✅ URLs exist, skipping library discovery');
          }
        } else {
          // Fallback to library discovery if data doesn't have URL fields
          console.log('⚠️ No URL fields in project data, discovering library');
          const name = await getProjectLibraryName(projectCode, projectTitle);
          setLibraryName(name);
        }
      } catch (error) {
        console.error('Error fetching project URLs:', error);
        // Fallback to library discovery
        const name = await getProjectLibraryName(projectCode, projectTitle);
        setLibraryName(name);
      }
    };
    void fetchProjectUrls();
  }, [projectId, projectCode, projectTitle]);

  useEffect(() => {
    // Load documents when library name exists OR when URLs exist
    const hasUrls = projectDocumentsUrl || bidDocumentsUrl || contractsDocumentsUrl;
    if (libraryName || hasUrls) {
      void loadDocuments();
      void updateBreadcrumbs();
    }
  }, [projectId, currentFolderPath, libraryName, activeDocTab, projectDocumentsUrl, bidDocumentsUrl, contractsDocumentsUrl]);

  useEffect(() => {
    void updateDisplayedDocuments();
  }, [pagination.currentPage, pagination.pageSize, allDocuments]);

  const updateBreadcrumbs = (): void => {
    const crumbs: BreadcrumbItem[] = [];

    // Add tab-specific root breadcrumb
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

  const loadDocuments = async (): Promise<void> => {
    try {
      setLoading(true);

      // Determine which URL/library to use based on active tab
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

      console.log('📂 Loading documents for tab:', activeDocTab, {
        targetUrl,
        targetLibrary,
        currentFolderPath
      });

      let docs: DocumentItem[] = [];

      // If URL exists, directly use server relative URL
      if (targetUrl) {
        console.log('🔗 Using server relative URL:', targetUrl, 'SubFolder:', currentFolderPath);
        docs = await getDocumentsByServerRelativeUrl(targetUrl, currentFolderPath);
      } else if (targetLibrary) {
        // Fallback to Code + Title library logic
        console.log('📚 Using library name:', targetLibrary, 'Folder:', currentFolderPath);
        docs = await getProjectDocuments(targetLibrary, currentFolderPath);
      } else {
        console.warn('⚠️ No URL or library name available for loading documents');
      }

      console.log('✅ Documents loaded:', docs.length);
      setAllDocuments(docs);

      setPagination({
        currentPage: 1,
        pageSize: 12,
        totalItems: docs.length,
        totalPages: Math.ceil(docs.length / 12),
      });
    } catch (error) {
      console.error("Error loading documents:", error);
    } finally {
      setLoading(false);
    }
  };

  const updateDisplayedDocuments = (): void => {
    const startIndex = (pagination.currentPage - 1) * pagination.pageSize;
    const endIndex = startIndex + pagination.pageSize;
    const paginatedDocs = allDocuments.slice(startIndex, endIndex);
    setDocuments(paginatedDocs);
  };

  const documentCount = React.useMemo(() => {
    return allDocuments.filter(item => item.FSObjType !== 1).length;
  }, [allDocuments]);

  const handlePageChange = (newPage: number): void => {
    if (newPage >= 1 && newPage <= pagination.totalPages) {
      setPagination(prev => ({
        ...prev,
        currentPage: newPage
      }));
      setSelectedItems(new Set());
    }
  };

  const handlePageSizeChange = (newPageSize: number): void => {
    setPagination({
      currentPage: 1,
      pageSize: newPageSize,
      totalItems: allDocuments.length,
      totalPages: Math.ceil(allDocuments.length / newPageSize),
    });
    setSelectedItems(new Set());
  };

  const formatDate = (dateString: string): string => {
    const date = new Date(dateString);
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const year = date.getFullYear();
    const hours = date.getHours().toString().padStart(2, "0");
    const minutes = date.getMinutes().toString().padStart(2, "0");
    return `${month}/${day}/${year} ${hours}:${minutes}...`;
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
      // Determine which URL/library to use based on active tab
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

      // If URL exists, use server relative URL upload
      if (targetUrl) {
        if (relativePath) {
          // This is a folder upload, use folder creation function
          await uploadDocumentWithFolderCreation(targetUrl, file, relativePath);
        } else {
          // This is a regular file upload, use original function
          await uploadDocumentByServerRelativeUrl(targetUrl, file, currentFolderPath);
        }
      } else if (targetLibrary) {
        // Fallback to library name upload (original function)
        const uploadPath = relativePath ? relativePath : currentFolderPath;
        await uploadProjectDocument(targetLibrary, file, uploadPath);
      }

      await loadDocuments();
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
      let targetUrl: string;
      let targetLibrary: string;

      // Determine target based on active tab
      switch (activeDocTab) {
        case 'project':
          targetUrl = projectDocumentsUrl;
          targetLibrary = 'Project Documents';
          break;
        case 'bid':
          targetUrl = bidDocumentsUrl;
          targetLibrary = 'Project Documents';
          break;
        case 'contract':
          targetUrl = contractsDocumentsUrl;
          targetLibrary = 'Project Documents';
          break;
      }

      // If URL exists, use server relative URL folder creation
      if (targetUrl) {
        const folderPath = currentFolderPath ? `${currentFolderPath}/${folderName}` : folderName;
        await createFolderByServerRelativeUrl(targetUrl, folderPath);
      } else {
        throw new Error('No document URL configured for this tab');
      }

      await loadDocuments();
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

  // Delete selected documents
  const handleDelete = async (): Promise<void> => {
    if (selectedItems.size === 0) return;

    const confirmed = window.confirm(
      `Are you sure you want to delete ${selectedItems.size} item(s)?`
    );

    if (!confirmed) return;

    try {
      setLoading(true);
      const selectedDocs = allDocuments.filter(doc => selectedItems.has(doc.Id));

      // Determine if using URL or library name
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

      // Delete each selected document
      if (targetUrl) {
        // Use server relative URL delete
        await Promise.all(
          selectedDocs.map(doc =>
            deleteDocumentByServerRelativeUrl(doc.FileRef || doc.ServerRelativeUrl || "")
          )
        );
      } else if (targetLibrary) {
        // Use library name delete
        await Promise.all(
          selectedDocs.map(doc =>
            deleteProjectDocument(targetLibrary, doc.FileRef || doc.ServerRelativeUrl || "")
          )
        );
      }

      // Reload documents after deletion
      await loadDocuments();
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

  // Download selected documents
  const handleDownload = async (): Promise<void> => {
    if (selectedItems.size === 0) return;

    try {
      const selectedDocs = allDocuments.filter(doc => selectedItems.has(doc.Id));
      const fileDocs = selectedDocs.filter(doc => doc.FSObjType !== 1); // Only download files, not folders

      if (fileDocs.length === 0) {
        alert("Cannot download folders. Please select files only.");
        return;
      }

      // Download each file
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
      // For URL-based navigation, use the folder name for relative navigation
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

  const startItem = (pagination.currentPage - 1) * pagination.pageSize + 1;
  const endItem = Math.min(
    pagination.currentPage * pagination.pageSize,
    pagination.totalItems
  );

  if (loading) return <GlobalLoader variant="content" />;

  // Check if any URLs exist to show tabs
  const hasAnyUrl = projectDocumentsUrl || bidDocumentsUrl || contractsDocumentsUrl;
  const isPipelineProject = projectStatus === 'Potential';

  return (
    <>
      <div className={styles.documentsWrapper}>
        {/* Document Type Tabs - Only show if URLs exist */}
        {hasAnyUrl && (
          <div className={styles.documentTabs}>
            {/* Hide Project Documents tab for Pipeline projects */}
            {!isPipelineProject && (projectDocumentsUrl || (!bidDocumentsUrl && !contractsDocumentsUrl)) && (
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
            
            {/* Show Bid Documents tab only for Pipeline projects - First tab */}
            {isPipelineProject && bidDocumentsUrl && (
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

            {/* Show Contract Documents tab for all projects */}
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

          {/* Action buttons - show when items are selected */}
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
            <div className={styles.documentCount}>
              {documentCount}
            </div>



            {!hideNewButton && canAdd && (
              <button className={styles.newButton} onClick={handleNewDocument}>
                <span className={styles.plusIcon}>+</span> New
              </button>
            )}
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

        {allDocuments.length > 0 && (
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

export default CusomDocumentsList;