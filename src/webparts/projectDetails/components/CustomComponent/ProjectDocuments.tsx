import * as React from "react";
import { useState, useEffect, useMemo } from "react";
import { toast } from "react-toastify";
import "react-toastify/dist/ReactToastify.css";
import styles from "./Documents.module.scss";
import { getProjectDocuments, uploadProjectDocument, deleteProjectDocument, downloadProjectDocument, initializePnP, getDocumentsByServerRelativeUrl, uploadDocumentByServerRelativeUrl, deleteDocumentByServerRelativeUrl, createFolderByServerRelativeUrl, uploadDocumentWithFolderCreation, getLibraryRootFolder, moveFileOrFolder, copyFileOrFolder } from "../../../../shared/services/projectService";
import { initLibraryDiscoveryService, getProjectLibraryName, getProjectById, findProjectLibraryByProject } from "../../../../shared/services/libraryDiscoveryService";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import GlobalLoader from "../../../../shared/component/GlobalLoader";
import FileUploadModal from "./FileUploadModal/FileUploadModal";
import FilePreview from "../../../../shared/component/FilePreview/FilePreview";
import MoveCopyModal from "../../../../shared/component/MoveCopyModal/MoveCopyModal";
import { usePermissionStore } from "../../../../Permission/PermissionStore";
import Pagination from "../../../../shared/component/Pagination/Pagination";
import { Eye, History, Share2, Key, Download, ExternalLink, Link, Folder, Copy } from 'lucide-react';
import { ActionMenu } from '../../../../shared/component/ActionMenu/ActionMenu';
import CommonShareComponent from "../../../../shared/Common/CommonShareComponent";
import { handleCopyLink, handleDownload as commonHandleDownload, isItemFolder } from "../../../../shared/utils/fileUtils";
import { getFileIcon as getSharedFileIcon } from "../../../../shared/utils/iconHelper";
import {
  DocumentItem,
  IProjectDocumentsProps,
  BreadcrumbItem,
  PaginationInfo,
  DocumentTabType,
  SortField,
  SortDirection
} from "../../../../shared/interfaces/IDocumentInterfaces";



const ProjectDocuments: React.FC<IProjectDocumentsProps> = ({
  projectId,
  projectCode,
  projectTitle,
  businessProjectId,
  hideNewButton = false,
  onNewClick,
  context,
  onShowShareAccess,
  isUserRestricted = false,
  isNonCmap = false,
  location = '',
  uncategorisedDocumentsUrl = ''
}) => {
  const { canAdd, canEdit } = usePermissionStore();
  const [documents, setDocuments] = useState<DocumentItem[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [selectedItems, setSelectedItems] = useState<Set<number>>(new Set());
  const [isUploadModalOpen, setIsUploadModalOpen] = useState<boolean>(false);
  const [currentFolderPath, setCurrentFolderPath] = useState<string>("");
  const [breadcrumbs, setBreadcrumbs] = useState<BreadcrumbItem[]>([]);
  const [sortField, setSortField] = useState<SortField>('name');
  const [sortDirection, setSortDirection] = useState<SortDirection>('asc');

  // Project URL states
  const [projectDocumentsUrl, setProjectDocumentsUrl] = useState<string>('');
  const [projectStatus, setProjectStatus] = useState<string>('');

  const [hasAutoNavigated, setHasAutoNavigated] = useState<boolean>(false);

  const [pagination, setPagination] = useState<PaginationInfo>({
    currentPage: 1,
    pageSize: 10,
    totalItems: 0,
    totalPages: 0,
  });

  const [filePreview, setFilePreview] = useState<{
    isOpen: boolean;
    fileUrl: string;
    fileName: string;
    filePath?: string;
  }>({
    isOpen: false,
    fileUrl: '',
    fileName: '',
    filePath: '',
  });

  const openFilePreview = (doc: DocumentItem): void => {
    setFilePreview({
      isOpen: true,
      fileUrl: doc.FileRef || doc.ServerRelativeUrl || '',
      fileName: doc.FileLeafRef,
      filePath: doc.FileRef || doc.ServerRelativeUrl || '',
    });
  };

  const closeFilePreview = (): void => {
    setFilePreview(prev => ({ ...prev, isOpen: false }));
  };

  // Move / Copy states and handlers
  const [moveCopyState, setMoveCopyState] = useState<{
    isOpen: boolean;
    type: 'Move' | 'Copy';
    item: DocumentItem | null;
  }>({
    isOpen: false,
    type: 'Move',
    item: null,
  });

  const [moveCopyPaths, setMoveCopyPaths] = useState<{
    initialPath: string;
    libraryName: string;
  }>({
    initialPath: '',
    libraryName: '',
  });

  const getCurrentLibraryRootPath = async (): Promise<string> => {
    let baseLibUrl = projectDocumentsUrl;

    if (baseLibUrl) {
      if (baseLibUrl.startsWith('http')) {
        try {
          return new URL(baseLibUrl).pathname;
        } catch (e) {
          return baseLibUrl;
        }
      }
      return baseLibUrl;
    }

    const libTitle = libraryName;
    if (libTitle) {
      try {
        const rootFolder = await getLibraryRootFolder(libTitle);
        return rootFolder.ServerRelativeUrl;
      } catch (e) {
        console.error("Failed to get library root folder:", e);
      }
    }
    return '';
  };

  const openMoveCopy = async (doc: DocumentItem, type: 'Move' | 'Copy'): Promise<void> => {
    setLoading(true);
    try {
      const libRoot = await getCurrentLibraryRootPath();
      const resolvedCurrentPath = resolveLibraryFolderPath(currentFolderPath);
      // Start from current folder path, not root
      const initialFolderPath = resolvedCurrentPath
        ? `${libRoot}/${resolvedCurrentPath}`.replace(/\/+/g, '/')
        : libRoot;

      setMoveCopyPaths({
        initialPath: initialFolderPath,
        libraryName: libRoot,
      });
      setMoveCopyState({
        isOpen: true,
        type: type,
        item: doc,
      });
    } catch (error) {
      console.error("Failed to initialize move/copy path:", error);
      toast.error("Failed to load destination folders. Please try again.");
    } finally {
      setLoading(false);
    }
  };

  const handleMoveCopyConfirm = async (destinationPath: string): Promise<void> => {
    const { type, item } = moveCopyState;
    if (!item || !item.FileRef) return;

    setMoveCopyState(prev => ({ ...prev, isOpen: false }));
    setLoading(true);

    try {
      const isFolder = item.FSObjType === 1;
      const sourcePath = item.FileRef;

      const decodedSource = decodeURIComponent(sourcePath);
      const decodedDest = decodeURIComponent(destinationPath);

      const sourceParentFolder = decodedSource.substring(0, decodedSource.lastIndexOf('/'));

      if (decodedSource === decodedDest || sourceParentFolder === decodedDest) {
        toast.warning(`Item is already in the selected destination folder!`);
        return;
      }

      if (type === 'Move') {
        await moveFileOrFolder(sourcePath, destinationPath, isFolder);
        toast.success(`${isFolder ? 'Folder' : 'File'} moved successfully!`);
      } else {
        await copyFileOrFolder(sourcePath, destinationPath, isFolder);
        toast.success(`${isFolder ? 'Folder' : 'File'} copied successfully!`);
      }

      // Refresh document list
      await loadDocuments(pagination.currentPage);
    } catch (error) {
      console.error(`Error during ${type} operation:`, error);
      toast.error(`Failed to ${type.toLowerCase()} item. Please try again.`);
    } finally {
      setLoading(false);
    }
  };

  const [libraryName, setLibraryName] = React.useState<string>('');

  const getFallbackBaseFolderForTab = (): string => {
    return '';
  };

  const resolveLibraryFolderPath = (relativePath: string): string => {
    const baseFolder = getFallbackBaseFolderForTab();
    const normalizedPath = (relativePath || '').replace(/^\/+|\/+$/g, '');

    if (!baseFolder) return normalizedPath;
    if (!normalizedPath) return baseFolder;

    const lowerPath = normalizedPath.toLowerCase();
    const lowerBase = baseFolder.toLowerCase();
    if (lowerPath === lowerBase || lowerPath.startsWith(`${lowerBase}/`)) {
      return normalizedPath;
    }

    return `${baseFolder}/${normalizedPath}`;
  };

  // Fetch project details and URLs
  React.useEffect(() => {
    const fetchProjectUrls = async (): Promise<void> => {
      try {
        const projectData = await getProjectById(projectId);
        console.log("CUSTOM_DOCUMENTS_LIST_DATA_RECEIVED:", projectData);

        if (projectData && typeof projectData === 'object' && 'ProjectDocumentsUrl' in projectData) {
          const cleanUrl = (url: any) => {
            if (!url) return '';
            const trimmed = String(url).trim();
            // Valid path should start with '/sites' or 'http'
            const isValidPath = trimmed.toLowerCase().startsWith('/sites') || trimmed.startsWith('http');
            return (trimmed === '-' || trimmed === '' || !isValidPath) ? '' : trimmed;
          };

          const projUrl = cleanUrl(projectData.ProjectDocumentsUrl);
          const status = projectData.Status || '';
          const bizProjId = projectData.ProjectID || (projectData as any).ProjectId || "";

          let fetchedLibraryName = '';
          if (!projUrl) {
            // Priority: Search by businessProjectId (e.g. 2075162)
            const found = await findProjectLibraryByProject(bizProjId || projectCode, projectTitle);
            fetchedLibraryName = found || '';
          }

          setProjectDocumentsUrl(projUrl);
          setProjectStatus(status);
          setLibraryName(fetchedLibraryName);

        } else {
          const found = await findProjectLibraryByProject(projectCode, projectTitle);
          setLibraryName(found || '');
        }
      } catch (error) {

        const found = await findProjectLibraryByProject(projectCode, projectTitle);
        setLibraryName(found || '');
      }
    };
    void fetchProjectUrls();
  }, [projectId, projectCode, projectTitle]);



  useEffect(() => {
    const hasUrls = projectDocumentsUrl;
    if (libraryName || hasUrls || businessProjectId) {
      void loadDocuments(1);
      void updateBreadcrumbs();
    }
  }, [projectId, currentFolderPath, libraryName, projectDocumentsUrl, businessProjectId]);

  const updateBreadcrumbs = (): void => {
    const crumbs: BreadcrumbItem[] = [];

    crumbs.push({
      name: "Home",
      path: getFallbackBaseFolderForTab()
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

      let targetUrl = projectDocumentsUrl;
      let targetLibrary = libraryName;

      let result: { items: DocumentItem[]; totalCount: number } = { items: [], totalCount: 0 };
      const shouldLoadSplitRoot = !targetUrl && currentFolderPath === '';
      const resolvedFolderPath = targetUrl
        ? currentFolderPath
        : (shouldLoadSplitRoot ? '' : resolveLibraryFolderPath(currentFolderPath));

      if (targetUrl) {
        result = await getDocumentsByServerRelativeUrl(targetUrl, currentFolderPath, page, pageSize);
      } else if (targetLibrary) {
        result = await getProjectDocuments(targetLibrary, resolvedFolderPath, page, pageSize, businessProjectId);
      }



      setDocuments(result.items);

      // Multi-level Auto-navigation (Double Jump)
      if (!hasAutoNavigated) {
        const targetFolders = ["Projects", "Bids", "Lost", "Closed"];

        // Level 1: At Root, look for organizational folders
        if (currentFolderPath === "") {
          const orgFolder = result.items.find(item => {
            if (item.FSObjType !== 1) return false;
            const folderName = item.FileLeafRef.toLowerCase();
            return targetFolders.some(target =>
              folderName === target.toLowerCase() ||
              folderName === (target.toLowerCase() + "s") ||
              folderName.includes(target.toLowerCase())
            );
          });

          if (orgFolder) {
            console.log("NESTED_NAV: Entering Org Folder:", orgFolder.FileLeafRef);
            setCurrentFolderPath(orgFolder.FileLeafRef);
            // Don't set hasAutoNavigated yet, we want to try Level 2 in the next load
            return;
          }
        }

        // Level 2: Inside an organizational folder, look for specific project folder
        else if (targetFolders.includes(currentFolderPath)) {
          const projectFolder = result.items.find(item =>
            item.FSObjType === 1 && (
              item.FileLeafRef === projectCode ||
              item.FileLeafRef === businessProjectId ||
              (projectTitle && item.FileLeafRef.includes(projectTitle))
            )
          );

          if (projectFolder) {
            console.log("NESTED_NAV: Entering Project Folder:", projectFolder.FileLeafRef);
            setCurrentFolderPath(currentFolderPath + "/" + projectFolder.FileLeafRef);
            setHasAutoNavigated(true); // Stop here
            return;
          } else {
            // No project folder found inside org folder, stop auto-nav
            setHasAutoNavigated(true);
          }
        } else {
          // Already deeper than org folder, stop auto-nav
          setHasAutoNavigated(true);
        }
      }

      setPagination({
        currentPage: page,
        pageSize: pageSize,
        totalItems: result.totalCount,
        totalPages: Math.ceil(result.totalCount / pageSize),
      });
      setSelectedItems(new Set());
    } catch (error) {

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

  const getFileIcon = (item: DocumentItem): React.ReactNode => {
    if (item.FSObjType === 1) return <Folder size={20} color="#0078d4" />;
    return getSharedFileIcon(item.File_x0020_Type);
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
      let targetUrl = projectDocumentsUrl;
      let targetLibrary = libraryName;

      if (targetUrl) {
        if (relativePath) {
          await uploadDocumentWithFolderCreation(targetUrl, file, relativePath);
        } else {
          await uploadDocumentByServerRelativeUrl(targetUrl, file, currentFolderPath);
        }
      } else if (targetLibrary) {
        const uploadPath = resolveLibraryFolderPath(relativePath ? relativePath : currentFolderPath);
        await uploadProjectDocument(targetLibrary, file, uploadPath);
      }

      await loadDocuments(1);
      toast.success("Document uploaded successfully!", {
        position: "top-right",
        autoClose: 3000,
      });
    } catch (error) {

      toast.error("Failed to upload document. Please try again.", {
        position: "top-right",
        autoClose: 5000,
      });
      throw error;
    }
  };

  const handleCreateFolder = async (folderName: string): Promise<void> => {
    try {
      let targetUrl: string = projectDocumentsUrl;
      let targetLibrary: string = libraryName;

      if (targetUrl) {
        // Correctly pass the new folder name and the current parent path
        await createFolderByServerRelativeUrl(targetUrl, folderName, currentFolderPath);
      } else if (targetLibrary) {
        // Fallback for library-based discovery
        const resolvedCurrentPath = resolveLibraryFolderPath(currentFolderPath);
        const uploadPath = resolvedCurrentPath ? `${resolvedCurrentPath}/${folderName}` : folderName;
        // Assuming we need a createFolderByLibrary function or similar
        // For now, let's use the server-relative path since we can resolve it
        const rootFolder = await getLibraryRootFolder(targetLibrary);
        await createFolderByServerRelativeUrl(rootFolder.ServerRelativeUrl, folderName, resolvedCurrentPath);
      } else {
        throw new Error('No document URL configured for this tab');
      }

      await loadDocuments(pagination.currentPage);
      toast.success("Folder created successfully!", {
        position: "top-right",
        autoClose: 3000,
      });
    } catch (error) {
      console.error("FOLDER_CREATE_ERROR:", error);
      toast.error(`Failed to create folder: ${error || "Please try again."}`, {
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

      let targetUrl: string = projectDocumentsUrl;
      let targetLibrary: string = libraryName;

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
      openFilePreview(item);
    }
  };

  const handleBreadcrumbClick = (path: string): void => {
    setCurrentFolderPath(path);
    setSelectedItems(new Set());
  };

  const handleSort = (field: SortField): void => {
    if (sortField === field) {
      setSortDirection(prev => prev === 'asc' ? 'desc' : 'asc');
      return;
    }

    setSortField(field);
    setSortDirection(field === 'modified' ? 'desc' : 'asc');
  };

  const getSortIndicator = (field: SortField): React.ReactNode => {
    if (sortField !== field) {
      return (
        <svg width="12" height="12" viewBox="0 0 12 12" fill="none" style={{ opacity: 0.3 }}>
          <path d="M6 2L9 5H3L6 2Z" fill="currentColor" />
          <path d="M6 10L3 7H9L6 10Z" fill="currentColor" />
        </svg>
      );
    }
    if (sortDirection === 'asc') {
      return (
        <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
          <path d="M6 2L10 7H2L6 2Z" fill="#1f4ed8" />
        </svg>
      );
    }
    return (
      <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
        <path d="M6 10L2 5H10L6 10Z" fill="#1f4ed8" />
      </svg>
    );
  };

  const sortedDocuments = useMemo(() => {
    const sorted = [...documents];

    sorted.sort((a, b) => {
      const aIsFolder = a.FSObjType === 1;
      const bIsFolder = b.FSObjType === 1;

      // Keep folders grouped before files.
      if (aIsFolder !== bIsFolder) {
        return aIsFolder ? -1 : 1;
      }

      let comparison = 0;

      switch (sortField) {
        case 'name':
          comparison = (a.FileLeafRef || '').localeCompare((b.FileLeafRef || ''), undefined, { numeric: true, sensitivity: 'base' });
          break;
        case 'size': {
          const aValue = aIsFolder ? (a.ItemCount ?? 0) : (a.Length ?? 0);
          const bValue = bIsFolder ? (b.ItemCount ?? 0) : (b.Length ?? 0);
          comparison = aValue - bValue;
          break;
        }
        case 'modified': {
          const aValue = new Date(a.Modified).getTime() || 0;
          const bValue = new Date(b.Modified).getTime() || 0;
          comparison = aValue - bValue;
          break;
        }
        case 'by':
          comparison = (a.Editor?.Title || '').localeCompare((b.Editor?.Title || ''), undefined, { sensitivity: 'base' });
          break;
      }

      if (comparison === 0) {
        comparison = (a.FileLeafRef || '').localeCompare((b.FileLeafRef || ''), undefined, { numeric: true, sensitivity: 'base' });
      }

      return sortDirection === 'asc' ? comparison : -comparison;
    });

    return sorted;
  }, [documents, sortField, sortDirection]);


  const hasAnyUrl = projectDocumentsUrl || libraryName;
  const isPipelineProject = projectStatus === 'Potential';

  // Override projectDocumentsUrl with uncategorisedDocumentsUrl when isNonCmap
  React.useEffect(() => {
    if (isNonCmap && uncategorisedDocumentsUrl) {
      // For Non-CMAP, use the stored folder URL directly.
      const baseUrl = uncategorisedDocumentsUrl.replace(/\/$/, '');
      setProjectDocumentsUrl(baseUrl);
      // Reset folder path to load from location root
      setCurrentFolderPath('');
    }
  }, [isNonCmap, uncategorisedDocumentsUrl]);

  return (
    <>
      <div className={styles.documentsWrapper}>
        <div className={styles.tabRow}>
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
                  {crumb.name === "Home" ? <i className="ms-Icon ms-Icon--Home" aria-hidden="true" style={{ fontSize: '14px' }}></i> : crumb.name}
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
                <th className={styles.nameCol}>
                  <button type="button" className={styles.sortHeaderButton} onClick={() => handleSort('name')}>
                    <span>Name</span>
                    <span className={styles.sortIndicator}>{getSortIndicator('name')}</span>
                  </button>
                </th>
                <th className={styles.modifiedCol}>
                  <button type="button" className={styles.sortHeaderButton} onClick={() => handleSort('size')}>
                    <span>Size / Items</span>
                    <span className={styles.sortIndicator}>{getSortIndicator('size')}</span>
                  </button>
                </th>
                <th className={styles.modifiedCol}>
                  <button type="button" className={styles.sortHeaderButton} onClick={() => handleSort('modified')}>
                    <span>Modified</span>
                    <span className={styles.sortIndicator}>{getSortIndicator('modified')}</span>
                  </button>
                </th>
                <th className={styles.byCol}>
                  <button type="button" className={styles.sortHeaderButton} onClick={() => handleSort('by')}>
                    <span>By</span>
                    <span className={styles.sortIndicator} style={{ display: 'inline-flex', alignItems: 'center' }}>{getSortIndicator('by')}</span>
                  </button>
                </th>
                <th className={styles.actionCol}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {sortedDocuments.length === 0 ? (
                <tr>
                  <td colSpan={6} className={styles.noData}>
                    No documents found
                  </td>
                </tr>
              ) : (
                sortedDocuments.map((doc) => (
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
                      <div className={styles.nameCell}>
                        <span className={styles.fileIcon}>{getFileIcon(doc)}</span>
                        <span className={styles.fileName}>{doc.FileLeafRef}</span>
                      </div>
                    </td>
                    <td className={styles.modifiedCol}>
                      {doc.FSObjType === 1 ? `${doc.ItemCount ?? 0} items` : doc.Length ? (doc.Length < 1024 ? `${doc.Length} B` : doc.Length < 1048576 ? `${(doc.Length / 1024).toFixed(1)} KB` : `${(doc.Length / 1048576).toFixed(1)} MB`) : ''}
                    </td>
                    <td className={styles.modifiedCol}>{formatDate(doc.Modified)}</td>
                    <td className={styles.byCol}>{doc.Editor?.Title || ""}</td>
                    <td className={styles.actionCol}>
                      <ActionMenu
                        items={[
                          {
                            key: 'share',
                            text: 'Share',
                            onRender: () => (
                              <CommonShareComponent
                                item={{
                                  ...doc,
                                  ServerRelativeUrl: doc.FileRef || doc.ServerRelativeUrl,
                                  IsFolder: doc.FSObjType === 1,
                                }}
                                context={context}
                                buttonType="menuItem"
                              />
                            ),
                            disabled: isUserRestricted || !canEdit,
                          },
                          {
                            key: 'copy_link',
                            text: 'Copy link',
                            icon: <Link size={16} />,
                            onClick: () => handleCopyLink(
                              doc.FileRef || doc.ServerRelativeUrl,
                              doc.FileLeafRef,
                              context.pageContext.web.absoluteUrl
                            ),
                            disabled: false,
                          },
                          {
                            key: 'download',
                            text: 'Download',
                            icon: <Download size={16} />,
                            onClick: () => commonHandleDownload(doc.FileRef, doc.FileLeafRef, isItemFolder(doc), downloadProjectDocument),
                            disabled: isItemFolder(doc),
                          },
                          {
                            key: 'view',
                            text: 'View',
                            icon: <Eye size={16} />,
                            onClick: () => openFilePreview(doc),
                            disabled: isItemFolder(doc),
                          },
                          {
                            key: 'open_in_new_tab',
                            text: 'Open in New Tab',
                            icon: <ExternalLink size={16} />,

                            onClick: () => {
                              const openInNewTabUrl = doc.FileRef || doc.ServerRelativeUrl || '';
                              if (!openInNewTabUrl) return;

                              const ext = doc.FileLeafRef.split('.').pop()?.toLowerCase() || '';
                              const officePreviewExts = ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'pdf'];

                              if (officePreviewExts.includes(ext)) {
                                let sourceDocPath = openInNewTabUrl;
                                let siteUrl = context.pageContext.web.absoluteUrl;

                                // Extract pathname if it's a full URL
                                if (openInNewTabUrl.startsWith('http')) {
                                  try {
                                    sourceDocPath = new URL(openInNewTabUrl).pathname;
                                  } catch {
                                    // Keep original if parsing fails
                                  }
                                }

                                // Detect if file is from external site and use correct site URL
                                if (sourceDocPath.startsWith('/sites/')) {
                                  const pathParts = sourceDocPath.split('/');
                                  if (pathParts.length >= 3 && pathParts[1] === 'sites') {
                                    const siteName = pathParts[2];
                                    siteUrl = `${window.location.origin}/sites/${siteName}`;
                                  }
                                }

                                const viewerUrl = `${siteUrl}/_layouts/15/Doc.aspx?sourcedoc=${encodeURIComponent(sourceDocPath)}&action=view`;
                                window.open(viewerUrl, '_blank', 'noopener,noreferrer');
                                return;
                              }

                              const url = openInNewTabUrl.startsWith('http')
                                ? openInNewTabUrl
                                : `${window.location.origin}${openInNewTabUrl}`;

                              window.open(url, '_blank', 'noopener,noreferrer');
                            },
                            disabled: isItemFolder(doc),
                          },
                          {
                            key: 'manage_access',
                            text: 'Manage Access',
                            icon: <Key size={16} />,
                            onClick: () => {
                              if (onShowShareAccess) {
                                onShowShareAccess([{
                                  ...doc,
                                  name: doc.FileLeafRef,
                                  fileRef: doc.FileRef || doc.ServerRelativeUrl,
                                  isFolder: doc.FSObjType === 1
                                }]);
                              }
                            },
                            disabled: false,
                          },
                          {
                            key: 'move_to',
                            text: 'Move to',
                            icon: <Folder size={16} />,
                            onClick: () => openMoveCopy(doc, 'Move'),
                            disabled: false,
                          },
                          {
                            key: 'copy_to',
                            text: 'Copy to',
                            icon: <Copy size={16} />,
                            onClick: () => openMoveCopy(doc, 'Copy'),
                            disabled: false,
                          }
                        ]}
                      />
                    </td>
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

      <FilePreview
        isOpen={filePreview.isOpen}
        onDismiss={closeFilePreview}
        fileUrl={filePreview.fileUrl}
        fileName={filePreview.fileName}
        filePath={filePreview.filePath}
        siteUrl={context.pageContext.web.absoluteUrl}
      />

      {moveCopyState.item && (
        <MoveCopyModal
          isOpen={moveCopyState.isOpen}
          onClose={() => setMoveCopyState(prev => ({ ...prev, isOpen: false }))}
          onConfirm={handleMoveCopyConfirm}
          title={moveCopyState.type}
          sourceItemName={moveCopyState.item.FileLeafRef}
          initialPath={moveCopyPaths.initialPath}
          libraryName={moveCopyPaths.libraryName}
        />
      )}
    </>
  );
};

export default ProjectDocuments;