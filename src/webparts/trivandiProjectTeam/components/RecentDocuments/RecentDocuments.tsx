import * as React from "react";
import { useState, useEffect } from "react";
import styles from "./RecentDocuments.module.scss";
import { getRecentDocuments } from "../../../../shared/services/projectService";
import GlobalLoader from "../../../../shared/component/GlobalLoader";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import FilePreview from "../../../projectDetails/components/FilePreview/FilePreview";
import { getFileIcon } from "../../../../shared/utils/iconHelper";

interface RecentDocument {
  Id: number;
  FileLeafRef: string;
  Modified: string;
  FileRef: string;
  File_x0020_Type?: string;
  ProjectTitle?: string;
  LibraryTitle?: string;
}

interface RecentDocumentsProps {
  context?: WebPartContext;
}

const RecentDocuments: React.FC<RecentDocumentsProps> = ({ context }) => {
  const [documents, setDocuments] = useState<RecentDocument[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [filePreview, setFilePreview] = useState<{
    isOpen: boolean;
    fileUrl: string;
    fileName: string;
    filePath?: string;
    canDownload?: boolean;
    siteUrl?: string;
  }>({
    isOpen: false,
    fileUrl: '',
    fileName: '',
    filePath: '',
    canDownload: true,
    siteUrl: '',
  });

  const loadRecentDocuments = async (): Promise<void> => {
    try {
      setLoading(true);
      const docs = await getRecentDocuments();
      setDocuments(docs);
    } catch (_error) {
      console.log("Failed to load recent documents", _error);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadRecentDocuments().catch((err: unknown) => console.log(err));
  }, []);

  const getRelativeTime = (dateString: string): string => {
    const date = new Date(dateString);
    const now = new Date();
    const diffMs = now.getTime() - date.getTime();
    const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
    const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));

    if (diffHours < 1) {
      const diffMins = Math.floor(diffMs / (1000 * 60));
      return diffMins < 1 ? "Just now" : `${diffMins} minute${diffMins > 1 ? "s" : ""} ago`;
    } else if (diffHours < 24) {
      return `${diffHours} hour${diffHours > 1 ? "s" : ""} ago`;
    } else if (diffDays === 1) {
      return "Yesterday";
    } else if (diffDays < 7) {
      return `${diffDays} days ago`;
    } else {
      return date.toLocaleDateString("en-US", { month: "short", day: "numeric" });
    }
  };

  const getSiteUrlForDocument = (fileRef: string): string => {
    const currentSiteUrl = context?.pageContext?.web?.absoluteUrl || window.location.origin;
    
    // 1. Resolve OneDrive files e.g. /personal/username/
    if (fileRef.toLowerCase().startsWith("/personal/")) {
      try {
        const url = new URL(currentSiteUrl);
        const hostParts = url.hostname.split(".");
        const tenantName = hostParts[0].replace("-my", "");
        
        const parts = fileRef.split("/");
        if (parts.length >= 3 && parts[1].toLowerCase() === "personal") {
          const personalUser = parts[2];
          return `https://${tenantName}-my.sharepoint.com/personal/${personalUser}`;
        }
        return `https://${tenantName}-my.sharepoint.com`;
      } catch (e) {
        // Fallback
      }
    }
    
    // 2. Resolve external SharePoint Site Collections e.g. /sites/PeopleHub
    if (fileRef.toLowerCase().startsWith("/sites/")) {
      const parts = fileRef.split("/");
      if (parts.length >= 3 && parts[1].toLowerCase() === "sites") {
        const siteName = parts[2];
        try {
          const url = new URL(currentSiteUrl);
          return `${url.protocol}//${url.hostname}/sites/${siteName}`;
        } catch (e) {
          return `${window.location.origin}/sites/${siteName}`;
        }
      }
    }

    return currentSiteUrl;
  };

  const handleDocumentClick = (doc: RecentDocument): void => {
    if (doc.FileRef) {
      const computedSiteUrl = getSiteUrlForDocument(doc.FileRef);
      setFilePreview({
        isOpen: true,
        fileUrl: doc.FileRef,
        fileName: doc.FileLeafRef,
        filePath: doc.FileRef,
        canDownload: true,
        siteUrl: computedSiteUrl,
      });
    }
  };

  const closeFilePreview = (): void => {
    setFilePreview({
      isOpen: false,
      fileUrl: '',
      fileName: '',
      filePath: '',
      canDownload: true,
      siteUrl: '',
    });
  };

  if (loading) {
    return (
      <div className={styles.wrapper}>
        {/* <h2 className={styles.title}>Recent Documents</h2> */}
        <GlobalLoader variant="content" />
      </div>
    );
  }

  return (
    <div className={styles.wrapper}>
      {/* <h2 className={styles.title}>Recent Documents</h2> */}
      <div className={styles.documentsList}>
        {documents.length === 0 ? (
          <div className={styles.emptyState}>No recent documents found</div>
        ) : (
          documents.map((doc) => (
            <div
              key={doc.Id}
              className={styles.documentItem}
              onClick={() => handleDocumentClick(doc)}
            >
              <div className={styles.iconWrapper}>
                <span className={styles.fileIcon}>{getFileIcon(doc.File_x0020_Type)}</span>
              </div>
              <div className={styles.documentInfo}>
                <div className={styles.fileName}>{doc.FileLeafRef}</div>
                <div className={styles.metadata}>
                  {doc.ProjectTitle || doc.LibraryTitle || "Unknown Project"} • {getRelativeTime(doc.Modified)}
                </div>
              </div>
            </div>
          ))
        )}
      </div>
      <FilePreview
        isOpen={filePreview.isOpen}
        onDismiss={closeFilePreview}
        fileUrl={filePreview.fileUrl}
        fileName={filePreview.fileName}
        filePath={filePreview.filePath}
        siteUrl={filePreview.siteUrl}
        canDownload={filePreview.canDownload}
      />
    </div>
  );
};

export default RecentDocuments;