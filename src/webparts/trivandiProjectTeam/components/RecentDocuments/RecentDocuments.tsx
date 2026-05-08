import * as React from "react";
import { useState, useEffect } from "react";
import styles from "./RecentDocuments.module.scss";
import { getRecentDocuments } from "../../../../shared/services/projectService";
import GlobalLoader from "../../../../shared/component/GlobalLoader";


interface RecentDocument {
  Id: number;
  FileLeafRef: string;
  Modified: string;
  FileRef: string;
  File_x0020_Type?: string;
  ProjectTitle?: string;
  LibraryTitle?: string;
}

const RecentDocuments: React.FC = () => {
  const [documents, setDocuments] = useState<RecentDocument[]>([]);
  const [loading, setLoading] = useState<boolean>(true);

  const loadRecentDocuments = async (): Promise<void> => {
    try {
      setLoading(true);
      const docs = await getRecentDocuments();
      setDocuments(docs);
    } catch (error) {
      console.error("Error loading recent documents:", error);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadRecentDocuments().catch((err: unknown) => console.error(err));
  }, []);

  const getFileIcon = (fileType?: string): string => {
    const ext = fileType?.toLowerCase();
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

  const handleDocumentClick = (doc: RecentDocument): void => {
    if (doc.FileRef) {
      const siteUrl = window.location.origin;
      window.open(`${siteUrl}${doc.FileRef}`, "_blank");
    }
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
    </div>
  );
};

export default RecentDocuments;