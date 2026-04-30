import * as React from "react";
import { useState } from "react";
import styles from "./FileUploadModal.module.scss";

interface IFileUploadModalProps {
  isOpen: boolean;
  onClose: () => void;
  onUpload: (file: File, relativePath?: string) => Promise<void>;
  onCreateFolder?: (folderName: string) => Promise<void>;
}

const FileUploadModal: React.FC<IFileUploadModalProps> = ({ isOpen, onClose, onUpload, onCreateFolder }) => {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [selectedFiles, setSelectedFiles] = useState<FileList | null>(null);
  const [uploading, setUploading] = useState<boolean>(false);
  const [dragActive, setDragActive] = useState<boolean>(false);
  const [showFolderInput, setShowFolderInput] = useState<boolean>(false);
  const [folderName, setFolderName] = useState<string>("");
  const [uploadMode, setUploadMode] = useState<'file' | 'folder'>('file');

  if (!isOpen) return null;

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>): void => {
    if (e.target.files && e.target.files.length > 0) {
      const file = e.target.files[0];
      setSelectedFile(file);
      setSelectedFiles(null);
      setUploadMode('file');
    }
  };

  const handleFolderSelect = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const files = event.target.files;
    if (files && files.length > 0) {
      setSelectedFiles(files);
      setSelectedFile(null);
      setUploadMode('folder');
    }
  };

  const handleCreateFolder = async (): Promise<void> => {
    if (!folderName.trim() || !onCreateFolder) return;

    try {
      setUploading(true);
      await onCreateFolder(folderName.trim());
      setFolderName("");
      setShowFolderInput(false);
      onClose();
    } catch (error) {
      console.error("Create folder failed:", error);
      alert("Failed to create folder. Please try again.");
    } finally {
      setUploading(false);
    }
  };

  const handleDrag = (e: React.DragEvent): void => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e: React.DragEvent): void => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);

    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      setSelectedFile(e.dataTransfer.files[0]);
    }
  };

  const handleUpload = async (): Promise<void> => {
    if (!selectedFile && !selectedFiles) return;

    try {
      setUploading(true);
      
      if (uploadMode === 'file' && selectedFile) {
        await onUpload(selectedFile);
      } else if (uploadMode === 'folder' && selectedFiles) {
        // Upload all files with their relative paths
        for (let i = 0; i < selectedFiles.length; i++) {
          const file = selectedFiles[i];
          const relativePath = (file as any).webkitRelativePath || file.name;
          const folderPath = relativePath.substring(0, relativePath.lastIndexOf('/'));
          await onUpload(file, folderPath);
        }
      }
      
      setSelectedFile(null);
      setSelectedFiles(null);
      onClose();
    } catch (error) {
      console.error("Upload failed:", error);
      alert("Failed to upload file. Please try again.");
    } finally {
      setUploading(false);
    }
  };

  const handleClose = (): void => {
    if (!uploading) {
      setSelectedFile(null);
      setSelectedFiles(null);
      setShowFolderInput(false);
      setFolderName("");
      setUploadMode('file');
      onClose();
    }
  };

  const getFileIcon = (file: File): string => {
    const ext = file.name.split('.').pop()?.toLowerCase();
    switch (ext) {
      case 'pdf':
        return '📄';
      case 'doc':
      case 'docx':
        return '📘';
      case 'xls':
      case 'xlsx':
        return '📊';
      case 'ppt':
      case 'pptx':
        return '📙';
      case 'jpg':
      case 'jpeg':
      case 'png':
      case 'gif':
        return '🖼️';
      case 'mp4':
      case 'avi':
      case 'mov':
        return '🎬';
      default:
        return '📄';
    }
  };

  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return "0 Bytes";
    const k = 1024;
    const sizes = ["Bytes", "KB", "MB", "GB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + " " + sizes[i];
  };

  return (
    <div className={styles.modalOverlay} onClick={handleClose}>
      <div className={styles.modalContent} onClick={(e) => e.stopPropagation()}>
        <div className={styles.modalHeader}>
          <h2 className={styles.modalTitle}>Upload Document</h2>
          <button 
            className={styles.closeButton} 
            onClick={handleClose}
            disabled={uploading}
          >
            ×
          </button>
        </div>

        <div className={styles.modalBody}>
          {showFolderInput ? (
            <div className={styles.folderInput}>
              <div className={styles.folderIcon}>📁</div>
              <h3>Create New Folder</h3>
              <input
                type="text"
                placeholder="Enter folder name"
                value={folderName}
                onChange={(e) => setFolderName(e.target.value)}
                className={styles.folderNameInput}
                disabled={uploading}
                autoFocus
              />
              <div className={styles.folderActions}>
                <button
                  className={styles.cancelButton}
                  onClick={() => {
                    setShowFolderInput(false);
                    setFolderName("");
                  }}
                  disabled={uploading}
                >
                  Cancel
                </button>
                <button
                  className={styles.createButton}
                  onClick={handleCreateFolder}
                  disabled={!folderName.trim() || uploading}
                >
                  {uploading ? "Creating..." : "Create Folder"}
                </button>
              </div>
            </div>
          ) : (
            <div
              className={`${styles.dropZone} ${dragActive ? styles.dragActive : ""}`}
              onDragEnter={handleDrag}
              onDragLeave={handleDrag}
              onDragOver={handleDrag}
              onDrop={handleDrop}
            >
              {selectedFile || selectedFiles ? (
                <div className={styles.fileSelected}>
                  {uploadMode === 'file' && selectedFile ? (
                    <>
                      <div className={styles.fileIcon}>{getFileIcon(selectedFile)}</div>
                      <div className={styles.fileInfo}>
                        <div className={styles.fileName}>{selectedFile.name}</div>
                        <div className={styles.fileSize}>{formatFileSize(selectedFile.size)}</div>
                      </div>
                      <button
                        className={styles.removeFile}
                        onClick={() => setSelectedFile(null)}
                        disabled={uploading}
                      >
                        ×
                      </button>
                    </>
                  ) : uploadMode === 'folder' && selectedFiles ? (
                    <>
                      <div className={styles.fileIcon}>📁</div>
                      <div className={styles.fileInfo}>
                        <div className={styles.fileName}>Folder ({selectedFiles.length} files)</div>
                        <div className={styles.fileSize}>{Array.from(selectedFiles).reduce((total, file) => total + file.size, 0)} bytes total</div>
                      </div>
                      <button
                        className={styles.removeFile}
                        onClick={() => setSelectedFiles(null)}
                        disabled={uploading}
                      >
                        ×
                      </button>
                    </>
                  ) : null}
                </div>
              ) : (
                <>
                  <div className={styles.uploadIcon}>📁</div>
                  <p className={styles.dropText}>Drag and drop file here</p>
                  <p className={styles.orText}>or</p>
                  <div className={styles.actionButtons}>
                    <label className={styles.browseButton}>
                      Browse Files
                      <input
                        type="file"
                        onChange={handleFileSelect}
                        style={{ display: "none" }}
                        disabled={uploading}
                      />
                    </label>
                    {/* <label className={styles.browseButton}>
                      Browse Folders
                      <input
                        type="file"
                        onChange={handleFolderSelect}
                        style={{ display: "none" }}
                        disabled={uploading}
                        {...({ webkitdirectory: "", directory: "" } as any)}
                        multiple
                      />
                    </label> */}
                    {onCreateFolder && (
                      <button
                        className={styles.folderButton}
                        onClick={() => setShowFolderInput(true)}
                        disabled={uploading}
                      >
                        📁 New Folder
                      </button>
                    )}
                  </div>
                </>
              )}
            </div>
          )}
        </div>

        <div className={styles.modalFooter}>
          {!showFolderInput && (
            <>
              <button
                className={styles.cancelButton}
                onClick={handleClose}
                disabled={uploading}
              >
                Cancel
              </button>
              <button
                className={styles.uploadButton}
                onClick={handleUpload}
                disabled={(!selectedFile && !selectedFiles) || uploading}
              >
                {uploading ? "Uploading..." : uploadMode === 'folder' && selectedFiles ? `Upload ${selectedFiles.length} files` : "Upload"}
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  );
};

export default FileUploadModal;