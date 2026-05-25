import { toast } from 'react-toastify';

/**
 * Copies a file URL to the clipboard for viewing (not downloading).
 * For Office/PDF files, uses the SharePoint Doc.aspx viewer URL.
 * Handles both internal and external SharePoint sites.
 * @param fileRef The server-relative URL of the file.
 * @param fileLeafRef The file name with extension.
 * @param webAbsoluteUrl The absolute URL of the web (site) - used as fallback if fileRef doesn't contain site info.
 */
export const handleCopyLink = (fileRef: string, fileLeafRef?: string, webAbsoluteUrl?: string): void => {
  if (!fileRef) {
    toast.error('Link not available');
    return;
  }

  let linkToCopy = fileRef;

  console.log('📋 handleCopyLink called:', { fileRef, fileLeafRef, webAbsoluteUrl });

  // Generate Doc.aspx viewer URL for Office/PDF files
  if (fileLeafRef) {
    const ext = fileLeafRef.split('.').pop()?.toLowerCase() || '';
    const officePreviewExts = ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'pdf'];

    console.log('📄 File extension:', ext, 'isOffice:', officePreviewExts.includes(ext));

    if (officePreviewExts.includes(ext)) {
      // Determine source doc path (convert full URL to server-relative if needed)
      let sourceDocPath = fileRef;
      if (fileRef.startsWith('http')) {
        try {
          sourceDocPath = new URL(fileRef).pathname;
          console.log('🔗 Extracted pathname from URL:', sourceDocPath);
        } catch (e) {
          console.error('❌ Failed to extract pathname:', e);
        }
      }

      // Determine the correct site URL for Doc.aspx viewer
      let siteUrl = webAbsoluteUrl || window.location.origin;

      // If FileRef contains site path (e.g., /sites/externalsite/...), extract and use that site's URL
      if (sourceDocPath.startsWith('/sites/')) {
        const pathParts = sourceDocPath.split('/');
        if (pathParts.length >= 3 && pathParts[1] === 'sites') {
          const siteName = pathParts[2]; // e.g., 'externalsite'
          const extractedSiteUrl = `${window.location.origin}/sites/${siteName}`;
          console.log('🌐 Detected external site:', siteName, '→', extractedSiteUrl);
          siteUrl = extractedSiteUrl;
        }
      }

      linkToCopy = `${siteUrl}/_layouts/15/Doc.aspx?sourcedoc=${encodeURIComponent(sourceDocPath)}&action=view`;
      console.log('🎯 Generated Doc.aspx URL:', linkToCopy);
    }
  }

  const absoluteUrl = linkToCopy.startsWith('http') 
    ? linkToCopy 
    : `${window.location.origin}${linkToCopy}`;

  console.log('📌 Final URL to copy:', absoluteUrl);

  navigator.clipboard.writeText(absoluteUrl).then(() => {
    toast.success('Link copied to clipboard');
    console.log('✅ Link copied successfully');
  }).catch(err => {
    console.error('Failed to copy: ', err);
    toast.error('Failed to copy link');
  });
};

/**
 * Checks if an item is a folder based on common document object shapes.
 */
export const isItemFolder = (doc: any): boolean => {
  if (!doc) return false;
  return doc.isFolder === true || doc.FSObjType === 1;
};

/**
 * Handles document download with folder check.
 * @param fileUrl Server-relative URL of the file.
 * @param fileName Name of the file.
 * @param isFolder Boolean indicating if the item is a folder.
 * @param downloadService Optional service function for handling the actual download logic.
 */
export const handleDownload = (
  fileUrl: string, 
  fileName: string, 
  isFolder: boolean, 
  downloadService?: (url: string, name: string) => Promise<void> | void
): void => {
  if (isFolder) {
    // Folders shouldn't be downloadable through this generic handler
    return;
  }

  if (!fileUrl) {
    toast.error('Download link not available');
    return;
  }

  if (downloadService) {
    try {
      void downloadService(fileUrl, fileName);
    } catch (error) {
      console.error('Download failed:', error);
      toast.error('Download failed');
    }
  } else {
    // Native browser download fallback
    const absoluteUrl = fileUrl.startsWith('http') ? fileUrl : `${window.location.origin}${fileUrl}`;
    const link = document.createElement('a');
    link.href = absoluteUrl;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }
};
