export interface ISearchResult {
  id: string;
  title: string;
  webUrl: string;
  fileType: string;         // e.g. "docx", "pdf", "xlsx", "pptx"
  lastModified: string;     // ISO date string
  author: string;
  originalAuthor?: string;  // The embedded document author (if different from uploader)
  size: number;             // bytes
  summary: string;          // content snippet / description
  siteUrl: string;
  siteName: string;
  isStarred?: boolean;
  thumbnailUrl?: string;       // small (96px) — for result card
  thumbnailUrlLarge?: string;  // large (800px) — for detail panel preview
  createdDate?: string;
  objectType?: string;
  relevanceScore?: number;
  authorEmail?: string;
  driveId?: string;
  itemId?: string;
  authorPhotoUrl?: string;
  libraryUrl?: string;   // Document library root URL e.g. https://tenant.sharepoint.com/sites/Site/Shared%20Documents
}
