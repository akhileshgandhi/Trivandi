export interface ISearchResult {
  id: string;
  title: string;
  webUrl: string;
  fileType: string;         // e.g. "docx", "pdf", "xlsx", "pptx"
  lastModified: string;     // ISO date string
  author: string;
  size: number;             // bytes
  summary: string;          // content snippet / description
  siteUrl: string;
  siteName: string;
  isStarred?: boolean;
  thumbnailUrl?: string;
  createdDate?: string;
  objectType?: string;
  relevanceScore?: number;
}
