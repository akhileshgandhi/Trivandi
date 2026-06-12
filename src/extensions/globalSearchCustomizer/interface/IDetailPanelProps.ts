export interface IDetailPanelProps {
  selectedFile: {
    id: string;
    title: string;
    author: string;
    originalAuthor?: string;
    date: string;
    size: string;
    description: string;
    url: string;
    type: string;
    starred: boolean;
    color: string;
    badgeColor: string;
    summary: string;
    siteName: string;
    siteUrl: string;
    webUrl: string;
    lastModified: string;
    libraryUrl?: string;
    thumbnailUrl?: string;
    thumbnailUrlLarge?: string;
    driveId?: string;
    itemId?: string;
  };
  searchQuery: string;
}
