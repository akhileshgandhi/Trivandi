export interface IDetailPanelProps {
  selectedFile: {
    id: string;
    title: string;
    author: string;
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
  };
  searchQuery: string;
}
