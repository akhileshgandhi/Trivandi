import * as React from 'react';

export interface IDocumentAnalysisProps {
  selectedFile: any;
  setSelectedFile: (file: any) => void;
  previewWidth: number;
  searchQuery: string;
  isResizingPreview: boolean;
  startResizingPreview: (e: React.MouseEvent) => void;
}
