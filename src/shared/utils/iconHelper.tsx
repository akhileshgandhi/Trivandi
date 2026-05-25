import * as React from 'react';
import {
  FileText,
  File as FileIcon,
  Image as ImageIcon,
  FileSpreadsheet
} from 'lucide-react';
import { BiSolidFilePdf } from "react-icons/bi";
import { BsFiletypeDoc } from "react-icons/bs";

export const getFileIcon = (fileType?: string, iconSize: number = 20): React.ReactNode => {
  const ext = fileType?.toLowerCase() || '';

  switch (ext) {
    case 'pdf':
      return React.createElement(BiSolidFilePdf as any, { size: iconSize, color: "#ef4444" });
    case 'doc':
      return React.createElement(BsFiletypeDoc as any, { size: iconSize, color: "#2563eb" });
    case 'docx':
      return <FileText size={iconSize} color="#2563eb" />;
    case 'xls':
    case 'xlsx':
    case 'csv':
      return <FileSpreadsheet size={iconSize} color="#16a34a" />;
    case 'ppt':
    case 'pptx':
      return <FileText size={iconSize} color="#ea580c" />;
    case 'jpg':
    case 'jpeg':
    case 'png':
    case 'gif':
    case 'svg':
      return <ImageIcon size={iconSize} color="#9333ea" />;
    default:
      return <FileIcon size={iconSize} color="#64748b" />;
  }
};
