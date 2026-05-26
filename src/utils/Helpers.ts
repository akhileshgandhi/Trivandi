/**
 * Format a byte count into a human-readable string.
 */
export function formatFileSize(bytes: number): string {
  if (!bytes || bytes === 0) return '–';
  const units = ['B', 'KB', 'MB', 'GB'];
  let value = bytes;
  let unit = 0;
  while (value >= 1024 && unit < units.length - 1) {
    value /= 1024;
    unit++;
  }
  return `${value.toFixed(1)} ${units[unit]}`;
}

/**
 * Format an ISO date string into a short readable form.
 */
export function formatDate(iso: string): string {
  if (!iso) return '–';
  try {
    const d = new Date(iso);
    return d.toLocaleDateString('en-GB', { year: 'numeric', month: 'short', day: '2-digit' });
  } catch {
    return iso;
  }
}

/**
 * Truncate text to a max length.
 */
export function truncate(text: string, maxLength: number = 120): string {
  if (!text) return '';
  return text.length > maxLength ? text.slice(0, maxLength - 1) + '…' : text;
}

/**
 * Get a colour class name based on file type.
 */
export function getFileTypeColor(fileType: string): string {
  const map: Record<string, string> = {
    docx: '#7B61FF',
    doc: '#7B61FF',
    pdf: '#E84040',
    xlsx: '#22B07D',
    xls: '#22B07D',
    pptx: '#FF6B35',
    ppt: '#FF6B35',
    png: '#0095E8',
    jpg: '#0095E8',
    jpeg: '#0095E8',
    gif: '#0095E8',
    txt: '#8C8C8C',
  };
  return map[fileType?.toLowerCase()] || '#555E6D';
}

/**
 * Get the display badge label for a file type.
 */
export function getFileTypeBadge(fileType: string): string {
  return (fileType || 'FILE').toUpperCase();
}

/**
 * Extract the project hub display name from a URL.
 */
export function getSiteDisplayUrl(url: string): string {
  if (!url) return '';
  try {
    const u = new URL(url);
    return u.hostname + u.pathname;
  } catch {
    return url;
  }
}
