import * as React from 'react';

/**
 * Format search result descriptions cleanly:
 * - Replaces <ddd/> tags with standard ellipsis "..."
 * - Wraps <c0>...</c0> highlighted query matches in beautiful themed highlight spans
 * - Dynamically matches background/color with the file brand accent color
 * - Debounces / handles manual highlights when no SharePoint tags are present (fallbacks)
 */
export const getCleanSharePointUrl = (webUrl: string, title: string): string => {
  if (!webUrl) return '';
  try {
    let cleaned = webUrl;
    // Replace SharePoint List Item detail page link (DispForm.aspx) with direct file path
    if (cleaned.includes('/Forms/DispForm.aspx')) {
      cleaned = cleaned.replace(/\/Forms\/DispForm\.aspx.*/i, `/${title}`);
    }
    // Remove sharing/redirect tokens like /:i:/r/ or /:f:/g/ or /:x:/r/ etc.
    cleaned = cleaned.replace(/\/:[a-z]:\/[a-z]\//i, '/');
    cleaned = cleaned.replace(/\/:[a-z]:\/r\//i, '/');
    cleaned = cleaned.replace(/\/:[a-z]:\/g\//i, '/');
    cleaned = cleaned.split('?')[0]; // Strip query parameters
    return cleaned;
  } catch (e) {
    return webUrl;
  }
};

export const getSharePointThumbnailUrl = (webUrl: string, title: string, resolution: number = 0): string => {
  try {
    if (!webUrl) return '';
    
    // Guard: do not attempt thumbnail if the URL contains spaces (plain or encoded)
    if (webUrl.includes(' ') || webUrl.includes('%20') || decodeURIComponent(webUrl).includes(' ')) {
      return '';
    }

    const cleanUrl = getCleanSharePointUrl(webUrl, title);
    const urlObj = new URL(cleanUrl);
    const host = urlObj.origin;
    const pathParts = urlObj.pathname.split('/');
    let sitePath = '';
    
    // Correctly resolve site collections and personal OneDrive personal paths
    if (host.includes('-my.sharepoint.com') && pathParts[1] === 'personal' && pathParts[2]) {
      sitePath = `/personal/${pathParts[2]}`;
    } else if (pathParts[1] === 'sites' && pathParts[2]) {
      sitePath = `/sites/${pathParts[2]}`;
    }
    
    const resParam = resolution ? `&resolution=${resolution}` : '&size=S';
    return `${host}${sitePath}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(cleanUrl)}${resParam}`;
  } catch (e) {
    return '';
  }
};

export const renderFormattedSummary = (
  summary: string, 
  query: string = '', 
  highlightColor: string = '#1a73e8'
): React.ReactNode => {
  if (!summary) return null;

  // First, replace <ddd/> and </ddd/> tags with standard ellipsis "..."
  let cleaned = summary.replace(/<\/?ddd\/?>/gi, '...');

  // Normalize consecutive ellipsis/spaces
  cleaned = cleaned.replace(/\s*\.{3,}\s*/g, ' ... ');
  cleaned = cleaned.replace(/^\s*\.\.\.\s+/, '... ');
  cleaned = cleaned.replace(/\s+\.\.\.\s*$/, ' ...');
  cleaned = cleaned.trim();

  // If there are highlighting tags (<c0>, etc.), parse those tags
  if (/<c\d+>/i.test(cleaned)) {
    const tagRegex = /(<c\d+>.*?<\/c\d+>)/gi;
    const splitParts = cleaned.split(tagRegex);

    return (
      <>
        {splitParts.map((part, index) => {
          const match = part.match(/<c\d+>(.*?)<\/c\d+>/i);
          if (match) {
            const highlightedText = match[1];
            return (
              <span
                key={index}
                style={{
                  backgroundColor: highlightColor === '#1a73e8' ? 'rgba(26, 115, 232, 0.15)' : `${highlightColor}15`,
                  color: highlightColor,
                  fontWeight: 'bold',
                  padding: '0 2px',
                  borderRadius: '2px',
                }}
              >
                {highlightedText}
              </span>
            );
          }
          return part;
        })}
      </>
    );
  }

  // If there are no highlight tags, but there is a query string, highlight the query string
  if (query && query.trim()) {
    // Escape special characters in query to prevent invalid RegEx
    const escapedQuery = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    // eslint-disable-next-line @rushstack/security/no-unsafe-regexp
    const parts = cleaned.split(new RegExp(`(${escapedQuery})`, 'gi'));
    return (
      <>
        {parts.map((part, i) => 
          part.toLowerCase() === query.toLowerCase() ? (
            <span 
              key={i} 
              style={{ 
                backgroundColor: highlightColor === '#1a73e8' ? 'rgba(26, 115, 232, 0.15)' : `${highlightColor}15`, 
                color: highlightColor, 
                fontWeight: 'bold', 
                padding: '0 2px', 
                borderRadius: '2px' 
              }}
            >
              {part}
            </span>
          ) : (
            <span key={i}>{part}</span>
          )
        )}
      </>
    );
  }

  return <>{cleaned}</>;
};
