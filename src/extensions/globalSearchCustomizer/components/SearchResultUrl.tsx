import * as React from 'react';
import { Link } from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';

export interface ISearchResultUrlProps {
  url?: string;
  currentTitle?: string;
  matchedProjectTitle?: string;
  searchService?: { getTitleForUrl: (webUrl: string) => Promise<string | null> };
}

const getReadableUrlLabel = (url: string): string => {
  try {
    const parsed = new URL(url);
    const segments = parsed.pathname
      .split('/')
      .filter(Boolean)
      .map(segment => decodeURIComponent(segment).replace(/^\d{4,}\s*/, '').replace(/[_-]+/g, ' ').trim())
      .filter(Boolean);

    const siteIndex = segments.findIndex(segment => segment.toLowerCase() === 'sites');
    const projectSegment = siteIndex >= 0 ? segments[siteIndex + 2] : '';
    const lastSegment = segments[segments.length - 1] || url;

    if (projectSegment && lastSegment && projectSegment.toLowerCase() !== lastSegment.toLowerCase()) {
      return `${lastSegment} in ${projectSegment}`;
    }

    return lastSegment || url;
  } catch (e) {
    return url;
  }
};


export const SearchResultUrl: React.FC<ISearchResultUrlProps> = ({ url, currentTitle, matchedProjectTitle, searchService }) => {
  // Only render if matchedProjectTitle exists
  if (!matchedProjectTitle) return null;

  if (!url) return null;

  const displayTitle = matchedProjectTitle.trim();

  return (
    <div className={styles.cardProjectHubRow} style={{ marginTop: '6px' }}>
      <Link size={12} />
      <a
        href={url}
        target="_blank"
        rel="noreferrer"
        style={{
          marginLeft: '6px',
          color: '#1a73e8',
          textDecoration: 'underline',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          whiteSpace: 'nowrap',
          display: 'inline-block',
          maxWidth: 'calc(100% - 24px)',
          fontSize: '12px',
        }}
        title={displayTitle}
      >
        {displayTitle}
      </a>
    </div>
  );
};
