import * as React from 'react';
import { Link } from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';

export interface ISearchResultUrlProps {
  url?: string;
  searchService?: any;
}

export const SearchResultUrl: React.FC<ISearchResultUrlProps> = ({ url, searchService }) => {
  const [title, setTitle] = React.useState<string | null>(null);
  const [loading, setLoading] = React.useState(false);

  React.useEffect(() => {
    let mounted = true;
    if (!url) return;
    const resolve = async () => {
      setLoading(true);
      try {
        if (searchService && typeof searchService.getTitleForUrl === 'function') {
          const t = await searchService.getTitleForUrl(url);
          if (mounted) setTitle(t || url);
        } else {
          // Fallback: show url when no service available
          if (mounted) setTitle(url);
        }
      } catch (e) {
        if (mounted) setTitle(url);
      } finally {
        if (mounted) setLoading(false);
      }
    };
    resolve();
    return () => { mounted = false; };
  }, [url, searchService]);

  if (!url) return null;

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
          maxWidth: 'calc(100% - 24px)'
        }}
        title={title || url}
      >
        {loading ? (title || 'Loading...') : (title || url)}
      </a>
    </div>
  );
};
