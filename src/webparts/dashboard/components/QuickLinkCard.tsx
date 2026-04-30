import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './Dashboard.module.scss';

interface QuickLinkCardProps {
    title: string;
    url: string;
    backgroundColor: string;
    icon: string;
    hasChildren?: boolean;
    fetchChildren?: (parentTitle: string) => Promise<Array<{ name: string; url: string }>>;
}

const QuickLinkCard: React.FC<QuickLinkCardProps> = ({ 
    title, 
    url, 
    backgroundColor, 
    icon, 
    hasChildren,
    fetchChildren
}) => {
    const [isDropdownOpen, setIsDropdownOpen] = useState(false);
    const [childLinks, setChildLinks] = useState<Array<{ name: string; url: string }>>([]);
    const [isLoadingChildren, setIsLoadingChildren] = useState(false);

    useEffect(() => {
        if (isDropdownOpen && hasChildren && fetchChildren && childLinks.length === 0) {
            setIsLoadingChildren(true);
            fetchChildren(title)
                .then(setChildLinks)
                .catch(error => {
                    console.error('Error loading child links:', error);
                    setChildLinks([]);
                })
                .then(() => setIsLoadingChildren(false));
        }
    }, [isDropdownOpen, hasChildren, fetchChildren, title, childLinks.length]);

    const handleClick = (urlToOpen?: string): void => {
        const targetUrl = urlToOpen || url;
        if (targetUrl && targetUrl !== '#') {
            window.open(targetUrl, '_blank', 'noopener,noreferrer');
        }
        setIsDropdownOpen(false);
    };

    const handleDropdownClick = (e: React.MouseEvent): void => {
        e.stopPropagation();
        setIsDropdownOpen(!isDropdownOpen);
    };

    const handleLocationSelect = (locationUrl: string, e: React.MouseEvent): void => {
        e.stopPropagation();
        handleClick(locationUrl);
    };

    const badgeColor = backgroundColor
        ? (backgroundColor.startsWith('#') ? backgroundColor : '#' + backgroundColor)
        : '#6366f1';

    return (
        <div
            className={styles.quickLinkCard}
            style={{ backgroundImage: `url(${icon})` }}
            onClick={() => handleClick()}
            role="button"
            tabIndex={0}
            onKeyPress={(e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    handleClick();
                }
            }}
        >
            {/* Dark gradient overlay */}
            <div className={styles.quickLinkOverlay} />

            {/* Dropdown button for child items */}
            {hasChildren && (
                <div className={styles.quickLinkDropdownButton}>
                    <button
                        className={styles.dropdownTrigger}
                        onClick={handleDropdownClick}
                        aria-label="Select option"
                        title="Select different location"
                    >
                        ⋯
                    </button>
                    {isDropdownOpen && (
                        <div className={styles.quickLinkDropdownMenu}>
                            {isLoadingChildren ? (
                                <div className={styles.dropdownOption} style={{ color: '#999' }}>
                                    Loading...
                                </div>
                            ) : childLinks.length > 0 ? (
                                childLinks.map((child, index) => (
                                    <button
                                        key={index}
                                        className={styles.dropdownOption}
                                        onClick={(e) => handleLocationSelect(child.url, e)}
                                    >
                                        {child.name}
                                    </button>
                                ))
                            ) : (
                                <div className={styles.dropdownOption} style={{ color: '#999' }}>
                                    No options available
                                </div>
                            )}
                        </div>
                    )}
                </div>
            )}

            {/* Bottom content: badge + title */}
            <div className={styles.quickLinkBottom}>
                {/* <div className={styles.quickLinkBadge} style={{ backgroundColor: badgeColor }}>
                    <img src={icon} alt={title} className={styles.quickLinkBadgeIcon} />
                </div> */}
                <h3 className={styles.quickLinkTitle}>{title}</h3>
            </div>
        </div>
    );
};

export default QuickLinkCard;
