import * as React from 'react';
import styles from './Dashboard.module.scss';

interface ToolCardProps {
  title: string;
  description: string;
  color: string;
  icon: string;
  link: string;
}

const ToolCard: React.FC<ToolCardProps> = ({ title, description, color, icon, link }) => {
  const handleClick = (): void => {
    if (link && link !== '#') {
      window.open(link, '_blank', 'noopener,noreferrer');
    }
  };

  const badgeColor = color
    ? (color.startsWith('#') ? color : '#' + color)
    : '#6366f1';

  return (
    <div
      className={styles.toolCard}
      style={{ backgroundImage: `url(${icon})` }}
      onClick={handleClick}
      role="button"
      tabIndex={0}
      onKeyPress={(e) => {
        if (e.key === 'Enter' || e.key === ' ') handleClick();
      }}
    >
      {/* Dark gradient overlay */}
      <div className={styles.toolOverlay} />

      {/* Bottom content: badge + title + description */}
      <div className={styles.toolBottom}>
        <div className={styles.toolBadge} style={{ backgroundColor: badgeColor }}>
          <img src={icon} alt={title} className={styles.toolBadgeIcon} />
        </div>
        <div className={styles.toolContent}>
          <h4 className={styles.toolTitle}>{title}</h4>
          <p className={styles.toolDescription}>{description}</p>
        </div>
      </div>
    </div>
  );
};

export default ToolCard;