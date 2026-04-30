import * as React from 'react';
import styles from './Dashboard.module.scss';

interface ComplianceCardProps {
  title: string;
  description: string;
  color: string;
  icon: string; // changed from React.ReactNode to string
  link?: string;
}

const ComplianceCard: React.FC<ComplianceCardProps> = ({
  title,
  description,
  color,
  icon,
  link,
}) => {
  const handleClick = (e: React.MouseEvent): void => {
    e.stopPropagation();
    console.log('Card clicked! Link:', link);
    if (link) {
      console.log('Opening link:', link);
      window.open(link, '_blank', 'noopener,noreferrer');
    } else {
      console.log('No link available for this card');
    }
  };

  React.useEffect(() => {
    console.log(`ComplianceCard ${title} - Link:`, link);
  }, [title, link]);

  return (
    <div
      className={styles.complianceCard}
      style={{ background: color, cursor: link ? 'pointer' : 'default' }}
      onClick={handleClick}
      role={link ? 'button' : undefined}
      tabIndex={link ? 0 : undefined}
      onKeyPress={(e): void => {
        if (link && (e.key === 'Enter' || e.key === ' ')) {
          e.preventDefault();
          handleClick(e as any);
        }
      }}
    >
      <div className={styles.iconWrapper} style={{ pointerEvents: 'none' }}>
       <img src={icon} alt={`${title} icon`} style={{ pointerEvents: 'none' }} />
      </div>

      <div className={styles.complianceContent} style={{ pointerEvents: 'none' }}>
        <h4 className={styles.complianceTitle}>{title}</h4>
        <p className={styles.complianceDescription}>{description}</p>
      </div>
    </div>
  );
};

export default ComplianceCard;
