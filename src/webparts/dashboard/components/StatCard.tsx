import * as React from 'react';
import styles from './Dashboard.module.scss';

interface StatCardProps {
  title: string;
  value: number;
  icon: string;
  color: string;
}

const StatCard: React.FC<StatCardProps> = ({ title, value, icon, color }) => {
  return (
    <div className={styles.statCard}>
      <div className={styles.statCardGradient} style={{ background: `#0511F2` }}></div>
      <div className={styles.statIcon} style={{ backgroundColor: `${color}20`, color: color }}>
      <img src={icon} alt={`${title} icon`} />
      </div>
      <div className={styles.statContent}>
        <h3 className={styles.statTitle} style={{ color: color }}>{title}</h3>
        <p className={styles.statValue}>{value}</p>
      </div>
    </div>
  );
};

export default StatCard;
