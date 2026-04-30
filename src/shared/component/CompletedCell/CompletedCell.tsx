import * as React from "react";
import styles from "./CompletedCell.module.scss";

interface CompletedCellProps {
  value: number; // percentage
}

const CompletedCell: React.FC<CompletedCellProps> = ({ value }) => {
  // Ensure value is a valid number
  const percentage = Math.max(0, Math.min(100, Number(value) || 0));
  
  // Determine color based on percentage
  const getColorClass = (percent: number): string => {
    if (percent >= 100) return styles.complete;
    if (percent >= 80) return styles.high;
    if (percent >= 60) return styles.medium;
    if (percent >= 40) return styles.low;
    return styles.critical;
  };

  const colorClass = getColorClass(percentage);

  return (
    <div className={styles.completedWrapper}>
      <span className={`${styles.percent} ${colorClass}`}>
        {percentage}%
      </span>

      <div className={styles.bar}>
        <div
          className={`${styles.fill} ${colorClass}`}
          style={{ width: `${percentage}%` }}
        />
      </div>
    </div>
  );
};

export default CompletedCell;
