import React from "react";
import styles from "./Section.module.scss";

interface SectionProps {
  title?: string;
  viewAll?: boolean;
  children: React.ReactNode;
}

const Section: React.FC<SectionProps> = ({ title, viewAll, children }) => {
  return (
    <div className={styles.bgW}>
      {title && (
        <div className={styles.sectionHeader} style={{margin:'0 0 10px 0'}}>
          <h4>{title}</h4>
          {viewAll && <span className={styles.viewAll}>View all</span>}
        </div>
      )}

      <div>{children}</div>
    </div>
  );
};

export default Section;
