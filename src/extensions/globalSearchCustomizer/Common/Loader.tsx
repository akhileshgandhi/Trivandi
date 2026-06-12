import * as React from "react";
import styles from "./Loader.module.scss";
 
interface GlobalLoaderProps {
  variant?: "overlay" | "content";
}
 
export const Loader: React.FC<GlobalLoaderProps> = ({
  variant = "overlay",
}) => {
  return (
    <div
      className={
        variant === "overlay"
          ? styles.overlay
          : styles.contentWrapper
      }
    >
      <div className={styles.saturationLoader}>
        <div className={styles.loaderRing}></div>
        <div className={styles.loaderRing}></div>
        <div className={styles.loaderRing}></div>
        <div className={styles.loaderCore}></div>
      </div>
    </div>
  );
};

export default Loader;
