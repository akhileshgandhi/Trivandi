/* eslint-disable react/self-closing-comp */
import * as React from "react";
import styles from "./GlobalLoader.module.scss";

interface GlobalLoaderProps {
  variant?: "overlay" | "content" | "bar";
}

const GlobalLoader: React.FC<GlobalLoaderProps> = ({
  variant = "overlay",
}) => {
  return (
    <div
      className={
        variant === "overlay"
          ? styles.overlay
          : variant === "bar"
            ? styles.barWrapper
            : styles.contentWrapper
      }
    >
      {variant === "bar" ? (
        <div className={styles.progressBar}>
          <div className={styles.progressFill}></div>
        </div>
      ) : (
        <div className={styles.saturationLoader}>
          <div className={styles.loaderRing}></div>
          <div className={styles.loaderRing}></div>
          <div className={styles.loaderRing}></div>
          <div className={styles.loaderCore}></div>
        </div>
      )}
    </div>
  );
};

export default GlobalLoader;
