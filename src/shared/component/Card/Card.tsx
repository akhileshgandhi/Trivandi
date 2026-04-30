import React from "react";
import styles from "./Card.module.scss";

import userIcon from "../../assets/profile-2user.png";
import clockIcon from "../../assets/loc.png";
import folderIcon from "../../assets/folder.png";
import bookIcon from "../../assets/book.png";
import airIcon from "../../assets/airg.png";

interface CardProps {
  key?: React.Key;
  title: string;
  subtitle?: string;
  meta?: string;
  color?: string;
  border?: string;
  isDocument?: boolean;
  iconRight?: boolean;
  hideIcon?: boolean;
  documentType?: "folder" | "book";
  variant?: "template" | "invoice";
  customIcon?: string;
  noIconBg?: boolean;
  onClick?: () => void;
}

const Card: React.FC<CardProps> = ({
  title,
  subtitle,
  meta,
  color,
  border,
  isDocument,
  iconRight,
  hideIcon,
  documentType = "folder",
  variant,
  customIcon,
  noIconBg,
  onClick,
}) => {
  const iconSrc = customIcon
    ? customIcon
    : isDocument
    ? documentType === "book"
      ? bookIcon
      : folderIcon
    : userIcon;

  const leftIconClass = [
    noIconBg || isDocument ? styles.iconNoBg : styles.icon,
    isDocument && documentType === "book" && variant === "template"
      ? styles.fileIcon
      : "",
    isDocument && documentType === "book" && variant === "invoice"
      ? styles.invoiceIcon
      : "",
  ]
    .filter(Boolean)
    .join(" ");

  return (
    <div
      className={`${styles.card} ${isDocument ? styles.cardDocument : ""} ${color ? styles.coloredCard : ""}`}
      style={{ ...(color ? { backgroundColor: color } : {}), ...(border ? { border } : {}), cursor: onClick ? "pointer" : "default" }}
      onClick={onClick}
      role={onClick ? "button" : undefined}
      tabIndex={onClick ? 0 : undefined}
      onKeyDown={onClick ? (e) => { if (e.key === 'Enter' || e.key === ' ') onClick(); } : undefined}
    >
      {/* ================= LEFT ICON ================= */}
      {!hideIcon && !iconRight && (
        <div className={leftIconClass}>
          <img src={iconSrc} alt="" />
        </div>
      )}

      {/* ================= CONTENT ================= */}
      <div className={styles.cardContent}>
        <h5>{title}</h5>

        {subtitle && <p>{subtitle}</p>}

        {meta && (
          <span>
            <img src={clockIcon} className={styles.timeIcon} alt="" />
            {meta}
          </span>
        )}
      </div>

      {/* ================= RIGHT ICON ================= */}
      {!hideIcon && iconRight && (
        <div
          className={`${styles.iconRight} ${isDocument ? styles.iconNoBg : ""}`}
        >
          <img src={iconSrc} alt="" />
        </div>
      )}
    </div>
  );
};

export default Card;
