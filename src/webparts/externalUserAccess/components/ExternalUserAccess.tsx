import * as React from "react";
import { useState, useEffect, useCallback } from "react";
import { spfi, SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { IExternalUserAccessProps } from "./IExternalUserAccessProps";
import styles from "./ExternalUserAccess.module.scss";

/* ─────────────────────────────────────────────
   Types
───────────────────────────────────────────── */
interface ISharedFile {
  // ── From SharedDocumentLog ──────────────────────────
  LogId: number;
  Permission: string;   // "Read" | "Edit" | "Full"
  SharedBy: string;
  ActionDate: string;
  ExpiryDate: string;
  SharingLink: string;
  GuestEmail: string;
  DocId: number;
  // ── From ExternalShareDocument library ─────────────
  FileId: number;
  FileName: string;
  FileRef: string;
  FileModified: string;
  FileModifiedBy: string;
  FileType: string;
}

interface IProjectInfo {
  Title: string;
  Company?: string;
}

/* ─────────────────────────────────────────────
   Permission helpers
───────────────────────────────────────────── */
const canDownload = (permission: string): boolean =>
  ["edit", "write", "full"].includes((permission || "").toLowerCase());

/* ─────────────────────────────────────────────
   URL / format helpers
───────────────────────────────────────────── */
const getProjectIdFromUrl = (): number | null => {
  try {
    const id = new URLSearchParams(window.location.search).get("projectId");
    return id ? parseInt(id, 10) : null;
  } catch { return null; }
};

const OFFICE_EXTS = ["xlsx", "xls", "xlsm", "docx", "doc", "pptx", "ppt"];
const PDF_EXTS = ["pdf"];

/**
 * Build preview URL strategy (in priority order):
 *
 * 1. SharingLink + &action=embedview  → works for file-level-only guests (no site membership needed)
 * 2. /_layouts/15/Doc.aspx embedview  → works for guests who ARE site members
 * 3. Direct file URL                  → PDF / fallback
 *
 * If the guest only has file access (not site access), option 1 is the ONLY working one.
 * SharingLink must be stored in SharedDocumentLog when the file is originally shared.
 */
const buildPreviewUrl = (sharingLink: string, fileRef: string): string => {
  const ext = (fileRef.split(".").pop() || "").toLowerCase();

  // Strategy 1 — sharing link embedview (works without site membership)
  if (OFFICE_EXTS.includes(ext) && sharingLink) {
    const base = sharingLink.split("?")[0];
    const params = new URLSearchParams(sharingLink.includes("?") ? sharingLink.split("?")[1] : "");
    params.set("action", "embedview");
    return `${base}?${params.toString()}`;
  }

  // Strategy 2 — Doc.aspx embedview (works if guest has site membership)
  if (OFFICE_EXTS.includes(ext)) {
    const encoded = encodeURIComponent(fileRef);
    return `${window.location.origin}/_layouts/15/Doc.aspx?sourcedoc=${encoded}&action=embedview`;
  }

  // Strategy 3 — direct URL for PDF / images
  return `${window.location.origin}${fileRef}`;
};

const formatDate = (iso: string): string => {
  if (!iso) return "";
  try {
    return new Date(iso).toLocaleDateString("en-GB", {
      day: "2-digit", month: "2-digit", year: "numeric",
      hour: "2-digit", minute: "2-digit",
    });
  } catch { return iso; }
};

const getFileExt = (name: string): string =>
  (name || "").split(".").pop()?.toLowerCase() || "";

const isExpired = (expiryDate: string): boolean => {
  if (!expiryDate) return false;
  return new Date(expiryDate) < new Date();
};

const getFolderLabel = (docPath: string): string => {
  if (!docPath) return "";
  const parts = docPath.split("/").filter(Boolean);
  return parts.length >= 2 ? parts[parts.length - 2] : parts[0] || "";
};

const fileEmoji = (ext: string): string => {
  if (["xlsx", "xls", "csv"].includes(ext)) return "📗";
  if (["docx", "doc"].includes(ext)) return "📘";
  if (["pptx", "ppt"].includes(ext)) return "📙";
  if (ext === "pdf") return "📕";
  return "📄";
};

/* ─────────────────────────────────────────────
   Icons
───────────────────────────────────────────── */
const FileIcon: React.FC<{ ext?: string }> = ({ ext }) => {
  const color =
    ext === "xlsx" || ext === "xls" ? "#217346" :
      ext === "docx" || ext === "doc" ? "#2B579A" :
        ext === "pptx" || ext === "ppt" ? "#D24726" :
          ext === "pdf" ? "#D32F2F" : "#94a3b8";
  return (
    <svg width="15" height="17" viewBox="0 0 16 18" fill="none">
      <path d="M9 1H3C2.47 1 1.96 1.21 1.59 1.59C1.21 1.96 1 2.47 1 3V15C1 15.53 1.21 16.04 1.59 16.41C1.96 16.79 2.47 17 3 17H13C13.53 17 14.04 16.79 14.41 16.41C14.79 16.04 15 15.53 15 15V7L9 1Z"
        stroke={color} strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
      <path d="M9 1V7H15" stroke={color} strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
};
const UsersIcon = () => (
  <svg width="22" height="22" viewBox="0 0 24 24" fill="none">
    <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" stroke="#0511F2" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
    <circle cx="9" cy="7" r="4" stroke="#0511F2" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
    <path d="M23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75" stroke="#0511F2" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);
const CheckCircleIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" style={{ flexShrink: 0 }}>
    <circle cx="8" cy="8" r="7" stroke="#16a34a" strokeWidth="1.5" />
    <path d="M5 8.5l2 2 4-4" stroke="#16a34a" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);
const ArrowLeftIcon = () => (
  <svg width="18" height="18" viewBox="0 0 18 18" fill="none">
    <path d="M15 9H3M3 9L8 4M3 9L8 14" stroke="#374151" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);
const ChevronRightIcon = () => (
  <svg width="13" height="13" viewBox="0 0 14 14" fill="none">
    <path d="M5 3l4 4-4 4" stroke="#94a3b8" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);
const DotsIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
    <circle cx="8" cy="3" r="1.2" fill="#64748b" />
    <circle cx="8" cy="8" r="1.2" fill="#64748b" />
    <circle cx="8" cy="13" r="1.2" fill="#64748b" />
  </svg>
);
const ChevronDownIcon = () => (
  <svg width="13" height="13" viewBox="0 0 14 14" fill="none">
    <path d="M3 5L7 9L11 5" stroke="#64748b" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);
const DownloadIcon: React.FC<{ color?: string }> = ({ color = "#0511F2" }) => (
  <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
    <path d="M7 1v8M4 6l3 3 3-3M2 11h10" stroke={color} strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);
const EyeIcon = () => (
  <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
    <path d="M1 7s2-4 6-4 6 4 6 4-2 4-6 4-6-4-6-4z" stroke="#64748b" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round" />
    <circle cx="7" cy="7" r="1.8" stroke="#64748b" strokeWidth="1.3" />
  </svg>
);
const LockSmallIcon = () => (
  <svg width="13" height="13" viewBox="0 0 14 14" fill="none">
    <rect x="2" y="6" width="10" height="7" rx="1.5" stroke="#92400e" strokeWidth="1.3" />
    <path d="M4.5 6V4.5a2.5 2.5 0 0 1 5 0V6" stroke="#92400e" strokeWidth="1.3" strokeLinecap="round" />
  </svg>
);
const LockRedIcon = () => (
  <svg width="22" height="22" viewBox="0 0 24 24" fill="none">
    <rect x="3" y="11" width="18" height="11" rx="2" stroke="#ef4444" strokeWidth="1.8" />
    <path d="M7 11V7a5 5 0 0 1 10 0v4" stroke="#ef4444" strokeWidth="1.8" strokeLinecap="round" />
  </svg>
);
const WarningIcon = () => (
  <svg width="20" height="20" viewBox="0 0 24 24" fill="none">
    <path d="M12 9v4M12 17h.01M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"
      stroke="#d97706" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);
const EmptyFolderIcon = () => (
  <svg width="48" height="48" viewBox="0 0 24 24" fill="none" style={{ margin: "0 auto 12px", display: "block" }}>
    <path d="M13 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9L13 2Z"
      stroke="#cbd5e1" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
    <path d="M13 2v7h7" stroke="#cbd5e1" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);
const CloseIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
    <path d="M12 4L4 12M4 4l8 8" stroke="white" strokeWidth="1.8" strokeLinecap="round" />
  </svg>
);
const ExternalLinkIcon = () => (
  <svg width="13" height="13" viewBox="0 0 14 14" fill="none">
    <path d="M6 2H2a1 1 0 0 0-1 1v9a1 1 0 0 0 1 1h9a1 1 0 0 0 1-1V8" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round" />
    <path d="M9 1h4m0 0v4m0-4L6 8" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);

/* ─────────────────────────────────────────────
   Permission Badge
───────────────────────────────────────────── */
const PermissionBadge: React.FC<{ permission: string }> = ({ permission }) => {
  const modMap: Record<string, string> = {
    read: styles["badge--read"],
    edit: styles["badge--edit"],
    write: styles["badge--write"],
    full: styles["badge--full"],
  };
  const mod = modMap[(permission || "").toLowerCase()] || styles["badge--default"];
  return <span className={`${styles.badge} ${mod}`}>{permission}</span>;
};

/* ─────────────────────────────────────────────
   Preview Modal
───────────────────────────────────────────── */
interface IPreviewModalProps {
  doc: ISharedFile;
  onClose: () => void;
  onDownload: (doc: ISharedFile) => void;
  siteUrl: string;  // context.pageContext.web.absoluteUrl
}

/**
 * PreviewModal — mirrors FilePreview.tsx exactly.
 * Guests are added to the SP Visitors group via addUserToSharePointGroup()
 * at invite time, so WopiFrame (which needs site membership) works for them.
 *
 * Read  → sandbox iframe: hides Download/Print buttons inside Office Online toolbar
 *       + iframeBlocker div covers the bottom bar as extra protection
 * Edit  → full iframe: no sandbox, guest can edit and save
 */
const PreviewModal: React.FC<IPreviewModalProps> = ({ doc, onClose, onDownload, siteUrl }) => {
  const [iframeLoaded, setIframeLoaded] = useState(false);

  const ext = getFileExt(doc.FileName);
  const readOnly = !canDownload(doc.Permission);
  const expired = isExpired(doc.ExpiryDate);

  const OFFICE_FILE_EXTS = ["doc", "docx", "xls", "xlsx", "ppt", "pptx", "pdf"];
  const isOfficePdf = OFFICE_FILE_EXTS.includes(ext);

  // Exact same URL builder as FilePreview._buildPreviewUrl
  const buildWopiUrl = (fileUrl: string): string => {
    const serverRelativeUrl = fileUrl.startsWith("http")
      ? new URL(fileUrl).pathname
      : fileUrl;
    return `${siteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(serverRelativeUrl)}&action=embedview`;
  };

  useEffect(() => {
    const handler = (e: KeyboardEvent) => { if (e.key === "Escape") onClose(); };
    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  }, [onClose]);

  useEffect(() => { setIframeLoaded(false); }, [doc.FileId]);

  // ── File icon colour ─────────────────────────────────────────────────────
  const iconColor =
    ext === "xlsx" || ext === "xls" ? "#217346" :
      ext === "docx" || ext === "doc" ? "#2B579A" :
        ext === "pptx" || ext === "ppt" ? "#D24726" :
          ext === "pdf" ? "#D32F2F" : "#64748b";

  const renderBody = (): React.ReactElement => {
    // ── Expired ─────────────────────────────────────────────────────────────
    if (expired) {
      return (
        <div className={styles.modalPreviewCard}>
          <div className={styles.modalPreviewCardIcon} style={{ background: "#fef2f2" }}>
            <LockRedIcon />
          </div>
          <h2 className={styles.modalPreviewCardTitle}>Access Revoked</h2>
          <p className={styles.modalPreviewCardDesc}>
            Your access to <strong>{doc.FileName}</strong> expired on{" "}
            <strong>{formatDate(doc.ExpiryDate)}</strong>.
            Contact <strong>{doc.SharedBy}</strong> to request renewed access.
          </p>
        </div>
      );
    }

    // ── Office / PDF — WopiFrame iframe (same as FilePreview.tsx) ──────────
    if (isOfficePdf) {
      const wopiUrl = buildWopiUrl(doc.FileRef);
      return (
        <div className={styles.modalIframeWrap}>
          {!iframeLoaded && (
            <div className={styles.modalIframeLoading}>
              <div className={styles.spinner} />
              <span>Loading preview…</span>
            </div>
          )}
          <iframe
            key={doc.FileRef}
            className={styles.modalIframe}
            src={wopiUrl}
            title={doc.FileName}
            onLoad={() => setIframeLoaded(true)}
            // Read-only: sandbox blocks popups (Download a Copy / Open in Desktop App)
            // allow-forms required: WopiFrame uses form POST for auth
            // allow-same-origin required: viewer JS needs to run
            {...(readOnly
              ? { sandbox: "allow-scripts allow-same-origin allow-forms" }
              : {}
            )}
          />
          {/* Extra blocker over the bottom toolbar where the … / Download button lives */}
          {readOnly && (
            <div
              className={styles.iframeBlocker}
              title="Download and print are disabled in read-only view"
            />
          )}
        </div>
      );
    }

    // ── Images ──────────────────────────────────────────────────────────────
    if (["jpg", "jpeg", "png", "gif", "bmp", "svg", "webp"].includes(ext)) {
      return (
        <div className={styles.modalImageWrap}>
          <img
            src={`${window.location.origin}${doc.FileRef}`}
            alt={doc.FileName}
            className={styles.modalImage}
          />
        </div>
      );
    }

    // ── Unsupported ──────────────────────────────────────────────────────────
    return (
      <div className={styles.modalPreviewCard}>
        <div className={styles.modalPreviewCardIcon} style={{ background: `${iconColor}18` }}>
          <FileIcon ext={ext} />
        </div>
        <p className={styles.modalPreviewCardTitle}>Preview not available</p>
        <p className={styles.modalPreviewCardDesc}>
          This file type cannot be previewed. Use the Download button to access the file.
        </p>
      </div>
    );
  };

  return (
    <div className={styles.modalOverlay} onClick={onClose}>
      <div className={styles.modalWindow} onClick={e => e.stopPropagation()}>

        {/* ── Title bar ── */}
        <div className={styles.modalTitleBar}>
          <span className={styles.modalFileEmoji}>{fileEmoji(ext)}</span>
          <div className={styles.modalFileInfo}>
            <div className={styles.modalFileName}>{doc.FileName}</div>
            <div className={styles.modalFilePath}>{doc.FileRef}</div>
          </div>
          <div className={styles.modalActions}>
            {/* Open in New Tab — Edit/Full only */}
            {!readOnly && !expired && (
              <button type="button"
                className={styles["modalNewTabBtn"]}
                title="Open in new tab"
                onClick={() => {
                  const url = `${siteUrl}/_layouts/15/Doc.aspx?sourcedoc=${encodeURIComponent(doc.FileRef)}&action=default`;
                  window.open(url, "_blank", "noopener,noreferrer");
                }}
              >
                <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
                  <path d="M6 1H2a1 1 0 0 0-1 1v9a1 1 0 0 0 1 1h9a1 1 0 0 0 1-1V7"
                    stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
                  <path d="M9 1h4m0 0v4m0-4L6 8"
                    stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
                </svg>
                New Tab
              </button>
            )}
            {/* Download — Edit/Full only */}
            {!readOnly && !expired && (
              <button type="button" className={styles.modalDownloadBtn} onClick={() => onDownload(doc)}>
                <DownloadIcon color="#fff" />
                Download
              </button>
            )}
            <button type="button" className={styles.modalCloseBtn} onClick={onClose} title="Close (Esc)">
              <CloseIcon />
            </button>
          </div>
        </div>

        {/* ── Read-only banner ── */}
        {readOnly && !expired && (
          <div className={styles.modalReadOnlyBanner}>
            <LockSmallIcon />
            <span>
              <strong>Read-only view</strong> — download and print are disabled for your access level.
            </span>
          </div>
        )}

        {/* ── Body ── */}
        {renderBody()}

      </div>
    </div>
  );
};

/* ─────────────────────────────────────────────
   Row Dropdown Menu
───────────────────────────────────────────── */
interface IRowMenuProps {
  doc: ISharedFile;
  onPreview: (doc: ISharedFile) => void;
  onDownload: (doc: ISharedFile) => void;
}

const RowMenu: React.FC<IRowMenuProps> = ({ doc, onPreview, onDownload }) => {
  const [open, setOpen] = useState(false);
  const downloadable = canDownload(doc.Permission);
  const expired = isExpired(doc.ExpiryDate);

  return (
    <div style={{ position: "relative", display: "inline-block" }}>
      <div className={styles.actionButtons}>
        {/* <button type="button" className={styles.actionBtn} onClick={() => setOpen(!open)}>
          <DotsIcon />
        </button> */}
        <button type="button" className={styles.actionBtn} onClick={() => setOpen(!open)}>
          <ChevronDownIcon />
        </button>
      </div>

      {open && (
        <>
          <div className={styles.dropdownBackdrop} onClick={() => setOpen(false)} />
          <div className={styles.dropdown}>

            {/* Preview — always available (modal handles all fallback logic) */}
            <button type="button"
              className={styles.dropdownItem}
              onClick={() => { onPreview(doc); setOpen(false); }}
            >
              <EyeIcon />
              Preview
            </button>

            <div className={styles.dropdownDivider} />

            {/* Download */}
            <button type="button"
              className={`${styles.dropdownItem} ${!downloadable ? styles["dropdownItem--disabled"] : ""}`}
              onClick={() => { if (downloadable) { onDownload(doc); setOpen(false); } }}
              disabled={!downloadable}
              title={!downloadable ? "Download restricted — Read access only" : "Download file"}
            >
              <DownloadIcon color={downloadable ? "#0511F2" : "#94a3b8"} />
              {downloadable ? "Download" : "Download (restricted)"}
            </button>

            {/* Expired */}
            {expired && (
              <>
                <div className={styles.dropdownDivider} />
                <button type="button" className={`${styles.dropdownItem} ${styles["dropdownItem--danger"]}`} disabled>
                  <LockSmallIcon />
                  Access Expired
                </button>
              </>
            )}

          </div>
        </>
      )}
    </div>
  );
};

/* ─────────────────────────────────────────────
   Main Component
───────────────────────────────────────────── */
const ExternalUserAccess: React.FC<IExternalUserAccessProps> = ({ context }) => {
  const [projectId, setProjectId] = useState<number | null>(null);
  const [projectInfo, setProjectInfo] = useState<IProjectInfo | null>(null);
  const [sharedFiles, setSharedFiles] = useState<ISharedFile[]>([]);
  const [currentUserEmail, setCurrentUserEmail] = useState<string>("");
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string>("");
  const [checkedAll, setCheckedAll] = useState<boolean>(false);
  const [checked, setChecked] = useState<Record<number, boolean>>({});
  const [previewDoc, setPreviewDoc] = useState<ISharedFile | null>(null);
  const [siteUrl] = useState<string>(
    () => (context as any)?.pageContext?.web?.absoluteUrl || window.location.origin
  );
  // Debug info — visible only in dev; helps diagnose email mismatch
  const [debugInfo, setDebugInfo] = useState<string>("");

  const sp = spfi().using(SPFx(context));

  //#region   
  /* ── Block right-click ONLY on this webpart's page ── */
  useEffect(() => {
    // Block right-click context menu
    const blockContextMenu = (e: MouseEvent): void => {
      e.preventDefault();
      e.stopPropagation();
    };

    // Block keyboard shortcuts: Ctrl+S, Ctrl+P, Ctrl+U, F12, Ctrl+Shift+I
    const blockShortcuts = (e: KeyboardEvent): void => {
      const key = e.key?.toLowerCase();
      const ctrl = e.ctrlKey || e.metaKey;
      if (
        (ctrl && key === 's') ||
        (ctrl && key === 'p') ||
        (ctrl && key === 'u') ||
        key === 'f12' ||
        (ctrl && e.shiftKey && key === 'i')
      ) {
        e.preventDefault();
        e.stopPropagation();
      }
    };

    document.addEventListener('contextmenu', blockContextMenu);
    document.addEventListener('keydown', blockShortcuts);

    // Cleanup when component unmounts (user navigates away)
    return () => {
      document.removeEventListener('contextmenu', blockContextMenu);
      document.removeEventListener('keydown', blockShortcuts);
    };
  }, []); // runs once on mount, cleans up on unmount
  //#endregion
  /* ── Load data ── */
  const loadData = useCallback(async () => {
    setLoading(true);
    setError("");
    setDebugInfo("");
    try {
      const pid = getProjectIdFromUrl();
      if (!pid) {
        setError("Kindly check your email. You will find the link to the file there.");
        setLoading(false);
        return;
      }
      setProjectId(pid);

      // ── 1. Current user ──────────────────────────────────────────────────
      // Prioritise .email — for B2B guests .loginName is the UPN hash
      // (e.g. spweb94_gmail.com#EXT#@tenant.onmicrosoft.com) which does NOT
      // match the plain email stored in SharedDocumentLog / ExternalGuestAccess.
      const user = await sp.web.currentUser.select("Email", "Title", "LoginName")();
      const rawEmail = (user.Email || user.LoginName || "")
        .toLowerCase()
        .replace(/^i:0#\.f\|membership\|/i, "")
        .trim();
      setCurrentUserEmail(rawEmail);

      // ── 2. Project info ──────────────────────────────────────────────────
      try {
        const projects = await sp.web.lists
          .getByTitle("ProjectsNew")
          .items
          .select("Title", "Company")
          .filter(`ID eq ${pid}`)
          .top(1)();
        if (projects.length > 0)
          setProjectInfo({ Title: projects[0].Title, Company: projects[0].Company });
      } catch { console.warn("Could not load project info:", pid); }

      // ── 3. SharedDocumentLog ─────────────────────────────────────────────
      //
      // WHY substringof instead of eq:
      //   GuestEmail stores MULTIPLE emails as semicolon-separated values,
      //   e.g. "alice@co.com; bob@co.com". Using `eq` only matches a single
      //   exact value. substringof() finds the email anywhere in the string.
      //
      // WHY client-side filter too:
      //   substringof is case-sensitive in some SP versions and also matches
      //   partial strings (e.g. "ali" matches "alice"). We do a precise
      //   case-insensitive check client-side after the server query.
      //
      // WHY fetch ALL logs then filter (fallback):
      //   If substringof fails (e.g. library throttle / SP version issue),
      //   we fall back to fetching all project logs and filtering client-side.
      let allLogs: any[] = [];
      const safeEmail = rawEmail.replace(/'/g, "''");

      try {
        // Primary: substringof filter (matches emails in semicolon-separated list)
        allLogs = await sp.web.lists
          .getByTitle("SharedDocumentLog")
          .items
          .select(
            "Id", "DocId", "Document", "DocumentName",
            "Permission", "SharedBy", "SharingLink",
            "ActionDate", "ExpiryDate", "GuestEmail", "ProjectId"
          )
          .filter(
            `ProjectId eq ${pid} and substringof('${safeEmail}', GuestEmail)`
          )
          .orderBy("ActionDate", false)
          .top(500)();
      } catch (logErr) {
        console.warn("substringof filter failed, falling back to full project fetch:", logErr);
        // Fallback: get all logs for project and filter client-side
        allLogs = await sp.web.lists
          .getByTitle("SharedDocumentLog")
          .items
          .select(
            "Id", "DocId", "Document", "DocumentName",
            "Permission", "SharedBy", "SharingLink",
            "ActionDate", "ExpiryDate", "GuestEmail", "ProjectId"
          )
          .filter(`ProjectId eq ${pid}`)
          .orderBy("ActionDate", false)
          .top(1000)();
      }

      // Client-side precise email check (handles semicolon-separated list + case)
      const userLogs = allLogs.filter((log: any) => {
        const emails = (log.GuestEmail || "")
          .split(";")
          .map((e: string) => e.trim().toLowerCase())
          .filter(Boolean);
        return emails.includes(rawEmail);
      });

      // setDebugInfo(
      //   `Email: ${rawEmail}`
      // );

      if (userLogs.length === 0) {
        setSharedFiles([]);
        return;
      }

      // ── 4. Build access map from logs ────────────────────────────────────
      //
      // Key strategy (mirrors filterDocumentsForRestrictedUser in service):
      //   a) DocId  → log  (join key when DocId was stored correctly)
      //   b) Document (full server-relative path) → log  (fallback)
      //   c) filename only → log  (last-resort fallback)
      //
      // For each key we keep only the MOST RECENT (DESC ordered) log entry.
      const now = new Date();

      const docIdMap = new Map<number, any>();   // DocId → log
      const pathMap = new Map<string, any>();   // lower server-relative path → log
      const fileNameMap = new Map<string, any>();   // lower filename → log

      userLogs.forEach((log: any) => {
        // a) DocId map
        if (log.DocId && !docIdMap.has(log.DocId)) {
          docIdMap.set(log.DocId, log);
        }
        // b) Full path map
        if (log.Document) {
          const pathKey = (log.Document as string).toLowerCase().trim();
          if (!pathMap.has(pathKey)) pathMap.set(pathKey, log);
        }
        // c) Filename map
        if (log.DocumentName) {
          const nameKey = (log.DocumentName as string).toLowerCase().trim();
          if (!fileNameMap.has(nameKey)) fileNameMap.set(nameKey, log);
        }
      });

      // Helper: is a log entry still valid (not expired)?
      const isLogValid = (log: any): boolean => {
        if (log.ExpiryDate) return new Date(log.ExpiryDate) > now;
        if (log.AccessDuration > 0 && log.ActionDate) {
          const exp = new Date(
            new Date(log.ActionDate).getTime() + log.AccessDuration * 24 * 60 * 60 * 1000
          );
          return exp > now;
        }
        return true; // permanent (duration 0 / -1 / no date)
      };

      // Find matching log for a file (tries all three keys)
      const findLog = (fileId: number, fileRef: string, fileName: string): any | null => {
        // 1. DocId exact match
        const byDocId = docIdMap.get(fileId);
        if (byDocId) return byDocId;
        // 2. Full server-relative path
        const byPath = pathMap.get(fileRef.toLowerCase().trim());
        if (byPath) return byPath;
        // 3. Filename
        const byName = fileNameMap.get(fileName.toLowerCase().trim());
        if (byName) return byName;
        return null;
      };

      // ── 5. ExternalShareDocument library ────────────────────────────────
      // Fetch project folder using rootFolder URL to get correct server-relative path
      let libraryItems: any[] = [];
      try {
        const rootFolder = await sp.web.lists
          .getByTitle("ExternalShareDocument")
          .rootFolder.select("ServerRelativeUrl")();
        const projectFolder = `${rootFolder.ServerRelativeUrl}/Project-${pid}`;

        libraryItems = await sp.web.lists
          .getByTitle("ExternalShareDocument")
          .items
          .select(
            "Id", "FileLeafRef", "FileRef", "FileDirRef",
            "Modified", "Editor/Title", "File_x0020_Type", "FSObjType"
          )
          .expand("Editor")
          // Files only (FSObjType=0), inside the project folder
          .filter(
            `startswith(FileDirRef,'${projectFolder}') and FSObjType eq 0`
          )
          .orderBy("Modified", false)
          .top(500)();
      } catch (libErr: any) {
        console.error("Failed to fetch from ExternalShareDocument:", libErr);
        setError(`Could not load library: ${libErr.message || libErr}`);
        return;
      }

      // ── 6. Join + access filter ──────────────────────────────────────────
      const merged: ISharedFile[] = [];

      libraryItems.forEach((file: any) => {
        const log = findLog(file.Id, file.FileRef, file.FileLeafRef);
        if (!log) return;            // not shared with this user
        if (!isLogValid(log)) return; // access expired

        merged.push({
          LogId: log.Id,
          Permission: log.Permission || "Read",
          SharedBy: log.SharedBy || "",
          ActionDate: log.ActionDate || "",
          ExpiryDate: log.ExpiryDate || "",
          SharingLink: log.SharingLink || "",
          GuestEmail: log.GuestEmail || "",
          DocId: log.DocId || 0,
          FileId: file.Id,
          FileName: file.FileLeafRef,
          FileRef: file.FileRef,
          FileModified: file.Modified,
          FileModifiedBy: file.Editor?.Title || "",
          FileType: (
            file.File_x0020_Type ||
            file.FileLeafRef?.split(".").pop() ||
            ""
          ).toLowerCase(),
        });
      });

      setSharedFiles(merged);
    } catch (err: any) {
      console.error("ExternalUserAccess error:", err);
      setError(`Failed to load: ${err.message || err}`);
    } finally {
      setLoading(false);
    }
  }, [context]);

  useEffect(() => { loadData().catch(console.error); }, [loadData]);

  /* ── Checkboxes ── */
  const toggleAll = (): void => {
    const next = !checkedAll;
    setCheckedAll(next);
    const all: Record<number, boolean> = {};
    sharedFiles.forEach(f => (all[f.FileId] = next));
    setChecked(all);
  };
  const toggleOne = (id: number): void =>
    setChecked(prev => ({ ...prev, [id]: !prev[id] }));

  /* ── Download ── */
  const handleDownload = (doc: ISharedFile): void => {
    if (!canDownload(doc.Permission)) return;
    const link = document.createElement("a");
    link.href = `${window.location.origin}${doc.FileRef}`;
    link.download = doc.FileName;
    link.target = "_blank";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  /* ─────────────── RENDER ─────────────── */
  return (
    <>
      <div className={styles.page}>

        {/* Header */}
        <div className={styles.header}>
          {/* <button type="button" className={styles.header__back} onClick={() => window.history.back()}>
            <ArrowLeftIcon />
          </button> */}
          <div>
            <h1 className={styles.header__title}>
              {projectInfo?.Title || (projectId ? `Project #${projectId}` : "Loading…")}
            </h1>
            <p className={styles.header__subtitle}>
              {projectInfo?.Company || "—"} • Project
            </p>
          </div>
        </div>

        {/* Tabs */}
        <div className={styles.tabBar}>
          <button type="button" className={styles.tab}>External Portal</button>
        </div>

        {/* Content */}
        <div className={styles.content}>

          {/* Portal info card */}
          <div className={styles.card}>
            <div className={styles.portalHeader}>
              <div className={styles.portalIconWrap}><UsersIcon /></div>
              <div>
                <h2 className={styles.portalTitle}>External Client Portal</h2>
                <p className={styles.portalDesc}>
                  Manage access and share documents securely with external partners.
                  All files placed here are visible to invited guests.
                </p>
                {currentUserEmail && (
                  <p className={styles.portalUserEmail}>
                    Viewing as: <strong>{currentUserEmail}</strong>
                  </p>
                )}
              </div>
            </div>
          </div>

          {/* Shared Files card */}
          <div className={styles.card}>
            <h3 className={styles.sectionTitle}>Shared Files</h3>

            <div className={styles.breadcrumb}>
              <span className={styles["breadcrumb__link"]}>Project-{projectId}</span>
              <ChevronRightIcon />
              <span className={styles["breadcrumb__current"]}>Shared Documents</span>
            </div>

            {/* <div className={styles.accessBanner}>
              <CheckCircleIcon />
              <span>
                <strong>Guest access:</strong> You can view documents shared with you.
                Files are only visible while your access is active.
              </span>
            </div> */}

            {/* Debug strip — shows email matching info; remove in production */}
            {debugInfo && (
              <div className={styles.debugBanner}>
                🔍 {debugInfo}
              </div>
            )}

            {loading && (
              <div className={styles.stateWrapper}>
                <div className={styles.spinner} />
                <p className={styles.stateMessage}>Loading shared documents…</p>
              </div>
            )}

            {!loading && error && (
              <div className={styles.errorBanner}>
                <WarningIcon /><span>{error}</span>
              </div>
            )}

            {!loading && !error && sharedFiles.length === 0 && (
              <div className={styles.emptyState}>
                <EmptyFolderIcon />
                <p>No documents have been shared with you for this project.</p>
                {debugInfo && (
                  <p style={{ fontSize: 11, color: "#94a3b8", marginTop: 8 }}>
                    {debugInfo}
                  </p>
                )}
              </div>
            )}

            {!loading && !error && sharedFiles.length > 0 && (
              <div className={styles.tableWrap}>
                <table className={styles.table}>
                  <thead className={styles.tableHead}>
                    <tr>
                      <th className={`${styles.th} ${styles["th--center"]}`} style={{ width: 40 }}>
                        <input type="checkbox" className={styles.checkbox}
                          checked={checkedAll} onChange={toggleAll} />
                      </th>
                      <th className={`${styles.th} ${styles["th--left"]}`}>NAME</th>
                      <th className={`${styles.th} ${styles["th--left"]}`}>PERMISSION</th>
                      <th className={`${styles.th} ${styles["th--left"]} ${styles["col--hide-sm"]}`}>SHARED BY</th>
                      <th className={`${styles.th} ${styles["th--left"]} ${styles["col--hide-sm"]}`}>MODIFIED</th>
                      <th className={`${styles.th} ${styles["th--left"]} ${styles["col--hide-xs"]}`}>EXPIRY DATE</th>
                      <th className={`${styles.th} ${styles["th--right"]}`}>ACTIONS</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sharedFiles.map((doc, i) => {
                      const ext = getFileExt(doc.FileName);
                      const expired = isExpired(doc.ExpiryDate);
                      const rowClass = [
                        styles.tr,
                        checked[doc.FileId]
                          ? styles["tr--checked"]
                          : i % 2 !== 0 ? styles["tr--alt"] : "",
                        expired ? styles["tr--expired"] : "",
                      ].filter(Boolean).join(" ");

                      return (
                        <tr key={doc.FileId} className={rowClass}>

                          <td className={`${styles.td} ${styles["td--center"]}`}>
                            <input type="checkbox" className={styles.checkbox}
                              checked={!!checked[doc.FileId]}
                              onChange={() => toggleOne(doc.FileId)} />
                          </td>

                          <td className={styles.td}>
                            <div className={styles.fileCell}>
                              <FileIcon ext={ext} />
                              <div>
                                <div className={styles.fileName}>
                                  {doc.FileName}
                                  {expired && <span className={styles.expiredBadge}>EXPIRED</span>}
                                </div>
                                {doc.FileRef && (
                                  <div className={styles.fileFolder}>{getFolderLabel(doc.FileRef)}</div>
                                )}
                              </div>
                            </div>
                          </td>

                          <td className={styles.td}>
                            <PermissionBadge permission={doc.Permission} />
                          </td>

                          <td className={`${styles.td} ${styles.tdText} ${styles["col--hide-sm"]}`}>{doc.SharedBy}</td>
                          <td className={`${styles.td} ${styles.tdMuted} ${styles["col--hide-sm"]}`}>{formatDate(doc.FileModified)}</td>
                          <td className={`${styles.td} ${expired ? styles.tdExpired : styles.tdMuted} ${styles["col--hide-xs"]}`}>
                            {formatDate(doc.ExpiryDate) || "—"}
                          </td>

                          <td className={`${styles.td} ${styles["td--right"]}`}>
                            <RowMenu
                              doc={doc}
                              onPreview={setPreviewDoc}
                              onDownload={handleDownload}
                            />
                          </td>

                        </tr>
                        
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}

            {!loading && sharedFiles.length > 0 && (
              <div className={styles.footerCount}>
                {sharedFiles.length} file{sharedFiles.length !== 1 ? "s" : ""} shared with you
              </div>
            )}
          </div>
        </div>
      </div>

      {previewDoc && (
        <PreviewModal
          doc={previewDoc}
          onClose={() => setPreviewDoc(null)}
          onDownload={handleDownload}
          siteUrl={siteUrl}
        />
      )}
    </>
  );
};

export default ExternalUserAccess;