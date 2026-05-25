// import * as React from "react";
// import styles from "./TemplatesResources.module.scss";

// import Hero from "../../../shared/component/Hero/Hero";
// import type { ITemplatesResourcesProps } from "./ITemplatesResourcesProps";
// import heroBg from "../../../shared/assets/hero_new.jpg";
// import Section from "../../../shared/component/Section/Section";
// import Card from "../../../shared/component/Card/Card";
// import arrow from "../../../shared/assets/arrow-down.png";
// import GlobalLoader from "../../../shared/component/GlobalLoader";
// import "../../../shared/globalcss/globalcss.scss"
// import "../../../shared/globalcss/globalcss.scss";
// import {
//   getTemplatesResourcesData,
//   getCompletefolderStructure,
//   getTemplatesResourcesDataFromCurrentSite,
//   getFileDownloadUrl,
//   getFilePreviewUrl,
//   getExternalTemplatesFolderUrl,
//   getTemplatesAndDocumentsListData,
//   getTemplatesChildItemsByParent,
//   getRootFoldersWithCounts
// } from "../../../shared/services/templatesResourcesService";
// import {
//   IDocumentLibraryResponse,
//   ITemplateDocument,
//   IResourceFolder,
//   IFolderStructure
// } from "../../../shared/interfaces/ITemplatesResourcesInterfaces";
// /* ===================== TABS ===================== */

// const tabs = ["Project Templates"];

// /* ===================== WORKFLOWS ===================== */

// const workflows = [
//   {
//     id: "kickoff",
//     title: "Standard Project Kickoff",
//     steps: [
//       {
//         step: 1,
//         title: "Identify Stakeholders",
//         type: "Planning",
//         meta: "Due +1 days from start",
//       },
//       {
//         step: 2,
//         title: "Schedule Kickoff Meeting",
//         type: "Meeting",
//         meta: "Due +2 days from start",
//       },
//       {
//         step: 3,
//         title: "Distribute Project Charter",
//         type: "Email",
//         meta: "Due +3 days from start",
//       },
//       {
//         step: 4,
//         title: "Set up Project Folder",
//         type: "Admin",
//         meta: "Starts on Day 0",
//       },
//     ],
//   },
//   {
//     id: "bid",
//     title: "Bid Submission Workflow",
//     steps: [],
//   },
//   {
//     id: "design",
//     title: "Design Review Cycle",
//     steps: [],
//   },
//   {
//     id: "closeout",
//     title: "Closeout Process",
//     steps: [],
//   },
// ];

// /* ===================== TEMPLATES & DOCUMENTS ===================== */

// const templatesAndDocuments = [
//   {
//     id: "trivandi-cvs",
//     title: "Trivandi CVs",
//     subtitle: "Official formatted CVs for all staff members",
//     link: null, // Will use external templates folder URL
//   },
//   {
//     id: "nda-templates",
//     title: "NDA Templates",
//     subtitle: "Non-disclosure agreement templates for Associates and Companies",
//     link: null, // Will use external templates folder URL
//   },
//   {
//     id: "contract-templates",
//     title: "Contract Templates",
//     subtitle: "Contract templates for Associates and Companies",
//     link: null, // Will use external templates folder URL
//   },
//   {
//     id: "case-studies",
//     title: "Case Studies Template",
//     subtitle: "Project case studies and past performance references",
//     link: "https://trivandildn.sharepoint.com/:f:/r/sites/TrivandiLondon/Shared%20Documents/8%20Marketing/03%20Templates/05%20Case%20Study?csf=1&web=1&e=5Bd1Mg",
//   },
//   {
//     id: "policies",
//     title: "Policies",
//     subtitle: "Internal policies relevant for bid submissions",
//     link: null, // Will use external templates folder URL
//   },
// ];

// /* ===================== RESOURCE LIBRARY ITEMS ===================== */

// const resourceLibraryItems = [
//   {
//     id: "rfp-guidance",
//     title: "RFP Guidance",
//     subtitle: "Best practice guides for responding to tenders",
//     link: null, // Will use external templates folder URL
//   },
//   {
//     id: "brand-assets",
//     title: "Brand Assets",
//     subtitle: "Logos, fonts, and presentation decks",
//     link: "https://trivandildn.sharepoint.com/:f:/r/sites/TrivandiLondon/Shared%20Documents/8%20Marketing/02%20Brand%20Library?csf=1&web=1&e=XZkX9U",
//   },
//   {
//     id: "sector-research",
//     title: "Sector Research",
//     subtitle: "Market analysis and sector insights for strategic planning",
//     link: null, // Will use external templates folder URL
//   },
// ];

// /* ===================== ALTERNATING COLORS ===================== */

// const alternatingColors = [
//   "#0b0f6b", // Dark blue
//   "#7c3aed", // Purple
//   "#059669", // Green
//   "#dc2626", // Red
//   "#f59e0b", // Orange
//   "#8fa3e8", // Light blue
//   "#ec4899", // Pink
//   "#10b981", // Emerald
// ];

// /* ===================== COMPONENT ===================== */

// const TemplatesResources: React.FC<ITemplatesResourcesProps> = (props) => {
//   const [activeTab, setActiveTab] = React.useState("Project Templates");
//   const [selectedWorkflow, setSelectedWorkflow] = React.useState<
//     (typeof workflows)[0] | null
//   >(null);
//   const [isLoading, setIsLoading] = React.useState(true);
//   const [documentsData, setDocumentsData] = React.useState<IDocumentLibraryResponse | null>(null);
//   const [templatesListData, setTemplatesListData] = React.useState<Array<{
//     id: string;
//     title: string;
//     description: string;
//     link: string | null;
//     hasChildren: boolean;
//   }> | null>(null);
//   const [openDropdownId, setOpenDropdownId] = React.useState<string | null>(null);
//   const [childItemsCache, setChildItemsCache] = React.useState<Record<string, Array<{ id: string; title: string; link: string | null }>>>({});
//   const [loadingChildId, setLoadingChildId] = React.useState<string | null>(null);
//   const [dataError, setDataError] = React.useState<string | null>(null);
//   const [expandedFolders, setExpandedFolders] = React.useState<Set<number>>(new Set());
//   const [templatesFolders, setTemplatesFolders] = React.useState<Array<{
//     id: number;
//     name: string;
//     itemCount: number;
//     modified: string;
//     serverRelativeUrl: string;
//     color: string;
//   }> | null>(null);
//   const [searchQuery, setSearchQuery] = React.useState<string>("");
//   const [viewMode, setViewMode] = React.useState<'grid' | 'list'>('grid');

//   // Toggle folder expansion
//   const toggleFolder = (folderId: number) => {
//     setExpandedFolders(prev => {
//       const newSet = new Set(prev);
//       if (newSet.has(folderId)) {
//         newSet.delete(folderId);
//       } else {
//         newSet.add(folderId);
//       }
//       return newSet;
//     });
//   };

//   // Render folder structure recursively
//   const renderFolderStructure = (folders: IFolderStructure[], level: number = 0) => {
//     return folders.map(folder => (
//       <div key={folder.Id} className={styles.folderItem} style={{ paddingLeft: `${level * 20}px` }}>
//         <div 
//           className={styles.folderHeader}
//           onClick={() => toggleFolder(folder.Id)}
//         >
//           <img 
//             src={arrow} 
//             alt=""
//             className={`${styles.folderArrow} ${expandedFolders.has(folder.Id) ? styles.expanded : ''}`}
//           />
//           <div 
//             className={styles.folderIcon}
//             style={{ backgroundColor: folder.Color || '#0b0f6b' }}
//           >
//             📁
//           </div>
//           <div className={styles.folderDetails}>
//             <h5>{folder.Title || folder.Name}</h5>
//             <p>{folder.Description || `${folder.files.length} files, ${folder.subFolders.length} folders`}</p>
//           </div>
//         </div>

//         {expandedFolders.has(folder.Id) && (
//           <div className={styles.folderContent}>
//             {/* Render files in this folder */}
//             {folder.files.map(file => (
//               <div 
//                 key={file.Id} 
//                 className={styles.fileItem}
//                 style={{ paddingLeft: `${(level + 1) * 20}px` }}
//                 onClick={() => window.open(getFilePreviewUrl(file.FileRef, file.File_x0020_Type), '_self')}
//               >
//                 <div className={styles.fileIcon}>📄</div>
//                 <div className={styles.fileDetails}>
//                   <span className={styles.fileName}>{file.Title || file.Name}</span>
//                   <span className={styles.fileType}>{file.File_x0020_Type?.toUpperCase()}</span>
//                 </div>
//               </div>
//             ))}

//             {/* Render subfolders recursively */}
//             {folder.subFolders.length > 0 && renderFolderStructure(folder.subFolders, level + 1)}
//           </div>
//         )}
//       </div>
//     ));
//   };

//   const handleDropdownToggle = (templateId: string, templateTitle: string, e: React.MouseEvent): void => {
//     e.stopPropagation();
//     if (openDropdownId === templateId) {
//       setOpenDropdownId(null);
//       return;
//     }
//     setOpenDropdownId(templateId);
//     // Load children if not cached
//     if (!childItemsCache[templateTitle]) {
//       setLoadingChildId(templateId);
//       getTemplatesChildItemsByParent(templateTitle)
//         .then(children => {
//           setChildItemsCache(prev => ({ ...prev, [templateTitle]: children }));
//           setLoadingChildId(null);
//         })
//         .catch(() => {
//           setChildItemsCache(prev => ({ ...prev, [templateTitle]: [] }));
//           setLoadingChildId(null);
//         });
//     }
//   };

//   const onTabChange = (tab: string) => {
//     setIsLoading(true);
//     setActiveTab(tab);
//     setSelectedWorkflow(null);
//     setTimeout(() => {
//       setIsLoading(false);
//     }, 600);
//   };

//   React.useEffect(() => {
//     const fetchData = async () => {
//       try {
//         setIsLoading(true);
//         setDataError(null);

//         

//         // Fetch folders from Templates library (try multiple library names)
//         try {
//           let folders = [];
//           const libraryNames = ["TemplatesandDocuments", "Templates", "TemplatesandResources", "Templates and Resources", "Documents"];

//           for (const libName of libraryNames) {
//             try {
//               
//               folders = await getRootFoldersWithCounts(libName);
//               if (folders && folders.length > 0) {
//                 
//                 setTemplatesFolders(folders);
//                 break;
//               }
//             } catch (err) {
//               
//             }
import * as React from "react";
import styles from "./TemplatesResources.module.scss";

import type { ITemplatesResourcesProps } from "./ITemplatesResourcesProps";
import GlobalLoader from "../../../shared/component/GlobalLoader";
import "../../../shared/globalcss/globalcss.scss";
import {
  getTemplatesAndDocumentsListData,
  getTemplatesChildItemsByParent,
  getRootFoldersWithCounts,
} from "../../../shared/services/templatesResourcesService";
import {
  Folder,
  Search,
  FileText,
  File as FileIcon,
  LayoutGrid,
  List,
  MoreVertical
} from 'lucide-react';

/* ===================== COMPONENT ===================== */

const TemplatesResources: React.FC<ITemplatesResourcesProps> = (props) => {
  const [isLoading, setIsLoading] = React.useState(true);
  const [templatesListData, setTemplatesListData] = React.useState<Array<{
    id: string;
    title: string;
    description: string;
    link: string | null;
    hasChildren: boolean;
  }> | null>(null);
  const [openDropdownId, setOpenDropdownId] = React.useState<string | null>(null);
  const [childItemsCache, setChildItemsCache] = React.useState<Record<string, Array<{ id: string; title: string; link: string | null }>>>({});
  const [loadingChildId, setLoadingChildId] = React.useState<string | null>(null);
  const [dataError, setDataError] = React.useState<string | null>(null);
  const [templatesFolders, setTemplatesFolders] = React.useState<Array<{
    id: number;
    name: string;
    itemCount: number;
    modified: string;
    serverRelativeUrl: string;
    color: string;
  }> | null>(null);
  const [searchQuery, setSearchQuery] = React.useState<string>("");
  const [viewMode, setViewMode] = React.useState<'grid' | 'list'>('grid');

  const handleDropdownToggle = (templateId: string, templateTitle: string, e: React.MouseEvent): void => {
    e.stopPropagation();
    if (openDropdownId === templateId) {
      setOpenDropdownId(null);
      return;
    }
    setOpenDropdownId(templateId);
    // Load children if not cached
    if (!childItemsCache[templateTitle]) {
      setLoadingChildId(templateId);
      getTemplatesChildItemsByParent(templateTitle)
        .then(children => {
          setChildItemsCache(prev => ({ ...prev, [templateTitle]: children }));
          setLoadingChildId(null);
        })
        .catch(() => {
          setChildItemsCache(prev => ({ ...prev, [templateTitle]: [] }));
          setLoadingChildId(null);
        });
    }
  };

  React.useEffect(() => {
    const fetchData = async (): Promise<void> => {
      try {
        setIsLoading(true);
        setDataError(null);

        // Fetch folders from Templates library
        try {
          const libraryNames = ["TemplatesandDocuments"];
          let folders = [];

          for (const libName of libraryNames) {
            try {
              folders = await getRootFoldersWithCounts(libName);
              if (folders && folders.length > 0) {
                setTemplatesFolders(folders);
                break;
              }
            } catch (_err) {
              // Ignore error for individual library fetch
            }
          }
        } catch (folderError) {
          // Ignore folder fetch error
        }

        // Fetch templates from the TemplatesandDocuments list
        try {
          const templatesListData = await getTemplatesAndDocumentsListData();
          setTemplatesListData(templatesListData);
        } catch (_listError) {
          // Ignore list fetch error
        }
      } catch (_error) {
        setDataError('Failed to load templates and resources data.');
      } finally {
        setIsLoading(false);
      }
    };

    fetchData().catch((err: unknown) => console.log(err));
  }, []);

  const formatDate = (dateStr: string): string => {
    if (!dateStr) return "--";
    const date = new Date(dateStr);
    return date.toLocaleDateString('en-US', {
      month: 'short',
      day: 'numeric',
      year: 'numeric'
    });
  };

  const filteredFolders = (templatesFolders || []).filter(folder =>
    folder.name.toLowerCase().includes(searchQuery.toLowerCase())
  );

  const filteredListItems = (templatesListData || []).filter(template =>
    template.title.toLowerCase().includes(searchQuery.toLowerCase())
  );

  const totalItems = filteredFolders.length + filteredListItems.length;

  return (
    <div className={styles.page}>
      {/* Page Title Section */}
      <div className={styles.pageTitleSection}>
        <div className={styles.libraryIcon}>
          <FileIcon size={22} />
        </div>
        <div className={styles.titleInfo}>
          <h1>Templates & Resources</h1>
          <span className={styles.itemCount}>{isLoading ? 'Loading...' : `${totalItems} items`}</span>
        </div>
      </div>

      {/* Header Bar */}
      <div className={styles.headerBar}>
        <div className={styles.leftSection}>
          <div className={styles.breadcrumbWrapper}>
            <span className={styles.breadcrumbItem}>Root</span>
          </div>
        </div>

        <div className={styles.rightSection}>
          <div className={styles.searchBox}>
            <div className={styles.searchIcon}><Search size={18} /></div>
            <input
              type="text"
              placeholder="Search..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
            />
          </div>
          <div className={styles.viewToggles}>
            <button
              className={viewMode === 'grid' ? styles.viewActive : ''}
              onClick={() => setViewMode('grid')}
            >
              <LayoutGrid size={18} />
            </button>
            <button
              className={viewMode === 'list' ? styles.viewActive : ''}
              onClick={() => setViewMode('list')}
            >
              <List size={18} />
            </button>
          </div>
        </div>
      </div>

      {/* Main Content Area */}
      <div className={styles.contentWrapper}>
        {isLoading ? (
          <div className={styles.loaderCell}>
            <GlobalLoader variant="content" />
          </div>
        ) : dataError ? (
          <div className={styles.errorMessage}>
            <p>{dataError}</p>
          </div>
        ) : (
          <>
            {viewMode === 'list' ? (
              <table className={styles.documentTable}>
                <thead>
                  <tr>
                    <th>NAME</th>
                    <th>MODIFIED</th>
                    <th>CATEGORY/DESCRIPTION</th>
                    <th>SIZE/ITEMS</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredFolders.map((folder, idx) => (
                    <tr key={`folder-${idx}`} className={styles.tableRow} onClick={() => {
                      const siteUrl = window.location.origin;
                      window.location.href = `${siteUrl}${folder.serverRelativeUrl}`;
                    }}>
                      <td>
                        <div className={styles.itemCell}>
                          <div className={styles.folderIconWrapper}><Folder size={20} /></div>
                          <span className={styles.itemName}>{folder.name}</span>
                        </div>
                      </td>
                      <td><span className={styles.dateText}>{formatDate(folder.modified)}</span></td>
                      <td><span className={styles.dateText}>Folder Library</span></td>
                      <td><span className={styles.sizeText}>{folder.itemCount} items</span></td>
                    </tr>
                  ))}

                  {filteredListItems.map((item, idx) => (
                    <tr key={`list-item-${idx}`} className={styles.tableRow} onClick={() => {
                      if (openDropdownId === item.id) return;
                      if (item.link) window.location.href = item.link;
                    }}>
                      <td>
                        <div className={styles.itemCell}>
                          <div className={styles.fileIconWrapper}><FileText size={20} /></div>
                          <span className={styles.itemName}>{item.title}</span>
                        </div>
                      </td>
                      <td><span className={styles.dateText}>--</span></td>
                      <td><span className={styles.dateText}>{item.description || "Template Document"}</span></td>
                      <td>
                        {item.hasChildren ? (
                          <div className={styles.dropdownWrapper}>
                            <button className={styles.dropdownTrigger} onClick={(e) => handleDropdownToggle(item.id, item.title, e)}>
                              <MoreVertical size={18} />
                            </button>
                            {openDropdownId === item.id && (
                              <div className={styles.dropdownMenu}>
                                {loadingChildId === item.id ? (
                                  <div className={styles.dropdownOption}>Loading...</div>
                                ) : (childItemsCache[item.title]?.map(child => (
                                  <button key={child.id} className={styles.dropdownOption} onClick={(e) => {
                                    e.stopPropagation();
                                    if (child.link) window.location.href = child.link;
                                    setOpenDropdownId(null);
                                  }}>{child.title}</button>
                                )))}
                              </div>
                            )}
                          </div>
                        ) : (
                          <span className={styles.sizeText}>Link</span>
                        )}
                      </td>
                    </tr>
                  ))}

                  {totalItems === 0 && (
                    <tr>
                      <td colSpan={4} className={styles.emptyCell}>No items found</td>
                    </tr>
                  )}
                </tbody>
              </table>
            ) : (
              <div className={styles.documentGrid}>
                {filteredFolders.map((folder, idx) => (
                  <div
                    key={`folder-${idx}`}
                    className={`${styles.gridCard} ${styles['card' + (idx % 6)]}`}
                    onClick={() => {
                      const siteUrl = window.location.origin;
                      window.location.href = `${siteUrl}${folder.serverRelativeUrl}`;
                    }}
                  >
                    <div className={styles.cardIcon}><Folder size={18} /></div>
                    <div className={styles.cardTitle}>{folder.name}</div>
                    <div className={styles.cardSub}>Folder Library</div>
                    <div className={styles.cardFooter}>
                      <div className={styles.metaRow}>
                        <span className={styles.label}>MODIFIED</span>
                        <span className={styles.value}>{formatDate(folder.modified)}</span>
                      </div>
                      <div className={styles.metaRow}>
                        <span className={styles.label}>ITEMS</span>
                        <span className={styles.value}>{folder.itemCount}</span>
                      </div>
                    </div>
                  </div>
                ))}
                {filteredListItems.map((item, idx) => (
                  <div
                    key={`list-item-${idx}`}
                    className={`${styles.gridCard} ${styles['card' + ((filteredFolders.length + idx) % 6)]}`}
                    onClick={() => {
                      if (openDropdownId === item.id) return;
                      if (item.link) window.location.href = item.link;
                    }}
                  >
                    <div className={styles.cardIcon}><FileText size={18} /></div>

                    {item.hasChildren && (
                      <div className={styles.dropdownWrapper}>
                        <button className={styles.dropdownTrigger} onClick={(e) => handleDropdownToggle(item.id, item.title, e)}>
                          <MoreVertical size={18} />
                        </button>
                        {openDropdownId === item.id && (
                          <div className={styles.dropdownMenu}>
                            {loadingChildId === item.id ? (
                              <div className={styles.dropdownOption}>Loading...</div>
                            ) : (childItemsCache[item.title]?.map(child => (
                              <button key={child.id} className={styles.dropdownOption} onClick={(e) => {
                                e.stopPropagation();
                                if (child.link) window.location.href = child.link;
                                setOpenDropdownId(null);
                              }}>{child.title}</button>
                            )))}
                          </div>
                        )}
                      </div>
                    )}

                    <div className={styles.cardTitle}>{item.title}</div>
                    <div className={styles.cardSub} title={item.description || "Template and Document"} key={`description-${item.id}`}>{item.description || "Template and Document"}</div>
                    <div className={styles.cardFooter}>
                      <div className={styles.metaRow}>
                        <span className={styles.label}>TYPE</span>
                        <span className={styles.value}>Template</span>
                      </div>
                      <div className={styles.metaRow}>
                        <span className={styles.label}>ACCESS</span>
                        <span className={styles.value}>{item.link ? "Direct Link" : "View"}</span>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
};

export default TemplatesResources;
