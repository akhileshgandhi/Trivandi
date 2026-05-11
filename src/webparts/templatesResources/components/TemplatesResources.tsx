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
//           }

//           if (!folders || folders.length === 0) {
//             
//           }
//         } catch (folderError) {
//           
//           // Fallback to list data if folder fetch fails
//         }

//         // Fetch templates from the TemplatesandDocuments list
//         const templatesData = await getTemplatesAndDocumentsListData();
//         
//         setTemplatesListData(templatesData);

//         // First try to get data from the specific SharePoint site and folder
//         let data = await getTemplatesResourcesData();

//         // If no data was found, try the current site's Documents library as fallback
//         if (!data.allFiles.length && !data.allFolders.length && !data.folderStructure.length) {
//           
//           setDataError('No templates and resources found in the SharePoint library. Please check if the documents exist or if you have the necessary permissions.');
//         }

//         
//         setDocumentsData(data);

//       } catch (error) {
//         
//         setDataError('Failed to load templates and resources data. Please check your permissions or try refreshing the page.');
//         setDocumentsData({
//           standardDocuments: [],
//           invoicingDocuments: [],
//           resourceFolders: [],
//           allFiles: [],
//           allFolders: [],
//           folderStructure: []
//         });
//       } finally {
//         setIsLoading(false);
//       }
//     };

//     fetchData();
//   }, []);
//   return (
//     <div className={styles.page}>
//       {/* ================= HERO ================= */}
//       <Hero
//         title={`Hey ${props.userDisplayName},`}
//         subtitle="Step into your project control hub!"
//         bg={heroBg}
//       />
//       {isLoading ? (
//         <GlobalLoader variant="content" /> 
//       ) : (
//         <>
//           {/* ================= PAGE LABEL ================= */}
//           <div className={styles.label}>
//             <span className={styles.labelLine}></span>
//             <h4>Templates and Resources</h4>
//           </div>

//           {/* ================= OUTER BLOCK ================= */}
//           <div className={styles.outerBlock}>
//             {/* ================= TABS ================= */}
//             <div className={styles.tabs}>
//               {tabs.map((tab) => (
//                 <button
//                   key={tab}
//                   className={`${styles.tab} ${
//                     activeTab === tab ? styles.activeTab : ""
//                   }`} onClick={() => onTabChange(tab)}>
//                   {tab}
//                 </button>
//               ))}
//             </div>

//             {/* ================= BACK LINK (ONLY FOR WORKFLOW DETAILS) ================= */}
//             {activeTab === "Activity Workflows" && selectedWorkflow && (
//               <div className={styles.backRow}>
//                 <button
//                   className={styles.backBtn}
//                   onClick={() => setSelectedWorkflow(null)}
//                 >
//                   <img src={arrow} alt="" className={styles.backIcon} />
//                   <span>Back to workflows</span>
//                 </button>
//               </div>
//             )}

//             {/* ================= PROJECT TEMPLATES ================= */}
//             {activeTab === "Project Templates" && (
//               <>
//                 {dataError && (
//                   <div className={styles.errorMessage}>
//                     <p>{dataError}</p>
//                   </div>
//                 )}

//                 {/* Show grid view if we have folders OR list data */}
//                 {((templatesFolders && templatesFolders.length > 0) || (templatesListData && templatesListData.length > 0)) ? (
//                   <div className={styles.foldersContainer}>
//                     {/* Header with count, view toggle, and search */}
//                     <div className={styles.foldersHeader}>
//                       <div className={styles.itemCount}>
//                         {templatesFolders && templatesFolders.length > 0 
//                           ? `${templatesFolders.filter(f => f.name.toLowerCase().includes(searchQuery.toLowerCase())).length} items`
//                           : templatesListData 
//                             ? `${templatesListData.filter(t => t.title.toLowerCase().includes(searchQuery.toLowerCase())).length} items`
//                             : '0 items'
//                         }
//                       </div>
//                       <div className={styles.headerActions}>
//                         <div className={styles.viewToggle}>
//                           <button 
//                             className={`${styles.viewToggleBtn} ${viewMode === 'grid' ? styles.active : ''}`}
//                             onClick={() => setViewMode('grid')}
//                             title="Grid view"
//                           >
//                             <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
//                               <rect x="2" y="2" width="5" height="5" rx="1"/>
//                               <rect x="9" y="2" width="5" height="5" rx="1"/>
//                               <rect x="2" y="9" width="5" height="5" rx="1"/>
//                               <rect x="9" y="9" width="5" height="5" rx="1"/>
//                             </svg>
//                           </button>
//                           <button 
//                             className={`${styles.viewToggleBtn} ${viewMode === 'list' ? styles.active : ''}`}
//                             onClick={() => setViewMode('list')}
//                             title="List view"
//                           >
//                             <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
//                               <rect x="2" y="3" width="12" height="2" rx="1"/>
//                               <rect x="2" y="7" width="12" height="2" rx="1"/>
//                               <rect x="2" y="11" width="12" height="2" rx="1"/>
//                             </svg>
//                           </button>
//                         </div>
//                         <div className={styles.searchContainer}>
//                           <svg className={styles.searchIcon} width="16" height="16" viewBox="0 0 16 16" fill="none">
//                             <path d="M7 12C9.76142 12 12 9.76142 12 7C12 4.23858 9.76142 2 7 2C4.23858 2 2 4.23858 2 7C2 9.76142 4.23858 12 7 12Z" stroke="#6B7280" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
//                             <path d="M14 14L10.5 10.5" stroke="#6B7280" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
//                           </svg>
//                           <input
//                             type="text"
//                             placeholder="Search..."
//                             className={styles.searchInput}
//                             value={searchQuery}
//                             onChange={(e) => setSearchQuery(e.target.value)}
//                           />
//                         </div>
//                       </div>
//                     </div>

//                     {/* Folder Grid or List View based on viewMode */}
//                     {viewMode === 'grid' ? (
//                       <div className={styles.folderGrid}>
//                         {templatesFolders && templatesFolders.length > 0 ? (
//                           /* Display actual folders from library */
//                           templatesFolders
//                             .filter(folder => 
//                               folder.name.toLowerCase().includes(searchQuery.toLowerCase())
//                             )
//                             .map((folder) => (
//                               <div 
//                                 key={folder.id} 
//                                 className={styles.folderCard}
//                                 onClick={() => {
//                                   const siteUrl = window.location.origin;
//                                   window.open(`${siteUrl}${folder.serverRelativeUrl}`, '_self');
//                                 }}
//                               >
//                                 <div className={styles.folderCardBorder} style={{ backgroundColor: folder.color }}></div>
//                                 <div className={styles.folderCardContent}>
//                                   <div className={styles.folderIconWrapper}>
//                                     <svg className={styles.folderIconSvg} style={{ color: folder.color }} width="40" height="40" viewBox="0 0 24 24" fill="none">
//                                       <path d="M3 7C3 5.89543 3.89543 5 5 5H9L11 7H19C20.1046 7 21 7.89543 21 9V17C21 18.1046 20.1046 19 19 19H5C3.89543 19 3 18.1046 3 17V7Z" fill="currentColor" opacity="0.2"/>
//                                       <path d="M3 7C3 5.89543 3.89543 5 5 5H9L11 7H19C20.1046 7 21 7.89543 21 9V17C21 18.1046 20.1046 19 19 19H5C3.89543 19 3 18.1046 3 17V7Z" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
//                                     </svg>
//                                   </div>
//                                   <div className={styles.folderCardDetails}>
//                                     <div className={styles.folderName}>{folder.name}</div>
//                                     <div className={styles.folderMeta}>{folder.itemCount} items</div>
//                                   </div>
//                                 </div>
//                                 <div className={styles.folderCardFooter}>
//                                   <span className={styles.folderModifiedLabel}>MODIFIED</span>
//                                   <span className={styles.folderModifiedDate}>
//                                     {new Date(folder.modified).toLocaleDateString('en-US', { 
//                                       month: 'short', 
//                                       day: 'numeric', 
//                                       year: 'numeric' 
//                                     })}
//                                   </span>
//                                 </div>
//                               </div>
//                             ))
//                         ) : templatesListData && templatesListData.length > 0 ? (
//                           /* Display list data as folder cards */
//                           templatesListData
//                             .filter(template => 
//                               template.title.toLowerCase().includes(searchQuery.toLowerCase())
//                             )
//                             .map((template, index) => (
//                               <div 
//                                 key={template.id} 
//                                 className={styles.folderCard}
//                                 onClick={() => {
//                                   if (openDropdownId === template.id) return;
//                                   if (template.link) {
//                                     window.open(template.link, '_self');
//                                   }
//                                 }}
//                               >
//                                 <div className={styles.folderCardBorder} style={{ backgroundColor: alternatingColors[index % alternatingColors.length] }}></div>
//                                 <div className={styles.folderCardContent}>
//                                   <div className={styles.folderIconWrapper}>
//                                     <svg className={styles.folderIconSvg} style={{ color: alternatingColors[index % alternatingColors.length] }} width="40" height="40" viewBox="0 0 24 24" fill="none">
//                                       <path d="M3 7C3 5.89543 3.89543 5 5 5H9L11 7H19C20.1046 7 21 7.89543 21 9V17C21 18.1046 20.1046 19 19 19H5C3.89543 19 3 18.1046 3 17V7Z" fill="currentColor" opacity="0.2"/>
//                                       <path d="M3 7C3 5.89543 3.89543 5 5 5H9L11 7H19C20.1046 7 21 7.89543 21 9V17C21 18.1046 20.1046 19 19 19H5C3.89543 19 3 18.1046 3 17V7Z" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
//                                     </svg>
//                                   </div>
//                                   <div className={styles.folderCardDetails}>
//                                     <div className={styles.folderName}>{template.title}</div>
//                                     <div className={styles.folderMeta}>{template.description}</div>
//                                   </div>
//                                   {template.hasChildren && (
//                                     <div className={styles.cardDropdownWrapper}>
//                                       <button
//                                         className={styles.cardDropdownTrigger}
//                                         onClick={(e) => handleDropdownToggle(template.id, template.title, e)}
//                                         title="Show options"
//                                       >
//                                         &#8943;
//                                       </button>
//                                       {openDropdownId === template.id && (
//                                         <div className={styles.cardDropdownMenu}>
//                                           {loadingChildId === template.id ? (
//                                             <div className={styles.cardDropdownOption} style={{ color: '#9ca3af' }}>Loading...</div>
//                                           ) : (childItemsCache[template.title] || []).length > 0 ? (
//                                             (childItemsCache[template.title] || []).map(child => (
//                                               <button
//                                                 key={child.id}
//                                                 className={styles.cardDropdownOption}
//                                                 onClick={(e) => {
//                                                   e.stopPropagation();
//                                                   if (child.link) window.open(child.link, '_self');
//                                                   setOpenDropdownId(null);
//                                                 }}
//                                               >
//                                                 {child.title}
//                                               </button>
//                                             ))
//                                           ) : (
//                                             <div className={styles.cardDropdownOption} style={{ color: '#9ca3af' }}>No options</div>
//                                           )}
//                                         </div>
//                                       )}
//                                     </div>
//                                   )}
//                                 </div>
//                                 <div className={styles.folderCardFooter}>
//                                   <span className={styles.folderModifiedLabel}>TEMPLATE</span>
//                                   <span className={styles.folderModifiedDate}>Click to open</span>
//                                 </div>
//                               </div>
//                             ))
//                         ) : null}
//                       </div>
//                     ) : (
//                       /* List View */
//                       <div className={styles.folderListView}>
//                         <div className={styles.listHeader}>
//                           <div className={styles.listHeaderName}>NAME</div>
//                           <div className={styles.listHeaderModified}>MODIFIED</div>
//                           <div className={styles.listHeaderSize}>SIZE/ITEMS</div>
//                         </div>
//                         <div className={styles.listBody}>
//                           {templatesFolders && templatesFolders.length > 0 ? (
//                             /* Display actual folders from library */
//                             templatesFolders
//                               .filter(folder => 
//                                 folder.name.toLowerCase().includes(searchQuery.toLowerCase())
//                               )
//                               .map((folder) => (
//                                 <div 
//                                   key={folder.id} 
//                                   className={styles.listRow}
//                                   onClick={() => {
//                                     const siteUrl = window.location.origin;
//                                     window.open(`${siteUrl}${folder.serverRelativeUrl}`, '_self');
//                                   }}
//                                 >
//                                   <div className={styles.listRowName}>
//                                     <svg width="24" height="24" viewBox="0 0 24 24" fill="none" style={{ color: folder.color, flexShrink: 0 }}>
//                                       <path d="M3 7C3 5.89543 3.89543 5 5 5H9L11 7H19C20.1046 7 21 7.89543 21 9V17C21 18.1046 20.1046 19 19 19H5C3.89543 19 3 18.1046 3 17V7Z" fill="currentColor" opacity="0.3"/>
//                                       <path d="M3 7C3 5.89543 3.89543 5 5 5H9L11 7H19C20.1046 7 21 7.89543 21 9V17C21 18.1046 20.1046 19 19 19H5C3.89543 19 3 18.1046 3 17V7Z" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
//                                     </svg>
//                                     <span>{folder.name}</span>
//                                   </div>
//                                   <div className={styles.listRowModified}>
//                                     {new Date(folder.modified).toLocaleDateString('en-US', { 
//                                       month: 'short', 
//                                       day: 'numeric', 
//                                       year: 'numeric' 
//                                     })}
//                                   </div>
//                                   <div className={styles.listRowSize}>{folder.itemCount} items</div>
//                                 </div>
//                               ))
//                           ) : templatesListData && templatesListData.length > 0 ? (
//                             /* Display list data */
//                             templatesListData
//                               .filter(template => 
//                                 template.title.toLowerCase().includes(searchQuery.toLowerCase())
//                               )
//                               .map((template, index) => (
//                                 <div 
//                                   key={template.id} 
//                                   className={styles.listRow}
//                                   onClick={() => {
//                                     if (openDropdownId === template.id) return;
//                                     if (template.link) {
//                                       window.open(template.link, '_self');
//                                     }
//                                   }}
//                                 >
//                                   <div className={styles.listRowName}>
//                                     <svg width="24" height="24" viewBox="0 0 24 24" fill="none" style={{ color: alternatingColors[index % alternatingColors.length], flexShrink: 0 }}>
//                                       <path d="M3 7C3 5.89543 3.89543 5 5 5H9L11 7H19C20.1046 7 21 7.89543 21 9V17C21 18.1046 20.1046 19 19 19H5C3.89543 19 3 18.1046 3 17V7Z" fill="currentColor" opacity="0.3"/>
//                                       <path d="M3 7C3 5.89543 3.89543 5 5 5H9L11 7H19C20.1046 7 21 7.89543 21 9V17C21 18.1046 20.1046 19 19 19H5C3.89543 19 3 18.1046 3 17V7Z" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
//                                     </svg>
//                                     <span>{template.title}</span>
//                                   </div>
//                                   <div className={styles.listRowModified}>-</div>
//                                   <div className={styles.listRowSize}>
//                                     {template.hasChildren && (
//                                       <div className={styles.cardDropdownWrapper}>
//                                         <button
//                                           className={styles.listDropdownTrigger}
//                                           onClick={(e) => handleDropdownToggle(template.id, template.title, e)}
//                                           title="Show options"
//                                         >
//                                           &#8943;
//                                         </button>
//                                         {openDropdownId === template.id && (
//                                           <div className={styles.cardDropdownMenu} style={{ right: 0, left: 'auto' }}>
//                                             {loadingChildId === template.id ? (
//                                               <div className={styles.cardDropdownOption} style={{ color: '#9ca3af' }}>Loading...</div>
//                                             ) : (childItemsCache[template.title] || []).length > 0 ? (
//                                               (childItemsCache[template.title] || []).map(child => (
//                                                 <button
//                                                   key={child.id}
//                                                   className={styles.cardDropdownOption}
//                                                   onClick={(e) => {
//                                                     e.stopPropagation();
//                                                     if (child.link) window.open(child.link, '_self');
//                                                     setOpenDropdownId(null);
//                                                   }}
//                                                 >
//                                                   {child.title}
//                                                 </button>
//                                               ))
//                                             ) : (
//                                               <div className={styles.cardDropdownOption} style={{ color: '#9ca3af' }}>No options</div>
//                                             )}
//                                           </div>
//                                         )}
//                                       </div>
//                                     )}
//                                   </div>
//                                 </div>
//                               ))
//                           ) : null}
//                         </div>
//                       </div>
//                     )}
//                   </div>
//                 ) : (
//                   /* Show message if no data */
//                   <div className={styles.noDataMessage}>
//                     <p>No templates or resources found.</p>
//                   </div>
//                 )}

//                 {/* <Section title="Standard Documents" viewAll>
//                   <div className={styles.grid}>
//                     {documentsData?.standardDocuments.length ? (
//                       documentsData.standardDocuments.map((doc) => (
//                         <Card
//                           key={doc.Id}
//                           title={doc.Title || doc.Name}
//                           isDocument
//                           documentType={doc.DocumentType || "book"}
//                           variant="template"
//                           onClick={() => window.open(getFilePreviewUrl(doc.FileRef, doc.File_x0020_Type), '_self')}
//                         />
//                       ))
//                     ) : (
//                       <div className={styles.noDataMessage}>
//                         <p>No standard documents found in the SharePoint library.</p>
//                         <p>Please ensure documents exist in the library and you have appropriate permissions.</p>
//                       </div>
//                     )}
//                   </div>
//                 </Section>

//                 <div className={styles.sectionGap}>
//                   <Section title="Invoicing & Reporting" viewAll>
//                     <div className={styles.grid}>
//                       {documentsData?.invoicingDocuments.length ? (
//                         documentsData.invoicingDocuments.map((doc) => (
//                           <Card
//                             key={doc.Id}
//                             title={doc.Title || doc.Name}
//                             isDocument
//                             documentType={doc.DocumentType || "book"}
//                             variant="invoice"
//                             onClick={() => window.open(getFilePreviewUrl(doc.FileRef, doc.File_x0020_Type), '_self')}
//                           />
//                         ))
//                       ) : (
//                         <div className={styles.noDataMessage}>
//                           <p>No invoicing documents found in the SharePoint library.</p>
//                           <p>Please ensure documents exist in the library and you have appropriate permissions.</p>
//                         </div>
//                       )}
//                     </div>
//                   </Section>
//                 </div> */}
//               </>
//             )}

//             {/* ================= RESOURCE LIBRARY ================= */}
//             {/* {activeTab === "Resource Library" && (
//               <Section title="Resources & Materials">
//                 {dataError && (
//                   <div className={styles.errorMessage}>
//                     <p>{dataError}</p>
//                   </div>
//                 )}

//                 <div className={styles.resourceGrid}>
//                   {documentsData?.resourceFolders.length ? (
//                     documentsData.resourceFolders.map((folder) => (
//                       <Card
//                         key={folder.Id}
//                         title={folder.Title || folder.Name}
//                         subtitle={folder.Description || `Resources and materials in ${folder.Name}`}
//                         color={folder.Color || "#0b0f6b"}
//                         iconRight
//                         onClick={() => {
//                           const folderUrl = `https://trivandildn.sharepoint.com${folder.FileRef}`;
//                           window.open(folderUrl, '_self');
//                         }}
//                       />
//                     ))
//                   ) : null}

//                   {documentsData?.allFiles.filter(file => 
//                     !documentsData.standardDocuments.find(std => std.Id === file.Id) &&
//                     !documentsData.invoicingDocuments.find(inv => inv.Id === file.Id)
//                   ).map((file) => (
//                     <Card
//                       key={`file-${file.Id}`}
//                       title={file.Title || file.Name}
//                       subtitle={`File • Modified ${new Date(file.Modified).toLocaleDateString()}`}
//                       color="#8fa3e8"
//                       iconRight
//                       onClick={() => window.open(getFilePreviewUrl(file.FileRef, file.File_x0020_Type), '_self')}
//                     />
//                   ))}
//                 </div>
//               </Section>
//             )} */}

//             {/* ================= FOLDER STRUCTURE ================= */}
//             {activeTab === "Folder Structure" && (
//               <Section title="Complete Folder Structure">
//                 {dataError && (
//                   <div className={styles.errorMessage}>
//                     <p>{dataError}</p>
//                   </div>
//                 )}

//                 <div className={styles.infoMessage}>
//                   <p>
//                     <strong>📂 Hierarchical View:</strong> 
//                     {" "}This shows the complete folder and file structure. Click folders to expand/collapse and click files to preview.
//                   </p>
//                 </div>

//                 <div className={styles.folderStructureContainer}>
//                   {documentsData?.folderStructure.length ? (
//                     renderFolderStructure(documentsData.folderStructure)
//                   ) : (
//                     <div className={styles.noDataMessage}>
//                       <p>No folder structure available.</p>
//                       <p>Please ensure the SharePoint library contains folders and files, and that you have appropriate permissions to access them.</p>
//                     </div>
//                   )}
//                 </div>
//               </Section>
//             )}

//             {/* ================= ACTIVITY WORKFLOWS (LIST) ================= */}
//             {activeTab === "Activity Workflows" && !selectedWorkflow && (
//               <Section title="Standard Activity Workflows">
//                 <div className={styles.workflowList}>
//                   {workflows.map((wf, index) => (
//                     <div
//                       key={wf.id}
//                       className={styles.workflowItem}
//                       onClick={() => setSelectedWorkflow(wf)}
//                     >
//                       <div className={styles.stepCircle}>{index + 1}</div>
//                       <span>{wf.title}</span>
//                       <span className={styles.arrow}>
//                         <img src={arrow} alt="" />
//                       </span>
//                     </div>
//                   ))}
//                 </div>
//               </Section>
//             )}

//             {/* ================= ACTIVITY WORKFLOWS (DETAILS) ================= */}
//             {activeTab === "Activity Workflows" && selectedWorkflow && (
//               <Section title="Standard Activity Workflows">
//                 <div className={styles.workflowList}>
//                   {selectedWorkflow.steps.map((step) => (
//                     <div key={step.step} className={styles.workflowItem}>
//                       <div className={styles.stepCircle}>{step.step}</div>

//                       <div className={styles.stepContent}>
//                         <h5>{step.title}</h5>
//                         <div className={styles.innerStepContent}>
//                           <p className={styles.innerStepContent1}>
//                             {" "}
//                             {step.type}{" "}
//                           </p>
//                           <p className={styles.innerStepContent2}>
//                             {" "}
//                             {step.meta}
//                           </p>
//                         </div>
//                       </div>

//                       <span className={styles.arrow}>
//                         <img src={arrow} alt="" />
//                       </span>
//                     </div>
//                   ))}
//                 </div>
//               </Section>
//             )}
//           </div>
//         </>
//       )}
//     </div>
//   );
// };

// export default TemplatesResources;
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
            } catch (err) {

            }
          }
        } catch (folderError) {

        }

        // Fetch templates from the TemplatesandDocuments list
        try {
          const templatesListData = await getTemplatesAndDocumentsListData();
          setTemplatesListData(templatesListData);
        } catch (listError) {

        }

      } catch (error) {

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
