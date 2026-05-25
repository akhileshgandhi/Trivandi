import * as React from "react";
import { useState, useEffect } from "react";
import styles from "./Documents.module.scss";
import ProjectDocuments from "./ProjectDocuments";
import BidDocuments from "./BidDocuments";
import { getProjectById } from "../../../../shared/services/projectService";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import { IDocumentsProps } from "../../../../shared/interfaces/IDocumentInterfaces";

const CustomDocumentsList: React.FC<IDocumentsProps> = (props) => {
  const [activeDocTab, setActiveDocTab] = useState<'project' | 'bid' | 'contract'>('project');
  const [hasSplitFolderTabs, setHasSplitFolderTabs] = useState<boolean>(true); // default true for now

  // Add initial tab logic
  useEffect(() => {
    const fetchProjectUrls = async () => {
      try {
        const projectData = await getProjectById(props.projectId);
        if (projectData && typeof projectData === 'object' && 'ProjectDocumentsUrl' in projectData) {
          const projUrl = projectData.ProjectDocumentsUrl;
          const bidUrl = projectData.BidDocumentsUrl;
          const contractUrl = projectData.ContractsDocumentsUrl;
          const status = projectData.Status || '';

          const searchParams = new URLSearchParams(window.location.search);
          const urlSource = searchParams.get('source');

          if (urlSource === 'bid' && bidUrl) {
            setActiveDocTab('bid');
          } else if (urlSource === 'project' && projUrl) {
            setActiveDocTab('project');
          } else if (urlSource === 'contract' && contractUrl) {
            setActiveDocTab('contract');
          } else {
            if (status === 'Lost' && bidUrl) {
              setActiveDocTab('bid');
            } else if (projUrl) {
              setActiveDocTab('project');
            } else if (bidUrl) {
              setActiveDocTab('bid');
            } else if (contractUrl) {
              setActiveDocTab('contract');
            } else {
              setActiveDocTab('project');
            }
          }
        }
      } catch (error) {
        console.error(error);
      }
    };
    void fetchProjectUrls();
  }, [props.projectId]);

  return (
    <div className={styles.documentsWrapper}>
      <div className={styles.documentTabs}>
        {hasSplitFolderTabs ? (
          <>
            <div
              className={`${styles.docTab} ${activeDocTab === 'project' ? styles.activeDocTab : ''}`}
              onClick={() => setActiveDocTab('project')}
            >
              Project Documents
            </div>
            <div
              className={`${styles.docTab} ${activeDocTab === 'bid' ? styles.activeDocTab : ''}`}
              onClick={() => setActiveDocTab('bid')}
            >
              Bid Documents
            </div>
            {/* 
            <div
              className={`${styles.docTab} ${activeDocTab === 'contract' ? styles.activeDocTab : ''}`}
              onClick={() => setActiveDocTab('contract')}
            >
              Contract Documents
            </div>
            */}
          </>
        ) : (
          <div className={`${styles.docTab} ${styles.activeDocTab}`}>Documents</div>
        )}
      </div>

      <div style={{ marginTop: "16px" }}>
        {activeDocTab === 'project' && (
          <ProjectDocuments {...props} />
        )}
        {activeDocTab === 'bid' && (
          <BidDocuments {...props} />
        )}
      </div>
    </div>
  );
};

export default CustomDocumentsList;