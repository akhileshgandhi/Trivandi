import React from 'react'
import { getProjectsByOwnerAndStatus, IDashboardProject } from '../../../shared/services/projectService';
import styles from "../components/MyWorkNew.module.scss";
import Hero from '../../../shared/component/Hero/Hero';
import GlobalLoader from '../../../shared/component/GlobalLoader';
import Section from '../../../shared/component/Section/Section';
import Card from '../../../shared/component/Card/Card';
import RecentDocuments from './RecentDocuments/RecentDocuments';
import heroBg from '../../../shared/assets/hero_new.jpg';
import '../../../shared/globalcss/globalcss.scss';
const MyWorkNew = ({ props }) => {

  const [isLoading, setIsLoading] = React.useState(true);
  const [liveProjects, setLiveProjects] = React.useState<IDashboardProject[]>(
    []
  );
  const [liveBids, setLiveBids] = React.useState<IDashboardProject[]>([]);

  React.useEffect(() => {
    const loadData = async () => {
      try {
        console.log("[Dashboard] Loading dashboard data...");
        setIsLoading(true);

        const ownerEmail = props?.context?.pageContext?.user?.email;
        if (!ownerEmail) {
          setLiveProjects([]);
          setLiveBids([]);
          return;
        }

        const [projects, bids] = await Promise.all([
          getProjectsByOwnerAndStatus(ownerEmail, "Project"),
          getProjectsByOwnerAndStatus(ownerEmail, ["Potential"]),
        ]);

        console.log("[Dashboard] Live Projects:", projects);
        console.log("[Dashboard] Live Bids:", bids);

        setLiveProjects(projects);
        setLiveBids(bids);
      } catch (error) {
        console.error("[Dashboard] Data load failed", error);
      } finally {
        setIsLoading(false);
      }
    };

    void loadData();
  }, []);

  // ✅ Resolve greeting from current user
  const displayName = props.userDisplayName?.trim();
  const firstName =
    displayName && displayName.length > 0 ? displayName.split(" ")[0] : "there";

  return (
    <div className={styles.dashboard}>


      {isLoading ? (
        <GlobalLoader variant="content" />
      ) : (
        <>
          {/* ================= MY WORK ================= */}
          <div className={styles.label}>
            <span className={styles.labelLine} />
            <h4>My Work</h4>
          </div>

          <div className={styles.topSection}>
            {/* ===== MY LIVE PROJECTS ===== */}
            <Section title="My Live Projects">
              {liveProjects.length === 0 ? (
                <div className={styles.emptyState}>
                  <img src={require("../../../shared/assets/Empty.png")} alt="No data" className={styles.emptyIcon} />
                  <p className={styles.emptyText}>No live projects available</p>
                </div>
              ) : (
                <div className={styles.scrollWrapper}>
                  <div className={`${styles.grid} ${styles.gridTwo}`}>
                    {liveProjects.map((item) => (
                      <Card
                        key={item.Id}
                        title={item.Title}
                        subtitle={item.Company}
                        meta={item.Office ?? "Office not specified"}
                        border="2px solid #0511F2"
                        onClick={() => {
                          window.location.href = `/sites/Projects/SitePages/ProjectDetails.aspx?projectId=${item.Id}`;
                        }}
                      />
                    ))}
                  </div>
                </div>
              )}
            </Section>

            {/* ===== MY LIVE BIDS ===== */}
            <Section title="My Live Pipeline">
              {liveBids.length === 0 ? (
                <div className={styles.emptyState}>
                  <img src={require("../../../shared/assets/Empty.png")} alt="No data" className={styles.emptyIcon} />
                  <p className={styles.emptyText}>No live bids available</p>
                </div>
              ) : (
                <div className={styles.scrollWrapper}>
                  <div className={`${styles.grid} ${styles.gridTwo}`}>
                    {liveBids.map((item) => (
                      <Card
                        key={item.Id}
                        title={item.Title}
                        subtitle={item.Company}
                        meta={item.Office ?? "Office not specified"}
                        customIcon={require("../../../shared/assets/airg.png")}
                        border="2px solid #EC008C"
                        noIconBg
                        onClick={() => {
                          window.location.href = `/sites/Projects/SitePages/PipelineDetails.aspx?projectId=${item.Id}`;
                        }}
                      />
                    ))}
                  </div>
                </div>
              )}
            </Section>
          </div>

          {/* ================= STANDARDS & COMPLIANCE ================= */}
          {/* <div className={styles.label}>
            <span className={styles.labelLine} />
            <h4>Standards & Compliance</h4>
          </div>

          <div className={styles.topSection2}>
            <Section>
              <div className={`${styles.grid} ${styles.gridFiveSmall}`}>
                <Card
                  title="Associate Onboarding"
                  subtitle="Process guide..."
                  color="#0A1B44"
                  iconRight
                />
                <Card
                  title="Mobilisations"
                  subtitle="Process guide..."
                  color="#F58220"
                  iconRight
                />
                <Card
                  title="Health & Safety"
                  subtitle="Process guide..."
                  color="#8B9CF6"
                  iconRight
                />
                <Card
                  title="RICS Standards"
                  subtitle="Process guide..."
                  color="#1E40FF"
                  iconRight
                />
                <Card
                  title="Project Delivery Rules"
                  subtitle="Process guide..."
                  color="#EC008C"
                  iconRight
                />
              </div>
            </Section>
          </div> */}

          {/* ================= RECENT DOCUMENTS ================= */}
          <div className={styles.label}>
            <span className={styles.labelLine} />
            <h4>Recent Documents</h4>
          </div>

          <div className={styles.topSection2}>
            <Section>
              <div className={styles.gridDocuments}>
                <RecentDocuments />
                {/* <Card title="Project_Kickoff_Deck_v2.pptx" subtitle="Event Readiness Programme • 2 hours ago" isDocument />
                <Card title="Budget_Tracking_Oct25.xlsx" subtitle="London Gatwick Airport • Yesterday" isDocument />
                <Card title="Meeting_Minutes_SteerCo.docx" subtitle="Driving Council Steering • Yesterday" isDocument />
                <Card title="RFP_Response_Draft.docx" subtitle="Wembley Security Review • 2 days ago" isDocument /> */}
              </div>
            </Section>
          </div>


        </>
      )}
    </div>
  );
};


export default MyWorkNew