/* eslint-disable no-void */
import * as React from "react";
import { useEffect, useMemo, useState, useRef } from "react";
import { toast, ToastContainer } from "react-toastify";
import "react-toastify/dist/ReactToastify.css";
import GlobalLoader from "../../../shared/component/GlobalLoader";
import { getProjectById, initializePnP } from "../../../shared/services/projectService";
import { initLibraryDiscoveryService, getProjectLibraryName } from "../../../shared/services/libraryDiscoveryService";
import {
  getTeamMembersByProjectId,
  addTeamMember,
  deleteTeamMember,
  initTeamMemberService,
  sendInvitation,
  ITeamMember
} from "../../../shared/services/teamMemberService";
import styles from "./ProjectDetails.module.scss";
import { IProjectDetailsProps } from "./IProjectDetailsProps";
import CustomDocumentsList from "./CustomComponent/CustomDocumentsList";
import AddTeamMemberModal from "./CustomComponent/AddTeamMemberModal/AddTeamMemberModal";
import "../../../shared/globalcss/globalcss.scss";
import { ArrowLeft, CalendarDays, MapPin, User } from "lucide-react";
import SharedFiles from "../components/SharedFiles/SharedFiles";
import ActiveGuestsComponent from "./ActiveGuests/ActiveGuests";
import InviteGuestDialog from "../components/Dialogs/InviteGuestDialog";
import ShareDocumentDialog from "../components/Dialogs/ShareDocumentDialog";
import ImportFromProjectDocsDialog from './Dialogs/ImportFromProjectDocsDialog';
import CustomManageAccess from '../../../shared/Common/CustomManageAccess';
import NewDocumentDialog from "../components/Dialogs/NewDocumentDialog";
import { IBreadcrumb } from './IProjectExternalPortalState';
import { ProjectExternalPortalService } from "../services/ProjectExternalPortalService";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import { Icon } from "@fluentui/react/lib/Icon";
import portalStyles from "./ProjectExternalPortal.module.scss";
import { usePermissionStore } from "../../../Permission/PermissionStore";

type ProjectItem = {
  Id?: number;
  Sector: string;
  Title?: string;
  Status?: string;
  Stage?: string;
  Code?: string;
  PercentComplete?: number;
  TotalProjectValue?: number;
  EstimatedPrice?: number;
  StartDate?: string;
  EndDate?: string;
  Company?: string;
  Office?: string;
  BusinessUnit?: string;
  Owner?: { Title?: string } | string;
  OwnerEmail?: string;
  ProjectManager?: { Title?: string } | string;
  ProjectManagerEmail?: string;
  ProjectId?: string;
  ProjectID?: string;
  EventName?: string;
  Location?: string;
  ContractedEntity?: string;
  TeamMembers?: Array<{ name: string; email: string; phone: string }>;
  TeamMembersLookup?: ITeamMember[];
  TotalFee?: number;
  TotalAdditionals?: number;
  TimeCostToDate?: number;
  AdditionalsCost?: number;
  FutureAdditionalsCosts?: number;
  EstimatedProfit?: number;
  EstimatedProfitMargin?: number;
  InvoicedToDate?: number;
  BudgetVariance?: number;
  LeadSource?: string;
  FeeType: string;
  Probability: number;
  TrivandiAcademyRevenuePercent?: number;
  NonCmap?: boolean;
  UncategorisedDocumentsUrl?: string;
  ContractType?: string;
  Country?: string;
};
const formatCurrency = (value?: number): string => {
  if (!value && value !== 0) return "-";
  try {
    return new Intl.NumberFormat("en-GB", {
      style: "currency",
      currency: "GBP",
      maximumFractionDigits: 0,
    }).format(value);
  } catch {
    return `${value}`;
  }
};

const formatNumber = (value?: number): string => {
  if (!value && value !== 0) return "-";
  return new Intl.NumberFormat("en-GB").format(value);
};

const formatDate = (value?: string): string => {
  if (!value) return "-";
  const d = new Date(value);
  if (isNaN(d.getTime())) return "-";
  return d.toLocaleDateString("en-GB", { year: "numeric", month: "2-digit", day: "2-digit" });
};



const STEPS = ["Initial Discussions", "Preparing Proposal", "Proposal Submitted", "Commitment"];

const ProjectDetails: React.FC<IProjectDetailsProps> = (props) => {
  const { canAdd, canEdit } = usePermissionStore();
  const [project, setProject] = useState<ProjectItem | null>(null);
  const [loading, setLoading] = useState<boolean>(true);

  // Initialize activeTab from sessionStorage or default to "Dashboard"
  const getInitialTab = (): string => {
    const params = new URLSearchParams(window.location.search);
    const projectId = params.get("projectId");

    // Check if user came from Projects page
    const navigationSource = sessionStorage.getItem('projectDetails_navigation_source');
    if (navigationSource === 'projects_page') {
      // Clear the flag so it doesn't affect future refreshes
      sessionStorage.removeItem('projectDetails_navigation_source');
      // Return Dashboard as default when coming from Projects page
      return "Dashboard";
    }

    // For refreshes, use the last active tab for this project
    const savedTab = sessionStorage.getItem(`projectDetails_tab_${projectId || 'default'}`);
    return savedTab || "Dashboard";
  };

  const [activeTab, setActiveTab] = useState<string>(getInitialTab());
  const [teamMembers, setTeamMembers] = useState<ITeamMember[]>([]);
  const [isAddMemberModalOpen, setIsAddMemberModalOpen] = useState<boolean>(false);
  const [isLoadingMembers, setIsLoadingMembers] = useState<boolean>(false);
  const [copiedMemberId, setCopiedMemberId] = useState<number | null>(null);

  // External Portal State
  const [showInviteGuestDialog, setShowInviteGuestDialog] = useState<boolean>(false);
  const [showShareDocumentDialog, setShowShareDocumentDialog] = useState<boolean>(false);
  const [showImportDialog, setShowImportDialog] = React.useState(false);
  const [showNewDocumentDialog, setShowNewDocumentDialog] = useState<boolean>(false);
  const [showShareAccessDialog, setShowShareAccessDialog] = useState<boolean>(false);
  const [selectedDocumentsForSharing, setSelectedDocumentsForSharing] = useState<any[]>([]);
  const [currentFolderPath, setCurrentFolderPath] = useState<string>('');
  const [breadcrumbs, setBreadcrumbs] = useState<Array<{ text: string; key: string; path: string }>>([{ text: 'Root', key: 'root', path: '' }]);
  const [portalService] = useState<ProjectExternalPortalService>(() => new ProjectExternalPortalService(props.context));
  const [guestRefreshTrigger, setGuestRefreshTrigger] = useState<number>(Date.now());

  // User restriction state
  const [isUserRestricted, setIsUserRestricted] = useState<boolean>(false);
  const [restrictionCheckLoading, setRestrictionCheckLoading] = useState<boolean>(true);

  // Ref for ActiveGuests component
  const activeGuestsRef = useRef<ActiveGuestsComponent>(null);
  const sharedFilesRef = React.useRef<any>(null);

  // Scope the External Portal folder to the current project's subfolder in ExternalShareDocument
  const initPortalFolderPath = (projectId: string | number): void => {
    const pid = String(projectId);
    const folderName = `Project-${pid}`;
    portalService.setProjectId(pid);
    setCurrentFolderPath(folderName);
    setBreadcrumbs([
      { text: folderName, key: folderName, path: folderName }
    ]);
  };

  useEffect((): void => {
    const params = new URLSearchParams(window.location.search);
    const projectId = params.get("projectId");
    const useMock = params.get("useMock") === "1";

    // Initialize PnP
    if (props.context) {
      const spInstance = initializePnP(props.context);
      initTeamMemberService(spInstance);
      initLibraryDiscoveryService(spInstance);
    }

    const loadProject = async (): Promise<void> => {
      try {
        if (!useMock && projectId) {
          const data = await getProjectById(Number(projectId));
          console.log("PROJECT_DATA_FROM_API:", data);
          setProject(data as ProjectItem);

          // Scope the External Portal to this project's subfolder
          const numericId = data?.Id || Number(projectId);
          if (numericId) {
            initPortalFolderPath(numericId);
          }

          // Dynamically find and set project library name
          if (data?.Title) {
            const libraryName = await getProjectLibraryName(data.Code, data.Title);
            // Note: The project library is for the Documents tab, not External Portal
            // External Portal uses 'ExternalShareDocument' library by default

          }

          // Load team members
          await loadTeamMembers(Number(projectId));
        } else {
          const mock: ProjectItem = {
            Id: 25128,
            ProjectId: "25128",
            Title: "Event Readiness Programme",
            EventName: "AFC Asian Cup 2027 LOC",
            Status: "Live",
            Stage: "2. Preparing Proposal",
            PercentComplete: 54,
            TotalProjectValue: 3114929,
            EstimatedPrice: 3114929,
            BudgetVariance: 2602853,
            StartDate: "2025-02-15T00:00:00Z",
            EndDate: "2027-07-30T00:00:00Z",
            Company: "El Seif",
            Office: "Trivandi Global",
            BusinessUnit: "Venue Design & Delivery",
            Owner: "Tibbetts, James",
            Location: "Riyadh",
            TotalFee: 3114929,
            TotalAdditionals: 3114929,
            TimeCostToDate: 5040,
            AdditionalsCost: 0,
            FutureAdditionalsCosts: 4350,
            EstimatedProfit: 5610,
            EstimatedProfitMargin: 37,
            InvoicedToDate: 10000,
            FeeType: "Fixed Fee",
            LeadSource: "Referral",
            Probability: 75,
            Sector: "Technology",
            TeamMembers: [
              { name: "Darrell Steward", email: "darrell@example.com", phone: "+1234567890" },
              { name: "Ronald Richards", email: "ronald@example.com", phone: "+1234567890" },
              { name: "Albert Flores", email: "albert@example.com", phone: "+1234567890" },
              { name: "Darrell Steward", email: "darrell@example.com", phone: "+1234567890" },
              { name: "Ronald Richards", email: "ronald@example.com", phone: "+1234567890" },
              { name: "Albert Flores", email: "albert@example.com", phone: "+1234567890" },
              { name: "Ronald Richards", email: "ronald@example.com", phone: "+1234567890" },
              { name: "Albert Flores", email: "albert@example.com", phone: "+1234567890" },
              { name: "Darrell Steward", email: "darrell@example.com", phone: "+1234567890" },
              { name: "Albert Flores", email: "albert@example.com", phone: "+1234567890" },
            ],
          };
          setProject(mock);

          // Scope portal to mock project ID
          initPortalFolderPath(mock.Id || 25128);

          // Note: External Portal uses 'ExternalShareDocument' by default, not project library

        }
      } finally {
        setLoading(false);
      }
    };

    void loadProject();
  }, [props.context]);

  // Handle tab switching for Non-CMAP projects
  useEffect(() => {
    if (project?.NonCmap && activeTab === "Dashboard") {
      // For Non-CMAP projects, switch to Documents tab if currently on Dashboard
      setActiveTab("Documents");
    }
  }, [project, activeTab]);

  // Sync portal folder when project data is loaded or from URL directly
  React.useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const urlProjectId = params.get("projectId");
    const id = urlProjectId || project?.Id || project?.ProjectId;

    if (id && currentFolderPath === "") {
      const targetFolder = `Project-${id}`;

      setCurrentFolderPath(targetFolder);
      setBreadcrumbs([
        { text: targetFolder, key: targetFolder, path: targetFolder }
      ]);
    }
  }, [project]);

  // Check if current user is restricted (in ExternalGuestAccess list)
  useEffect(() => {
    const checkUserRestriction = async (): Promise<void> => {
      if (!props.context) {
        setRestrictionCheckLoading(false);
        return;
      }

      try {
        const isRestricted = await portalService.isCurrentUserRestrictedGuest();
        setIsUserRestricted(isRestricted);

        if (isRestricted) {

        }
      } catch (error) {

        // Fail safely - if we cannot verify, don't restrict
        setIsUserRestricted(false);
      } finally {
        setRestrictionCheckLoading(false);
      }
    };

    checkUserRestriction();
  }, [props.context, portalService]);

  // Load team members from SharePoint
  const loadTeamMembers = async (projectId: number): Promise<void> => {
    setIsLoadingMembers(true);
    try {
      const members = await getTeamMembersByProjectId(projectId);
      setTeamMembers(members);
    } catch (error) {

      setTeamMembers([]);
    } finally {
      setIsLoadingMembers(false);
    }
  };

  // Handle adding a new team member
  const handleAddMember = async (member: { name: string; email: string; role: string; mobileNumber: string }): Promise<void> => {
    // Use Id (from SharePoint) or ProjectId (from mock or custom field)
    const projectId = project?.Id || (project?.ProjectId ? Number(project.ProjectId) : null);

    if (!projectId) {
      throw new Error("Project ID not found. Please ensure the project is loaded correctly.");
    }

    try {
      const newMember = await addTeamMember(projectId, member);

      // Send invitation email
      await sendInvitation(
        member.email,
        member.name,
        project.Title || "Project",
        String(projectId)
      );

      // Reload team members
      await loadTeamMembers(projectId);


    } catch (error) {

      toast.error(error instanceof Error ? error.message : "Failed to add team member", {
        position: "top-right",
        autoClose: 5000,
      });
      throw error;
    }
  };

  // Handle deleting a team member
  const handleDeleteMember = async (memberId: number): Promise<void> => {
    if (!confirm("Are you sure you want to remove this team member?")) {
      return;
    }

    try {
      await deleteTeamMember(memberId);
      setTeamMembers(prevMembers => prevMembers.filter(m => m.Id !== memberId));

      toast.success("Team member removed successfully!", {
        position: "top-right",
        autoClose: 3000,
      });
    } catch (error) {

      toast.error("Failed to remove team member. Please try again.", {
        position: "top-right",
        autoClose: 5000,
      });
    }
  };

  // Handle copying mobile number
  const handleCopyMobile = async (memberId: number, mobileNumber: string): Promise<void> => {
    if (mobileNumber) {
      try {
        await navigator.clipboard.writeText(mobileNumber);
        setCopiedMemberId(memberId);
        setTimeout(() => setCopiedMemberId(null), 2000);
      } catch (err) {

        // Fallback for older browsers
        const textArea = document.createElement('textarea');
        textArea.value = mobileNumber;
        textArea.style.position = 'fixed';
        textArea.style.left = '-999999px';
        document.body.appendChild(textArea);
        textArea.select();
        try {
          document.execCommand('copy');
          setCopiedMemberId(memberId);
          setTimeout(() => setCopiedMemberId(null), 2000);
        } catch (e) {

        }
        document.body.removeChild(textArea);
      }
    }
  };

  // Handle calling mobile number
  const handleCallMobile = (mobileNumber: string): void => {
    if (mobileNumber) {
      window.location.href = `tel:${mobileNumber}`;
    }
  };

  const currentStepIndex = useMemo(() => {
    const stg = (project?.Stage || "").toLowerCase();
    for (let i = 0; i < STEPS.length; i++) {
      const s = STEPS[i].toLowerCase();
      if (stg.includes(s)) return i;
      if (stg.startsWith(`${i + 1}. `) && stg.includes(s)) return i;
    }
    return -1;
  }, [project?.Stage]);

  const completion = Math.max(0, Math.min(100, project?.PercentComplete ?? 0));

  if (loading && !project) return <GlobalLoader variant="content" />;
  if (!project) return <div className={styles.noData}>No project found</div>;

  const goBack = (): void => {
    window.history.back();
  }
  // Add this helper function at the top of your component (after the formatDate function)
  const getPersonName = (person?: { Title?: string } | string): string | null => {
    if (!person) return null;
    if (typeof person === "string") return person || null;
    return person.Title || null;
  };

  const getInitials = (person?: { Title?: string } | string): string => {
    const name = getPersonName(person);
    if (!name) return "?";
    return name.split(" ").map(n => n[0]).join("").substring(0, 2).toUpperCase();
  };

  // Handler for tab change with persistence
  const handleTabChange = (tab: string): void => {
    setActiveTab(tab);
    const params = new URLSearchParams(window.location.search);
    const projectId = params.get("projectId");
    sessionStorage.setItem(`projectDetails_tab_${projectId || 'default'}`, tab);
  };

  // Handler for opening documents tab
  const handleOpenFolder = (): void => {
    handleTabChange("Documents");
  };

  // Handler for refresh after document operations
  const handleRefreshDocuments = (): void => {

    if (sharedFilesRef.current && sharedFilesRef.current.refreshDocuments) {

      sharedFilesRef.current.refreshDocuments();
    } else {

    }
  };

  // Handler for showing new document dialog
  const handleShowNewDocument = (): void => {
    setShowNewDocumentDialog(true);
  };

  // Handler for showing share access dialog
  const handleShowShareAccess = (selectedDocs: any[]): void => {
    setSelectedDocumentsForSharing(selectedDocs);
    setShowShareAccessDialog(true);
  };

  // Handler for showing import dialog
  const handleShowImportDialog = (): void => {
    setShowImportDialog(true);
  };

  // Handler for successful operations
  const handleOperationSuccess = (): void => {
    handleRefreshDocuments();
  };

  // Handler for successful guest invitation
  const handleGuestInvited = (): void => {


    // Update refresh trigger to force ActiveGuests component to refresh
    setGuestRefreshTrigger(Date.now());

    // Also call the ref method as backup

    if (activeGuestsRef.current) {

      activeGuestsRef.current.refreshGuests();
    } else {

    }
  };


  // Then update the Team Members section:
  return (
    <div className={styles.wrapper}>
      {loading && <GlobalLoader variant="bar" />}
      <ToastContainer />
      <div className={styles.header}>
        <button type="button" className={styles.backButton} onClick={() => { goBack() }}>
          <ArrowLeft />
        </button>
        <div className={styles.headerContent}>
          <h1 className={styles.mainTitle}>
            {project.Code} - {project.Title}
          </h1>
          <div className={styles.subtitle}>
            {project.Company} • {project.Status}
          </div>
        </div>
      </div>

      <div className={styles.tabs}>

        <div className={styles.tabhead}>
          {/* For Non-CMAP projects, only show Documents and External Portal tabs */}
          {project?.NonCmap ? (
            <>
              <button
                className={`${styles.tab} ${activeTab === "Documents" ? styles.active : ""}`}
                onClick={() => handleTabChange("Documents")}
              >
                Documents
              </button>
              {/* <button
                className={`${styles.tab} ${activeTab === "External Portal" ? styles.active : ""}`}
                onClick={() => handleTabChange("External Portal")}
              >
                External Portal
              </button> */}
            </>
          ) : (
            <>
              {/* For regular CMAP projects, show all three tabs */}
              <button
                className={`${styles.tab} ${activeTab === "Dashboard" ? styles.active : ""}`}
                onClick={() => handleTabChange("Dashboard")}
              >
                Dashboard
              </button>
              <button
                className={`${styles.tab} ${activeTab === "Documents" ? styles.active : ""}`}
                onClick={() => handleTabChange("Documents")}
              >
                Documents
              </button>
              {/* <button
                className={`${styles.tab} ${activeTab === "External Portal" ? styles.active : ""}`}
                onClick={() => handleTabChange("External Portal")}
              >
                External Portal
              </button> */}
            </>
          )}
        </div>
        <div className={styles.addmemberButtonholder}>
          {/* Dashboard Tab: Show only Add Member button */}
          {activeTab === "Dashboard" && canAdd && (
            <button
              className={styles.newButton}
              onClick={() => setIsAddMemberModalOpen(true)}
              type="button"
            >
              + Add Member
            </button>
          )}

          {/* Documents Tab: Show New button and Action dropdown */}
          {/* {activeTab === "Documents" && (
            <>
              <button
                className={styles.newButton}
                onClick={() => ()}
                type="button"
              >
                + New
              </button>
              <button
                className={styles.actionButton}
                onClick={() => { }}
                type="button"
              >
                Actions ▼
              </button>
            </>
          )} */}

          {/* External Portal Tab: No buttons shown */}
        </div>


      </div>

      {activeTab === "Dashboard" && !project?.NonCmap && (
        <>
          {/* <div className={styles.stepper}>
            {STEPS.map((s, i) => {
              const isActive = currentStepIndex >= 0 && i <= currentStepIndex;
              let colorClass = styles.stepGray;

              if (isActive) {
                if (i === 0) colorClass = styles.stepBlue;
                else if (i === 1) colorClass = styles.stepPink;
                else if (i === 2) colorClass = styles.stepOrange;
                else if (i === 3) colorClass = styles.stepGreen;
              }

              const lastClass = i === STEPS.length - 1 ? styles.lastStep : "";
              return (
                <div key={s} className={`${styles.step} ${colorClass} ${lastClass}`}>
                  <span className={styles.stepText}>{s}</span>
                </div>
              );
            })}
          </div> */}

          <div className={styles.mainCardsGrid}>
            <div className={styles.mainCard}>
              <div className={styles.cardHeader}>
                <div>
                  <div className={styles.cardSubtitle}>{project.Company}</div>
                  <h3 className={styles.cardMainTitle}>{project.Title}</h3>
                  <div className={styles.cardMeta}>
                    <span><User size={16} /> {project.Owner}</span>
                    <span><MapPin size={16} /> {project.ContractedEntity}</span>
                    <span><CalendarDays size={16} /> {formatDate(project.EndDate)}</span>
                  </div>
                </div>
                <div className={styles.firstcard}>
                  <span className={styles.badgeLive} >{project.Status}</span>

                  <div className={styles.metricSection}>
                    <div className={styles.metricLabel}>Project Completion</div>
                    <div className={styles.progressBar}>
                      <div className={styles.progressFillRed} style={{ width: `${completion}%` }} />
                      <span className={styles.progressLabel}>{completion}%</span>
                    </div>
                  </div>
                </div>

              </div>

            </div>

            <div className={styles.mainCard}>
              <div className={styles.cardHeader}>
                <div>
                  <h3 className={styles.cardMainTitle}>Financials</h3>
                  <div className={styles.budgetLabel}>Total Budget</div>
                  <div className={styles.budgetAmount}>{formatCurrency(project.TotalProjectValue)}</div>
                </div>
                <span className={styles.badgeSuccess}>On Track</span>
              </div>
              <div className={styles.metricSection}>
                {/* <div className={styles.varianceHeader}>
                  <span className={styles.metricLabel}>Total Budget Variance</span>
                  <span className={styles.varianceAmount}>{formatCurrency(project.BudgetVariance)}</span>
                </div> */}
                {/* <div className={styles.progressBar}>
                  <div className={styles.progressFillGreen} style={{ width: `${completion}%` }} />
                  <span className={styles.progressLabel}>{completion}%</span>
                </div> */}
              </div>
            </div>
          </div>

          <div className={styles.bottomGrid}>
            <div className={styles.panel}>
              <h3 className={styles.panelTitle}>Project Details</h3>
              <div className={styles.updatesGrid}>
                <div className={styles.updateItem}>
                  <div className={styles.updateLabel}>Project Code</div>
                  <div className={styles.updateValue}>{project.Code || "-"}</div>
                </div>
                <div className={styles.updateItem}>
                  <div className={styles.updateLabel}>Status</div>
                  <div className={styles.updateValue}>{project.Status || "-"}</div>
                </div>
                {/* <div className={styles.updateItem}>
                  <div className={styles.updateLabel}>Stage</div>
                  <div className={styles.updateValue}>{project.Stage || "-"}</div>
                </div> */}
                <div className={styles.updateItem}>
                  <div className={styles.updateLabel}>Contract Type</div>
                  <div className={styles.updateValue}>{project.ContractType || "-"}</div>
                </div>
                <div className={styles.updateItem}>
                  <div className={styles.updateLabel}>Country</div>
                  <div className={styles.updateValue}>{project.Country || "-"}</div>
                </div>
                <div className={styles.updateItem}>
                  <div className={styles.updateLabel}>Company</div>
                  <div className={styles.updateValue}>{project.Company || "-"}</div>
                </div>
                <div className={styles.updateItem}>
                  <div className={styles.updateLabel}>Contracted Entity</div>
                  <div className={styles.updateValue}>{project.ContractedEntity || "-"}</div>
                </div>
                <div className={styles.updateItem}>
                  <div className={styles.updateLabel}>Business Unit</div>
                  <div className={styles.updateValue}>{project.BusinessUnit || "-"}</div>
                </div>
                <div className={styles.updateItem}>
                  <div className={styles.updateLabel}>Sector</div>
                  <div className={styles.updateValue}>{project.Sector || "-"}</div>
                </div>
                <div className={styles.updateItem}>
                  <div className={styles.updateLabel}>Start Date</div>
                  <div className={styles.updateValue}>{formatDate(project.StartDate)}</div>
                </div>
                <div className={styles.updateItem}>
                  <div className={styles.updateLabel}>End Date</div>
                  <div className={styles.updateValue}>{formatDate(project.EndDate)}</div>
                </div>
                <div className={styles.updateItem}>
                  <div className={styles.updateLabel}>Percent Completed</div>
                  <div className={styles.updateValue}>{completion}%</div>
                </div>
                {/* <div className={styles.updateItem}>
                  <div className={styles.updateLabel}>Total Project Value</div>
                  <div className={styles.updateValue}>{formatCurrency(project.TotalProjectValue)}</div>
                </div> */}
                <div className={styles.updateItem}>
                  <div className={styles.updateLabel}>Estimated Price</div>
                  <div className={styles.updateValue}>{formatCurrency(project.EstimatedPrice)}</div>
                </div>
                <div className={styles.updateItem}>
                  <div className={styles.updateLabel}>Fee Type</div>
                  <div className={styles.updateValue}>{project.FeeType || "-"}</div>
                </div>
                <div className={styles.updateItem}>
                  <div className={styles.updateLabel}>Lead Source</div>
                  <div className={styles.updateValue}>{project.LeadSource || "-"}</div>
                </div>
              </div>
            </div>

            {/* Team Members */}
            <div className={styles.panel}>
              <div className={styles.tabsRow}>
                <h3 className={styles.panelTitle}>Team Members</h3>

              </div>

              {isLoadingMembers ? (
                <div className={styles.noData}>Loading team members...</div>
              ) : (
                <div className={styles.teamList}>
                  {/* Owner - Only show if assigned */}
                  {getPersonName(project.Owner) && (
                    <div className={styles.teamMemberRow}>
                      <div className={styles.memberAvatarCircle}>
                        {getInitials(project.Owner)}
                      </div>
                      <div className={styles.memberInfo}>
                        <div className={styles.memberNameText}>{getPersonName(project.Owner)}</div>
                        <div className={styles.memberRoleText}>Owner</div>
                      </div>
                      <div className={styles.memberActions}>
                        {project.OwnerEmail && (
                          <a
                            href={`mailto:${project.OwnerEmail}`}
                            className={styles.iconButton}
                            title="Send Email"
                          >
                            <svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
                              <path d="M3 4h14a1 1 0 0 1 1 1v10a1 1 0 0 1-1 1H3a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1z" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                              <path d="m2 5 8 6 8-6" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                            </svg>
                          </a>
                        )}
                      </div>
                    </div>
                  )}

                  {/* Project Manager - Only show if assigned */}
                  {getPersonName(project.ProjectManager) && (
                    <div className={styles.teamMemberRow}>
                      <div className={styles.memberAvatarCircle}>
                        {getInitials(project.ProjectManager)}
                      </div>
                      <div className={styles.memberInfo}>
                        <div className={styles.memberNameText}>{getPersonName(project.ProjectManager)}</div>
                        <div className={styles.memberRoleText}>Project Manager</div>
                      </div>
                      <div className={styles.memberActions}>
                        {project.ProjectManagerEmail && (
                          <a
                            href={`mailto:${project.ProjectManagerEmail}`}
                            className={styles.iconButton}
                            title="Send Email"
                          >
                            <svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
                              <path d="M3 4h14a1 1 0 0 1 1 1v10a1 1 0 0 1-1 1H3a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1z" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                              <path d="m2 5 8 6 8-6" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                            </svg>
                          </a>
                        )}
                      </div>
                    </div>
                  )}

                  {/* Dynamic Team Members from SharePoint */}
                  {teamMembers.map((member) => (
                    <div key={member.Id} className={styles.teamMemberRow}>
                      <div className={styles.memberAvatarCircle}>
                        {member.Name.split(" ").map(n => n[0]).join("").substring(0, 2).toUpperCase()}
                      </div>
                      <div className={styles.memberInfo}>
                        <div className={styles.memberNameText}>{member.Name}</div>
                        <div className={styles.memberRoleText}>{member.Role}</div>
                      </div>
                      <div className={styles.memberActions}>
                        <a
                          href={`mailto:${member.EmailAddress}`}
                          className={styles.iconButton}
                          title="Send Email"
                        >
                          <svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
                            <path d="M3 4h14a1 1 0 0 1 1 1v10a1 1 0 0 1-1 1H3a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1z" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                            <path d="m2 5 8 6 8-6" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                          </svg>
                        </a>
                        {member.MobileNumber && (
                          <button
                            className={styles.iconButton}
                            title="Copy Mobile Number"
                            onClick={() => handleCopyMobile(member.Id!, member.MobileNumber!)}
                          >
                            <svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
                              <path d="M18 13.5v3a2 2 0 0 1-2.18 2A19.79 19.79 0 0 1 7 13.65a19.5 19.5 0 0 1-5-5.81A19.79 19.79 0 0 1 2 1.18 2 2 0 0 1 4 .5h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L8.09 8.29a16 16 0 0 0 5.81 5.81l1.15-1.15a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 18 13.5z" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                            </svg>
                          </button>
                        )}
                      </div>
                      {copiedMemberId === member.Id && (
                        <div className={styles.memberCopiedToast}>
                          Copied!
                        </div>
                      )}
                    </div>
                  ))}

                  {teamMembers.length === 0 && !getPersonName(project.Owner) && !getPersonName(project.ProjectManager) && (
                    <div className={styles.noData}>
                      No team members added yet. Click "Add Member" to get started.
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        </>
      )}

      {activeTab === "Documents" && (
        <CustomDocumentsList
          projectId={project.Id || Number(project.ProjectId) || 0}
          businessProjectId={project.ProjectID}
          projectCode={project.Code || project.ProjectId || ""}
          projectTitle={project.Title}
          context={props.context}
          onShowShareAccess={handleShowShareAccess}
          isUserRestricted={isUserRestricted}
          isNonCmap={project?.NonCmap || false}
          location={project?.Location || ""}
          uncategorisedDocumentsUrl={project?.UncategorisedDocumentsUrl || ""}
        />
      )}

      {/* {activeTab === "Activities" && (
        <div className={styles.comingSoon}>Activities - In-progress</div>
      )} */}

      {/* {activeTab === "External Portal" && (
        <div className={portalStyles.externalPortal}>
          <div className={portalStyles.content}>
            <SharedFiles
              ref={sharedFilesRef}
              context={props.context}
              onShowNewDocument={handleShowNewDocument}
              onShowShareAccess={handleShowShareAccess}
              onShowImportDialog={handleShowImportDialog}
              currentPath={currentFolderPath}
              breadcrumbs={breadcrumbs}
              onFolderNavigate={(path, crumbs) => {
                setCurrentFolderPath(path);
                setBreadcrumbs(crumbs);
              }}
              onRefresh={handleOperationSuccess}
              service={portalService}
              isUserRestricted={isUserRestricted}
              onFolderChange={(path) => setCurrentFolderPath(path)}
            />
          </div>
        </div>
      )} */}

      {/* Add Team Member Modal */}
      <AddTeamMemberModal
        isOpen={isAddMemberModalOpen}
        onClose={() => setIsAddMemberModalOpen(false)}
        onAddMember={handleAddMember}
        projectId={project?.Id || project?.ProjectId || null}
        projectCode={project.Code}
        ProjectTitle={project.Title}
      />

      {/* External Portal Dialogs */}
      {showInviteGuestDialog && (
        <InviteGuestDialog
          isOpen={showInviteGuestDialog}
          onClose={() => setShowInviteGuestDialog(false)}
          context={props.context}
          projectId={String(project?.Id || project?.ProjectId || "")}
          onGuestInvited={handleGuestInvited}
        />
      )}

      {showShareDocumentDialog && (
        <ShareDocumentDialog
          isOpen={showShareDocumentDialog}
          onClose={() => setShowShareDocumentDialog(false)}
          context={props.context}
          currentPath={currentFolderPath}
        />
      )}

      <NewDocumentDialog
        isOpen={showNewDocumentDialog}
        onClose={() => setShowNewDocumentDialog(false)}
        onSuccess={() => {
          setShowNewDocumentDialog(false);
          handleOperationSuccess();
        }}
        context={props.context}
        currentPath={currentFolderPath}
        service={portalService}
      />

      <CustomManageAccess
        isOpen={showShareAccessDialog}
        onDismiss={() => setShowShareAccessDialog(false)}
        item={selectedDocumentsForSharing.length > 0 ? {
          ...selectedDocumentsForSharing[0],
          Name: selectedDocumentsForSharing[0].name,
          ServerRelativeUrl: selectedDocumentsForSharing[0].fileRef,
          IsFolder: selectedDocumentsForSharing[0].isFolder
        } : null}
        context={props.context}
        siteUrl={props.context.pageContext.web.absoluteUrl}
      />

      <ImportFromProjectDocsDialog
        isOpen={showImportDialog}
        onClose={() => setShowImportDialog(false)}
        onSuccess={() => {
          // Refresh the SharedFiles list in background while dialog shows success message
          handleOperationSuccess();
        }}
        context={props.context}
        targetPath={`ExternalShareDocument/${currentFolderPath}`}
        service={portalService}
        projectCode={project?.Code || project?.ProjectId || ""}
        projectId={project?.Id || Number(project?.ProjectId) || 0}
        businessProjectId={project?.ProjectID}
        projectTitle={project?.Title}
      />
    </div>
  );
};

export default ProjectDetails;