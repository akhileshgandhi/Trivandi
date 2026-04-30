/* eslint-disable @typescript-eslint/no-explicit-any */
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";

let sp: SPFI;

export const initDashboardService = (spInstance: SPFI): void => {
  sp = spInstance;
};

// List names - customize these to match your SharePoint lists
const LISTS = {
  PROJECTS: "ProjectsNew",
  COMPLIANCE: "StandardsandCompliance",
  KEY_PEOPLE: "KeyPeople",
  EVENTS: "UpComingEvents",
  TOOLS: "KeyTools",
  QUICK_LINKS: "QuickLinks"
};

// Get project statistics
export const getProjectStats = async (): Promise<any> => {
  try {
    // Get current user email
    const currentUser = await sp.web.currentUser();
    // const userEmail =  "petergriffin@trivandi.com";//currentUser.Email ||
 const userEmail = currentUser.Email;
    // Get all projects
    const allProjects = await sp.web.lists
      .getByTitle(LISTS.PROJECTS)
      .items
      .select("Id", "Title", "Status", "OwnerEmail","NonCmap")
      .getAll();
   console.log(allProjects,'getProjectStats>>>');
    
    // Calculate stats
    const totalProjects = allProjects.length;
    
    // Status == "Live" means live projects
    const liveProjects = allProjects.filter(p => p.Status === "Project"&& p.NonCmap!=true).length;
    console.log(liveProjects,'liveProjects');
    
    // Status == "Potential" means bids
    const myBids = allProjects.filter(p => p.Status === "Potential").length;
    
    // My projects: Status == "Live" AND OwnerEmail matches current user
    const myProjects = allProjects.filter(p => 
      p.Status === "Project" && p.OwnerEmail === userEmail
    ).length;

    return {
      totalProjects,
      liveProjects,
      myProjects,
      myBids
    };
  } catch (err) {
    console.error("Error fetching project stats:", err);
    // Return default values if lists don't exist
    return {
      totalProjects: 0,
      liveProjects: 0,
      myProjects: 0,
      myBids: 0
    };
  }
};

// Get compliance items
export const getComplianceItems = async (): Promise<any[]> => {
  try {
    const items = await sp.web.lists
      .getByTitle(LISTS.COMPLIANCE)
      .items
      .select("Id", "Title", "Description", "Color", "Icon", "Links")
      .top(4)
      .getAll();
   console.log(items,'getComplianceItems>>>');
    return items.map(item => {
      let iconUrl = "";
      if (item.Icon) {
        try {
          const iconData = typeof item.Icon === 'string' ? JSON.parse(item.Icon) : item.Icon;
          if (iconData.serverRelativeUrl && iconData.serverUrl) {
            iconUrl = `${iconData.serverUrl}${iconData.serverRelativeUrl}`;
          }
        } catch (e) {
          console.warn("Error parsing icon data:", e);
        }
      }
      let linkUrl = "";
      if (item.Links) {
        try {
          const linkData = typeof item.Links === 'string' ? JSON.parse(item.Links) : item.Links;
          linkUrl = linkData.Url || linkData.url || "";
        } catch (e) {
          console.warn("Error parsing link data:", e);
        }
      }
      return {
        id: item.Id,
        title: item.Title,
        description: item.Description || "Procedures & terms...",
        color: item.Color || "#FF8C42",
        icon: iconUrl,
        link: item.Links ? linkUrl : ""
      };
    });
  } catch (error) {
    console.error("Error fetching compliance items:", error);
    // Return sample data if list doesn't exist
    return [
      {
        id: 1,
        title: "Associate Onboarding",
        description: "Procedures & terms...",
        color: "#4A5AFF",
        icon: require("../assets/firstFrame.png")
      },
      {
        id: 2,
        title: "Mobilisations",
        description: "Procedures & terms...",
        color: "#FF8C42",
        icon: require("../assets/secFrame.png")
      },
      {
        id: 3,
        title: "Health & Safety",
        description: "Procedures & forms...",
        color: "#FF4DB8",
        icon: require("../assets/thirdFrame.png")
      },
      {
        id: 4,
        title: "BCS Standards",
        description: "Procedures & forms...",
        color: "#4A5AFF",
        icon: require("../assets/forthFrame.png")
      }
    ];
  }
};

// Get key people
export const getKeyPeople = async (): Promise<any[]> => {
  try {
    const items = await sp.web.lists
      .getByTitle(LISTS.KEY_PEOPLE)
      .items
      .select("Id", "Title", "JobTitle", "Email", "WorkPhone", "Picture")
      .top(8)
      .getAll();
   console.log(items,'itemskeypeople>>>');
    return items.map(item => {
      let imageUrl = "";
      if (item.Picture) {
        try {
          const pictureData = typeof item.Picture === 'string' ? JSON.parse(item.Picture) : item.Picture;
          if (pictureData.serverRelativeUrl && pictureData.serverUrl) {
            imageUrl = `${pictureData.serverUrl}${pictureData.serverRelativeUrl}`;
          } else if (pictureData.Url) {
            imageUrl = pictureData.Url;
          }
        } catch (e) {
          // If Picture is already a URL string
          imageUrl = item.Picture;
        }
      }
      return {
        id: item.Id,
        name: item.Title,
        title: item.JobTitle || "Project Head",
        imageUrl: imageUrl,
        email: item.Email || "",
        phone: item.WorkPhone || ""
      };
    });
  } catch (error) {
    console.error("Error fetching key people:", error);
    // Return sample data if list doesn't exist
    return [
      { id: 1, name: "Darrell Steward", title: "Project Head", imageUrl: "", email: "", phone: "" },
      { id: 2, name: "Ronald Richards", title: "Project Head", imageUrl: "", email: "", phone: "" },
      { id: 3, name: "Jane Flores", title: "Project Head", imageUrl: "", email: "", phone: "" },
      { id: 4, name: "Darrell Howard", title: "Project Head", imageUrl: "", email: "", phone: "" },
      { id: 5, name: "Ronald Richards", title: "Project Head", imageUrl: "", email: "", phone: "" },
      { id: 6, name: "Albert Flores", title: "Project Head", imageUrl: "", email: "", phone: "" },
      { id: 7, name: "Ronald Richards", title: "Project Head", imageUrl: "", email: "", phone: "" },
      { id: 8, name: "Albert Flores", title: "Project Head", imageUrl: "", email: "", phone: "" }
    ];
  }
};

// Get upcoming events
export const getUpcomingEvents = async (): Promise<any[]> => {
  try {
    const today = new Date();
    const items = await sp.web.lists
      .getByTitle(LISTS.EVENTS)
      .items
      .select("Id", "Title", "EventDate", "EndDate")
      .filter(`EventDate ge datetime'${today.toISOString()}'`)
      .orderBy("EventDate", true)
      .top(5)
      .getAll();

      console.log(items,'getUpcomingEvents>>>');
      

    return items.map(item => {
      const eventDate = new Date(item.EventDate);
      const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      const days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

      return {
        id: item.Id,
        title: item.Title,
        date: `${days[eventDate.getDay()]}, ${months[eventDate.getMonth()]} ${eventDate.getDate()}, ${eventDate.getHours() > 12 ? eventDate.getHours() - 12 : eventDate.getHours()}:${String(eventDate.getMinutes()).padStart(2, '0')} ${eventDate.getHours() >= 12 ? 'PM' : 'AM'}`,
        time: eventDate.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' }),
        day: eventDate.getDate().toString(),
        month: months[eventDate.getMonth()]
      };
    });
  } catch (error) {
    console.error("Error fetching upcoming events:", error);
    // Return sample data if list doesn't exist
    return [
      { id: 1, title: "New Partnership Announcement", date: "Monday, Nov 4, 2024", time: "10:00 AM", day: "4", month: "Nov" },
      { id: 2, title: "Town Hall Meeting", date: "Sunday, Nov 3, 2024", time: "12:00 PM", day: "3", month: "Nov" },
      { id: 3, title: "Timesheet Deadline", date: "Sunday, Nov 3, 2024", time: "11:59 PM", day: "3", month: "Nov" },
      { id: 4, title: "Webinar Series Launch", date: "Tuesday, Nov 12, 2024", time: "9:00 PM", day: "12", month: "Nov" },
      { id: 5, title: "Project Review", date: "Sunday, Dec 22, 2024", time: "12:03 PM", day: "22", month: "Dec" }
    ];
  }
};

// Get key tools
export const getKeyTools = async (): Promise<any[]> => {
  try {
    const items = await sp.web.lists
      .getByTitle(LISTS.TOOLS)
      .items
      .select("Id", "Title", "Description", "Color", "Icon","Link")
      .getAll();
  console.log(items,'getKeyTools>>>');
    return items.map(item => {
      let iconUrl = "";
      if (item.Icon) {
        try {
          const iconData = typeof item.Icon === 'string' ? JSON.parse(item.Icon) : item.Icon;
          if (iconData.serverRelativeUrl && iconData.serverUrl) {
            iconUrl = `${iconData.serverUrl}${iconData.serverRelativeUrl}`;
          }
        } catch (e) {
          console.warn("Error parsing icon data:", e);
        }
      }
      return {
        id: item.Id,
        title: item.Title,
        description: item.Description || "",
        color: item.Color || "#4A5AFF",
        icon: iconUrl,
        link:item.Link
      };
    });
  } catch (error) {
    console.error("Error fetching key tools:", error);
    // Return sample data if list doesn't exist
    return [
      { id: 1, title: "CMAP", description: "Project Management", color: "#FF8C42", icon: require("../assets/secFrame.png") },
      { id: 2, title: "Timesheets", description: "Log Hours", color: "#4A5AFF", icon: require("../assets/firstFrame.png") },
      { id: 3, title: "Expenses", description: "Submit Claims", color: "#4A9EFF", icon: require("../assets/forthFrame.png") },
      { id: 4, title: "BenefitX HR", description: "Imagine Benefit...", color: "#FF4DB8", icon: require("../assets/thirdFrame.png") }
    ];
  }
};

// Get child items by parent name
export const getChildLinksByParent = async (parentName: string): Promise<Array<{ name: string; url: string }>> => {
  try {
    const items = await sp.web.lists
      .getByTitle(LISTS.QUICK_LINKS)
      .items
      .select("Id", "Title", "URL", "ParentName")
      .filter(`ParentName eq '${parentName}'`)
      .orderBy("SortOrder", true)();

    return items.map((item: any) => ({
      name: item.Title,
      url: item.URL?.Url || item.URL || "#"
    }));
  } catch (error) {
    console.error(`Error fetching child links for parent "${parentName}":`, error);
    return [];
  }
};

// Get quick links
export const getQuickLinks = async (): Promise<any[]> => {
  try {
    const items = await sp.web.lists
      .getByTitle(LISTS.QUICK_LINKS)
      .items
      .select("Id", "Title", "URL", "Icon", "BackgroundColor", "ParentName", "SortOrder")
      .filter("ParentName eq null or ParentName eq ''")
      .orderBy("SortOrder", true)();
      // .top(10)();

    // For each parent item, check if it has children
    const itemsWithChildren = await Promise.all(items.map(async (item: any) => {
      let iconUrl = require("../assets/firstFrame.png"); // default icon
      
      if (item.Icon) {
        try {
          const iconData = JSON.parse(item.Icon);
          if (iconData.serverUrl && iconData.serverRelativeUrl) {
            iconUrl = `${iconData.serverUrl}${iconData.serverRelativeUrl}`;
          }
        } catch (e) {
          console.warn("Error parsing icon data:", e);
        }
      }

      // Check if this parent has any children
      let hasChildren = false;
      try {
        const childItems = await sp.web.lists
          .getByTitle(LISTS.QUICK_LINKS)
          .items
          .select("Id")
          .filter(`ParentName eq '${item.Title}'`).orderBy("SortOrder", true)
          .top(1)();
        hasChildren = childItems.length > 0;
      } catch (e) {
        console.warn(`Error checking children for "${item.Title}":`, e);
      }

      return {
        id: item.Id,
        title: item.Title,
        url: item.URL?.Url || item.URL || "#",
        backgroundColor: item.BackgroundColor || "#E8EEFF",
        icon: iconUrl,
        hasChildren: hasChildren
      };
    }));

    return itemsWithChildren;
  } catch (error) {
    console.error("Error fetching quick links:", error);
    // Return sample data if list doesn't exist
    return [
      { 
        id: 1, 
        title: "RFP Resources", 
        url: "#", 
        backgroundColor: "#E8EEFF", 
        icon: require("../assets/firstFrame.png"),
        hasChildren: false
      },
      { 
        id: 2, 
        title: "Bid Templates", 
        url: "#", 
        backgroundColor: "#F3E8FF", 
        icon: require("../assets/secFrame.png"),
        hasChildren: true
      },
      { 
        id: 3, 
        title: "UK Projects", 
        url: "#", 
        backgroundColor: "#E0F7FA", 
        icon: require("../assets/thirdFrame.png"),
        hasChildren: false
      },
      { 
        id: 4, 
        title: "Team Resources", 
        url: "#", 
        backgroundColor: "#FFF8E1", 
        icon: require("../assets/forthFrame.png"),
        hasChildren: false
      }
    ];
  }
};

