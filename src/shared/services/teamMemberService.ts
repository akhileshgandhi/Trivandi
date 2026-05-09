/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { SPFI } from "@pnp/sp/presets/all";

let sp: SPFI | null = null;

export interface ITeamMember {
  Id?: number;
  Title: string;
  Name: string;
  EmailAddress: string;
  MobileNumber?: string;
  Role: string;
  ProjectId?: string;
  ProjectIdId?: number;
}

/**
 * Initialize the team member service with PnP instance
 */
export const initTeamMemberService = (spInstance: SPFI): void => {
  sp = spInstance;
};

/**
 * Get all team members for a specific project
 */
export const getTeamMembersByProjectId = async (projectId: number): Promise<ITeamMember[]> => {
  if (!sp) throw new Error("PnPjs not initialized. Call initTeamMemberService first.");

  try {
    const items = await sp.web.lists
      .getByTitle("TeamMember")
      .items
      .select("Id", "Title", "Name", "EmailAddress", "MobileNumber", "Role", "ProjectId/Id", "ProjectId/Title")
      .expand("ProjectId")
      .filter(`ProjectId/Id eq ${projectId}`)
      .orderBy("Created", false)
      ();

    return items.map((item: any) => ({
      Id: item.Id,
      Title: item.Title || item.Name,
      Name: item.Name,
      EmailAddress: item.EmailAddress,
      MobileNumber: item.MobileNumber,
      Role: item.Role,
      ProjectIdId: item.ProjectId?.Id,
    }));
  } catch (error) {
    
    throw error;
  }
};

/**
 * Add a new team member to the project
 */
export const addTeamMember = async (
  projectId: number,
  member: { name: string; email: string; role: string; mobileNumber?: string }
): Promise<ITeamMember> => {
  if (!sp) throw new Error("PnPjs not initialized. Call initTeamMemberService first.");

  try {
    const itemData: any = {
      Title: member.name,
      Name: member.name,
      EmailAddress: member.email,
      MobileNumber: member.mobileNumber || "",
      Role: member.role,
      ProjectIdId: projectId, // Lookup field requires Id suffix
    };

    const result = await sp.web.lists
      .getByTitle("TeamMember")
      .items
      .add(itemData);

    

    return {
      Id: result.data.Id,
      Title: member.name,
      Name: member.name,
      EmailAddress: member.email,
      MobileNumber: member.mobileNumber,
      Role: member.role,
      ProjectIdId: projectId,
    };
  } catch (error) {
    
    throw new Error("Failed to add team member. Please try again.");
  }
};

/**
 * Update an existing team member
 */
export const updateTeamMember = async (
  itemId: number,
  updates: Partial<{ name: string; email: string; role: string }>
): Promise<void> => {
  if (!sp) throw new Error("PnPjs not initialized. Call initTeamMemberService first.");

  try {
    const itemData: any = {};
    if (updates.name) {
      itemData.Title = updates.name;
      itemData.Name = updates.name;
    }
    if (updates.email) itemData.EmailAddress = updates.email;
    if (updates.role) itemData.Role = updates.role;

    await sp.web.lists
      .getByTitle("TeamMember")
      .items
      .getById(itemId)
      .update(itemData);

    
  } catch (error) {
    
    throw new Error("Failed to update team member. Please try again.");
  }
};

/**
 * Delete a team member
 */
export const deleteTeamMember = async (itemId: number): Promise<void> => {
  if (!sp) throw new Error("PnPjs not initialized. Call initTeamMemberService first.");

  try {
    await sp.web.lists
      .getByTitle("TeamMember")
      .items
      .getById(itemId)
      .delete();

    
  } catch (error) {
    
    throw new Error("Failed to delete team member. Please try again.");
  }
};

/**
 * Send invitation email to team member (simulated - you may need to implement actual email sending)
 */
export const sendInvitation = async (
  memberEmail: string,
  memberName: string,
  projectTitle: string,
  projectId: string
): Promise<void> => {
  try {
    // In a real implementation, you would call a Flow/Power Automate endpoint
    // or use SharePoint's email capabilities
    
    
    
    

    // Simulate email sending
    // You can implement actual email sending using:
    // 1. Microsoft Graph API
    // 2. Power Automate Flow trigger
    // 3. SharePoint REST API email endpoint

    // For now, just log and return success
    return Promise.resolve();
  } catch (error) {
    
    throw new Error("Failed to send invitation email.");
  }
};

/**
 * Ensure TeamMember list exists with proper columns
 */
export const ensureTeamMemberList = async (): Promise<void> => {
  if (!sp) throw new Error("PnPjs not initialized. Call initTeamMemberService first.");

  try {
    // Check if list exists
    const lists = await sp.web.lists();
    const teamMemberList = lists.find((l: any) => l.Title === "TeamMember");

    if (!teamMemberList) {
      
      
      // Create the list
      await sp.web.lists.add("TeamMember", "List to store project team members", 100, false);

      // Add custom columns
      const list = sp.web.lists.getByTitle("TeamMember");
      
      // Add Name column (Single line of text)
      await list.fields.addText("Name", { MaxLength: 255 });

      // Add EmailAddress column (Single line of text)
      await list.fields.addText("EmailAddress", { MaxLength: 255 });

      // Add Role column (Choice)
      await list.fields.addChoice("Role", {
        Choices: [
          "Contributor",
          "Reader",
          "Administrator",
          "Project Manager",
          "Team Member"
        ],
      });

      // Add ProjectId lookup column (requires Projects list to exist)
      await list.fields.addLookup("ProjectId", {
        LookupListId: (await sp.web.lists.getByTitle("ProjectsNew").select("Id")()).Id,
        LookupFieldName: "Title",
      });

      
    } else {
      
    }
  } catch (error) {
    
    // Don't throw - list might exist already
  }
};
