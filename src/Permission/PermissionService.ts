import { getSP, getContext } from "../pnpjsConfig";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { usePermissionStore } from "./PermissionStore";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export const OWNER_GROUP = "Projects Owners";
export const CONTRIBUTOR_GROUP = "Projects Members";
export const VIEWER_GROUP = "Projects Visitors";
export const AAD_VIEWER_GROUP_ID = "4fd8e847-82b2-4cbc-b376-d4a36abd5316";

// 🔹 Check if AAD group is inside SP group
const isAADGroupInSPGroup = async (sp: any, groupName: string): Promise<boolean> => {
    try {
        const users = await sp.web.siteGroups.getByName(groupName).users();
        console.log(`[DEBUG] Members of SP Group "${groupName}":`, users.map((u: any) => u.LoginName));
 
        const exists = users.some((u: any) =>
            u.LoginName &&
            u.LoginName.toLowerCase().includes(AAD_VIEWER_GROUP_ID.toLowerCase())
        );
 
        if (exists) {
            console.log(`✅ [NESTING] AAD Group (${AAD_VIEWER_GROUP_ID}) is NESTED inside SP Group: "${groupName}"`);
        } else {
            console.log(`❌ [NESTING] AAD Group (${AAD_VIEWER_GROUP_ID}) NOT FOUND in SP Group: "${groupName}"`);
        }
 
        return exists;
    } catch (err) {
        console.error(`[Permission] Error reading ${groupName}:`, err);
        return false;
    }
};
 
// 🔹 Robust AAD Membership (handles pagination)
const checkAADMembership = async (context: WebPartContext): Promise<boolean> => {
    try {
        const client = await context.msGraphClientFactory.getClient("3");
        let url = "/me/memberOf";
        let isMember = false;
 
        while (url) {
            const res = await client.api(url).get();
            const groups = res.value || [];
 
            if (groups.some((g: any) => g.id === AAD_VIEWER_GROUP_ID)) {
                isMember = true;
                break;
            }
            url = res["@odata.nextLink"] || null;
        }
 
        if (isMember) {
            console.log(`✅ [AAD] Current user IS a member of AAD Group: ${AAD_VIEWER_GROUP_ID}`);
        } else {
            console.log(`❌ [AAD] Current user IS NOT a member of AAD Group: ${AAD_VIEWER_GROUP_ID}`);
        }
 
        return isMember;
    } catch (err) {
        console.error("[Permission] AAD check failed:", err);
        return false;
    }
};
 
export const checkPermissions = async (context?: WebPartContext): Promise<void> => {
    const ctx = context || getContext();
    if (!ctx) {
        console.error("[PermissionService] No context available");
        return;
    }
 
    const store = usePermissionStore.getState();
    store.setPermissions({ isLoading: true });
 
    try {
        const sp = getSP(ctx);
 
        // 🔹 1. Direct SharePoint membership
        const [currentUser, groups] = await Promise.all([
            sp.web.currentUser(),
            sp.web.currentUser.groups()
        ]);
 
        const groupNames = groups.map(g => g.Title.toLowerCase());
        const isSiteAdmin = currentUser.IsSiteAdmin;
 
        console.log(`[USER GROUPS] ${groups.map(g => g.Title).join(", ")}`);
        if (isSiteAdmin) console.log("👑 [USER] IS SITE ADMIN");
 
        // ✅ SP ROLE
        let spRole: "Owner" | "Contributor" | "Viewer" | "None" = "None";
 
        if (isSiteAdmin || groupNames.includes(OWNER_GROUP.toLowerCase())) {
            spRole = "Owner";
        } else if (groupNames.includes(CONTRIBUTOR_GROUP.toLowerCase())) {
            spRole = "Contributor";
        } else if (groupNames.includes(VIEWER_GROUP.toLowerCase())) {
            spRole = "Viewer";
        }
 
        console.log("[SP ROLE]", spRole);
 
        // 🔹 2. AAD → SP group mapping (Check Nesting First)
        const [inOwner, inContributor, inViewer] = await Promise.all([
            isAADGroupInSPGroup(sp, OWNER_GROUP),
            isAADGroupInSPGroup(sp, CONTRIBUTOR_GROUP),
            isAADGroupInSPGroup(sp, VIEWER_GROUP)
        ]);
 
        const hasNesting = inOwner || inContributor || inViewer;
        let aadRole: "Owner" | "Contributor" | "Viewer" | "None" = "None";
 
        if (hasNesting) {
            // Only check AAD membership if nesting exists in SP groups
            const isAADMember = await checkAADMembership(ctx);
            if (isAADMember) {
                if (inOwner) {
                    aadRole = "Owner";
                    console.log(`💡 [INFO] User has AAD-based "Owner" role via nesting in: "${OWNER_GROUP}"`);
                } else if (inContributor) {
                    aadRole = "Contributor";
                    console.log(`💡 [INFO] User has AAD-based "Contributor" role via nesting in: "${CONTRIBUTOR_GROUP}"`);
                } else if (inViewer) {
                    aadRole = "Viewer";
                    console.log(`💡 [INFO] User has AAD-based "Viewer" role via nesting in: "${VIEWER_GROUP}"`);
                }
            }
        }
 
        console.log("[AAD ROLE]", aadRole);
 
        // 🔹 3. FINAL ROLE (priority)
        const priority = ["Owner", "Contributor", "Viewer", "None"];
 
        const role =
            priority.indexOf(aadRole) < priority.indexOf(spRole)
                ? aadRole
                : spRole;
 
        console.log("[FINAL ROLE]", role);
 
        // 🔹 4. Permissions
        const isOwner = role === "Owner";
        const isContributor = role === "Contributor" || isOwner;
        const isViewer = role === "Viewer" || isContributor;
 
        const permissions = {
            role,
            canAdd: isContributor,
            canEdit: isContributor,
            canDelete: isOwner,
            canView: isViewer,
            isOwner,
            isContributor,
            isViewer,
            isLoading: false
               
        };
 
        console.log("🔓 [FINAL PERMISSIONS]", permissions);
        store.setPermissions(permissions);
 
    } catch (err) {
        console.error("[PermissionService] Error:", err);
        store.setPermissions({ isLoading: false });
    }
}; 
 
