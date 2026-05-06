import { getSP, getContext } from "../pnpjsConfig";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { usePermissionStore } from "./PermissionStore";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export const OWNER_GROUP = "Projects Owners";
export const CONTRIBUTOR_GROUP = "Projects Members";
export const VIEWER_GROUP = "Projects Visitors";
export const AAD_VIEWER_GROUP_ID = "4fd8e847-82b2-4cbc-b376-d4a36abd5316";

export const checkPermissions = async (context?: WebPartContext): Promise<void> => {
    const activeContext = context || getContext();
    if (!activeContext) {
        console.error("[PermissionService] No context provided or found in pnpjsConfig.");
        return;
    }
    const store = usePermissionStore.getState();
    store.setPermissions({ isLoading: true });

    try {
        const sp = getSP();
        
        // 1. SharePoint Permissions Check
        const [currentUser, groups] = await Promise.all([
            sp.web.currentUser(),
            sp.web.currentUser.groups()
        ]);
        
        const groupNames = groups.map(g => g.Title.toLowerCase());
        const isSiteAdmin = currentUser.IsSiteAdmin;
        
        const isOwner = groupNames.indexOf(OWNER_GROUP.toLowerCase()) > -1 || isSiteAdmin;
        const isContributor = groupNames.indexOf(CONTRIBUTOR_GROUP.toLowerCase()) > -1 || isOwner;
        const isSPViewer = groupNames.indexOf(VIEWER_GROUP.toLowerCase()) > -1 || isContributor;

        // 2. AAD Group Check (for specific View permission)
        let isAADMember = false;
        try {
            const client = await activeContext.msGraphClientFactory.getClient('3');
            const response = await client.api('/me/memberOf').get();
            const aadGroups = response.value;
            isAADMember = aadGroups.some((group: any) => group.id === AAD_VIEWER_GROUP_ID);
            console.log(`[Permission] AAD Group Check (${AAD_VIEWER_GROUP_ID}): ${isAADMember}`);
        } catch (aadErr) {
            console.error("[Permission] Error checking AAD group:", aadErr);
        }

        const isViewer = isSPViewer || isAADMember;

        console.log(`[Permission] Groups: ${groups.map(g => g.Title).join(", ")}`);
        console.log(`[Permission] Access -> View: ${isViewer}, Add: ${isContributor}, Edit: ${isContributor}, Delete: ${isOwner}`);

        let role: 'Owner' | 'Contributor' | 'Viewer' | 'None' = 'None';
        if (isOwner) role = 'Owner';
        else if (isContributor) role = 'Contributor';
        else if (isViewer) role = 'Viewer';

        store.setPermissions({
            role,
            canAdd: isContributor,
            canEdit: isContributor,
            canDelete: isOwner,
            canView: isViewer,
            isOwner,
            isContributor,
            isViewer,
            isLoading: false
        });
    } catch (err) {
        console.error("[PermissionService] Error checking permissions:", err);
        store.setPermissions({ isLoading: false });
    }
};
