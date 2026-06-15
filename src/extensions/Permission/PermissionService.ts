import { getSP } from "./pnpjsConfig";
import { spfi, SPFx } from "@pnp/sp";

import "@pnp/sp/webs";

import "@pnp/sp/site-users/web";

import "@pnp/sp/site-groups/web";

import { usePermissionStore } from "./PermissionStore";




// step id comapny ,project ,maindashboard


export const OWNER_GROUP = ["CompanyHub Owners", "Projects Owners","TrivandiHub Owners", "Branding & Marketing Owners","PeopleHub Owners"];

export const CONTRIBUTOR_GROUP = ["CompanyHub_Contributor", "Projects Members","TrivandiHub_Contributer","Marketing_Contributor","PeopleHub_Contributor"];

export const VIEWER_GROUP = ["CompanyHub_Viewer", "Projects Visitors","TrivandiHub_Visitor","Marketing_Viewer","PeopleHub_Viewer"];

export const AAD_VIEWER_GROUP_ID = ["7917dfd4-64c9-4202-aa87-9f89822f8f5d","4fd8e847-82b2-4cbc-b376-d4a36abd5316","8c393c7c-587a-48ab-944a-df35b012459e","33e5005e-aa10-4ebb-98cb-85e94d46919d","21448698-f734-4f99-b8a3-e64c8a5de658"];

const SITE_PERMISSIONS_CONFIG = [
  {
    siteKey: "CompanyHub",
    siteUrlPart: "CompanyHub",
    ownerGroup: "CompanyHub Owners",
    contributorGroup: "CompanyHub_Contributor",
    viewerGroup: "CompanyHub_Viewer",
    aadGroupId: "7917dfd4-64c9-4202-aa87-9f89822f8f5d"
  },
  {
    siteKey: "Projects",
    siteUrlPart: "Projects",
    ownerGroup: "Projects Owners",
    contributorGroup: "Projects Members",
    viewerGroup: "Projects Visitors",
    aadGroupId: "4fd8e847-82b2-4cbc-b376-d4a36abd5316"
  },
  {
    siteKey: "TrivandiHub",
    siteUrlPart: "TrivandiHub",
    ownerGroup: "TrivandiHub Owners",
    contributorGroup: "TrivandiHub_Contributer",
    viewerGroup: "TrivandiHub_Visitor",
    aadGroupId: "8c393c7c-587a-48ab-944a-df35b012459e"
  },
  {
    siteKey: "Branding & Marketing",
    siteUrlPart: "BrandingMarketing",
    ownerGroup: "Branding & Marketing Owners",
    contributorGroup: "Marketing_Contributor",
    viewerGroup: "Marketing_Viewer",
    aadGroupId: "33e5005e-aa10-4ebb-98cb-85e94d46919d"
  },
  {
    siteKey: "PeopleHub",
    siteUrlPart: "PeopleHub",
    ownerGroup: "PeopleHub Owners",
    contributorGroup: "PeopleHub_Contributor",
    viewerGroup: "PeopleHub_Viewer",
    aadGroupId: "21448698-f734-4f99-b8a3-e64c8a5de658"
  }
];

// export const OWNER_GROUP = ["OperationsHub Owners", "Freudiger Owners", "moreYeahsdepartmentsDMS Owners", "PembePortal Owners"];
// export const CONTRIBUTOR_GROUP = ["OperationsHub Members", "Freudiger Members", "moreYeahsdepartmentsDMS Members", "PembePortal Members"];
// export const VIEWER_GROUP = ["OperationsHub Visitors", "Freudiger Visitors", "moreYeahsdepartmentsDMS Visitors", "PembePortal Visitors"];
// export const AAD_VIEWER_GROUP_ID = ["00000000-0000-0000-0000-000000000000"];
// 
// const SITE_PERMISSIONS_CONFIG = [
//   {
//     siteKey: "OperationsHub",
//     siteUrlPart: "OperationsHub",
//     ownerGroup: "OperationsHub Owners",
//     contributorGroup: "OperationsHub Members",
//     viewerGroup: "OperationsHub Visitors",
//     aadGroupId: "00000000-0000-0000-0000-000000000000"
//   },
//   {
//     siteKey: "Freudiger",
//     siteUrlPart: "Freudiger",
//     ownerGroup: "Freudiger Owners",
//     contributorGroup: "Freudiger Members",
//     viewerGroup: "Freudiger Visitors",
//     aadGroupId: "00000000-0000-0000-0000-000000000000"
//   },
//   {
//     siteKey: "moreYeahsdepartmentsDMS",
//     siteUrlPart: "moreYeahsdepartmentsDMS",
//     ownerGroup: "moreYeahsdepartmentsDMS Owners",
//     contributorGroup: "moreYeahsdepartmentsDMS Members",
//     viewerGroup: "moreYeahsdepartmentsDMS Visitors",
//     aadGroupId: "00000000-0000-0000-0000-000000000000"
//   },
//   {
//     siteKey: "PembePortal",
//     siteUrlPart: "PembePortal",
//     ownerGroup: "PembePortal Owners",
//     contributorGroup: "PembePortal Members",
//     viewerGroup: "PembePortal Visitors",
//     aadGroupId: "00000000-0000-0000-0000-000000000000"
//   }
// ];
 
// 🔹 Check if any AAD group is inside any of the SP groups
const isAnyAADGroupInSPGroups = async (sp: any, groupNames: string[], siteGroupTitles?: Set<string>): Promise<boolean> => {
    let titles = siteGroupTitles;
    if (!titles) {
        try {
            const groups = await sp.web.siteGroups.select("Title")();
            titles = new Set(groups.map((g: any) => g.Title.toLowerCase()));
        } catch (e) {
            titles = new Set();
        }
    }

    for (const groupName of groupNames) {
        if (!titles.has(groupName.toLowerCase())) {
            // Group doesn't exist on this site, skip to avoid 404 error in console
            continue;
        }
        try {
            const users = await sp.web.siteGroups.getByName(groupName).users();
            const exists = users.some((u: any) =>
                u.LoginName &&
                AAD_VIEWER_GROUP_ID.some(id => u.LoginName.toLowerCase().includes(id.toLowerCase()))
            );
            if (exists) {
                return true;
            }
        } catch (err) {
            // Ignore error for specific group if it doesn't exist on this site, and try the next one
            console.warn(`[Permission] Group ${groupName} not found or error reading:`, err);
        }
    }
    return false;
};

// 🔹 Check if specific AAD group is nested in a specific SP group
const isAADGroupInSPGroupSingle = async (sp: any, groupName: string, aadGroupId: string, siteGroupTitles?: Set<string>): Promise<boolean> => {
    let titles = siteGroupTitles;
    if (!titles) {
        try {
            const groups = await sp.web.siteGroups.select("Title")();
            titles = new Set(groups.map((g: any) => g.Title.toLowerCase()));
        } catch (e) {
            titles = new Set();
        }
    }

    if (!titles.has(groupName.toLowerCase())) {
        // Group doesn't exist on this site, skip to avoid 404 error in console
        return false;
    }

    try {
        const users = await sp.web.siteGroups.getByName(groupName).users();
        return users.some((u: any) =>
            u.LoginName &&
            u.LoginName.toLowerCase().includes(aadGroupId.toLowerCase())
        );
    } catch (err) {
        return false;
    }
};

// 🔹 Robust AAD Membership (handles pagination)
const checkAADMembership = async (context: any): Promise<{ isMember: boolean; matchedAADGroupIds: string[]; allUserAADGroups: { id: string; displayName?: string }[] }> => {
    const matchedAADGroupIds: string[] = [];
    const allUserAADGroups: { id: string; displayName?: string }[] = [];
    try {
        // Check if msGraphClientFactory is available
        if (!context || !context.msGraphClientFactory) {
            console.warn("[Permission] MSGraphClientFactory not available");
            return { isMember: false, matchedAADGroupIds, allUserAADGroups };
        }

        const client = await context.msGraphClientFactory.getClient("3");
       
        // Check if client was successfully created
        if (!client) {
            console.warn("[Permission] MSGraphClient is not available");
            return { isMember: false, matchedAADGroupIds, allUserAADGroups };
        }

        let url = "/me/memberOf";
        let isMember = false;

        while (url) {
            const res = await client.api(url).get();
           
            // Validate response structure
            if (!res || typeof res !== 'object') {
                console.warn("[Permission] Invalid response from MSGraph API:", res);
                break;
            }

            const groups = res.value || [];

            groups.forEach((g: any) => {
                if (g.id) {
                    allUserAADGroups.push({ id: g.id, displayName: g.displayName });
                    if (AAD_VIEWER_GROUP_ID.some(id => g.id === id)) {
                        isMember = true;
                        if (!matchedAADGroupIds.includes(g.id)) {
                            matchedAADGroupIds.push(g.id);
                        }
                    }
                }
            });
            url = res["@odata.nextLink"] || null;
        }

        return { isMember, matchedAADGroupIds, allUserAADGroups };
    } catch (err) {
        console.error("[Permission] AAD check failed:", err);
        return { isMember: false, matchedAADGroupIds, allUserAADGroups };
    }
};

let _context: any;

export const checkPermissions = async (context?: any): Promise<void> => {
    if (context) {
        _context = context;
    }
   
    const ctx = context || _context;
    if (!ctx) {
        console.error("[PermissionService] No context available");
        return;
    }

    const store = usePermissionStore.getState();
    store.setPermissions({ isLoading: true });

    try {
        const sp = getSP(ctx);

        // Validate SP context
        if (!sp) {
            console.error("[PermissionService] SP context is not available");
            store.setPermissions({ isLoading: false });
            return;
        }

        // 🔹 1. Direct SharePoint membership
        const [currentUser, groups, allSiteGroups] = await Promise.all([
            sp.web.currentUser(),
            sp.web.currentUser.groups(),
            sp.web.siteGroups.select("Title")().catch(() => [])
        ]);

        // Validate responses
        if (!currentUser || !groups) {
            console.error("[PermissionService] Invalid response from SharePoint API");
            store.setPermissions({ isLoading: false, role: "None", canView: false, canAdd: false, canEdit: false, canDelete: false, isOwner: false, isContributor: false, isViewer: false });
            return;
        }

        const siteGroupTitles = new Set((allSiteGroups || []).map((g: any) => (g?.Title || "").toLowerCase()).filter(Boolean));

        const groupNames = (groups || []).map((g: any) => (g?.Title || "").toLowerCase()).filter(Boolean);
        const isSiteAdmin = currentUser?.IsSiteAdmin || false;

        // ✅ SP ROLE
        let spRole: "Owner" | "Contributor" | "Viewer" | "None" = "None";

        const isOwnerMember = OWNER_GROUP.some(g => groupNames.includes(g.toLowerCase()));
        const isContributorMember = CONTRIBUTOR_GROUP.some(g => groupNames.includes(g.toLowerCase()));
        const isViewerMember = VIEWER_GROUP.some(g => groupNames.includes(g.toLowerCase()));

        const matchedSPGroups: string[] = [];
        OWNER_GROUP.forEach(g => {
            if (groupNames.includes(g.toLowerCase())) matchedSPGroups.push(g);
        });
        CONTRIBUTOR_GROUP.forEach(g => {
            if (groupNames.includes(g.toLowerCase())) matchedSPGroups.push(g);
        });
        VIEWER_GROUP.forEach(g => {
            if (groupNames.includes(g.toLowerCase())) matchedSPGroups.push(g);
        });

        if (isSiteAdmin || isOwnerMember) {
            spRole = "Owner";
        } else if (isContributorMember) {
            spRole = "Contributor";
        } else if (isViewerMember) {
            spRole = "Viewer";
        }

        // 🔹 2. AAD → SP group mapping (Check Nesting First)
        const [inOwner, inContributor, inViewer] = await Promise.all([
            isAnyAADGroupInSPGroups(sp, OWNER_GROUP, siteGroupTitles),
            isAnyAADGroupInSPGroups(sp, CONTRIBUTOR_GROUP, siteGroupTitles),
            isAnyAADGroupInSPGroups(sp, VIEWER_GROUP, siteGroupTitles)
        ]);

        const hasNesting = inOwner || inContributor || inViewer;
        let aadRole: "Owner" | "Contributor" | "Viewer" | "None" = "None";
        let isAADMember = false;
        let matchedAADGroupIds: string[] = [];
        let allUserAADGroups: { id: string; displayName?: string }[] = [];

        // Always check AAD membership to get all user's AAD groups for logging/debugging
        const aadResult = await checkAADMembership(ctx);
        isAADMember = aadResult.isMember;
        matchedAADGroupIds = aadResult.matchedAADGroupIds;
        allUserAADGroups = aadResult.allUserAADGroups;

        if (hasNesting && isAADMember) {
            if (inOwner) {
                aadRole = "Owner";
            } else if (inContributor) {
                aadRole = "Contributor";
            } else if (inViewer) {
                aadRole = "Viewer";
            }
        }

        const priority = ["Owner", "Contributor", "Viewer", "None"];
        const role = priority.indexOf(aadRole) < priority.indexOf(spRole) ? aadRole : spRole;

        // 🔹 Determine Permission Source
        let permissionSource = "Not assigned";
        if (isSiteAdmin) {
            permissionSource = "Site Administrator";
        } else if (role !== "None") {
            if (priority.indexOf(aadRole) < priority.indexOf(spRole)) {
                permissionSource = `AAD Nesting (via ${aadRole} group)`;
            } else {
                permissionSource = `SharePoint Group (via ${spRole} group)`;
            }
        }

        // 🔹 4. Permissions
        const isOwner = role === "Owner";
        const isContributor = role === "Contributor" || isOwner;
        const isViewer = role === "Viewer" || isContributor;

        const tenantBaseUrl = ctx.pageContext.web.absoluteUrl.split("/sites/")[0];

        // Compute site breakdown
        const siteBreakdown = await Promise.all(SITE_PERMISSIONS_CONFIG.map(async (site) => {
            let sRole: "Owner" | "Contributor" | "Viewer" | "None" = "None";
            let matchedSPGroup = "";
            let matchedAADGroup = "";
            let siteGroupNames: string[] = [];
            let siteGroupTitlesForBreakdown = new Set<string>();

            const siteUrl = `${tenantBaseUrl}/sites/${site.siteUrlPart}`;
            const siteSp = spfi(siteUrl).using(SPFx(ctx));

            try {
                // Get user's SharePoint groups on the target site collection
                const [groups, allSiteGroups] = await Promise.all([
                    siteSp.web.currentUser.groups(),
                    siteSp.web.siteGroups.select("Title")().catch(() => [])
                ]);
                siteGroupNames = groups.map((g: any) => g.Title.toLowerCase());
                siteGroupTitlesForBreakdown = new Set((allSiteGroups || []).map((g: any) => (g?.Title || "").toLowerCase()).filter(Boolean));
            } catch (e) {
                // Ignore if unable to access site
            }

            // 1. Direct SP membership
            const matchedSPGroupsForSite: string[] = [];
            if (siteGroupNames.includes(site.ownerGroup.toLowerCase())) {
                sRole = "Owner";
                matchedSPGroupsForSite.push(site.ownerGroup);
            }
            if (siteGroupNames.includes(site.contributorGroup.toLowerCase())) {
                if (sRole !== "Owner") sRole = "Contributor";
                matchedSPGroupsForSite.push(site.contributorGroup);
            }
            if (siteGroupNames.includes(site.viewerGroup.toLowerCase())) {
                if (sRole !== "Owner" && sRole !== "Contributor") sRole = "Viewer";
                matchedSPGroupsForSite.push(site.viewerGroup);
            }

            // 2. AAD Nesting membership
            const [inOwnerG, inContributorG, inViewerG] = await Promise.all([
                isAADGroupInSPGroupSingle(siteSp, site.ownerGroup, site.aadGroupId, siteGroupTitlesForBreakdown),
                isAADGroupInSPGroupSingle(siteSp, site.contributorGroup, site.aadGroupId, siteGroupTitlesForBreakdown),
                isAADGroupInSPGroupSingle(siteSp, site.viewerGroup, site.aadGroupId, siteGroupTitlesForBreakdown)
            ]);

            const userInAAD = allUserAADGroups.some(g => g.id === site.aadGroupId);
            const matchedAADGroupsForSite: string[] = [];

            if (userInAAD) {
                let aadSRole: "Owner" | "Contributor" | "Viewer" | "None" = "None";
                if (inOwnerG) {
                    aadSRole = "Owner";
                    matchedAADGroupsForSite.push(`${site.ownerGroup} (via AAD)`);
                }
                if (inContributorG) {
                    if (aadSRole !== "Owner") aadSRole = "Contributor";
                    matchedAADGroupsForSite.push(`${site.contributorGroup} (via AAD)`);
                }
                if (inViewerG) {
                    if (aadSRole !== "Owner" && aadSRole !== "Contributor") aadSRole = "Viewer";
                    matchedAADGroupsForSite.push(`${site.viewerGroup} (via AAD)`);
                }

                if (priority.indexOf(aadSRole) < priority.indexOf(sRole)) {
                    sRole = aadSRole;
                }
            }

            return {
                siteKey: site.siteKey,
                role: sRole,
                matchedSPGroups: matchedSPGroupsForSite,
                matchedAADGroups: matchedAADGroupsForSite,
                configuredAADId: site.aadGroupId,
                userInAAD,
                nestingInSP: inOwnerG || inContributorG || inViewerG
            };
        }));

        let finalAllowedSites: string[] = [];
        if (isSiteAdmin) {
            finalAllowedSites = [
                'TrivandiHub',
                'PeopleHub',
                'CompanyHub',
                'BrandingMarketing',
                'Projects',
                'TrivandiLondon',
                'TDMCC',
                'TrivandiUSA',
                'TrivandiAustralia',
                'TrivandiKSA'
            ];
        } else {
            const allowedSites: string[] = [];
            siteBreakdown.forEach(s => {
                if (s.role === "Owner" || s.role === "Contributor") {
                    let mappedPath = "";
                    if (s.siteKey === "CompanyHub") mappedPath = "CompanyHub";
                    else if (s.siteKey === "Projects") mappedPath = "Projects";
                    else if (s.siteKey === "TrivandiHub") mappedPath = "TrivandiHub";
                    else if (s.siteKey === "Branding & Marketing") mappedPath = "BrandingMarketing";
                    else if (s.siteKey === "PeopleHub") mappedPath = "PeopleHub";
                    
                    if (mappedPath) {
                        allowedSites.push(mappedPath);
                    }
                }
            });
            const unrestrictedSites = ['TrivandiLondon', 'TDMCC', 'TrivandiUSA', 'TrivandiAustralia', 'TrivandiKSA'];
            finalAllowedSites = [...allowedSites, ...unrestrictedSites];
        }

        // const finalAllowedSites = [
        //     'OperationsHub',
        //     'Freudiger',
        //     'moreYeahsdepartmentsDMS',
        //     'PembePortal'
        // ];

        const permissions = {
            role,
            canAdd: isContributor,
            canEdit: isContributor,
            canDelete: isOwner,
            canView: isViewer,
            isOwner,
            isContributor,
            isViewer,
            isLoading: false,
            allowedSites: finalAllowedSites
        };

        // 📊 CONSOLIDATED PERMISSIONS SUMMARY
        const nestingFound = [];
        if (inOwner) nestingFound.push("Owner Nesting");
        if (inContributor) nestingFound.push("Contributor Nesting");
        if (inViewer) nestingFound.push("Viewer Nesting");

        const allUserAADGroupsLog = allUserAADGroups.map(g => `${g.displayName || 'Unnamed Group'} (${g.id})`).join(", ");

        console.log("%c--- PERMISSIONS SUMMARY ---", "color: #2563eb; font-weight: bold; font-size: 13px; padding: 4px;");
        console.log(`👤 User: %c${currentUser.Title} (${currentUser.Email || 'No Email'})`, "color: #0f172a; font-weight: 600;");
        console.log(`📂 SharePoint Groups User Is In: %c${groupNames.length > 0 ? groupNames.join(", ") : "None"}`, "color: #475569;");
        console.log(`🎯 Configured SP Groups Matched: %c${matchedSPGroups.length > 0 ? matchedSPGroups.join(", ") : "None"}`, "color: #0284c7; font-weight: bold;");
        console.log(`🔗 AAD Nesting Found in Site: %c${nestingFound.length > 0 ? nestingFound.join(", ") : "None"}`, "color: #475569;");
        console.log(`🌐 ALL User AAD Groups: %c${allUserAADGroups.length > 0 ? allUserAADGroupsLog : "None (or no AAD groups returned)"}`, "color: #7c3aed;");
        console.log(`✅ Member of Configured AAD Groups: %c${matchedAADGroupIds.length > 0 ? matchedAADGroupIds.join(", ") : "No matched groups (User is not a member of AAD permission groups)"}`, `color: ${isAADMember ? "#16a34a" : "#dc2626"}; font-weight: bold;`);
        
        console.log("🏢 --- SITE-BY-SITE ROLES BREAKDOWN ---");
        siteBreakdown.forEach(s => {
            const allMatches = [...s.matchedSPGroups, ...s.matchedAADGroups];
            const detail = allMatches.length > 0 ? allMatches.join(", ") : "No groups match";
            const userInAADStr = s.userInAAD ? "Yes" : "No";
            const nestingInSPStr = s.nestingInSP ? "Yes" : "No";
            
            console.log(`  📍 %c${s.siteKey}: %c${s.role} %c(${detail})`, "font-weight: bold; color: #1e293b;", `font-weight: bold; color: ${s.role === 'None' ? '#64748b' : s.role === 'Owner' ? '#16a34a' : '#2563eb'}`, "color: #64748b;");
            console.log(`     └─ Configured AAD ID: %c${s.configuredAADId}%c | User Member: %c${userInAADStr}%c | SP Nesting Active: %c${nestingInSPStr}`, "color: #7c3aed;", "color: #64748b;", `color: ${s.userInAAD ? "#16a34a" : "#dc2626"}; font-weight: bold;`, "color: #64748b;", `color: ${s.nestingInSP ? "#16a34a" : "#dc2626"}; font-weight: bold;`);
        });
        console.log(`🔒 Search-Allowed Sites: %c${finalAllowedSites.join(", ")}`, "color: #0284c7; font-weight: bold;");
        console.log("---------------------------------------");

        console.log(`🎯 Assigned Role (Current Site): %c${role}`, "color: #2563eb; font-weight: bold;");
        console.log(`📍 Permission Source (Current Site): %c${permissionSource}`, "color: #7c3aed; font-weight: 600;");
       
        console.log(`🛠️ Permissions -> Add: ${permissions.canAdd}, Edit: ${permissions.canEdit}, Delete: ${permissions.canDelete}, View: ${permissions.canView}`);
        console.log("%c---------------------------", "color: #cbd5e1;");

        store.setPermissions(permissions);
 
    } catch (err) {
        console.error("[PermissionService] Error checking permissions:", err);
        // Set default safe permissions on error
        store.setPermissions({
            isLoading: false,
            role: "None",
            canView: false,
            canAdd: false,
            canEdit: false,
            canDelete: false,
            isOwner: false,
            isContributor: false,
            isViewer: false,
            allowedSites: []
        });
    }
};