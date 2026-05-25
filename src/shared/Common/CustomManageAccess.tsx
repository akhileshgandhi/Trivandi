import * as React from 'react';
import {
  Panel,
  PanelType,
  IconButton,
  Dropdown,
  IDropdownOption,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Persona,
  PersonaSize,
  Pivot,
  PivotItem,
  SearchBox,
  Stack,
  Text,
  mergeStyleSets,
  FontWeights
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/security";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { UserMinus, Ban, Folder, FileText } from 'lucide-react';
import { useToastStore } from '../Global/ToastStore';

// Version: 2.0.0 - Uses SharePoint Sharing API (works for Member/Edit role too)

export interface ICustomManageAccessProps {
  isOpen: boolean;
  onDismiss: () => void;
  item: any;
  context: WebPartContext;
  siteUrl: string;
}

interface IAccessEntry {
  principalId: number;
  title: string;
  email: string;
  permissionsText: string;
  roleBindings: number[];
  principalType: number;
  isLink: boolean;
  loginName?: string;
}

const classNames = mergeStyleSets({
  headerSection: {
    padding: '20px 24px',
    borderBottom: '1px solid #edebe9'
  },
  itemHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: 12,
    marginBottom: 8
  },
  itemName: {
    fontSize: 16,
    fontWeight: FontWeights.semibold,
    color: '#323130'
  },
  actionLink: {
    color: '#0078d4',
    cursor: 'pointer',
    fontSize: 14,
    display: 'flex',
    alignItems: 'center',
    gap: 6,
    textDecoration: 'none',
    selectors: {
      '&:hover': { textDecoration: 'underline' }
    }
  },
  stopSharing: {
    color: '#605e5c',
    fontSize: 14,
    display: 'flex',
    alignItems: 'center',
    gap: 6,
    cursor: 'pointer',
    selectors: {
      '&:hover': { color: '#323130' }
    }
  },
  tabSection: {
    padding: '0 24px'
  },
  searchSection: {
    padding: '16px 24px'
  },
  listSection: {
    padding: '0 8px',
    maxHeight: '400px',
    overflowY: 'auto'
  },
  accessRow: {
    display: 'flex',
    alignItems: 'center',
    padding: '12px 16px',
    borderRadius: 4,
    selectors: {
      '&:hover': { background: '#f3f2f1' }
    }
  }
});

export const CustomManageAccess: React.FunctionComponent<ICustomManageAccessProps> = (
  props: ICustomManageAccessProps
): React.ReactElement<ICustomManageAccessProps> => {
  const { isOpen, onDismiss, item, context, siteUrl } = props;
  const { showToast } = useToastStore();

  const [loading, setLoading] = React.useState(true);
  const [accessList, setAccessList] = React.useState<IAccessEntry[]>([]);
  const [roleDefinitions, setRoleDefinitions] = React.useState<IDropdownOption[]>([]);
  const [error, setError] = React.useState<string | null>(null);
  const [hasUniquePermissions, setHasUniquePermissions] = React.useState(false);
  const [canManagePermissions, setCanManagePermissions] = React.useState(true);
  const [searchQuery, setSearchQuery] = React.useState('');
  const [selectedTab, setSelectedTab] = React.useState('all');

  // ─── Helpers ────────────────────────────────────────────────────────────────

  const normalizeServerRelativeUrl = React.useCallback((rawUrl: string): string => {
    if (!rawUrl) return '';
    let normalized = rawUrl;
    try { normalized = decodeURIComponent(rawUrl); } catch { normalized = rawUrl; }
    if (/^https?:\/\//i.test(normalized)) {
      try { normalized = new URL(normalized).pathname; } catch { /* keep */ }
    }
    if (!normalized.startsWith('/')) normalized = `/${normalized}`;
    return normalized.replace(/\/+/g, '/');
  }, []);

  const getSiteRootFromServerRelativeUrl = React.useCallback((serverRelativeUrl: string): string => {
    const segments = serverRelativeUrl.split('/').filter(Boolean);
    if (
      segments.length >= 2 &&
      (segments[0].toLowerCase() === 'sites' || segments[0].toLowerCase() === 'teams')
    ) {
      return `/${segments[0]}/${segments[1]}`;
    }
    return '/';
  }, []);

  const getTargetWeb = React.useCallback((serverRelativeUrl: string) => {
    const normalizedPath = normalizeServerRelativeUrl(serverRelativeUrl);
    const siteRoot = getSiteRootFromServerRelativeUrl(normalizedPath);
    const webUrl = siteRoot === '/' ? siteUrl : `${window.location.origin}${siteRoot}`;
    return { web: spfi(webUrl).using(SPFx(context)).web, normalizedPath, webUrl };
  }, [context, getSiteRootFromServerRelativeUrl, normalizeServerRelativeUrl, siteUrl]);

  const resolveItemContext = React.useCallback(async (): Promise<{
    web: any;
    spItem: any;
    normalizedPath: string;
    webUrl: string;
    listId: string;
    itemId: number;
  }> => {
    const defaultWeb = spfi(siteUrl).using(SPFx(context)).web;

    // Case 1: ListId + IntegerId/Id directly available on item prop
    if (item.ListId && (item.IntegerId || item.Id)) {
      const itemId = Number(item.IntegerId || item.Id);
      return {
        web: defaultWeb,
        spItem: defaultWeb.lists.getById(item.ListId).items.getById(itemId),
        normalizedPath: item.ServerRelativeUrl || item.FileRef || item.Path || '',
        webUrl: siteUrl,
        listId: item.ListId,
        itemId
      };
    }

    // Case 2: Resolve via ServerRelativeUrl / FileRef / Path
    const rawPath = item.ServerRelativeUrl || item.FileRef || item.Path || '';
    const { web, normalizedPath, webUrl } = getTargetWeb(rawPath);

    if (!normalizedPath) throw new Error('Missing server relative URL for selected item.');

    const fileOrFolder = item.IsFolder
      ? web.getFolderByServerRelativePath(normalizedPath)
      : web.getFileByServerRelativePath(normalizedPath);

    const itemInfo = await fileOrFolder.listItemAllFields
      .select('Id', 'ParentList/Id')
      .expand('ParentList')();

    const resolvedListId: string = itemInfo.ParentList?.Id || '';
    const resolvedItemId: number = Number(itemInfo.Id);

    return {
      web,
      spItem: web.lists.getById(resolvedListId).items.getById(resolvedItemId),
      normalizedPath,
      webUrl,
      listId: resolvedListId,
      itemId: resolvedItemId
    };
  }, [context, getTargetWeb, item, siteUrl]);

  const getUserPhotoUrl = React.useCallback((email: string): string | undefined => {
    if (!email) return undefined;
    return `${window.location.origin}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(email)}`;
  }, []);

  // ─── GetSharingInformation API (exact same endpoint SharePoint UI uses) ────────
  // Works for Member/Edit role — confirmed from Network tab

  const getSharingInfo = React.useCallback(async (
    webUrl: string,
    listId: string,
    itemId: number
  ): Promise<any> => {
    const digest = await getRequestDigest(webUrl);
    // Exact URL pattern confirmed from SharePoint Network tab
    const url = `${webUrl}/_api/web/Lists(@a1)/GetItemById(@a2)/GetSharingInformation` +
      `?@a1='${listId}'&@a2='${itemId}'` +
      `&$Expand=pickerSettings,permissionsInformation,addressBarLinkSettings`;

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': digest
      },
      credentials: 'include',
      body: ''
    });

    if (!response.ok) {
      const errText = await response.text().catch(() => '');
      throw new Error(`GetSharingInformation error: ${response.status} ${errText}`);
    }
    return response.json();
  }, []);

  const getRequestDigest = async (webUrl: string): Promise<string> => {
    const resp = await fetch(`${webUrl}/_api/contextinfo`, {
      method: 'POST',
      headers: { 'Accept': 'application/json;odata=verbose' },
      credentials: 'include'
    });
    const data = await resp.json();
    return data?.d?.GetContextWebInformation?.FormDigestValue || '';
  };

  // ─── Load access data using Sharing API ─────────────────────────────────────

  const loadAccessData = async (): Promise<void> => {
    try {
      setLoading(true);
      setError(null);

      const { web, spItem, normalizedPath, webUrl, listId: resolvedListId, itemId: resolvedItemId } = await resolveItemContext();

      console.log('Resolved ListId:', resolvedListId, 'ItemId:', resolvedItemId);

      // Check unique permissions
      const itemData = await spItem.select("HasUniqueRoleAssignments")();
      setHasUniquePermissions(itemData.HasUniqueRoleAssignments);

      // ── GetSharingInformation API (exact endpoint SharePoint UI uses) ──
      let mappedAccess: IAccessEntry[] = [];
      let usedSharingApi = false;

      try {
        const sharingInfo = await getSharingInfo(webUrl, resolvedListId, resolvedItemId);
        console.log('GetSharingInformation response:', JSON.stringify(sharingInfo).substring(0, 800));

        // Response structure: d.permissionsInformation.principals / links
        const permInfo = sharingInfo?.d?.permissionsInformation;
        const principals: any[] = [
          ...(permInfo?.principals?.results || permInfo?.principals || []),
        ];
        const links: any[] = [
          ...(permInfo?.links?.results || permInfo?.links || [])
        ];

        // Map principals (users & groups)
        const mappedPrincipals: IAccessEntry[] = principals
          .filter((p: any) => p?.principal)
          .map((p: any) => {
            const principal = p.principal;
            const roleMap: Record<number, string> = { 1: 'Can view', 2: 'Can edit', 3: 'Owner' };
            return {
              principalId: principal.id,
              title: principal.name || principal.loginName || 'Unknown',
              email: principal.email || '',
              permissionsText: roleMap[p.role] || 'Custom',
              roleBindings: [p.role],
              principalType: principal.principalType,
              isLink: false,
              loginName: principal.loginName
            };
          });

        // Map sharing links
        const mappedLinks: IAccessEntry[] = links.map((l: any, idx: number) => ({
          principalId: -(idx + 1), // negative id for links
          title: 'Sharing Link (External/Guest)',
          email: '',
          permissionsText: l.linkDetails?.Url ? 'Shared Link' : 'Link',
          roleBindings: [],
          principalType: 8,
          isLink: true,
          loginName: ''
        }));

        mappedAccess = [...mappedPrincipals, ...mappedLinks];
        setCanManagePermissions(true);
        usedSharingApi = true;
      } catch (sharingErr) {
        console.warn('GetSharingInformation failed, falling back to roleAssignments:', sharingErr);
      }

      // ── Fallback: roleAssignments (needs Full Control) ──
      if (!usedSharingApi) {
        try {
          const assignments = await spItem.roleAssignments
            .expand("Member", "RoleDefinitionBindings")();

          mappedAccess = assignments
            .filter((a: any) => a.Member)
            .map((a: any) => {
              const isLink = a.Member.Title && a.Member.Title.includes('SharingLinks');
              let displayTitle = isLink ? 'Sharing Link (External/Guest)' : (a.Member.Title || 'Unknown');

              return {
                principalId: a.Member.Id,
                title: displayTitle,
                email: a.Member.Email || a.Member.UserPrincipalName || '',
                roleBindings: a.RoleDefinitionBindings
                  ? a.RoleDefinitionBindings.map((r: any) => r.Id)
                  : [],
                permissionsText: a.RoleDefinitionBindings
                  ? a.RoleDefinitionBindings.map((r: any) => {
                    if (r.Name === 'Full Control') return 'Owner';
                    if (r.Name === 'Design') return 'Editor';
                    if (r.Name === 'Contribute' || r.Name === 'Edit') return 'Can edit';
                    if (r.Name === 'Read') return 'Can view';
                    return r.Name;
                  }).join(', ')
                  : 'No Access',
                principalType: a.Member.PrincipalType,
                isLink,
                loginName: a.Member.LoginName
              };
            });

          setCanManagePermissions(true);
        } catch (roleErr: any) {
          const errText = `${roleErr?.message || ''} ${roleErr?.status || ''} ${JSON.stringify(roleErr || {})}`;
          const isForbidden =
            roleErr?.status === 403 ||
            roleErr?.status === 401 ||
            errText.indexOf('403') > -1 ||
            errText.indexOf('401') > -1 ||
            errText.indexOf('UnauthorizedAccessException') > -1 ||
            errText.indexOf('AccessDenied') > -1;

          if (isForbidden) {
            setCanManagePermissions(false);
            setAccessList([]);
            setRoleDefinitions([]);
            return;
          }
          throw roleErr;
        }
      }

      setAccessList(mappedAccess);

      // Load role definitions for dropdown (needs Full Control; gracefully degrade)
      try {
        const roles = await web.roleDefinitions();
        const readRole = roles.find((r: any) => r.Name === 'Read');
        const editRole = roles.find((r: any) => r.Name === 'Edit' || r.Name === 'Contribute');
        const options: IDropdownOption[] = [];
        if (readRole) options.push({ key: readRole.Id, text: 'Can view' });
        if (editRole) options.push({ key: editRole.Id, text: 'Can edit' });
        setRoleDefinitions(options);
      } catch {
        // If role definitions can't be loaded, use sharing role numbers as keys
        setRoleDefinitions([
          { key: 1, text: 'Can view' },
          { key: 2, text: 'Can edit' }
        ]);
      }
    } catch (err) {
      console.error('Error loading access data:', err);
      setError('Failed to load access information.');
    } finally {
      setLoading(false);
    }
  };

  React.useEffect(() => {
    if (isOpen && item) {
      setCanManagePermissions(true);
      loadAccessData();
    }
  }, [isOpen, item]);

  // ─── Actions ─────────────────────────────────────────────────────────────────

  const handleBreakInheritance = async (): Promise<void> => {
    try {
      setLoading(true);
      const { spItem } = await resolveItemContext();
      await spItem.breakRoleInheritance(true);
      showToast('Inheritance broken successfully.', 'success');
      await loadAccessData();
    } catch (err) {
      console.error('Error breaking inheritance:', err);
      showToast(`Failed to break inheritance: ${err}`, 'error');
      setLoading(false);
    }
  };

  const handleResetInheritance = async (): Promise<void> => {
    try {
      setLoading(true);
      const { spItem } = await resolveItemContext();
      await spItem.resetRoleInheritance();
      showToast('Sharing reset successfully.', 'success');
      await loadAccessData();
    } catch (err) {
      console.error('Error resetting inheritance:', err);
      showToast(`Failed to reset sharing: ${err}`, 'error');
      setLoading(false);
    }
  };

  // Remove access via Sharing API (POST to UpdateDocumentSharingInfo with role=0)
  const handleRemoveAccessViaShareApi = async (
    webUrl: string,
    normalizedPath: string,
    loginName: string
  ): Promise<void> => {
    const digest = await getRequestDigest(webUrl);
    const escapedPath = normalizedPath.replace(/'/g, "''");
    const endpoint = `${webUrl}/_api/SP.Sharing.UpdateDocumentSharingInfo` +
      `@target('${escapedPath}')`;

    const body = JSON.stringify({
      userRoleAssignments: [{
        Role: 0, // 0 = remove
        UserId: loginName
      }],
      validateExistingPermissions: false,
      additiveMode: false,
      sendServerManagedNotification: false,
      customMessage: '',
      includeAnonymousLinksInNotification: false
    });

    const response = await fetch(endpoint, {
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': digest
      },
      credentials: 'include',
      body
    });

    if (!response.ok) throw new Error(`Remove access failed: ${response.status}`);
  };

  const handleRemoveAccess = async (principalId: number): Promise<void> => {
    if (!principalId) {
      setError('Cannot remove access: Principal ID is missing.');
      return;
    }
    try {
      setLoading(true);
      const { spItem, normalizedPath, webUrl } = await resolveItemContext();
      const accessItem = accessList.find(a => a.principalId === principalId);

      // Try Sharing API first (works for Members)
      if (accessItem?.loginName) {
        try {
          if (!hasUniquePermissions) await spItem.breakRoleInheritance(true);
          await handleRemoveAccessViaShareApi(webUrl, normalizedPath, accessItem.loginName);
          showToast('Access removed successfully.', 'success');
          await loadAccessData();
          return;
        } catch (sharingErr) {
          console.warn('Sharing API remove failed, trying roleAssignments:', sharingErr);
        }
      }

      // Fallback to roleAssignments
      if (!hasUniquePermissions) await spItem.breakRoleInheritance(true);
      if (accessItem?.roleBindings) {
        for (const roleId of accessItem.roleBindings) {
          await spItem.roleAssignments.remove(principalId, roleId);
        }
      }
      showToast('Access removed successfully.', 'success');
      await loadAccessData();
    } catch (err) {
      console.error('Error removing access:', err);
      showToast(`Failed to remove access: ${err}`, 'error');
      setLoading(false);
    }
  };

  const handleUpdatePermission = async (principalId: number, newRoleId: number): Promise<void> => {
    if (!principalId) {
      setError('Cannot update permission: Missing Principal ID.');
      return;
    }
    try {
      setLoading(true);
      const { spItem } = await resolveItemContext();
      if (!hasUniquePermissions) await spItem.breakRoleInheritance(true);

      const currentAccess = accessList.find(a => a.principalId === principalId);
      if (currentAccess?.roleBindings) {
        for (const roleId of currentAccess.roleBindings) {
          try { await spItem.roleAssignments.remove(principalId, roleId); } catch { /* ignore */ }
        }
      }
      await spItem.roleAssignments.add(principalId, newRoleId);
      showToast('Permissions updated successfully.', 'success');
      await loadAccessData();
    } catch (err) {
      console.error('Error updating permission:', err);
      showToast(`Failed to update permission: ${err}`, 'error');
      setLoading(false);
    }
  };

  // ─── Filtering ───────────────────────────────────────────────────────────────

  const filteredAccess = accessList.filter(a => {
    const matchesSearch = a.title.toLowerCase().indexOf(searchQuery.toLowerCase()) > -1;
    if (!matchesSearch) return false;
    if (selectedTab === 'all') return true;
    if (selectedTab === 'people') return a.principalType !== 4 && a.principalType !== 8 && !a.isLink;
    if (selectedTab === 'groups') return (a.principalType === 4 || a.principalType === 8) && !a.isLink;
    if (selectedTab === 'links') return a.isLink;
    return true;
  });

  const peopleCount = accessList.filter(a => a.principalType !== 4 && a.principalType !== 8 && !a.isLink).length;
  const groupsCount = accessList.filter(a => (a.principalType === 4 || a.principalType === 8) && !a.isLink).length;
  const linksCount = accessList.filter(a => a.isLink).length;

  // ─── Dropdown selected key helper ───────────────────────────────────────────

  const getSelectedRoleKey = (accessItem: IAccessEntry): number | string | undefined => {
    if (!accessItem.roleBindings.length) return undefined;
    const canViewOpt = roleDefinitions.find(opt => opt.text === 'Can view');
    const isRead = accessItem.roleBindings.some(id => canViewOpt && canViewOpt.key === id);
    if (isRead) return canViewOpt?.key;
    return roleDefinitions.find(opt => opt.text === 'Can edit')?.key;
  };

  // ─── Render ──────────────────────────────────────────────────────────────────

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText="Manage Access"
      closeButtonAriaLabel="Close"
      styles={{
        main: { background: '#ffffff' },
        content: { padding: 0 }
      }}
    >
      {/* Item header */}
      <div className={classNames.headerSection}>
        <div className={classNames.itemHeader}>
          {item?.IsFolder
            ? <Folder size={20} color="#f8d12d" fill="#f8d12d" />
            : <FileText size={20} color="#0078d4" />}
          <Text className={classNames.itemName}>{item?.Name}</Text>
        </div>
        <Stack horizontal horizontalAlign="end" verticalAlign="center">
          {canManagePermissions && (hasUniquePermissions ? (
            <div className={classNames.stopSharing} onClick={handleResetInheritance}>
              <UserMinus size={16} />
              Stop sharing
            </div>
          ) : (
            <div className={classNames.actionLink} onClick={handleBreakInheritance}>
              <Ban size={16} />
              Stop Inheriting
            </div>
          ))}
        </Stack>
      </div>

      {/* Tabs */}
      <div className={classNames.tabSection}>
        <Pivot
          selectedKey={selectedTab}
          onLinkClick={(pivotItem) => setSelectedTab(pivotItem?.props.itemKey || 'all')}
          styles={{ root: { borderBottom: '1px solid #edebe9' } }}
        >
          <PivotItem headerText={`All • ${accessList.length}`} itemKey="all" />
          <PivotItem headerText={`People • ${peopleCount}`} itemKey="people" />
          <PivotItem headerText={`Groups • ${groupsCount}`} itemKey="groups" />
          {linksCount > 0 && <PivotItem headerText={`Links • ${linksCount}`} itemKey="links" />}
        </Pivot>
      </div>

      {/* Search */}
      <div className={classNames.searchSection}>
        <SearchBox
          placeholder="Search displayed names"
          value={searchQuery}
          onChange={(_, val) => setSearchQuery(val || '')}
          underlined={true}
        />
      </div>

      {/* Error / permission messages */}
      {error && (
        <MessageBar
          messageBarType={MessageBarType.error}
          onDismiss={() => setError(null)}
          style={{ margin: '0 24px 10px' }}
        >
          {error}
        </MessageBar>
      )}
      {!canManagePermissions && !loading && (
        <MessageBar messageBarType={MessageBarType.info} style={{ margin: '0 24px 10px' }}>
          You don&apos;t have permission to view or manage access for this item.
          Please use SharePoint&apos;s built-in sharing options instead.
        </MessageBar>
      )}

      {/* Access list */}
      <div className={classNames.listSection} style={{ position: 'relative', minHeight: 100 }}>
        {loading && (
          <div style={{
            position: 'absolute', top: 0, left: 0, right: 0, bottom: 0,
            background: 'rgba(255,255,255,0.6)', zIndex: 10,
            display: 'flex', justifyContent: 'center', alignItems: 'center'
          }}>
            <Spinner size={SpinnerSize.medium} />
          </div>
        )}

        {canManagePermissions && accessList.length === 0 && !loading && (
          <div style={{ padding: 40, textAlign: 'center', color: '#605e5c' }}>
            No access information found.
          </div>
        )}

        {canManagePermissions && accessList.length > 0 && (
          <div>
            {filteredAccess.map((accessItem) => (
              <div key={accessItem.principalId} className={classNames.accessRow}>
                <Persona
                  text={accessItem.title}
                  secondaryText={
                    accessItem.isLink
                      ? 'Shared Link'
                      : (accessItem.principalType !== 4 && accessItem.principalType !== 8
                        ? accessItem.email
                        : 'Group')
                  }
                  imageUrl={
                    accessItem.isLink || accessItem.principalType === 4 || accessItem.principalType === 8
                      ? undefined
                      : getUserPhotoUrl(accessItem.email)
                  }
                  size={PersonaSize.size32}
                  styles={{ root: { flexGrow: 1 } }}
                />
                <Dropdown
                  options={roleDefinitions}
                  selectedKey={getSelectedRoleKey(accessItem)}
                  onChange={(_, opt) => handleUpdatePermission(accessItem.principalId, opt?.key as number)}
                  disabled={loading}
                  styles={{
                    root: { width: 90 },
                    title: {
                      border: 'none',
                      background: 'transparent',
                      color: '#605e5c',
                      padding: '0 8px',
                      textAlign: 'right',
                      fontSize: 13
                    },
                    caretDownWrapper: { display: 'none' }
                  }}
                />
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <IconButton
                    iconProps={{ iconName: 'Cancel' }}
                    title="Stop sharing"
                    onClick={() => handleRemoveAccess(accessItem.principalId)}
                    styles={{ root: { color: '#605e5c' } }}
                  />
                  <Text
                    style={{ color: '#0078d4', cursor: 'pointer', fontSize: 12 }}
                    onClick={() => handleRemoveAccess(accessItem.principalId)}
                  >
                    Stop sharing
                  </Text>
                </Stack>
              </div>
            ))}
          </div>
        )}
      </div>
    </Panel>
  );
};

export default CustomManageAccess;