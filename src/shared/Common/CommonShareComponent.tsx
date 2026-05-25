import * as React from 'react';
import { Share2, Globe } from 'lucide-react';
import { Icon } from '@fluentui/react/lib/Icon';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

const SharedDialog = {
  show: async (options: { listId: string; itemId: number; siteRoot: string; isFolder: boolean }): Promise<void> => {
    const { listId, itemId, siteRoot } = options;

    try {
      // Ensure listId has braces {...} if missing - SharePoint often requires this format
      let formattedListId = listId || "";
      if (formattedListId && !formattedListId.startsWith('{')) formattedListId = `{${formattedListId}`;
      if (formattedListId && !formattedListId.endsWith('}')) formattedListId = `${formattedListId}}`;

      const shareUrl = `${siteRoot}/_layouts/15/sharedialog.aspx?listId=${encodeURIComponent(formattedListId)}&listItemId=${itemId}&IsDlg=1`;

      console.log("FINAL SHARE URL =>", shareUrl);

      window.open(
        shareUrl,
        "_blank",
        "width=900,height=700"
      );

      console.log("Share dialog opened");
    } catch (error) {
      console.error("SHARE ERROR =>", error);
      throw error;
    }
  }
};

export interface ICommonShareComponentProps {
  item: any;
  context: WebPartContext;
  buttonType?: 'icon' | 'button' | 'menuItem';
  className?: string;
  iconSize?: number;
  onAfterClick?: () => void;
}

const CommonShareComponent: React.FC<ICommonShareComponentProps> = (props) => {
  const { item, context, buttonType = 'icon', className, iconSize = 20, onAfterClick } = props;

  const getEffectiveSiteUrl = (serverRelativeUrl: string): string => {
    try {
      const hostUrl = new URL(context.pageContext.web.absoluteUrl);
      const match = serverRelativeUrl.match(/^(\/(?:sites|teams|personal)\/[^/]+)/i);
      if (match) return `${hostUrl.protocol}//${hostUrl.host}${match[1]}`;
    } catch {
      // Ignore and fallback to current web URL
    }
    return context.pageContext.web.absoluteUrl;
  };

  const toServerRelativeUrl = (value: string): string => {
    if (!value) return '';
    if (value.startsWith('http')) {
      try {
        return new URL(value).pathname;
      } catch {
        return value;
      }
    }
    return value;
  };

  const handleShareClick = async (e: React.MouseEvent): Promise<void> => {
    e.stopPropagation();

    console.log("========== SHARE DEBUG ==========");
    console.log("ITEM =>", item);

    const serverRelativeUrl = toServerRelativeUrl(item?.ServerRelativeUrl || item?.FileRef || '');
    const isFolder = item?.IsFolder === true;

    if (!serverRelativeUrl) {
      console.error("Missing serverRelativeUrl");
      return;
    }

    const siteRoot = getEffectiveSiteUrl(serverRelativeUrl);

    try {
      // Resolve actual ListItemAllFields before opening native sharing dialog
      // This is the official Microsoft-supported way for folder sharing.
      const escapedServerRelativeUrl = serverRelativeUrl.replace(/'/g, "''");
      const apiPath = isFolder
        ? `_api/web/GetFolderByServerRelativeUrl('${escapedServerRelativeUrl}')/ListItemAllFields`
        : `_api/web/GetFileByServerRelativeUrl('${escapedServerRelativeUrl}')/ListItemAllFields`;

      const apiUrl = `${siteRoot}/${apiPath.startsWith('/') ? apiPath.substring(1) : apiPath}`;
      console.log("API URL =>", apiUrl);

      const response = await context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`API failed with status ${response.status}`);
      }

      const data = await response.json();
      console.log("LIST ITEM DATA =>", data);

      // The API returns the list ID and the item ID we need
      let listId = data.ParentListId || item.ListId;
      const itemId = data.Id;

      // Extract listId from editLink if missing (common for folders)
      const editLink = data['@odata.editLink'] || data['odata.editLink'] || data['@odata.id'];
      if (!listId && editLink) {
        const match = editLink.match(/Lists\(guid'([^']+)'\)/i);
        if (match && match[1]) {
          listId = match[1];
        }
      }

      if (!listId || !itemId) {
        throw new Error('Could not determine List ID or Item ID for sharing');
      }

      console.log("RESOLVED IDs =>", { listId, itemId });

      await SharedDialog.show({
        listId,
        itemId,
        siteRoot,
        isFolder
      });

    } catch (error) {
      console.error("SHARE ERROR =>", error);

      // Fallback: If API fails, try with existing IDs and 'listItemId' parameter
      if (item.ListId && item.IntegerId) {
        console.log("Attempting fallback with existing IDs...");
        await SharedDialog.show({
          listId: item.ListId,
          itemId: item.IntegerId,
          siteRoot,
          isFolder
        });
      }
    } finally {
      if (onAfterClick) onAfterClick();
    }
  };

  if (buttonType === 'button') {
    return (
      <button
        className={className}
        onClick={handleShareClick}
        style={{
          height: '48px',
          borderRadius: '12px',
          fontWeight: 600,
          fontSize: '14px',
          cursor: 'pointer',
          transition: 'all 0.2s ease',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          gap: '8px',
          border: 'none',
          background: '#059669', // Emerald/Green for Share
          color: 'white',
          width: '100%'
        }}
        onMouseEnter={(e) => {
          (e.currentTarget as HTMLButtonElement).style.background = '#047857';
          (e.currentTarget as HTMLButtonElement).style.transform = 'translateY(-2px)';
          (e.currentTarget as HTMLButtonElement).style.boxShadow = '0 4px 12px rgba(5, 150, 105, 0.3)';
        }}
        onMouseLeave={(e) => {
          (e.currentTarget as HTMLButtonElement).style.background = '#059669';
          (e.currentTarget as HTMLButtonElement).style.transform = 'translateY(0)';
          (e.currentTarget as HTMLButtonElement).style.boxShadow = 'none';
        }}
      >
        <Share2 size={iconSize} />
        Share
      </button>
    );
  }

  if (buttonType === 'menuItem') {
    return (
      <div
        className={className}
        onClick={handleShareClick}
        style={{
          padding: '8px 12px',
          display: 'flex',
          alignItems: 'center',
          cursor: 'pointer',
          fontSize: '14px',
          color: '#323130',
          borderRadius: '4px',
          transition: 'all 0.15s ease',
          width: '100%',
          boxSizing: 'border-box'
        }}
        onMouseEnter={(e) => e.currentTarget.style.backgroundColor = '#f3f2f1'}
        onMouseLeave={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
      >
        <span style={{ marginRight: '10px', display: 'flex', alignItems: 'center', color: '#605e5c' }}>
          <Share2 size={16} />
        </span>
        <span style={{ flexGrow: 1 }}>Share</span>
      </div>
    );
  }

  return (
    <div
      className={className}
      onClick={handleShareClick}
      title="Share"
      style={{ cursor: 'pointer', display: 'inline-flex', alignItems: 'center' }}
    >
      <Share2
        size={iconSize}
        style={{
          color: '#5f6368',
          transition: 'color 0.2s'
        }}
        onMouseEnter={(e) => e.currentTarget.style.color = '#059669'}
        onMouseLeave={(e) => e.currentTarget.style.color = '#5f6368'}
      />
    </div>
  );
};

export default CommonShareComponent;
