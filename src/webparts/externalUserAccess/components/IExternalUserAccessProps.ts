import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IExternalUserAccessProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
   context: WebPartContext; // SPFx context for PnP initialization
}
