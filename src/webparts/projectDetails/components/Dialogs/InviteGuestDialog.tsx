import * as React from 'react';
import { toast } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';
import styles from './InviteGuestDialog.module.scss';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Icon } from '@fluentui/react/lib/Icon';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { ProjectExternalPortalService } from '../../services/ProjectExternalPortalService';

export interface IInviteGuestDialogProps {
  isOpen: boolean;
  onClose: () => void;
  context: WebPartContext;
  projectId?: string;
  onGuestInvited?: () => void;
}

export interface IInviteGuestDialogState {
  guestName: string;
  guestEmail: string;
  inviting: boolean;
  error?: string;
}

export default class InviteGuestDialog extends React.Component<IInviteGuestDialogProps, IInviteGuestDialogState> {
  private portalService: ProjectExternalPortalService;

  constructor(props: IInviteGuestDialogProps) {
    super(props);
    
    this.state = {
      guestName: '',
      guestEmail: '',
      inviting: false
    };

    this.portalService = new ProjectExternalPortalService(props.context);
  }

  private _onGuestNameChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ guestName: newValue || '' });
  }

  private _onGuestEmailChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ guestEmail: newValue || '' });
  }

  private _validateEmail = (email: string): boolean => {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
  }

  private inviteExternalUser = async (guestName: string, userEmail: string): Promise<void> => {
    const redirectUrl = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/ExternalPortal.aspx`;
    
    try {
      const graphClient: MSGraphClientV3 = await this.props.context.msGraphClientFactory.getClient("3");
      const invitation = {
        invitedUserEmailAddress: userEmail,
        inviteRedirectUrl: redirectUrl,
        sendInvitationMessage: true,
        invitedUserDisplayName: guestName,
        inviteRedeemUrl: redirectUrl
      };

      
      const inviteResponse = await graphClient.api("/invitations").post(invitation);
      
      
      // Add invited user to SharePoint group via PnP (no sp-http import needed)
      if (inviteResponse && inviteResponse.invitedUser) {
        await this.portalService.addUserToSharePointGroup(userEmail, "ExternalUsers");
      }
    
      
      // Track the invitation in SharePoint for local records
      await this.portalService.inviteGuest({
        name: guestName,
        email: userEmail,
        projectId: this.props.projectId
      });
      
      return inviteResponse;
    } catch (error) {
      
      throw new Error(`Failed to send invitation: ${error.message}`);
    }
  }

  private _onInviteGuest = async (): Promise<void> => {
    const { guestName, guestEmail } = this.state;
    
    if (!guestName || !guestEmail) {
      this.setState({ error: 'Please enter name and email' });
      return;
    }

    if (!this._validateEmail(guestEmail)) {
      this.setState({ error: 'Please enter a valid email address' });
      return;
    }

    this.setState({ inviting: true, error: undefined });

    try {
      
      
      // Use Graph API to send real B2B invitation
      await this.inviteExternalUser(guestName, guestEmail);
      
      
      
      toast.success('External user invitation sent successfully! User has been added to the ExternalUsers group.', {
        position: "top-right",
        autoClose: 5000,
      });
      
      // Longer delay to ensure SharePoint processes the item creation and cache clearing
      await new Promise(resolve => setTimeout(resolve, 500));
      
      // Refresh the guests list
      
      if (this.props.onGuestInvited) {
        
        this.props.onGuestInvited();
      } else {
        
      }
      
      // Small delay before closing to ensure refresh completes
      await new Promise(resolve => setTimeout(resolve, 100));
      this.props.onClose();
    } catch (error) {
      this.setState({ error: error.message, inviting: false });
      toast.error(`Failed to invite guest: ${error.message}`, {
        position: "top-right",
        autoClose: 5000,
      });
    }
  }

  public render(): React.ReactElement<IInviteGuestDialogProps> {
    const { isOpen, onClose } = this.props;
    const { guestName, guestEmail, inviting, error } = this.state;



    return (
      <Dialog
        hidden={!isOpen}
        onDismiss={onClose}
        dialogContentProps={{
          type: DialogType.close,
          title: 'Invite Guest',
          showCloseButton: true
        }}
        modalProps={{
          isBlocking: false,
          className: styles.dialog
        }}
      >
        <div className={styles.dialogContent}>
          <TextField
            label="Guest Name"
            placeholder="Enter guest name"
            value={guestName}
            onChange={this._onGuestNameChange}
            required
            className={styles.field}
          />

          <TextField
            label="Email Address"
            placeholder="guest@company.com"
            value={guestEmail}
            onChange={this._onGuestEmailChange}
            required
            type="email"
            className={styles.field}
          />



          <div className={styles.infoBox}>
            <Icon iconName="Info" className={styles.infoIcon} />
            <p className={styles.infoText}>
              The external user will receive an email invitation to join your organization and access the project portal. 
              This creates a secure B2B collaboration account for the guest user.
            </p>
          </div>

          {error && (
            <div className={styles.errorMessage}>
              <Icon iconName="Error" className={styles.errorIcon} />
              {error}
            </div>
          )}
        </div>

        <DialogFooter>
          <PrimaryButton
            text="Send Invitation"
            onClick={this._onInviteGuest}
            disabled={inviting}
            className={styles.primaryButton}
          />
          <DefaultButton
            text="Cancel"
            onClick={onClose}
            disabled={inviting}
          />
        </DialogFooter>
      </Dialog>
    );
  }
}
