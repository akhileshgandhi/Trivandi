import create from 'zustand';

export type UserRole = 'Owner' | 'Contributor' | 'Viewer' | 'None';

export interface IPermissionState {
    role: UserRole;
    canAdd: boolean;
    canEdit: boolean;
    canDelete: boolean;
    canView: boolean;
    isOwner: boolean;
    isContributor: boolean;
    isViewer: boolean;
    isLoading: boolean;
    allowedSites: string[];
    setPermissions: (permissions: Partial<IPermissionState>) => void;
    reset: () => void;
}

export const usePermissionStore = create<IPermissionState>((set) => ({
    role: 'None',
    canAdd: false,
    canEdit: false,
    canDelete: false,
    canView: false,
    isOwner: false,
    isContributor: false,
    isViewer: false,
    isLoading: true,
    allowedSites: [],
    setPermissions: (permissions) => set((state) => ({ ...state, ...permissions })),
    reset: () => set({ 
        role: 'None', 
        canAdd: false, 
        canEdit: false, 
        canDelete: false, 
        canView: false, 
        isOwner: false, 
        isContributor: false, 
        isViewer: false, 
        isLoading: true,
        allowedSites: []
    }),
}));