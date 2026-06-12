import { useEffect } from 'react';
import { checkPermissions } from './PermissionService';
import { usePermissionStore } from './PermissionStore';
import { UserRole } from './PermissionStore';
export interface ICheckPermissionsResult {
    role: UserRole;
    canAdd: boolean;
    canEdit: boolean;
    canDelete: boolean;
    canView: boolean;
    isLoading: boolean;
}
export const useCheckPermissions = (context?: any): ICheckPermissionsResult => {
    const { role, canAdd, canEdit, canDelete, canView, isLoading } = usePermissionStore();
    useEffect(() => {
        // Only trigger if we are in the initial loading state and context is provided
        if (isLoading && role === 'None' && context) {
            checkPermissions(context).catch(err => console.error(err));
        }
    }, [isLoading, role, context]);
    return { role, canAdd, canEdit, canDelete, canView, isLoading };
};