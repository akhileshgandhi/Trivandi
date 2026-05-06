
import { useEffect } from 'react';
import { checkPermissions } from './PermissionService';
import { usePermissionStore } from './PermissionStore';


export const useCheckPermissions = () => {
    const { role, canAdd, canEdit, canView, isLoading } = usePermissionStore();

    useEffect(() => {
        // Only trigger if we are in the initial loading state
        if (isLoading && role === 'None') {
            void checkPermissions();
        }
    }, [isLoading, role]);

    return { role, canAdd, canEdit, canView, isLoading };
};
