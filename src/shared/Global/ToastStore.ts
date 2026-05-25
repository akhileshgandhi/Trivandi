import { toast } from 'react-toastify';
import create from 'zustand';

interface ToastState {
  message: string;
  type: 'success' | 'error' | 'warning' | 'info';
  isOpen: boolean;
  showToast: (message: string, type?: 'success' | 'error' | 'warning' | 'info') => void;
  hideToast: () => void;
}

export const useToastStore = create<ToastState>((set) => ({
  message: '',
  type: 'info',
  isOpen: false,
  showToast: (message: string, type: 'success' | 'error' | 'warning' | 'info' = 'info') => {
    // Bridge to react-toastify since the project already uses it
    if (type === 'success') toast.success(message);
    else if (type === 'error') toast.error(message);
    else if (type === 'warning') toast.warn(message);
    else toast.info(message);

    set({ message, type, isOpen: true });
    // Auto-hide local state after 5 seconds
    setTimeout(() => {
      set({ isOpen: false });
    }, 5000);
  },
  hideToast: () => set({ isOpen: false }),
}));
