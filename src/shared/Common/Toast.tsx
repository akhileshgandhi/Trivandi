import * as React from 'react';
import { useToastStore } from '../Global/ToastStore';
import styles from './Toast.module.scss';
import { CheckCircle, XCircle, AlertCircle, Info, X } from 'lucide-react';

const Toast = () => {
    const { message, type, isOpen, hideToast } = useToastStore();

    if (!isOpen) return null;

    const getIcon = () => {
        switch (type) {
            case 'success': return <CheckCircle size={20} />;
            case 'error': return <XCircle size={20} />;
            case 'warning': return <AlertCircle size={20} />;
            case 'info': return <Info size={20} />;
        }
    };

    return (
        <div className={`${styles.toastContainer} ${styles[type]}`}>
            <div className={styles.icon}>{getIcon()}</div>
            <div className={styles.message}>{message}</div>
            <button className={styles.closeBtn} onClick={hideToast}>
                <X size={16} />
            </button>
        </div>
    );
};

export default Toast;
