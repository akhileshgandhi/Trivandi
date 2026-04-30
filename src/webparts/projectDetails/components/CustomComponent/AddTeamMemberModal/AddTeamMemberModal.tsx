import * as React from "react";
import { useState, useRef, useEffect } from "react";
import { toast } from "react-toastify";
import "react-toastify/dist/ReactToastify.css";
import styles from "./AddTeamMemberModal.module.scss";

interface IAddTeamMemberModalProps {
    isOpen: boolean;
    onClose: () => void;
    onAddMember: (member: { name: string; email: string; role: string; mobileNumber: string }) => Promise<void>;
    projectId: any;
    projectCode: string;
    ProjectTitle: string;
}

const AddTeamMemberModal: React.FC<IAddTeamMemberModalProps> = ({
    isOpen,
    onClose,
    onAddMember,
    projectId,
    projectCode,
    ProjectTitle
}) => {
    const [name, setName] = useState("");
    const [email, setEmail] = useState("");
    const [mobileNumber, setMobileNumber] = useState("");
    const [role, setRole] = useState("Contributor (Edit Files)");
    const [isSubmitting, setIsSubmitting] = useState(false);
    const [error, setError] = useState("");
    const [showCopied, setShowCopied] = useState(false);
    const modalRef = useRef<HTMLDivElement>(null);

    const roleOptions = [
        "Contributor",
        "Reader",
        "Administrator",
        "Project Manager",
        "Team Member"
    ];

    useEffect(() => {
        if (isOpen) {
            // Reset form when modal opens
            setName("");
            setEmail("");
            setMobileNumber("");
            setRole("Contributor (Edit Files)");
            setError("");
        }
    }, [isOpen]);

    const handleClose = (): void => {
        if (!isSubmitting) {
            onClose();
        }
    };

    const validateEmail = (email: string): boolean => {
        const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        return re.test(email);
    };

    const handleCopyMobile = async (): Promise<void> => {
        if (mobileNumber) {
            try {
                await navigator.clipboard.writeText(mobileNumber);
                setShowCopied(true);
                setTimeout(() => setShowCopied(false), 2000);
            } catch (err) {
                console.error('Failed to copy mobile number:', err);
                // Fallback for older browsers
                const textArea = document.createElement('textarea');
                textArea.value = mobileNumber;
                textArea.style.position = 'fixed';
                textArea.style.left = '-999999px';
                document.body.appendChild(textArea);
                textArea.select();
                try {
                    document.execCommand('copy');
                    setShowCopied(true);
                    setTimeout(() => setShowCopied(false), 2000);
                } catch (e) {
                    console.error('Fallback copy failed:', e);
                }
                document.body.removeChild(textArea);
            }
        }
    };

    const handleCallMobile = (): void => {
        if (mobileNumber) {
            window.location.href = `tel:${mobileNumber}`;
        }
    };

    const handleSubmit = async (e: React.FormEvent): Promise<void> => {
        e.preventDefault();
        setError("");

        // Validation
        if (!name.trim()) {
            setError("Name is required");
            return;
        }

        if (!email.trim()) {
            setError("Email address is required");
            return;
        }

        if (!validateEmail(email)) {
            setError("Please enter a valid email address");
            return;
        }

        setIsSubmitting(true);

        try {
            await onAddMember({ name: name.trim(), email: email.trim(), role, mobileNumber: mobileNumber.trim() });
            toast.success("Team member added and invitation sent successfully!", {
                position: "top-right",
                autoClose: 3000,
            });
            onClose();
        } catch (err) {
            const errorMessage = err instanceof Error ? err.message : "Failed to add team member";
            setError(errorMessage);
            toast.error(errorMessage, {
                position: "top-right",
                autoClose: 5000,
            });
        } finally {
            setIsSubmitting(false);
        }
    };

    if (!isOpen) return null;

    return (
        <div className={styles.modalOverlay} onClick={handleClose}>
            <div
                ref={modalRef}
                className={styles.modalContent}
                onClick={(e) => e.stopPropagation()}
            >
                <div className={styles.modalHeader}>
                    <h2 className={styles.modalTitle}>Add Team Member</h2>
                    <button
                        className={styles.closeButton}
                        onClick={handleClose}
                        disabled={isSubmitting}
                        type="button"
                    >
                        ×
                    </button>
                </div>

                <form onSubmit={handleSubmit}>
                    <div className={styles.modalBody}>
                        <div className={styles.formGroup}>
                            <label htmlFor="memberName" className={styles.label}>
                                Name
                            </label>
                            <input
                                id="memberName"
                                type="text"
                                className={styles.input}
                                placeholder="Enter name"
                                value={name}
                                onChange={(e) => setName(e.target.value)}
                                disabled={isSubmitting}
                            />
                        </div>

                        <div className={styles.formGroup}>
                            <label htmlFor="memberEmail" className={styles.label}>
                                Email Address
                            </label>
                            <input
                                id="memberEmail"
                                type="email"
                                className={styles.input}
                                placeholder="Enter email"
                                value={email}
                                onChange={(e) => setEmail(e.target.value)}
                                disabled={isSubmitting}
                            />
                        </div>

                        <div className={styles.formGroup}>
                            <label htmlFor="memberMobile" className={styles.label}>
                                Mobile Number
                            </label>
                            <div className={styles.inputWithButton}>
                                <input
                                    id="memberMobile"
                                    type="tel"
                                    className={styles.input}
                                    placeholder="Enter mobile number"
                                    value={mobileNumber}
                                    onChange={(e) => setMobileNumber(e.target.value)}
                                    disabled={isSubmitting}
                                />
                                {mobileNumber && (
                                    <div className={styles.actionButtons}>
                                        <button
                                            type="button"
                                            className={styles.callButton}
                                            onClick={handleCallMobile}
                                            disabled={isSubmitting}
                                            title="Call mobile number"
                                        >
                                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                                                <path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"/>
                                            </svg>
                                        </button>
                                        <button
                                            type="button"
                                            className={styles.copyButton}
                                            onClick={handleCopyMobile}
                                            disabled={isSubmitting}
                                            title="Copy mobile number"
                                        >
                                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                                                <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
                                                <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
                                            </svg>
                                        </button>
                                    </div>
                                )}
                            </div>
                            {showCopied && (
                                <div className={styles.copiedToast}>
                                    Copied!
                                </div>
                            )}
                        </div>

                        <div className={styles.formGroup}>
                            <label htmlFor="memberRole" className={styles.label}>
                                Role
                            </label>
                            <div className={styles.selectWrapper}>
                                <select
                                    id="memberRole"
                                    className={styles.select}
                                    value={role}
                                    onChange={(e) => setRole(e.target.value)}
                                    disabled={isSubmitting}
                                >
                                    {roleOptions.map((option) => (
                                        <option key={option} value={option}>
                                            {option}
                                        </option>
                                    ))}
                                </select>
                                <svg
                                    className={styles.selectIcon}
                                    width="12"
                                    height="8"
                                    viewBox="0 0 12 8"
                                    fill="none"
                                    xmlns="http://www.w3.org/2000/svg"
                                >
                                    <path
                                        d="M1 1.5L6 6.5L11 1.5"
                                        stroke="currentColor"
                                        strokeWidth="1.5"
                                        strokeLinecap="round"
                                        strokeLinejoin="round"
                                    />
                                </svg>
                            </div>
                        </div>

                        <div className={styles.infoMessage}>
                            This user will be added to {ProjectTitle} {projectCode}
                        </div>

                        {error && <div className={styles.errorMessage}>{error}</div>}
                    </div>

                    <div className={styles.modalFooter}>
                        <button
                            type="submit"
                            className={styles.submitButton}
                            disabled={isSubmitting}
                        >
                            {isSubmitting ? "Sending..." : "Send Invitation"}
                        </button>
                    </div>
                </form>
            </div>
        </div>
    );
};

export default AddTeamMemberModal;
