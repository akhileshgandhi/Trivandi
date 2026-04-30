import * as React from 'react';
import { useState } from 'react';
import styles from './Dashboard.module.scss';

interface KeyPersonCardProps {
  name: string | null;
  title: string | null;
  imageUrl: string | null;
  email: string | null;
  phone: string | null;
}

const KeyPersonCard: React.FC<KeyPersonCardProps> = ({ name, title, imageUrl, email, phone }) => {
  const [showCopied, setShowCopied] = useState(false);

  const handleEmailClick = (): void => {
    if (email) {
      window.location.href = `mailto:${email}`;
    }
  };

  const handlePhoneClick = async (): Promise<void> => {
    if (phone) {
      try {
        await navigator.clipboard.writeText(phone);
        setShowCopied(true);
        setTimeout(() => setShowCopied(false), 2000);
      } catch (err) {
        console.error('Failed to copy phone number:', err);
        // Fallback for older browsers
        const textArea = document.createElement('textarea');
        textArea.value = phone;
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

  return (
    <div className={styles.personCard}>
      {showCopied && (
        <div className={styles.copiedToast}>
          Copied!
        </div>
      )}
      <div className={styles.personImageWrapper}>
        {imageUrl ? (
          <img src={imageUrl} alt={name || ''} className={styles.personImage} />
        ) : (
          <div className={styles.personImagePlaceholder}>
            {name ? name.split(' ').map(n => n[0]).join('') : '?'}
          </div>
        )}
      </div>
      <div className={styles.personInfo}>
        <h4 className={styles.personName}>{name || 'Unknown'}</h4>
        <p className={styles.personTitle}>{title || ''}</p>
        <div className={styles.personActions}>
          {email && (
            <button 
              className={styles.actionIcon} 
              title="Send Email" 
              aria-label="Send email"
              onClick={handleEmailClick}
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <rect x="2" y="4" width="20" height="16" rx="2"/>
                <path d="m2 7 10 7 10-7"/>
              </svg>
            </button>
          )}
          {phone && (
            <button 
              className={styles.actionIcon} 
              title="Copy Phone Number" 
              aria-label="Copy phone number"
              onClick={handlePhoneClick}
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"/>
              </svg>
            </button>
          )}
        </div>
      </div>
    </div>
  );
};

export default KeyPersonCard;
