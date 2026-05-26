import * as React from 'react';
import styles from '../../../styles/PremiumSearch.module.scss';

import { ISkeletonLoaderProps } from '../interface/ISkeletonLoaderProps';

export const SkeletonLoader: React.FC<ISkeletonLoaderProps> = ({ count = 4 }) => {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', width: '100%', padding: '4px' }}>
      {Array.from({ length: count }).map((_, i) => (
        <div key={i} className={styles.resultCard} style={{ opacity: 0.6, pointerEvents: 'none', borderLeftColor: '#cbd5e1' }}>
          <div className={styles.cardLeftBlock}>
            <div className={styles.cardIconBox} style={{ backgroundColor: '#e2e8f0', width: '48px', height: '48px' }} />
          </div>
          <div className={styles.cardBody} style={{ display: 'flex', flexDirection: 'column', gap: '8px', padding: '4px 0' }}>
            <div style={{ height: '20px', width: '50%', backgroundColor: '#e2e8f0', borderRadius: '4px' }} />
            <div style={{ height: '14px', width: '35%', backgroundColor: '#f1f5f9', borderRadius: '4px', marginTop: '4px' }} />
            <div style={{ height: '14px', width: '90%', backgroundColor: '#f1f5f9', borderRadius: '4px', marginTop: '8px' }} />
            <div style={{ height: '14px', width: '80%', backgroundColor: '#f1f5f9', borderRadius: '4px' }} />
          </div>
        </div>
      ))}
    </div>
  );
};
