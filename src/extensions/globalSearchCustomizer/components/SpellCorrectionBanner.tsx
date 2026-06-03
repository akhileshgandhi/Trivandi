/* eslint-disable @rushstack/no-new-null */
import * as React from 'react';

export interface ISpellCorrectionBannerProps {
  correctedQuery: string | null;
  originalQuery: string;
  onUseOriginal: () => void;
}

export const SpellCorrectionBanner: React.FC<ISpellCorrectionBannerProps> = ({
  correctedQuery,
  originalQuery,
  onUseOriginal
}) => {
  if (!correctedQuery) return null;
  
  const cleanOriginal = originalQuery.trim();
  const cleanCorrected = correctedQuery.trim();

  // Only render if they differ (case-insensitively)
  if (cleanOriginal.toLowerCase() === cleanCorrected.toLowerCase()) {
    return null;
  }

  return (
    <div style={{
      backgroundColor: '#F0F4FF',
      color: '#0C447C',
      fontSize: '13px',
      padding: '8px 16px',
      borderRadius: '8px',
      margin: '8px 24px 16px 24px',
      display: 'flex',
      alignItems: 'center',
      gap: '8px',
      fontWeight: 500,
      fontFamily: 'sans-serif'
    }}>
      <span>Showing results for <strong>{cleanCorrected}</strong>.</span>
      <span>
        Search instead for{' '}
        <button 
          onClick={onUseOriginal}
          style={{
            background: 'none',
            border: 'none',
            color: '#1a73e8',
            textDecoration: 'underline',
            cursor: 'pointer',
            padding: 0,
            font: 'inherit',
            fontWeight: 600
          }}
        >
          {cleanOriginal}
        </button>
      </span>
    </div>
  );
};
