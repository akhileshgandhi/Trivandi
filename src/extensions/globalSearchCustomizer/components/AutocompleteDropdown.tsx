import * as React from 'react';
import { ISuggestion } from '../hooks/useAutocomplete';
import { Clock, Search, FileText } from 'lucide-react';

export interface IAutocompleteDropdownProps {
  suggestions: ISuggestion[];
  onSelect: (label: string) => void;
  activeIndex: number;
}

export const AutocompleteDropdown: React.FC<IAutocompleteDropdownProps> = ({
  suggestions,
  onSelect,
  activeIndex
}) => {
  if (!suggestions || suggestions.length === 0) return null;

  return (
    <div style={{
      position: 'absolute',
      top: '100%',
      left: 0,
      width: '100%',
      backgroundColor: '#ffffff',
      border: '1px solid #e0e0e0',
      borderRadius: '8px',
      boxShadow: '0 4px 12px rgba(0, 0, 0, 0.1)',
      zIndex: 1000,
      marginTop: '4px',
      overflow: 'hidden'
    }}>
      {suggestions.map((suggestion, idx) => {
        const isHighlighted = idx === activeIndex;
        
        let pillColor = '';
        let pillBg = '';
        let pillLabel = '';
        let IconComponent = Search;

        if (suggestion.type === 'history') {
          pillColor = '#6b7280';
          pillBg = '#f3f4f6';
          pillLabel = 'Recent';
          IconComponent = Clock;
        } else if (suggestion.type === 'file') {
          pillColor = '#7b1fa2';
          pillBg = '#f3e5f5';
          pillLabel = 'FILES';
          IconComponent = FileText;
        } else if (suggestion.type === 'suggested') {
          pillColor = '#16a34a';
          pillBg = '#f0fdf4';
          pillLabel = 'Suggested';
          IconComponent = Search;
        }

        return (
          <button
            type="button"
            key={`${suggestion.type}-${suggestion.label}`}
            onMouseDown={(e) => {
              e.preventDefault();
              onSelect(suggestion.label);
            }}
            onClick={() => onSelect(suggestion.label)}
            onMouseEnter={(e) => {
              (e.currentTarget as HTMLButtonElement).style.backgroundColor = '#F0F4FF';
            }}
            onMouseLeave={(e) => {
              if (!isHighlighted) {
                (e.currentTarget as HTMLButtonElement).style.backgroundColor = 'transparent';
              }
            }}
            style={{
              width: '100%',
              height: '36px',
              padding: '12px 16px',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              cursor: 'pointer',
              fontSize: '13px',
              fontFamily: 'sans-serif',
              backgroundColor: isHighlighted ? '#F0F4FF' : 'transparent',
              transition: 'background-color 0.2s ease',
              border: 'none',
              borderBottom: idx < suggestions.length - 1 ? '1px solid #f1f5f9' : 'none',
              boxSizing: 'border-box',
              textAlign: 'left'
            }}
          >
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
              <IconComponent size={14} color="#64748b" />
              <span style={{ color: '#334155', fontWeight: 500 }}>{suggestion.label}</span>
            </div>
            
            <div style={{
              backgroundColor: pillBg,
              color: pillColor,
              padding: '2px 8px',
              borderRadius: '12px',
              fontSize: '11px',
              fontWeight: 600,
              textTransform: 'uppercase',
              letterSpacing: '0.5px'
            }}>
              {pillLabel}
            </div>
          </button>
        );
      })}
    </div>
  );
};
