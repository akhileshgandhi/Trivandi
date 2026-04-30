import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';

export interface ILoaderProps {
    variant?: 'full' | 'content';
    message?: string;
}

const Loader: React.FC<ILoaderProps> = (props) => {
    const { variant = 'content', message = 'Loading...' } = props;
    
    const containerStyle: React.CSSProperties = variant === 'full' 
        ? { 
            position: 'fixed', 
            top: 0, 
            left: 0, 
            right: 0, 
            bottom: 0, 
            background: 'rgba(255,255,255,0.7)', 
            display: 'flex', 
            alignItems: 'center', 
            justifyContent: 'center',
            zIndex: 9999
          }
        : {
            padding: '40px',
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            width: '100%'
          };

    return (
        <div style={containerStyle}>
            <Spinner size={SpinnerSize.large} label={message} />
        </div>
    );
};

export default Loader;
