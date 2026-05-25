import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';

interface ILoaderProps {
    variant?: 'content' | 'full';
}

const Loader: React.FC<ILoaderProps> = ({ variant = 'content' }) => {
    return (
        <div style={{ padding: variant === 'full' ? '100px 0' : '40px 0', textAlign: 'center' }}>
            <Spinner size={SpinnerSize.large} label="Loading..." />
        </div>
    );
};

export default Loader;
