import * as React from 'react';

interface ISearchErrorBoundaryState {
  hasError: boolean;
  error: Error | null;
}

export class SearchErrorBoundary extends React.Component<React.PropsWithChildren<{}>, ISearchErrorBoundaryState> {
  constructor(props: React.PropsWithChildren<{}>) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  public static getDerivedStateFromError(error: Error): ISearchErrorBoundaryState {
    return { hasError: true, error };
  }

  public componentDidCatch(error: Error, errorInfo: React.ErrorInfo): void {
    console.error('SearchModal crashed:', error, errorInfo);
  }

  public render(): React.ReactNode {
    if (this.state.hasError) {
      return (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          width: '100vw',
          height: '100vh',
          backgroundColor: 'rgba(255, 255, 255, 0.98)',
          zIndex: 999999,
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          justifyContent: 'center',
          fontFamily: '"Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif',
          color: '#334155',
          padding: '24px',
          boxSizing: 'border-box',
          textAlign: 'center'
        }}>
          <div style={{
            fontSize: '48px',
            marginBottom: '16px'
          }}>⚠️</div>
          <h2 style={{
            fontSize: '22px',
            fontWeight: 700,
            margin: '0 0 10px 0',
            color: '#0f172a'
          }}>Search Service is Temporarily Offline</h2>
          <p style={{
            fontSize: '14px',
            color: '#64748b',
            maxWidth: '480px',
            margin: '0 0 24px 0',
            lineHeight: 1.5
          }}>
            An unexpected error occurred while loading the search interface. Please try refreshing the page.
          </p>
          <button 
            onClick={() => window.location.reload()}
            style={{
              padding: '10px 24px',
              backgroundColor: '#1a73e8',
              color: '#ffffff',
              border: 'none',
              borderRadius: '6px',
              fontSize: '14px',
              fontWeight: 600,
              cursor: 'pointer',
              boxShadow: '0 2px 4px rgba(26, 115, 232, 0.2)',
              transition: 'background-color 0.2s'
            }}
            onMouseOver={(e) => (e.currentTarget.style.backgroundColor = '#1557b0')}
            onMouseOut={(e) => (e.currentTarget.style.backgroundColor = '#1a73e8')}
          >
            Refresh Page
          </button>
        </div>
      );
    }

    return this.props.children;
  }
}
