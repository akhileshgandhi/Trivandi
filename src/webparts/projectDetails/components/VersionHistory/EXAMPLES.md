# Version History Component - Code Examples

## Complete Integration Example

### Step 1: Update SharedFiles.tsx Imports

```typescript
import * as React from 'react';
import { useState, useEffect } from 'react';
import VersionHistory from './VersionHistory/VersionHistory';
import { ISelectedDocument, IFileVersion } from './VersionHistory/IVersionHistory';
import { ISharedDocument } from '../IProjectExternalPortalState';
```

### Step 2: Add State

```typescript
const [showVersionHistory, setShowVersionHistory] = useState(false);
const [selectedDocumentForVersions, setSelectedDocumentForVersions] = useState<ISelectedDocument | undefined>();
```

### Step 3: Add Handlers

```typescript
const handleOpenVersionHistory = (doc: ISharedDocument): void => {
  setSelectedDocumentForVersions({
    fileRef: doc.fileRef,
    name: doc.name,
    serverRelativeUrl: doc.fileRef,
    listItemId: doc.id
  });
  setShowVersionHistory(true);
};

const handleCloseVersionHistory = (): void => {
  setShowVersionHistory(false);
  setSelectedDocumentForVersions(undefined);
};

const handleVersionRestored = (version: IFileVersion): void => {
  // Show notification
  alert(`Version ${version.displayNumber} has been restored`);
  // Optionally refresh documents
  // _loadDocuments();
};

const handleVersionDeleted = (): void => {
  // Show notification
  alert('Version has been deleted');
  // Optionally refresh documents
  // _loadDocuments();
};
```

### Step 4: Add to Menu Items

```typescript
{
  key: 'version_history',
  text: 'Version History',
  iconProps: { iconName: 'History' },
  onClick: () => handleOpenVersionHistory(doc),
  disabled: doc.isFolder || isUserRestricted,
  title: 'View file version history'
},
{
  key: 'divider_1',
  divider: true
},
{
  key: 'preview',
  text: 'Preview',
  iconProps: { iconName: 'View' },
  onClick: () => _openFilePreview(doc),
  disabled: doc.isFolder || !canPreviewDoc(doc)
},
// ... other items
```

### Step 5: Render Component

```typescript
// Add to your JSX return statement
<VersionHistory
  context={context}
  isOpen={showVersionHistory}
  document={selectedDocumentForVersions}
  onDismiss={handleCloseVersionHistory}
  onVersionRestored={handleVersionRestored}
  onVersionDeleted={handleVersionDeleted}
/>
```

## Service Usage Examples

### Example 1: Basic Version Fetching

```typescript
import { VersionHistoryService } from './VersionHistory/VersionHistoryService';

const service = new VersionHistoryService(context);

// Fetch versions using PnPjs
const versions = await service.getFileVersionsFromPnPjs(
  '/sites/contoso/Shared Documents/Report.docx'
);

console.log(`Found ${versions.length} versions`);
versions.forEach(v => {
  console.log(`${v.displayNumber}: ${v.modifiedBy} - ${v.size} bytes`);
});
```

### Example 2: With Error Handling

```typescript
try {
  const versions = await service.getFileVersionsFromPnPjs(fileUrl);
  setVersions(versions);
} catch (error) {
  // Fallback to REST API
  try {
    const versions = await service.getFileVersionsFromREST(fileUrl);
    setVersions(versions);
  } catch (restError) {
    setError('Failed to load versions from both PnPjs and REST API');
    console.error('Version loading failed:', restError);
  }
}
```

### Example 3: Delete Single Version

```typescript
const deleteVersion = async (versionId: string) => {
  try {
    await service.deleteVersion(fileUrl, versionId);
    // Reload versions
    const updated = await service.getFileVersionsFromPnPjs(fileUrl);
    setVersions(updated);
  } catch (error) {
    console.error('Failed to delete version:', error);
  }
};
```

### Example 4: Delete All Versions

```typescript
const deleteAllVersions = async () => {
  const confirmed = window.confirm(
    'This will delete all previous versions. This action cannot be undone. Continue?'
  );

  if (!confirmed) return;

  try {
    await service.deleteAllVersions(fileUrl);
    // Reload versions
    const updated = await service.getFileVersionsFromPnPjs(fileUrl);
    setVersions(updated);
  } catch (error) {
    console.error('Failed to delete all versions:', error);
  }
};
```

### Example 5: File Icon Mapping

```typescript
const fileExtensions = ['report.pdf', 'data.xlsx', 'image.png'];

fileExtensions.forEach(fileName => {
  const icon = service.getFileIcon(fileName);
  console.log(`${fileName} -> ${icon}`);
});

// Output:
// report.pdf -> FilePDF
// data.xlsx -> ExcelDocument
// image.png -> ImagePixel
```

### Example 6: Format Utilities

```typescript
// Format file size
const sizes = [0, 512, 1024, 1048576, 1073741824];
sizes.forEach(bytes => {
  console.log(`${bytes} bytes = ${service.formatFileSize(bytes)}`);
});

// Output:
// 0 bytes = 0 Bytes
// 512 bytes = 512 Bytes
// 1024 bytes = 1 KB
// 1048576 bytes = 1 MB
// 1073741824 bytes = 1 GB

// Format date
const date = new Date('2026-05-14T10:30:00Z');
console.log(service.formatDate(date));
// Output: 05/14/2026, 10:30:00 AM
```

## Advanced Examples

### Example 7: Custom Notification System

```typescript
interface INotification {
  id: string;
  message: string;
  type: 'success' | 'error' | 'warning' | 'info';
  duration?: number;
}

const [notifications, setNotifications] = useState<INotification[]>([]);

const showNotification = (message: string, type: 'success' | 'error' = 'success') => {
  const id = Math.random().toString(36);
  const notification: INotification = { id, message, type, duration: 3000 };
  
  setNotifications(prev => [...prev, notification]);
  
  if (notification.duration) {
    setTimeout(() => {
      setNotifications(prev => prev.filter(n => n.id !== id));
    }, notification.duration);
  }
};

// Use in handlers
const handleVersionDeleted = () => {
  showNotification('Version deleted successfully', 'success');
};

const handleVersionError = (error: string) => {
  showNotification(`Error: ${error}`, 'error');
};
```

### Example 8: Filtering Versions

```typescript
interface VersionFilter {
  author?: string;
  startDate?: Date;
  endDate?: Date;
  minSize?: number;
  maxSize?: number;
}

const filterVersions = (
  versions: IFileVersion[],
  filter: VersionFilter
): IFileVersion[] => {
  return versions.filter(v => {
    if (filter.author && v.modifiedBy !== filter.author) return false;
    if (filter.startDate && v.created < filter.startDate) return false;
    if (filter.endDate && v.created > filter.endDate) return false;
    if (filter.minSize && v.size < filter.minSize) return false;
    if (filter.maxSize && v.size > filter.maxSize) return false;
    return true;
  });
};

// Usage
const recentVersions = filterVersions(versions, {
  startDate: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000) // Last 7 days
});
```

### Example 9: Version Comparison

```typescript
interface IVersionComparison {
  before: IFileVersion;
  after: IFileVersion;
  sizeChange: number;
  sizeChangePercent: number;
  daysBetween: number;
}

const compareVersions = (
  before: IFileVersion,
  after: IFileVersion
): IVersionComparison => {
  const sizeChange = after.size - before.size;
  const sizeChangePercent = (sizeChange / before.size) * 100;
  const daysBetween = Math.floor(
    (after.created.getTime() - before.created.getTime()) / (1000 * 60 * 60 * 24)
  );

  return {
    before,
    after,
    sizeChange,
    sizeChangePercent,
    daysBetween
  };
};

// Usage
const comparison = compareVersions(versions[1], versions[0]);
console.log(`File grew by ${comparison.sizeChange} bytes in ${comparison.daysBetween} days`);
```

### Example 10: Export Versions to CSV

```typescript
const exportVersionsToCSV = (versions: IFileVersion[], fileName: string) => {
  const csv = [
    ['Version', 'Modified Date', 'Modified By', 'Size (KB)', 'Comments'],
    ...versions.map(v => [
      v.displayNumber,
      v.created.toISOString(),
      v.modifiedBy,
      (v.size / 1024).toFixed(2),
      v.comments
    ])
  ]
    .map(row => row.join(','))
    .join('\n');

  const blob = new Blob([csv], { type: 'text/csv' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${fileName}-versions.csv`;
  a.click();
  window.URL.revokeObjectURL(url);
};

// Usage
exportVersionsToCSV(versions, 'report');
```

## Testing Examples

### Test 1: Unit Test for Service

```typescript
import { VersionHistoryService } from './VersionHistoryService';

describe('VersionHistoryService', () => {
  let service: VersionHistoryService;
  let mockContext: any;

  beforeEach(() => {
    mockContext = {
      pageContext: {
        web: { absoluteUrl: 'https://contoso.sharepoint.com/sites/demo' }
      }
    };
    service = new VersionHistoryService(mockContext);
  });

  test('formatFileSize formats bytes correctly', () => {
    expect(service.formatFileSize(0)).toBe('0 Bytes');
    expect(service.formatFileSize(1024)).toBe('1 KB');
    expect(service.formatFileSize(1048576)).toBe('1 MB');
  });

  test('getFileIcon returns correct icon', () => {
    expect(service.getFileIcon('document.pdf')).toBe('FilePDF');
    expect(service.getFileIcon('sheet.xlsx')).toBe('ExcelDocument');
    expect(service.getFileIcon('image.png')).toBe('ImagePixel');
  });

  test('formatDate returns formatted date string', () => {
    const date = new Date('2026-05-14T10:30:00Z');
    const formatted = service.formatDate(date);
    expect(formatted).toContain('2026');
    expect(formatted).toContain('05');
  });
});
```

### Test 2: Integration Test

```typescript
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import VersionHistory from './VersionHistory';

describe('VersionHistory Component', () => {
  const mockContext = { /* ... */ };
  const mockDocument = {
    fileRef: '/sites/demo/Shared Documents/file.pdf',
    name: 'file.pdf',
    serverRelativeUrl: '/sites/demo/Shared Documents/file.pdf'
  };

  test('renders dialog when isOpen is true', () => {
    render(
      <VersionHistory
        context={mockContext}
        isOpen={true}
        document={mockDocument}
        onDismiss={() => {}}
      />
    );

    expect(screen.getByText(/Version History/i)).toBeInTheDocument();
  });

  test('loads and displays versions', async () => {
    render(
      <VersionHistory
        context={mockContext}
        isOpen={true}
        document={mockDocument}
        onDismiss={() => {}}
      />
    );

    await waitFor(() => {
      expect(screen.queryByText(/Loading/i)).not.toBeInTheDocument();
    });

    // Expect version rows
    expect(screen.getByText(/Version/i)).toBeInTheDocument();
  });

  test('calls onDismiss when close button clicked', () => {
    const onDismiss = jest.fn();
    render(
      <VersionHistory
        context={mockContext}
        isOpen={true}
        document={mockDocument}
        onDismiss={onDismiss}
      />
    );

    fireEvent.click(screen.getByText('Close'));
    expect(onDismiss).toHaveBeenCalled();
  });
});
```

## Real-World Scenario

### Scenario: Document Management Portal

```typescript
// DocumentListView.tsx
import VersionHistory from '@/components/VersionHistory/VersionHistory';

const DocumentListView: React.FC = () => {
  const [selectedDoc, setSelectedDoc] = useState<ISelectedDocument>();
  const [showVersions, setShowVersions] = useState(false);
  const [lastAction, setLastAction] = useState<string>('');

  const handleViewVersions = (doc: ISharedDocument) => {
    setSelectedDoc({
      fileRef: doc.fileRef,
      name: doc.name,
      serverRelativeUrl: doc.fileRef
    });
    setShowVersions(true);
  };

  const handleVersionAction = (action: string) => {
    setLastAction(action);
    setShowVersions(false);
    // Refresh document list
    refreshDocuments();
    // Show notification
    showNotification(`${action} completed successfully`);
  };

  return (
    <>
      <div className="document-list">
        {documents.map(doc => (
          <div key={doc.id} className="document-item">
            <span>{doc.name}</span>
            <button onClick={() => handleViewVersions(doc)}>
              📜 View Versions
            </button>
          </div>
        ))}
      </div>

      <VersionHistory
        context={context}
        isOpen={showVersions}
        document={selectedDoc}
        onDismiss={() => setShowVersions(false)}
        onVersionRestored={() => handleVersionAction('Version restored')}
        onVersionDeleted={() => handleVersionAction('Version deleted')}
      />
    </>
  );
};

export default DocumentListView;
```

---

## More Examples?

Check the comment sections in each source file for additional examples and explanations:
- **VersionHistory.tsx** - Component patterns and hooks
- **VersionHistoryService.ts** - PnPjs and REST API examples
- **VersionHistory.module.scss** - Styling patterns

