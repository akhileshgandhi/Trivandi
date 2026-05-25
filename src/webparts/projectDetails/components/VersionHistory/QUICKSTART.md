# Version History Component - Quick Start Guide

## What's Included?

✅ **VersionHistory.tsx** - Main React component with SharePoint-style UI  
✅ **VersionHistory.module.scss** - Professional styling  
✅ **VersionHistoryService.ts** - PnPjs + REST API service  
✅ **IVersionHistory.ts** - TypeScript interfaces  
✅ **INTEGRATION.md** - Detailed integration steps  
✅ **API.md** - Complete API documentation  

## 5-Minute Integration

### 1. Import Component
```typescript
import VersionHistory from '@/components/VersionHistory/VersionHistory';
import { ISelectedDocument } from '@/components/VersionHistory/IVersionHistory';
```

### 2. Add State
```typescript
const [showVersionHistory, setShowVersionHistory] = useState(false);
const [selectedDocumentForVersions, setSelectedDocumentForVersions] = useState<ISelectedDocument | undefined>();
```

### 3. Add Handlers
```typescript
const _openVersionHistory = (doc: ISharedDocument) => {
  setSelectedDocumentForVersions({
    fileRef: doc.fileRef,
    name: doc.name,
    serverRelativeUrl: doc.fileRef,
    listItemId: doc.id
  });
  setShowVersionHistory(true);
};

const _closeVersionHistory = () => setShowVersionHistory(false);
```

### 4. Add Menu Item
```typescript
{
  key: 'version_history',
  text: 'Version History',
  iconProps: { iconName: 'History' },
  onClick: () => _openVersionHistory(doc),
  disabled: doc.isFolder
}
```

### 5. Render Component
```typescript
<VersionHistory
  context={context}
  isOpen={showVersionHistory}
  document={selectedDocumentForVersions}
  onDismiss={_closeVersionHistory}
/>
```

## Features

| Feature | Status |
|---------|--------|
| Display version history | ✅ |
| Sort by date/version | ✅ |
| File icons | ✅ |
| Current version highlight | ✅ |
| Delete single version | ✅ |
| Delete all versions | ✅ |
| Restore version | ✅ |
| Loading state | ✅ |
| Error handling | ✅ |
| Empty state | ✅ |
| Responsive design | ✅ |
| PnPjs support | ✅ |
| REST API fallback | ✅ |

## Component Props

```typescript
interface IVersionHistoryProps {
  context: WebPartContext;        // SPFx context
  isOpen: boolean;                // Show/hide dialog
  document?: ISelectedDocument;    // File to show versions for
  onDismiss: () => void;          // Close handler
  onVersionRestored?: (version) => void;  // Restore callback
  onVersionDeleted?: () => void;  // Delete callback
}
```

## What Gets Fetched?

```typescript
IFileVersion[] = [
  {
    id: 'v512',                           // Version ID
    versionNumber: 3,                     // Major version
    displayNumber: '3.0',                 // Display format
    created: Date,                        // When created
    modifiedBy: 'John Doe',               // Who made it
    modifiedByEmail: 'john@contoso.com',  // Their email
    size: 52000,                          // File size in bytes
    comments: 'Final review complete',    // Check-in comment
    isCurrentVersion: true,               // Is this the latest?
    isMinorVersion: false,                // Is this a minor version?
  },
  // ... more versions
]
```

## Usage Examples

### Basic Usage
```typescript
<VersionHistory
  context={context}
  isOpen={isOpen}
  document={selectedDoc}
  onDismiss={() => setIsOpen(false)}
/>
```

### With Callbacks
```typescript
<VersionHistory
  context={context}
  isOpen={isOpen}
  document={selectedDoc}
  onDismiss={() => setIsOpen(false)}
  onVersionRestored={(version) => {
    console.log('Restored:', version.displayNumber);
    refreshDocumentList();
  }}
  onVersionDeleted={() => {
    console.log('Version deleted');
    refreshDocumentList();
  }}
/>
```

## API Methods Used

The service provides these methods:

```typescript
service.getFileVersionsFromPnPjs(url)       // Fetch versions (PnPjs)
service.getFileVersionsFromREST(url)        // Fetch versions (REST)
service.deleteVersion(url, versionId)       // Delete one version
service.deleteAllVersions(url)               // Delete all except current
service.restoreVersion(url, versionId)      // Restore a version
service.getFileIcon(fileName)               // Get icon for file type
service.formatFileSize(bytes)               // Format bytes to KB/MB
service.formatDate(date)                    // Format date
```

## Key Features Explained

### 1. Sorting
Click "📅 Date" or "#️⃣ Version" to sort, click again to reverse order.

### 2. Current Version Highlight
The latest version is highlighted in blue with a "Current" badge.

### 3. File Icons
Automatically shows the correct icon based on file extension:
- PDF → FilePDF
- Word → WordDocument
- Excel → ExcelDocument
- PowerPoint → PowerPointDocument
- Images → ImagePixel
- Others → Page

### 4. Delete Operations
- **Delete Version**: Remove a specific version
- **Delete All**: Remove all versions except current
- Confirmation dialog prevents accidental deletion

### 5. Version Restore
Click "↩️ Restore" to restore a previous version.
(Note: This marks it as a new version in SharePoint)

## Styling

The component uses SharePoint colors and design patterns:
- **Primary Blue**: #0078d4
- **Error Red**: #e81123
- **Success Green**: #107c41
- **Warning Yellow**: #ffb900

Customize colors in `VersionHistory.module.scss`.

## File Size Format

```
0 Bytes
512 Bytes
1 KB
2.5 MB
1.2 GB
```

## Date Format

```
05/14/2026, 10:30:00 AM
```

## Common Tasks

### Task 1: Add Version History to Menu
```typescript
// In your menu items array
{
  key: 'version_history',
  text: 'Version History',
  iconProps: { iconName: 'History' },
  onClick: () => _openVersionHistory(doc)
}
```

### Task 2: Refresh After Version Change
```typescript
const _onVersionDeleted = () => {
  // Reload the document list
  _loadDocuments();
};
```

### Task 3: Show Notification on Restore
```typescript
const _onVersionRestored = (version: IFileVersion) => {
  showNotification(`Restored version ${version.displayNumber}`);
};
```

### Task 4: Check If File Has Versions
```typescript
const hasVersions = versions && versions.length > 1;
const canShowVersionHistory = !doc.isFolder && hasVersions;
```

## Troubleshooting

### Versions Not Loading
```
✓ Check serverRelativeUrl format: /sites/site/Shared Documents/file.pdf
✓ Verify user has Read permission
✓ Check browser console for error messages
```

### Delete Not Working
```
✓ Ensure user has Edit permission
✓ Current version cannot be deleted
✓ Check for version locks or retention policies
```

### Styling Issues
```
✓ Import SCSS module correctly
✓ Check Fluent UI is installed
✓ Verify CSS-in-JS settings
```

## Performance Tips

1. **For large version lists** (>100 versions): Consider adding pagination
2. **For slow networks**: Add retry logic with exponential backoff
3. **For better UX**: Show success/error toast notifications
4. **For caching**: Store versions in memory if accessing frequently

## Next Steps

1. ✅ Copy component files to your project
2. ✅ Update imports in SharedFiles.tsx
3. ✅ Add state and handlers
4. ✅ Render the component
5. ✅ Test version operations
6. ✅ Deploy to production

## File Structure

```
VersionHistory/
├── VersionHistory.tsx              (Main component - 380 lines)
├── VersionHistory.module.scss      (Styles - 520 lines)
├── VersionHistoryService.ts        (Service - 290 lines)
├── IVersionHistory.ts              (Interfaces - 70 lines)
├── index.ts                        (Exports)
├── INTEGRATION.md                  (Integration guide)
├── API.md                          (API documentation)
└── QUICKSTART.md                   (This file)
```

## Support Resources

📖 **INTEGRATION.md** - Step-by-step integration  
📚 **API.md** - Complete API reference  
🔧 **VersionHistoryService.ts** - Service implementation  
💅 **VersionHistory.module.scss** - Styling details  

## Questions?

Check the comments in each file for detailed explanations of:
- PnPjs methods
- REST API endpoints
- State management
- Error handling
- Styling approaches

## Version

**Current Version**: 1.0.0  
**Last Updated**: May 2026  
**Status**: ✅ Production Ready  
**TypeScript**: ✅ Full support  
**SPFx**: ✅ Compatible  

---

Ready to integrate? Start with **INTEGRATION.md** for detailed steps!
