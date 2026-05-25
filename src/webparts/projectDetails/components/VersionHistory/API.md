# Version History API Documentation

## Overview
This document describes all API endpoints, PnPjs methods, and REST calls used in the Version History component.

## SharePoint Version History APIs

### 1. Get All File Versions

#### PnPjs Method
```typescript
const file = sp.web.getFileByServerRelativePath('/sites/site/lib/file.docx');
const versions = await file.versions();
```

**Response Type**: `IFileVersion[]`

#### REST API
```
GET /_api/web/GetFileByServerRelativePath(decodedurl='/sites/site/lib/file.docx')/Versions
```

**Headers**:
```
Accept: application/json
Content-Type: application/json
X-RequestDigest: [FORM DIGEST TOKEN]
```

**Response**:
```json
{
  "value": [
    {
      "ID": 512,
      "VersionLabel": "2.0",
      "Created": "2026-05-14T10:30:00Z",
      "CreatedBy": {
        "Title": "John Doe",
        "EMail": "john@contoso.com"
      },
      "CheckInComment": "Final version",
      "Size": 45678
    },
    {
      "ID": 256,
      "VersionLabel": "1.0",
      "Created": "2026-05-13T15:45:00Z",
      "CreatedBy": {
        "Title": "Jane Smith",
        "EMail": "jane@contoso.com"
      },
      "CheckInComment": "Initial upload",
      "Size": 42000
    }
  ]
}
```

### 2. Get File Properties

#### PnPjs Method
```typescript
const file = sp.web.getFileByServerRelativePath('/sites/site/lib/file.docx');
const item = await file
  .select('Length', 'TimeLastModified', 'Author/Title', 'Author/EMail')
  .expand('Author')
  .getItem();
```

**Response**:
```typescript
{
  Length: 45678,
  TimeLastModified: "2026-05-14T10:30:00Z",
  Author: {
    Title: "John Doe",
    EMail: "john@contoso.com"
  }
}
```

#### REST API
```
GET /_api/web/GetFileByServerRelativePath(decodedurl='/sites/site/lib/file.docx')/ListItemAllFields?$select=Length,TimeLastModified,Author/Title,Author/EMail&$expand=Author
```

### 3. Delete a Specific Version

#### PnPjs Method
```typescript
const file = sp.web.getFileByServerRelativePath('/sites/site/lib/file.docx');
await file.versions.getById(256).delete();
```

**Returns**: 204 No Content

#### REST API
```
POST /_api/web/GetFileByServerRelativePath(decodedurl='/sites/site/lib/file.docx')/Versions(256)/Delete
```

**Headers**:
```
Accept: application/json
Content-Type: application/json
X-HTTP-Method: DELETE
X-RequestDigest: [FORM DIGEST TOKEN]
```

### 4. Delete All Versions

#### PnPjs Method
```typescript
const file = sp.web.getFileByServerRelativePath('/sites/site/lib/file.docx');
const versions = await file.versions();

// Delete all except current (index 0)
for (let i = 1; i < versions.length; i++) {
  await file.versions.getById(versions[i].ID).delete();
}
```

#### REST API
```
// Make individual DELETE calls for each version
POST /_api/web/GetFileByServerRelativePath(decodedurl='/sites/site/lib/file.docx')/Versions(256)/Delete
POST /_api/web/GetFileByServerRelativePath(decodedurl='/sites/site/lib/file.docx')/Versions(128)/Delete
// ... for each version ID except current
```

### 5. Get File by List Item ID

#### PnPjs Method
```typescript
const item = sp.web.lists.getByTitle('Shared Documents').items.getById(5);
const file = item.getFile();
const versions = await file.versions();
```

#### REST API
```
GET /_api/web/lists/GetByTitle('Shared Documents')/items(5)/File/Versions
```

## Data Models

### IFileVersion Interface
```typescript
export interface IFileVersion {
  id: string;                    // Unique version ID (e.g., "v256")
  versionNumber: number;         // Numeric version (2, 1, 0)
  displayNumber: string;         // Display format (2.0, 1.0)
  created: Date;                 // Creation timestamp
  modifiedBy: string;            // User name
  modifiedByEmail?: string;      // User email
  size: number;                  // File size in bytes
  comments: string;              // Check-in comment
  isCurrentVersion: boolean;     // Is this the latest version?
  isMinorVersion: boolean;       // Is this a minor version?
  checkInComment?: string;       // Optional check-in note
  url?: string;                  // Version file URL
  fileRef: string;               // Server relative URL
  expiryInfo?: {                 // Optional expiry info
    expiryDate?: Date;
    accessDuration?: number;
    status: 'Active' | 'Expired' | 'Permanent';
  };
}
```

## Version History Service API

### Constructor
```typescript
const service = new VersionHistoryService(context: WebPartContext);
```

### Methods

#### getFileVersionsFromPnPjs(serverRelativeUrl: string)
```typescript
// Fetches using PnPjs library
const versions = await service.getFileVersionsFromPnPjs('/sites/site/lib/file.pdf');
// Returns: Promise<IFileVersion[]>
```

#### getFileVersionsFromREST(serverRelativeUrl: string)
```typescript
// Fetches using SharePoint REST API
const versions = await service.getFileVersionsFromREST('/sites/site/lib/file.pdf');
// Returns: Promise<IFileVersion[]>
```

#### deleteVersion(serverRelativeUrl: string, versionId: string)
```typescript
// Deletes a specific version
await service.deleteVersion('/sites/site/lib/file.pdf', 'v256');
// Returns: Promise<void>
```

#### deleteAllVersions(serverRelativeUrl: string)
```typescript
// Deletes all versions except current
await service.deleteAllVersions('/sites/site/lib/file.pdf');
// Returns: Promise<void>
```

#### restoreVersion(serverRelativeUrl: string, versionId: string)
```typescript
// Restores a previous version
await service.restoreVersion('/sites/site/lib/file.pdf', 'v256');
// Returns: Promise<void>
```

#### getFileIcon(fileName: string)
```typescript
// Gets Fluent UI icon name based on file extension
const icon = service.getFileIcon('document.pdf');
// Returns: 'FilePDF'
```

#### formatFileSize(bytes: number)
```typescript
// Formats bytes to human-readable format
const size = service.formatFileSize(45678);
// Returns: '44.61 KB'
```

#### formatDate(date: Date)
```typescript
// Formats date to locale string
const formatted = service.formatDate(new Date());
// Returns: '05/14/2026, 10:30:00 AM'
```

## Error Handling

### PnPjs Error Responses
```typescript
try {
  const versions = await service.getFileVersionsFromPnPjs(url);
} catch (error) {
  // Error types:
  // - File not found (404)
  // - Access denied (403)
  // - Throttling (429)
  // - Server error (500)
  console.error(error.message);
}
```

### REST API Error Responses
```json
{
  "error": {
    "code": "-2147024894, Microsoft.SharePoint.SPException",
    "message": {
      "lang": "en-US",
      "value": "File not found."
    }
  }
}
```

## Request/Response Examples

### Example 1: Fetch Versions for a Document

**Request**:
```typescript
const file = sp.web.getFileByServerRelativePath('/sites/contoso/Shared Documents/Report.docx');
const versions = await file.versions();
```

**Response**:
```typescript
[
  {
    ID: 512,
    VersionLabel: "3.0",
    Created: "2026-05-14T10:30:00Z",
    CreatedBy: {
      Title: "Sarah Johnson",
      EMail: "sarah@contoso.com"
    },
    CheckInComment: "Final review complete",
    Size: 52000,
    url: "/sites/contoso/_vti_history/512/Shared Documents/Report.docx"
  },
  {
    ID: 256,
    VersionLabel: "2.0",
    Created: "2026-05-13T15:45:00Z",
    CreatedBy: {
      Title: "Mike Chen",
      EMail: "mike@contoso.com"
    },
    CheckInComment: "Updated analysis",
    Size: 49500
  },
  {
    ID: 128,
    VersionLabel: "1.0",
    Created: "2026-05-12T09:15:00Z",
    CreatedBy: {
      Title: "Lisa Park",
      EMail: "lisa@contoso.com"
    },
    CheckInComment: "Initial upload",
    Size: 45000
  }
]
```

### Example 2: Delete a Version

**Request**:
```typescript
const file = sp.web.getFileByServerRelativePath('/sites/contoso/Shared Documents/Report.docx');
await file.versions.getById(256).delete();
```

**Response**: 204 No Content

### Example 3: Format File Size

**Input**: 52000 bytes

**Output**: "50.78 KB"

## Permissions Required

| Operation | Permission Level |
|-----------|-----------------|
| View versions | Read |
| Delete version | Edit |
| Delete all versions | Edit |
| Restore version | Edit |

## Performance Considerations

### Version Loading
- **Max versions**: 1000+ (SharePoint default)
- **Performance impact**: Minimal for <100 versions
- **Recommendation**: Add pagination for >200 versions

### Deletion Performance
- **Single version delete**: ~200-400ms
- **Bulk delete**: 100ms delay between each version
- **Expected time for 50 versions**: ~5 seconds

### Optimization Strategies
```typescript
// Use batch requests for bulk operations
const versions = await file.versions();
const batchRequests = sp.batched(async (batch) => {
  for (let i = 1; i < versions.length; i++) {
    await batch(file.versions.getById(versions[i].ID)).delete();
  }
});
```

## Throttling & Retry Logic

### SharePoint Throttling
- **Threshold**: 2000 API requests per 5 minutes per user
- **Response**: 429 Too Many Requests

### Recommended Retry Policy
```typescript
async function retryWithBackoff(
  operation: () => Promise<any>,
  maxRetries: number = 3
): Promise<any> {
  for (let i = 0; i < maxRetries; i++) {
    try {
      return await operation();
    } catch (error) {
      if (error.status === 429 && i < maxRetries - 1) {
        // Exponential backoff
        await new Promise(resolve =>
          setTimeout(resolve, Math.pow(2, i) * 1000)
        );
      } else {
        throw error;
      }
    }
  }
}
```

## Deployment Checklist

- [ ] Test with files <10MB
- [ ] Test with files >100MB
- [ ] Verify delete permissions
- [ ] Test error scenarios
- [ ] Test on mobile devices
- [ ] Verify CORS headers
- [ ] Check X-RequestDigest handling
- [ ] Load test with large version lists

## References

- [Microsoft Docs: File Versions REST API](https://docs.microsoft.com/en-us/sharepoint/dev/apis/rest/working-with-file-versions-rest)
- [PnPjs Documentation](https://pnp.github.io/pnpjs/)
- [SharePoint REST API Reference](https://docs.microsoft.com/en-us/sharepoint/dev/apis/rest/)

---

**Last Updated**: May 2026  
**API Version**: SharePoint Online
