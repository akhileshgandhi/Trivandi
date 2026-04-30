# Project Dashboard Web Part

A dynamic SharePoint Framework (SPFx) dashboard web part that displays project statistics, compliance information, key people, upcoming events, and tools - all powered by SharePoint lists.

## Features

- **Project Statistics**: Displays Total Projects, Live Projects, My Projects, and My Bids
- **Standards & Compliance**: Shows compliance categories with color-coded cards
- **Key People**: Team member directory with contact information
- **Coming Up Events**: Calendar of upcoming deadlines and events
- **Key Tools**: Quick access to important tools and resources
- **Fully Dynamic**: All data pulled from SharePoint lists
- **Responsive Design**: Works on desktop, tablet, and mobile devices

## SharePoint Lists Required

You need to create the following SharePoint lists in your site:

### 1. Projects List
**List Name**: `Projects`

| Column Name | Type | Required | Description |
|------------|------|----------|-------------|
| Title | Single line of text | Yes | Project name |
| Status | Choice | Yes | Project status (Live, Active, Completed, etc.) |
| AssignedTo | Person or Group | No | Project team member |

### 2. Bids List (Optional)
**List Name**: `Bids`

| Column Name | Type | Required | Description |
|------------|------|----------|-------------|
| Title | Single line of text | Yes | Bid name |
| AssignedTo | Person or Group | No | Person responsible for bid |

### 3. Standards and Compliance List
**List Name**: `Standards and Compliance`

| Column Name | Type | Required | Description |
|------------|------|----------|-------------|
| Title | Single line of text | Yes | Compliance item name |
| Description | Multiple lines of text | No | Item description |
| Color | Single line of text | No | Hex color code (e.g., #FF8C42) |
| Icon | Single line of text | No | Emoji or icon (e.g., 📋) |

**Sample Data**:
```
Title: Associate Onboarding
Description: Procedures & terms...
Color: #4A5AFF
Icon: 📋

Title: Mobilisations
Description: Procedures & terms...
Color: #FF8C42
Icon: 🚀

Title: Health & Safety
Description: Procedures & forms...
Color: #FF4DB8
Icon: ⚕️

Title: BCS Standards
Description: Procedures & forms...
Color: #4A5AFF
Icon: ✓
```

### 4. Key People List
**List Name**: `Key People`

| Column Name | Type | Required | Description |
|------------|------|----------|-------------|
| Title | Single line of text | Yes | Person's full name |
| JobTitle | Single line of text | No | Job title/role |
| Email | Single line of text | No | Email address |
| WorkPhone | Single line of text | No | Phone number |
| Picture | Hyperlink or Picture | No | Profile picture URL |

**Sample Data**:
```
Title: Darrell Steward
JobTitle: Project Head
Email: darrell.steward@company.com
WorkPhone: (555) 123-4567
```

### 5. Coming Up Events List
**List Name**: `Coming Up Events`

| Column Name | Type | Required | Description |
|------------|------|----------|-------------|
| Title | Single line of text | Yes | Event name |
| EventDate | Date and Time | Yes | Event start date/time |
| EndDate | Date and Time | No | Event end date/time |

**Sample Data**:
```
Title: New Partnership Announcement
EventDate: 11/04/2024 10:00 AM

Title: Town Hall Meeting
EventDate: 11/03/2024 12:00 PM

Title: Timesheet Deadline
EventDate: 11/03/2024 11:59 PM
```

### 6. Key Tools List
**List Name**: `Key Tools`

| Column Name | Type | Required | Description |
|------------|------|----------|-------------|
| Title | Single line of text | Yes | Tool name |
| Description | Multiple lines of text | No | Tool description |
| Color | Single line of text | No | Hex color code (e.g., #FF8C42) |
| Icon | Single line of text | No | Emoji or icon (e.g., 📊) |

**Sample Data**:
```
Title: CMAP
Description: Project Management
Color: #FF8C42
Icon: 📊

Title: Timesheets
Description: Log Hours
Color: #4A5AFF
Icon: ⏰

Title: Expenses
Description: Submit Claims
Color: #4A9EFF
Icon: 💰

Title: BenefitX HR
Description: Imagine Benefit...
Color: #FF4DB8
Icon: 👥
```

## Installation & Deployment

### Prerequisites
- Node.js v18.17.1 or higher
- SharePoint Online environment
- SPFx development environment set up

### Build and Deploy

1. **Install dependencies**:
   ```bash
   npm install
   ```

2. **Build the solution**:
   ```bash
   gulp clean
   gulp build
   gulp bundle --ship
   gulp package-solution --ship
   ```

3. **Deploy to SharePoint**:
   - Go to your SharePoint App Catalog site
   - Upload the `.sppkg` file from `sharepoint/solution` folder
   - Check "Make this solution available to all sites in the organization"
   - Click "Deploy"

4. **Add web part to a page**:
   - Go to any SharePoint modern page
   - Click "Edit"
   - Click "+" to add a web part
   - Search for "Project Dashboard"
   - Add the web part to the page

### Development

To test locally:

```bash
gulp serve
```

Then add the web part to your SharePoint workbench.

## Customization

### Changing List Names

If you want to use different list names, edit the `LISTS` constant in:
`src/shared/services/dashboardService.ts`

```typescript
const LISTS = {
  PROJECTS: "Your Projects List Name",
  COMPLIANCE: "Your Compliance List Name",
  KEY_PEOPLE: "Your People List Name",
  EVENTS: "Your Events List Name",
  TOOLS: "Your Tools List Name"
};
```

### Styling

Modify the colors and styling in:
`src/webparts/dashboard/components/Dashboard.module.scss`

Key color variables:
- Primary Blue: `#4A5AFF`
- Secondary Purple: `#7B68EE`
- Orange: `#FF8C42`
- Pink: `#FF4DB8`

## Features Details

### Dynamic Data Loading
- All data is loaded from SharePoint lists using PnP JS
- Falls back to sample data if lists don't exist
- Loads all sections in parallel for better performance

### User-Specific Data
- "My Projects" shows only projects assigned to the current user
- "My Bids" shows only bids assigned to the current user

### Responsive Layout
- Grid layout adapts to screen size
- Mobile-friendly design
- Touch-friendly buttons and cards

## Troubleshooting

**Web part shows sample data instead of real data**:
- Ensure all SharePoint lists are created with correct names
- Check that lists have the required columns
- Verify the current user has read permissions on all lists

**No data showing**:
- Open browser console (F12) to check for errors
- Verify SharePoint list names match exactly (case-sensitive)
- Ensure lists contain at least one item

**Build errors**:
- Run `npm install` to ensure all dependencies are installed
- Check Node.js version: `node --version`
- Clear temp folder: `gulp clean`

## Support

For issues or questions, please contact your SharePoint administrator or the development team.

## Version History

- **1.0.0** - Initial release
  - Project statistics
  - Compliance tracking
  - Key people directory
  - Events calendar
  - Tools quick access

## License

This project is part of the Trivendi Projects solution.
