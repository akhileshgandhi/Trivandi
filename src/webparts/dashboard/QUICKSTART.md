# 🚀 Quick Start Guide - Dashboard Web Part

## Overview
You now have a complete, dynamic SharePoint dashboard web part that displays:
- Project statistics (Total, Live, My Projects, My Bids)
- Standards & Compliance cards
- Key People directory
- Upcoming Events calendar
- Key Tools quick access

All data is pulled dynamically from SharePoint lists!

## 📋 Step-by-Step Setup

### Step 1: Create SharePoint Lists (Easiest Method)

**Option A: Using PowerShell Script (Recommended)**
```powershell
# Navigate to the dashboard folder
cd src/webparts/dashboard

# Run the setup script (replace with your site URL)
.\setup-lists.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite"
```

This script will:
- ✅ Create all 6 required SharePoint lists
- ✅ Add all necessary columns
- ✅ Populate sample data
- ✅ Configure list templates

**Option B: Manual Creation**
See [README.md](README.md) for detailed manual list creation instructions.

### Step 2: Build the Solution

```bash
# From project root
npm install

# Clean and build
gulp clean
gulp build
gulp bundle --ship
gulp package-solution --ship
```

### Step 3: Deploy to SharePoint

1. Navigate to your SharePoint App Catalog:
   `https://yourtenant.sharepoint.com/sites/appcatalog`

2. Go to **Apps for SharePoint**

3. Upload the package:
   - File location: `sharepoint/solution/trivandi-project-team.sppkg`
   - Check: "Make this solution available to all sites"
   - Click: **Deploy**

4. Wait for deployment to complete (usually 1-2 minutes)

### Step 4: Add to Page

1. Go to your SharePoint site
2. Create or edit a page
3. Click **"+"** to add a web part
4. Search for **"Project Dashboard"**
5. Add it to the page
6. **Save** and **Publish** the page

## 🎨 What You Get

### Dashboard Sections:

**1. Hero Banner**
- Personalized greeting with user's name
- "Step into your project control hub!"

**2. Project Stats Cards**
- Total Projects (all projects in the Projects list)
- Live Projects (projects with Status = "Live" or "Active")
- My Projects (projects assigned to current user)
- My Bids (bids assigned to current user)

**3. Standards & Compliance**
- Colorful cards for different compliance areas
- Each card shows title, description, custom color, and icon

**4. Key People**
- Team member cards with photos
- Shows name, job title
- Email and phone icons for quick contact

**5. Coming Up Events**
- Calendar-style event cards
- Shows upcoming events sorted by date
- Displays event name, date, and time

**6. Key Tools**
- Quick access tool cards
- Each tool has custom color and icon
- Shows tool name and description

## 📝 SharePoint Lists Created

1. **Projects** - Stores all project information
2. **Bids** - Stores bid information
3. **Standards and Compliance** - Compliance categories
4. **Key People** - Team member directory
5. **Coming Up Events** - Calendar events
6. **Key Tools** - Tool quick links

## 🔧 Customization

### Change Colors
Edit [Dashboard.module.scss](components/Dashboard.module.scss)
```scss
// Primary colors
$primary-blue: #4A5AFF;
$primary-purple: #7B68EE;
$accent-orange: #FF8C42;
$accent-pink: #FF4DB8;
```

### Change List Names
Edit [dashboardService.ts](../../../shared/services/dashboardService.ts)
```typescript
const LISTS = {
  PROJECTS: "Your Custom Projects List Name",
  COMPLIANCE: "Your Custom Compliance List Name",
  // ... etc
};
```

### Modify Layout
Edit [Dashboard.tsx](components/Dashboard.tsx)
- Reorder sections
- Add/remove components
- Change grid layouts

## 🧪 Testing Locally

```bash
# Start local workbench
gulp serve

# Add web part to workbench:
# https://yourtenant.sharepoint.com/sites/yoursite/_layouts/workbench.aspx
```

## 🐛 Troubleshooting

### Web part shows sample data instead of real data
**Solution**: Lists don't exist or have incorrect names
- Run the PowerShell setup script
- Verify list names match exactly (case-sensitive)

### "Cannot find list" errors in console
**Solution**: Check permissions and list names
```javascript
// Open browser console (F12)
// Look for red error messages
// Verify you have read access to all lists
```

### Build errors
**Solution**: Dependencies or Node version
```bash
# Check Node version (should be 18.x)
node --version

# Reinstall dependencies
npm install

# Clear build cache
gulp clean
```

### Web part doesn't appear in SharePoint
**Solution**: Re-deploy solution
1. Go to App Catalog
2. Find "trivandi-project-team"
3. Click "..." → **Remove**
4. Re-upload the .sppkg file
5. Click **Deploy**

## 📊 Sample Data

The PowerShell script adds sample data automatically. To add more:

**Compliance Items**:
- Go to "Standards and Compliance" list
- Click **New**
- Fill in: Title, Description, Color (hex), Icon (emoji)

**Key People**:
- Go to "Key People" list
- Click **New**
- Fill in: Name, Job Title, Email, Phone, Picture URL

**Events**:
- Go to "Coming Up Events" list
- Click **New**
- Fill in: Title, Event Date, End Date

**Tools**:
- Go to "Key Tools" list
- Click **New**
- Fill in: Title, Description, Color (hex), Icon (emoji)

## 🎯 Best Practices

1. **Keep Data Current**: Update lists regularly
2. **Use High-Quality Images**: For people profiles (minimum 200x200px)
3. **Consistent Colors**: Use the same color palette across items
4. **Clear Descriptions**: Keep text concise and meaningful
5. **Test on Mobile**: Verify responsive layout works

## 📚 Files Structure

```
src/webparts/dashboard/
├── DashboardWebPart.ts           # Main web part entry
├── DashboardWebPart.manifest.json # Web part manifest
├── components/
│   ├── Dashboard.tsx              # Main component
│   ├── Dashboard.module.scss      # Styles
│   ├── IDashboardProps.ts         # Props interface
│   ├── StatCard.tsx               # Stats display
│   ├── ComplianceCard.tsx         # Compliance items
│   ├── KeyPersonCard.tsx          # People cards
│   ├── EventCard.tsx              # Event items
│   └── ToolCard.tsx               # Tool cards
├── loc/
│   ├── en-us.js                   # Localization strings
│   └── mystrings.d.ts             # String types
├── README.md                      # Full documentation
├── setup-lists.ps1                # PowerShell setup script
└── QUICKSTART.md                  # This file

src/shared/services/
└── dashboardService.ts            # SharePoint data service
```

## 🎉 Success Checklist

- ✅ All 6 SharePoint lists created
- ✅ Sample data added to lists
- ✅ Solution built successfully
- ✅ Package deployed to App Catalog
- ✅ Web part added to SharePoint page
- ✅ Data displaying correctly
- ✅ All sections showing real data

## 🆘 Support

**Common Issues:**
1. Lists not found → Run setup-lists.ps1 script
2. No data showing → Check browser console for errors
3. Build fails → Run `npm install` and `gulp clean`
4. Can't deploy → Check App Catalog permissions

**Need Help?**
- Check console logs (F12 in browser)
- Review README.md for detailed info
- Verify SharePoint list permissions
- Ensure Node.js version is 18.x

## 🔄 Updating the Web Part

To make changes and redeploy:

```bash
# Make your code changes
# Then rebuild and redeploy:

gulp clean
gulp build
gulp bundle --ship
gulp package-solution --ship

# Upload new .sppkg to App Catalog
# SharePoint will prompt to replace existing version
```

## 🌟 Next Steps

1. **Customize Colors**: Match your company branding
2. **Add More Data**: Populate lists with real information
3. **Configure Permissions**: Set up list-level permissions
4. **Add Features**: Extend functionality as needed
5. **Train Users**: Show team how to update list data

---

**Happy Dashboarding! 🎊**

Your dynamic SharePoint dashboard is ready to use!
