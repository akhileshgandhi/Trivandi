# ✅ Dashboard Web Part - Deployment Checklist

Use this checklist to ensure a smooth deployment of your new Dashboard web part.

---

## Pre-Deployment

### Environment Setup
- [ ] Node.js v18.17.1 or higher installed
  - Check with: `node --version`
- [ ] SharePoint Online tenant access
- [ ] App Catalog site available
- [ ] Modern SharePoint site for testing
- [ ] Site Collection Admin or Owner permissions

### Code Verification
- [ ] All files created successfully in `src/webparts/dashboard/`
- [ ] Service file exists: `src/shared/services/dashboardService.ts`
- [ ] No critical syntax errors in code
- [ ] Dependencies installed: `npm install`

---

## Phase 1: SharePoint List Setup (Choose One Method)

### Method A: Automated (Recommended) ⚡
- [ ] PnP PowerShell module installed
  - Install with: `Install-Module -Name PnP.PowerShell -Scope CurrentUser`
- [ ] Navigate to dashboard folder: `cd src/webparts/dashboard`
- [ ] Run script: `.\setup-lists.ps1 -SiteUrl "YOUR_SITE_URL"`
- [ ] Script completed successfully
- [ ] All 6 lists created (check Site Contents)
- [ ] Sample data visible in lists

### Method B: Manual 📝
Follow [LIST_CREATION_GUIDE.md](LIST_CREATION_GUIDE.md)

- [ ] Created **Projects** list with columns:
  - [ ] Status (Choice)
  - [ ] AssignedTo (Person)
- [ ] Created **Bids** list with columns:
  - [ ] AssignedTo (Person)
  - [ ] Status (Choice)
- [ ] Created **Standards and Compliance** list with columns:
  - [ ] Description (Multiple lines)
  - [ ] Color (Single line)
  - [ ] Icon (Single line)
- [ ] Created **Key People** list with columns:
  - [ ] JobTitle (Single line)
  - [ ] Email (Single line)
  - [ ] WorkPhone (Single line)
  - [ ] Picture (Hyperlink/Picture)
- [ ] Created **Coming Up Events** list (Events template)
- [ ] Created **Key Tools** list with columns:
  - [ ] Description (Multiple lines)
  - [ ] Color (Single line)
  - [ ] Icon (Single line)

### Verify Lists
- [ ] Navigate to **Site contents**
- [ ] All 6 lists visible
- [ ] Each list has at least 1 item
- [ ] Column names spelled correctly
- [ ] Current user has Read permissions

---

## Phase 2: Build & Package

### From Project Root Directory

1. **Clean Previous Builds**
   ```bash
   gulp clean
   ```
   - [ ] No errors displayed
   - [ ] `temp/` folder cleared

2. **Build Solution**
   ```bash
   gulp build
   ```
   - [ ] Build completed successfully
   - [ ] No TypeScript errors
   - [ ] Files created in `lib/` folder

3. **Bundle for Production**
   ```bash
   gulp bundle --ship
   ```
   - [ ] Bundle created successfully
   - [ ] `release/` folder populated
   - [ ] JavaScript files generated

4. **Create Package**
   ```bash
   gulp package-solution --ship
   ```
   - [ ] Package created successfully
   - [ ] File exists: `sharepoint/solution/trivandi-project-team.sppkg`
   - [ ] No warnings or errors

### Verify Package
- [ ] Navigate to `sharepoint/solution/`
- [ ] File present: `trivandi-project-team.sppkg`
- [ ] File size > 0 KB (typically 500KB - 5MB)
- [ ] Creation timestamp is recent

---

## Phase 3: Deploy to SharePoint

### Upload to App Catalog

1. **Navigate to App Catalog**
   - [ ] Go to: `https://[tenant].sharepoint.com/sites/appcatalog`
   - [ ] Or: SharePoint admin center → More features → Apps → Open

2. **Upload Package**
   - [ ] Click **Apps for SharePoint** (left navigation)
   - [ ] Click **Upload** (or drag and drop)
   - [ ] Select `trivandi-project-team.sppkg`
   - [ ] Upload dialog appears

3. **Configure Deployment**
   - [ ] Check: **"Make this solution available to all sites in the organization"**
   - [ ] Review the details shown
   - [ ] Click **Deploy**

4. **Wait for Deployment**
   - [ ] Status changes to "Deployed"
   - [ ] No error messages
   - [ ] Deployment typically takes 1-2 minutes

### Verify Deployment
- [ ] App shows in catalog with green checkmark
- [ ] Status shows "Deployed"
- [ ] Version number matches package.json

---

## Phase 4: Add to SharePoint Site

### Install App on Site

1. **Navigate to Target Site**
   - [ ] Go to your SharePoint site
   - [ ] You have Edit permissions

2. **Add App to Site**
   - [ ] Click **Settings (gear)** → **Add an app**
   - [ ] Search for **"Project Dashboard"** or **"trivandi-project-team"**
   - [ ] Click the app tile
   - [ ] Click **Add**
   - [ ] Wait for "App is being added..." to complete

### Add Web Part to Page

1. **Create or Edit Page**
   - [ ] Click **New** → **Page** (or edit existing page)
   - [ ] Enter page name (e.g., "Dashboard")
   - [ ] Click **Create**

2. **Add Web Part**
   - [ ] Click **+** icon to add web part
   - [ ] Search for **"Project Dashboard"** or **"Dashboard"**
   - [ ] Click the web part in results
   - [ ] Web part appears on page

3. **Verify Display**
   - [ ] Hero banner shows with your name
   - [ ] Stats cards visible (may show 0 if no data)
   - [ ] All sections loaded (no error messages)
   - [ ] Data appears or sample data shows

4. **Save Page**
   - [ ] Click **Save as draft** (to test)
   - [ ] Or **Publish** (to make live)

---

## Phase 5: Testing & Verification

### Functional Testing

1. **Data Display**
   - [ ] Hero banner shows correct user name
   - [ ] Project stats showing (Total, Live, My Projects, My Bids)
   - [ ] Compliance cards displaying with colors
   - [ ] Key people cards showing
   - [ ] Upcoming events listed
   - [ ] Tool cards visible

2. **Visual Testing**
   - [ ] Colors match design (#4A5AFF blue, #FF8C42 orange, etc.)
   - [ ] Rounded corners on cards
   - [ ] Proper spacing and alignment
   - [ ] No overlapping elements
   - [ ] Icons/emojis displaying correctly

3. **Interactive Testing**
   - [ ] Hover effects work on cards (lift up, shadow)
   - [ ] "View All" buttons visible
   - [ ] No console errors (F12 → Console tab)
   - [ ] Page loads within 2-3 seconds

### Responsive Testing

1. **Desktop View**
   - [ ] Full layout visible
   - [ ] Two-column layout (content + sidebar)
   - [ ] Stats in 4 columns

2. **Tablet View**
   - [ ] Layout adjusts appropriately
   - [ ] Readable text sizes
   - [ ] Touch-friendly buttons

3. **Mobile View**
   - [ ] Single column layout
   - [ ] Stats in 2 columns
   - [ ] Scrollable content
   - [ ] No horizontal scrolling

### Browser Testing
- [ ] Microsoft Edge (Chromium)
- [ ] Google Chrome
- [ ] Mozilla Firefox
- [ ] Safari (if available)

---

## Phase 6: Data Population

### Replace Sample Data

1. **Projects List**
   - [ ] Add real project records
   - [ ] Set correct Status values
   - [ ] Assign team members
   - [ ] Delete sample items (optional)

2. **Bids List**
   - [ ] Add actual bids
   - [ ] Assign owners
   - [ ] Update statuses

3. **Standards and Compliance**
   - [ ] Update titles and descriptions
   - [ ] Verify colors match branding
   - [ ] Update icons if needed

4. **Key People**
   - [ ] Add all team members
   - [ ] Upload profile photos
   - [ ] Add contact information
   - [ ] Verify emails and phones

5. **Coming Up Events**
   - [ ] Add real upcoming events
   - [ ] Set accurate dates/times
   - [ ] Delete past events
   - [ ] Add recurring events

6. **Key Tools**
   - [ ] Add actual tools used
   - [ ] Update descriptions
   - [ ] Set tool URLs (if applicable)

### Refresh Dashboard
- [ ] Reload page to see new data
- [ ] Verify stats updated correctly
- [ ] Confirm user-specific data shows

---

## Phase 7: Permissions & Security

### List Permissions
- [ ] All users have **Read** access to all 6 lists
- [ ] Designated users have **Contribute** to update lists
- [ ] Site Owners have **Full Control**

### Test User Access
- [ ] Log in as regular user
- [ ] Dashboard loads and displays
- [ ] User sees their own projects/bids
- [ ] No permission errors

---

## Phase 8: Documentation & Training

### Internal Documentation
- [ ] Share list names with team
- [ ] Document who can edit lists
- [ ] Create process for updates
- [ ] Share color codes for consistency

### User Training
- [ ] Show team where dashboard is
- [ ] Explain what each section shows
- [ ] Train on updating lists
- [ ] Share contact for issues

---

## Phase 9: Monitoring & Maintenance

### First Week
- [ ] Monitor for errors daily
- [ ] Check page load performance
- [ ] Gather user feedback
- [ ] Fix any reported issues

### Ongoing
- [ ] Weekly: Update events
- [ ] Monthly: Review project stats
- [ ] Quarterly: Update compliance info
- [ ] As needed: Add/remove team members

---

## Troubleshooting

### Web Part Not Showing
**Problem**: Can't find web part when adding
- [ ] Verify app deployed in App Catalog
- [ ] Check app added to site (Site contents)
- [ ] Try in different browser
- [ ] Clear browser cache

### No Data Showing
**Problem**: Web part loads but shows 0s or sample data
- [ ] Verify SharePoint lists exist (check Site contents)
- [ ] Check list names match exactly (case-sensitive)
- [ ] Ensure lists have data
- [ ] Check browser console for errors (F12)
- [ ] Verify current user has Read access to lists

### Build Errors
**Problem**: Errors during gulp build
- [ ] Run `npm install` to reinstall dependencies
- [ ] Check Node.js version: `node --version`
- [ ] Run `gulp clean` to clear old builds
- [ ] Delete `node_modules` and run `npm install` again

### Deployment Errors
**Problem**: Can't deploy to App Catalog
- [ ] Verify you're a Site Collection Admin
- [ ] Check App Catalog is configured
- [ ] Try deploying in different browser
- [ ] Contact SharePoint admin

### Style Issues
**Problem**: Colors or layout incorrect
- [ ] Rebuild solution: `gulp clean` then `gulp build`
- [ ] Clear browser cache (Ctrl+F5)
- [ ] Check for conflicting CSS on page
- [ ] Verify SCSS compiled correctly

---

## Success Criteria

✅ **Deployment Successful When:**
- [ ] Web part visible in SharePoint page
- [ ] All 6 sections displaying
- [ ] Real data from SharePoint lists showing
- [ ] User name personalized in hero banner
- [ ] Stats calculating correctly
- [ ] No errors in browser console
- [ ] Mobile responsive layout works
- [ ] Load time under 3 seconds
- [ ] Users can access and view
- [ ] Team knows how to update lists

---

## Rollback Plan (If Needed)

If deployment fails or issues occur:

1. **Remove from Page**
   - [ ] Edit page
   - [ ] Delete web part
   - [ ] Save page

2. **Remove from Site**
   - [ ] Site contents
   - [ ] Find app
   - [ ] Remove

3. **Retract from Catalog**
   - [ ] App Catalog
   - [ ] Find package
   - [ ] Retract

4. **Fix Issues**
   - [ ] Review error messages
   - [ ] Fix code
   - [ ] Rebuild
   - [ ] Redeploy

---

## Post-Deployment

### Optimization
- [ ] Monitor page load times
- [ ] Optimize images if slow
- [ ] Review list query performance
- [ ] Consider caching strategies

### Enhancement Ideas
- [ ] Add filters to sections
- [ ] Create detail pages for items
- [ ] Add export functionality
- [ ] Implement search
- [ ] Add mobile app integration

---

## Contact & Support

**For Issues:**
- Check browser console (F12)
- Review SharePoint ULS logs
- Contact SharePoint admin
- Check [README.md](README.md) for detailed info

**For Customization:**
- See [QUICKSTART.md](QUICKSTART.md) for customization guide
- Review component code in `src/webparts/dashboard/components/`
- Modify styles in `Dashboard.module.scss`

---

## ✅ Final Checklist

- [ ] All SharePoint lists created and populated
- [ ] Solution built successfully
- [ ] Package deployed to App Catalog
- [ ] Web part added to SharePoint page
- [ ] Data displaying correctly
- [ ] Tested on multiple devices
- [ ] Users have access
- [ ] Team trained on updates
- [ ] Documentation shared
- [ ] Monitoring plan in place

---

**🎉 Congratulations! Your Dashboard is Live!**

**Deployment Date**: ______________

**Deployed By**: ______________

**Site URL**: ______________

---

_Save this checklist for future reference and updates._
