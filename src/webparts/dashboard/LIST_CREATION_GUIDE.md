# 📋 SharePoint List Creation Guide

## Visual Guide to Creating Lists

This guide provides detailed, step-by-step instructions for manually creating the SharePoint lists needed for the Dashboard web part.

---

## List 1: Projects

**Purpose**: Store all project information and track project status

### Steps to Create:
1. Go to your SharePoint site
2. Click **Settings (gear icon)** → **Site contents**
3. Click **+ New** → **List**
4. Name it: `Projects`
5. Click **Create**

### Columns to Add:

| Column Name | Type | Settings | How to Add |
|------------|------|----------|------------|
| Title | Single line of text | *(Already exists)* | Default column |
| Status | Choice | **Choices**: Live, Active, Completed, On Hold, Cancelled<br>**Default**: Live | + Add column → Choice |
| AssignedTo | Person or Group | Allow multiple: No<br>Show field: Name with presence | + Add column → Person |

### Sample Data:
```
Title: Website Redesign Project
Status: Live
AssignedTo: [Your Name]

Title: Mobile App Development
Status: Active
AssignedTo: [Your Name]

Title: Infrastructure Upgrade
Status: Completed
AssignedTo: [Team Member]
```

**⚠️ Note**: The "Status" column should be created as "ProjectStatus" internally to avoid conflicts.

---

## List 2: Bids

**Purpose**: Track bid submissions and assignments

### Steps to Create:
1. **Settings** → **Site contents**
2. **+ New** → **List**
3. Name: `Bids`
4. Click **Create**

### Columns to Add:

| Column Name | Type | Settings |
|------------|------|----------|
| Title | Single line of text | *(Already exists)* |
| AssignedTo | Person or Group | Allow multiple: No |
| Status | Choice | **Choices**: Draft, Submitted, Won, Lost<br>**Default**: Draft |

### Sample Data:
```
Title: Corporate Training Bid
AssignedTo: [Your Name]
Status: Submitted

Title: IT Services Proposal
AssignedTo: [Your Name]
Status: Draft
```

---

## List 3: Standards and Compliance

**Purpose**: Display compliance categories with custom colors and icons

### Steps to Create:
1. **Settings** → **Site contents**
2. **+ New** → **List**
3. Name: `Standards and Compliance`
4. Click **Create**

### Columns to Add:

| Column Name | Type | Settings | How to Add |
|------------|------|----------|------------|
| Title | Single line of text | *(Already exists)* | Default |
| Description | Multiple lines of text | Number of lines: 6<br>Type: Plain text | + Add column → Multiple lines of text |
| Color | Single line of text | Max length: 255 | + Add column → Single line of text |
| Icon | Single line of text | Max length: 255 | + Add column → Single line of text |

### Sample Data:

**Item 1:**
```
Title: Associate Onboarding
Description: Procedures & terms...
Color: #4A5AFF
Icon: 📋
```

**Item 2:**
```
Title: Mobilisations
Description: Procedures & terms...
Color: #FF8C42
Icon: 🚀
```

**Item 3:**
```
Title: Health & Safety
Description: Procedures & forms...
Color: #FF4DB8
Icon: ⚕️
```

**Item 4:**
```
Title: BCS Standards
Description: Procedures & forms...
Color: #4A5AFF
Icon: ✓
```

**💡 Tips:**
- For **Color**: Use hex color codes (e.g., #FF8C42)
- For **Icon**: You can use emojis or Font Awesome icon codes

---

## List 4: Key People

**Purpose**: Team member directory with contact information

### Steps to Create:
1. **Settings** → **Site contents**
2. **+ New** → **List**
3. Name: `Key People`
4. Click **Create**

### Columns to Add:

| Column Name | Type | Settings | How to Add |
|------------|------|----------|------------|
| Title | Single line of text | *(Full Name)* | Default |
| JobTitle | Single line of text | Max length: 255 | + Add column → Single line of text |
| Email | Single line of text | Max length: 255 | + Add column → Single line of text |
| WorkPhone | Single line of text | Max length: 255 | + Add column → Single line of text |
| Picture | Hyperlink or Picture | Format: Picture | + Add column → Hyperlink or Picture |

### Sample Data:

**Person 1:**
```
Title: Darrell Steward
JobTitle: Project Head
Email: darrell.steward@company.com
WorkPhone: (555) 123-4567
Picture: [URL to profile image or leave blank]
```

**Person 2:**
```
Title: Ronald Richards
JobTitle: Project Manager
Email: ronald.richards@company.com
WorkPhone: (555) 123-4568
Picture: [URL or leave blank]
```

**Person 3:**
```
Title: Jane Flores
JobTitle: Team Lead
Email: jane.flores@company.com
WorkPhone: (555) 123-4569
Picture: [URL or leave blank]
```

**💡 Tips:**
- If Picture is blank, the web part will show initials in a colored circle
- Use direct image URLs (e.g., from SharePoint document library)
- Recommended image size: 200x200 pixels or larger

---

## List 5: Coming Up Events

**Purpose**: Display upcoming events and deadlines

### Steps to Create:
1. **Settings** → **Site contents**
2. **+ New** → **List**
3. Choose: **Events** (use the Events template)
4. Name: `Coming Up Events`
5. Click **Create**

### Columns Included (with Events template):

| Column Name | Type | Notes |
|------------|------|-------|
| Title | Single line of text | Event name |
| EventDate | Date and Time | Start date/time |
| EndDate | Date and Time | End date/time (optional) |
| Location | Single line of text | (optional) |
| Description | Multiple lines of text | (optional) |

### Sample Data:

**Event 1:**
```
Title: New Partnership Announcement
EventDate: 11/04/2024 10:00 AM
EndDate: 11/04/2024 11:00 AM
```

**Event 2:**
```
Title: Town Hall Meeting
EventDate: 11/03/2024 12:00 PM
EndDate: 11/03/2024 1:00 PM
```

**Event 3:**
```
Title: Timesheet Deadline
EventDate: 11/03/2024 11:59 PM
```

**Event 4:**
```
Title: Webinar Series Launch
EventDate: 11/12/2024 9:00 PM
```

**Event 5:**
```
Title: Project Review
EventDate: 12/22/2024 12:00 PM
```

**💡 Tips:**
- Use future dates for events to appear in dashboard
- Events are automatically sorted by date
- Only upcoming events (future dates) are shown

---

## List 6: Key Tools

**Purpose**: Quick access to important tools and resources

### Steps to Create:
1. **Settings** → **Site contents**
2. **+ New** → **List**
3. Name: `Key Tools`
4. Click **Create**

### Columns to Add:

| Column Name | Type | Settings | How to Add |
|------------|------|----------|------------|
| Title | Single line of text | *(Tool name)* | Default |
| Description | Multiple lines of text | Number of lines: 6<br>Type: Plain text | + Add column → Multiple lines of text |
| Color | Single line of text | Max length: 255 | + Add column → Single line of text |
| Icon | Single line of text | Max length: 255 | + Add column → Single line of text |

### Sample Data:

**Tool 1:**
```
Title: CMAP
Description: Project Management
Color: #FF8C42
Icon: 📊
```

**Tool 2:**
```
Title: Timesheets
Description: Log Hours
Color: #4A5AFF
Icon: ⏰
```

**Tool 3:**
```
Title: Expenses
Description: Submit Claims
Color: #4A9EFF
Icon: 💰
```

**Tool 4:**
```
Title: BenefitX HR
Description: Imagine Benefit...
Color: #FF4DB8
Icon: 👥
```

---

## 🎨 Color Palette Reference

Use these hex codes for consistent colors across your dashboard:

| Color Name | Hex Code | Usage |
|-----------|----------|-------|
| Primary Blue | `#4A5AFF` | Main theme color |
| Purple | `#7B68EE` | Secondary theme |
| Orange | `#FF8C42` | Accent color |
| Pink | `#FF4DB8` | Accent color |
| Light Blue | `#4A9EFF` | Accent color |
| Green | `#10B981` | Success/Active |
| Red | `#EF4444` | Error/Urgent |
| Gray | `#6B7280` | Neutral |

---

## 📝 Quick Setup Checklist

- [ ] Created **Projects** list
  - [ ] Added Status column (Choice)
  - [ ] Added AssignedTo column (Person)
  - [ ] Added sample projects

- [ ] Created **Bids** list
  - [ ] Added AssignedTo column (Person)
  - [ ] Added Status column (Choice)
  - [ ] Added sample bids

- [ ] Created **Standards and Compliance** list
  - [ ] Added Description column (Multiple lines)
  - [ ] Added Color column (Single line)
  - [ ] Added Icon column (Single line)
  - [ ] Added 4 compliance items

- [ ] Created **Key People** list
  - [ ] Added JobTitle column (Single line)
  - [ ] Added Email column (Single line)
  - [ ] Added WorkPhone column (Single line)
  - [ ] Added Picture column (Hyperlink/Picture)
  - [ ] Added team members

- [ ] Created **Coming Up Events** list (Events template)
  - [ ] Added future events with dates
  - [ ] Set EventDate for each item

- [ ] Created **Key Tools** list
  - [ ] Added Description column (Multiple lines)
  - [ ] Added Color column (Single line)
  - [ ] Added Icon column (Single line)
  - [ ] Added tool items

---

## 🚀 After Creating Lists

1. **Verify List Names**: Make sure names match exactly (case-sensitive)
2. **Check Permissions**: Ensure all users have Read access
3. **Add More Data**: Populate with real information
4. **Deploy Web Part**: Follow build and deployment steps
5. **Test Dashboard**: Verify data appears correctly

---

## 🔄 Alternative: Automated Setup

Instead of manual creation, use the PowerShell script:

```powershell
.\setup-lists.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite"
```

This creates all lists, columns, and sample data automatically in under 2 minutes!

---

## ❓ Troubleshooting

**Problem**: Column not showing in list
- **Solution**: Click **+ Add column** → Choose correct type → Fill settings

**Problem**: Can't add Choice column
- **Solution**: Use **+ Add column** → **Choice** → Enter choices one per line

**Problem**: Date format issues
- **Solution**: Use format: MM/DD/YYYY HH:MM AM/PM

**Problem**: Colors not working
- **Solution**: Use hex format with #: `#FF8C42`

**Problem**: Icons not showing
- **Solution**: Use emojis or copy/paste icons: 📊 ⏰ 💰 👥

---

## 📞 Need Help?

- Check that list names are exact (including spaces and capitalization)
- Verify column types match this guide
- Ensure at least one item exists in each list
- Test by navigating to: Site contents → [List Name]

---

**List Creation Time**: ~20 minutes manually or ~2 minutes with PowerShell script

**Ready to proceed?** After creating lists, continue with build and deployment!
