# PowerShell script to create SharePoint lists for Dashboard Web Part
# Prerequisites: PnP PowerShell module installed
# Install with: Install-Module -Name PnP.PowerShell -Scope CurrentUser

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl
)

# Connect to SharePoint site
Write-Host "Connecting to SharePoint site: $SiteUrl" -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive

# Function to create list if it doesn't exist
function Create-ListIfNotExists {
    param(
        [string]$ListName,
        [string]$Template = "GenericList"
    )
    
    $list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
    if ($null -eq $list) {
        Write-Host "Creating list: $ListName" -ForegroundColor Green
        New-PnPList -Title $ListName -Template $Template
        return $true
    } else {
        Write-Host "List already exists: $ListName" -ForegroundColor Yellow
        return $false
    }
}

# 1. Create Projects List
Write-Host "`n=== Creating Projects List ===" -ForegroundColor Cyan
if (Create-ListIfNotExists -ListName "Projects") {
    Add-PnPField -List "Projects" -DisplayName "Status" -InternalName "ProjectStatus" -Type Choice -Choices "Live","Active","Completed","On Hold","Cancelled" -AddToDefaultView
    Add-PnPField -List "Projects" -DisplayName "AssignedTo" -InternalName "AssignedTo" -Type User -AddToDefaultView
    Write-Host "Projects list created successfully!" -ForegroundColor Green
}

# 2. Create Bids List
Write-Host "`n=== Creating Bids List ===" -ForegroundColor Cyan
if (Create-ListIfNotExists -ListName "Bids") {
    Add-PnPField -List "Bids" -DisplayName "AssignedTo" -InternalName "AssignedTo" -Type User -AddToDefaultView
    Add-PnPField -List "Bids" -DisplayName "Status" -InternalName "BidStatus" -Type Choice -Choices "Draft","Submitted","Won","Lost" -AddToDefaultView
    Write-Host "Bids list created successfully!" -ForegroundColor Green
}

# 3. Create Standards and Compliance List
Write-Host "`n=== Creating Standards and Compliance List ===" -ForegroundColor Cyan
if (Create-ListIfNotExists -ListName "Standards and Compliance") {
    Add-PnPField -List "Standards and Compliance" -DisplayName "Description" -InternalName "Description" -Type Note -AddToDefaultView
    Add-PnPField -List "Standards and Compliance" -DisplayName "Color" -InternalName "Color" -Type Text -AddToDefaultView
    Add-PnPField -List "Standards and Compliance" -DisplayName "Icon" -InternalName "Icon" -Type Text -AddToDefaultView
    Add-PnPField -List "Standards and Compliance" -DisplayName "Links" -InternalName "Links" -Type URL -AddToDefaultView
    
    # Add sample data
    Write-Host "Adding sample compliance items..." -ForegroundColor Yellow
    Add-PnPListItem -List "Standards and Compliance" -Values @{"Title"="Associate Onboarding";"Description"="Procedures & terms...";"Color"="#4A5AFF";"Icon"="📋";"Links"="https://www.example.com/onboarding, Onboarding"}
    Add-PnPListItem -List "Standards and Compliance" -Values @{"Title"="Mobilisations";"Description"="Procedures & terms...";"Color"="#FF8C42";"Icon"="🚀";"Links"="https://www.example.com/mobilisations, Mobilisations"}
    Add-PnPListItem -List "Standards and Compliance" -Values @{"Title"="Health & Safety";"Description"="Procedures & forms...";"Color"="#FF4DB8";"Icon"="⚕️";"Links"="https://www.example.com/safety, Safety"}
    Add-PnPListItem -List "Standards and Compliance" -Values @{"Title"="BCS Standards";"Description"="Procedures & forms...";"Color"="#4A5AFF";"Icon"="✓";"Links"="https://www.example.com/standards, Standards"}
    Write-Host "Standards and Compliance list created successfully!" -ForegroundColor Green
}

# 4. Create Key People List
Write-Host "`n=== Creating Key People List ===" -ForegroundColor Cyan
if (Create-ListIfNotExists -ListName "Key People") {
    Add-PnPField -List "Key People" -DisplayName "Job Title" -InternalName "JobTitle" -Type Text -AddToDefaultView
    Add-PnPField -List "Key People" -DisplayName "Email" -InternalName "Email" -Type Text -AddToDefaultView
    Add-PnPField -List "Key People" -DisplayName "Work Phone" -InternalName "WorkPhone" -Type Text -AddToDefaultView
    Add-PnPField -List "Key People" -DisplayName "Picture" -InternalName "Picture" -Type URL -AddToDefaultView
    
    # Add sample data
    Write-Host "Adding sample people..." -ForegroundColor Yellow
    Add-PnPListItem -List "Key People" -Values @{"Title"="Darrell Steward";"JobTitle"="Project Head";"Email"="darrell@company.com";"WorkPhone"="(555) 123-4567"}
    Add-PnPListItem -List "Key People" -Values @{"Title"="Ronald Richards";"JobTitle"="Project Head";"Email"="ronald@company.com";"WorkPhone"="(555) 123-4568"}
    Add-PnPListItem -List "Key People" -Values @{"Title"="Jane Flores";"JobTitle"="Project Manager";"Email"="jane@company.com";"WorkPhone"="(555) 123-4569"}
    Add-PnPListItem -List "Key People" -Values @{"Title"="Darrell Howard";"JobTitle"="Team Lead";"Email"="howard@company.com";"WorkPhone"="(555) 123-4570"}
    Write-Host "Key People list created successfully!" -ForegroundColor Green
}

# 5. Create Coming Up Events List
Write-Host "`n=== Creating Coming Up Events List ===" -ForegroundColor Cyan
if (Create-ListIfNotExists -ListName "Coming Up Events" -Template "Events") {
    # Events template already has EventDate and EndDate columns
    
    # Add sample data
    Write-Host "Adding sample events..." -ForegroundColor Yellow
    $today = Get-Date
    Add-PnPListItem -List "Coming Up Events" -Values @{"Title"="New Partnership Announcement";"EventDate"=$today.AddDays(5).ToString("yyyy-MM-dd 10:00:00")}
    Add-PnPListItem -List "Coming Up Events" -Values @{"Title"="Town Hall Meeting";"EventDate"=$today.AddDays(3).ToString("yyyy-MM-dd 12:00:00")}
    Add-PnPListItem -List "Coming Up Events" -Values @{"Title"="Timesheet Deadline";"EventDate"=$today.AddDays(2).ToString("yyyy-MM-dd 23:59:00")}
    Add-PnPListItem -List "Coming Up Events" -Values @{"Title"="Webinar Series Launch";"EventDate"=$today.AddDays(10).ToString("yyyy-MM-dd 21:00:00")}
    Add-PnPListItem -List "Coming Up Events" -Values @{"Title"="Project Review";"EventDate"=$today.AddDays(30).ToString("yyyy-MM-dd 12:00:00")}
    Write-Host "Coming Up Events list created successfully!" -ForegroundColor Green
}

# 6. Create Key Tools List
Write-Host "`n=== Creating Key Tools List ===" -ForegroundColor Cyan
if (Create-ListIfNotExists -ListName "Key Tools") {
    Add-PnPField -List "Key Tools" -DisplayName "Description" -InternalName "Description" -Type Note -AddToDefaultView
    Add-PnPField -List "Key Tools" -DisplayName "Color" -InternalName "Color" -Type Text -AddToDefaultView
    Add-PnPField -List "Key Tools" -DisplayName "Icon" -InternalName "Icon" -Type Text -AddToDefaultView
    
    # Add sample data
    Write-Host "Adding sample tools..." -ForegroundColor Yellow
    Add-PnPListItem -List "Key Tools" -Values @{"Title"="CMAP";"Description"="Project Management";"Color"="#FF8C42";"Icon"="📊"}
    Add-PnPListItem -List "Key Tools" -Values @{"Title"="Timesheets";"Description"="Log Hours";"Color"="#4A5AFF";"Icon"="⏰"}
    Add-PnPListItem -List "Key Tools" -Values @{"Title"="Expenses";"Description"="Submit Claims";"Color"="#4A9EFF";"Icon"="💰"}
    Add-PnPListItem -List "Key Tools" -Values @{"Title"="BenefitX HR";"Description"="Imagine Benefit...";"Color"="#FF4DB8";"Icon"="👥"}
    Write-Host "Key Tools list created successfully!" -ForegroundColor Green
}

Write-Host "`n=== All Lists Created Successfully! ===" -ForegroundColor Green
Write-Host "You can now deploy and use the Dashboard web part." -ForegroundColor Cyan
Write-Host "Site URL: $SiteUrl" -ForegroundColor Cyan

# Disconnect
Disconnect-PnPOnline
