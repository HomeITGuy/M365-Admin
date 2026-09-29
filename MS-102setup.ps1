<#  
Lab 1 - Exercise 1 : Initialize your Microsoft 365 Tenant
#>

# Install Microsoft Graph PowerShell
Install-Module Microsoft.Graph -Scope CurrentUser

# You will now display the list of sub-modules that were installed. To do so, run the following command:

Get-InstalledModule Microsoft.Graph.* | select *name*

<#
Verify the list of installed sub-modules includes the following four sub-modules that will be used in later lab exercises:
Microsoft.Graph.Groups
Microsoft.Graph.Identity.DirectoryManagement
Microsoft.Graph.Identity.Governance
Microsoft.Graph.Users
If all four sub-modules appear in the list of installed sub-modules, then proceed to the next step. However, 
if any of these sub-modules do not appear in the list, then run the following PowerShell command to manually install the missing sub-module:
Install-Module -Name <module name> -Scope CurrentUser
#>

# Run these commands
Install-Module -Name Microsoft.Graph.Identity.DirectoryManagement
Install-Module -Name Microsoft.Graph.Groups
Install-Module -Name Microsoft.Graph.Identity.Governance
Install-Module -Name Microsoft.Graph.Users

<#
PowerShell's execution policy settings dictate what PowerShell scripts can be run on a Windows system. Setting this policy to Unrestricted
#>
Set-ExecutionPolicy RemoteSigned

# Connect to your M365 Tenant
Connect-MgGraph -Scopes "Organization.ReadWrite.All"

# Set the Organization information
Update-MgOrganization -OrganizationId (Get-MgOrganization).Id `
    -Street "555 Main Street" `
    -City "Redmond" `
    -State "Washington" `
    -PostalCode "98052" `
    -PreferredLanguage "en-US"

# If you wanted to try updating the name (likely blocked):
Update-MgOrganization -OrganizationId (Get-MgOrganization).Id `
    -DisplayName "Adatum Corporation"
<#
Create the Microsoft 365 Group
Name: M365 pilot project
Description: Members of the Microsoft 365 pilot project team
Privacy: Private
Group email: m365pilotproject
Create a Team: Yes
#>
$group = New-MgGroup -DisplayName "M365 pilot project" `
    -Description "Members of the Microsoft 365 pilot project team" `
    -MailNickname "m365pilotproject" `
    -MailEnabled $true `
    -SecurityEnabled $false `
    -GroupTypes "Unified" `
    -Visibility "Private"

# Add the owner (MOD Administrator)
$owner = Get-MgUser -Filter "displayName eq 'MOD Administrator'"
Add-MgGroupOwner -GroupId $group.Id -DirectoryObjectId $owner.Id

# Add the 10 members
$members = @(
    "Alex Wilber",
    "Allan Deyoung",
    "Diego Siciliani",
    "Isaiah Langer",
    "Joni Sherman",
    "Lynne Robbins",
    "Megan Bowen",
    "MOD Administrator",
    "Nestor Wilke",
    "Patti Fernandez"
)

foreach ($m in $members) {
    $user = Get-MgUser -Filter "displayName eq '$m'"
    Add-MgGroupMember -GroupId $group.Id -DirectoryObjectId $user.Id
}

# Create a Team for the group
New-MgTeam -GroupId $group.Id

# Learning Path 1 - Lab 1 - Exercise 2
# Graph Command to Remove Licenses
#You must look up the SkuPartNumber for each license.
Get-MgSubscribedSku | Select SkuId, SkuPartNumber

# Once you identify the correct SkuPartNumbers, pull their SkuIds:
# Microsoft 365 E5 (no Teams) → often SPE_E5_NOPSTNCONF
$e5NoTeamsSku = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq "SPE_E5_NOPSTNCONF"

# Remove both licenses from Christie Cline
Set-MgUserLicense -UserId "ChristieCline@yourdomain.com" `
    -RemoveLicenses @($teamsSku.SkuId, $e5NoTeamsSku.SkuId) `
    -AddLicenses @{}

# Create Holly Dickson Account
Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All"

New-MgUser `
  -DisplayName "Holly Dickson" `
  -GivenName "Holly" `
  -Surname "Dickson" `
  -UserPrincipalName "Holly@yourtenant.onmicrosoft.com" `
  -MailNickname "Holly" `
  -AccountEnabled $true `
  -PasswordProfile @{ 
        Password = "54)(J^+NEaMg4w:QgE985(={ni0TvW5w"
        ForceChangePasswordNextSignIn = $false
  }
# Assign the license
# Once you identify the correct SkuPartNumbers, run:
#Microsoft 365 E5 (no Teams) → often SPE_E5_NOPSTNCONF
Get-MgSubscribedSku | Select SkuId, SkuPartNumber
$e5Sku    = Get-MgSubscribedSku | Where-Object SkuPartNumber -eq "<E5NoTeamsSkuPartNumber>"

# Make Holly a Global Administrator
# Find the Global Admin role:
$role = Get-MgDirectoryRole | Where-Object DisplayName -eq "Global Administrator"

# Add Holly to the role:
$user = Get-MgUser -UserPrincipalName "Holly@yourtenant.onmicrosoft.com"

New-MgDirectoryRoleMember `
  -DirectoryRoleId $role.Id `
  -DirectoryObjectId $user.Id

# Update the Microsoft 365 pilot project group
# Get the group ID for M365 pilot project
$group = Get-MgGroup -Filter "displayName eq 'M365 pilot project'"
# Get Holly Dickson’s user object ID
$holly = Get-MgUser -Filter "displayName eq 'Holly Dickson'"
# Add Holly to the group
New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $holly.Id

# Create additional groups for testing
# SECTION 1 — Create the “Inside Sales” Microsoft 365 Group
$group = New-MgGroup `
  -DisplayName "Inside Sales" `
  -Description "Collaboration group for the Inside Sales team" `
  -MailNickname "insidesales" `
  -MailEnabled $true `
  -SecurityEnabled $false `
  -GroupTypes "Unified" `
  -Visibility "Public"

# SECTION 2 — Add Owners (Allan Deyoung & Patti Fernandez)
$allan = Get-MgUser -Filter "displayName eq 'Allan Deyoung'"
$patti = Get-MgUser -Filter "displayName eq 'Patti Fernandez'"

Add-MgGroupOwner -GroupId $group.Id -DirectoryObjectId $allan.Id
Add-MgGroupOwner -GroupId $group.Id -DirectoryObjectId $patti.Id

# SECTION 3 — Add Members (Diego Siciliani & Lynne Robbins)
$diego = Get-MgUser -Filter "displayName eq 'Diego Siciliani'"
$lynne = Get-MgUser -Filter "displayName eq 'Lynne Robbins'"

Add-MgGroupMember -GroupId $group.Id -DirectoryObjectId $diego.Id
Add-MgGroupMember -GroupId $group.Id -DirectoryObjectId $lynne.Id

# SECTION 4 — Create a Team for the Group
New-MgTeam -GroupId $group.Id

# SECTION 5 — Create the “Accounting” Microsoft 365 Group
$acct = New-MgGroup `
  -DisplayName "Accounting" `
  -Description "List of all Accounting staff participating in the Microsoft 365 pilot project" `
  -MailNickname "accounting" `
  -MailEnabled $true `
  -SecurityEnabled $false `
  -GroupTypes "Unified" `
  -Visibility "Public"

# SECTION 6 — Add Owner (Joni Sherman)
$joni = Get-MgUser -Filter "displayName eq 'Joni Sherman'"
Add-MgGroupOwner -GroupId $acct.Id -DirectoryObjectId $joni.Id

# SECTION 7 — Add Members to Accounting (Alex, Joni, Lynne)
$alex = Get-MgUser -Filter "displayName eq 'Alex Wilber'"
$lynne = Get-MgUser -Filter "displayName eq 'Lynne Robbins'"

Add-MgGroupMember -GroupId $acct.Id -DirectoryObjectId $alex.Id
Add-MgGroupMember -GroupId $acct.Id -DirectoryObjectId $joni.Id
Add-MgGroupMember -GroupId $acct.Id -DirectoryObjectId $lynne.Id

# SECTION 8 — Create a Security Group (“IT Admins”)
$itadmins = New-MgGroup `
  -DisplayName "IT Admins" `
  -Description "IT administrative personnel" `
  -MailEnabled $false `
  -SecurityEnabled $true

# SECTION 9 — Add Members to IT Admins (Isaiah, Megan, Nestor)
$isaiah = Get-MgUser -Filter "displayName eq 'Isaiah Langer'"
$megan  = Get-MgUser -Filter "displayName eq 'Megan Bowen'"
$nestor = Get-MgUser -Filter "displayName eq 'Nestor Wilke'"

Add-MgGroupMember -GroupId $itadmins.Id -DirectoryObjectId $isaiah.Id
Add-MgGroupMember -GroupId $itadmins.Id -DirectoryObjectId $megan.Id
Add-MgGroupMember -GroupId $itadmins.Id -DirectoryObjectId $nestor.Id

# SECTION 10 — Delete the “Inside Sales” Group (Team + M365 Group)
Remove-MgGroup -GroupId $group.Id

# SECTION 11 — Verify Deleted Groups
Get-MgGroup -Filter "displayName eq 'Inside Sales'" -ConsistencyLevel eventual

# SECTION 12 — Verify Users Still Exist
Get-MgUser -Filter "displayName eq 'Allan Deyoung'"
Get-MgUser -Filter "displayName eq 'Patti Fernandez'"
Get-MgUser -Filter "displayName eq 'Diego Siciliani'"
Get-MgUser -Filter "displayName eq 'Lynne Robbins'"

