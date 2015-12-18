# Import the SPO PowerShell module
Import-Module 'C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell'

# Define credentials and connect to the tenant
$cred = Get-Credential -Message "Enter credentials with the SharePoint Online Global Administrator role."
Connect-SPOService -Url https://yourtenant-admin.sharepoint.com -Credential $cred

# Define the site tht this script is looking at
$site = https://yourtenant.sharepoint.com/sites/yoursite

# Identify who the site owner is
$siteOwner = Get-SPOSite -Identity $site | Select-Object Owner

# Get the list of external users and store them in variable $externalUsers
$externalUsers = Get-SPOExternalUser -Site $site | Select-Object DisplayName, Email

