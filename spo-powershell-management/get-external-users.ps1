# Import the SPO PowerShell module
Import-Module 'C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell'

# Define credentials and connect to the tenant
$spoCred = Get-Credential -Message "Enter credentials with the SharePoint Online 'Global Administrator' role."
Connect-SPOService -Url https://yourtenant-admin.sharepoint.com -Credential $spoCred

# Define the site that this script is looking at
$site = "https://yourtenant.sharepoint.com/sites/yoursite"

# Identify who the site owner is
$siteOwner = Get-SPOSite -Identity $site | Select-Object Owner

# Get the list of external users and store them in variable $externalUsers, and then convert to HTML
$externalUsers = Get-SPOExternalUser -Site $site | ConvertTo-Html -Fragment -Property DisplayName, Email

# Compose the body of the HTML email
$messageBody = @"
<!DOCTYPE html>
<html>
    <style type="text/css"> 
        body {
            font-family: sans-serif;
            font-size: 14px;
        }
		table {
			border-collapse: collapse;
		}
		table, th, td {
			border: black 1px solid;
			text-align: left;
			padding: 4px;
		}
    </style>
    <body>
        <h2>External User Report of $site</h2>
        <p>
            The following external users have been identitifed as having access to the "$site" SharePoint site collection that you are responsible for.
            Please review these users to ensure that it is still appropriate for them to have access to your resources.
        </p>
        <p>
            $externalUsers
        </p>
        <p>
            If you need to remove any users from the site collection, please refer to the operational user guide.
        </p>
    </body>
</html>
"@

# Send an email to the site owner with a list of external users
Send-MailMessage -UseSsl -BodyAsHtml -To $siteOwner -From "noreply@contoso.com" -SmtpServer smtp.contoso.com -Subject "External Users Report of $site" -Body $messageBody