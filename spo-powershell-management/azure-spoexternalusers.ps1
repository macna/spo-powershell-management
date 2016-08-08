# Stop on error, so we don't accidentally send a load of junk emails to people
$ErrorActionPreference = "Stop"

# Define credentials and connect to the tenant
$spoCred = Get-AutomationPSCredential -Name 'spoAdminCred'
Connect-SPOService -Url https://contoso-admin.sharepoint.com -Credential $spoCred

# Get a list of sites where sharing is enabled, excluding the public site and any others we want to exclude
$sharedSites = Get-SPOSite | Where-Object {$_.SharingCapability -ne "Disabled" -and $_.Url -ne "https://contoso.sharepoint.com/sites/somesite" -and $_.Url -ne "http://contoso-public.sharepoint.com/"}

# Define credentials for the SMTP relay
$smtpCred = Get-AutomationPSCredential -Name 'smtpCred'

# Run our loop, getting the site owners for each $sharedSite
ForEach ($site In $sharedSites)
    {
        $siteUrl = $site.Url
        $siteTitle = $site.Title
        $siteUrl
        $siteOwners = Get-SPOUser -Site $siteUrl -Limit ALL | Where-Object {$_.Groups -like "*Owners*" -and $_.DisplayName -ne "System Account"} | Select-Object DisplayName, LoginName
        $siteOwners

        # Send the reminder email to each person in $siteOwners
        ForEach ($owner In $siteOwners)
            {
                $siteUrl
                $owner
                $ownerEmail = $owner.LoginName
                $messageSubject = "$siteTitle - External Users (Quarterly Review)"
                $messageBody = @"
                <!DOCTYPE html>
                <html>
                    <style type="text/css"> 
                        body {
                            font-family: sans-serif;
                            font-size: 14px;
                        }
                    </style>
                    <body>
                        <p>Dear site owner,</p>
                        <p>
                            Your SharePoint site "$siteTitle" has external sharing enabled, and as an owner you are responsible for ensuring that it
                            continues to be appropriate for external users to access your site.
                        </p>
                        <p>
                            To check who has access to your site, please follow these steps:
                        </p>
                        <p>
                            Click the <b>Settings</b> icon (cog in the top-right of the window), and then select <b>Site Settings</b>.<br>
                            Select the <b>People and Groups</b> option, where the users and groups who have access to your site will be displayed.
                        </p>
                        <p>
                            This is an automated message, so please do not reply. You are receiving this message because you are listed in the <b>Owners</b> group of the <a href="$siteURL">$siteTitle</a> SharePoint site.
                        </p>
                        <p>
                            If you require any further assistance, please contact the <a href="mailto:ithelpdesk@contoso.com?Subject=SharePoint%20External%20Users%20Email" target="_top">IT Helpdesk</a>.
                        </p>
                        <p>
                            Kind Regards,
                        </p>
                        <p>
                            SharePoint Team
                        </p>
                    </body>
                </html>
"@
                
				# Send the email to the relay 
				Send-MailMessage -UseSsl -BodyAsHtml -From noreply@contoso.com -To $ownerEmail -Subject $messageSubject -Body $messageBody -SmtpServer smtp.contoso.com -Credential $smtpCred
            }
    }