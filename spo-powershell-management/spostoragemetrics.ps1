# Get through the proxy
$webclient=New-Object System.Net.WebClient
$creds=Get-Credential -Message "Proxy Creds"
$webclient.Proxy.Credentials=$creds

# Import the SPO PowerShell module
Import-Module 'C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell'

# Define credentials and connect to the tenant
$spoCred = Get-Credential -Message "Enter credentials with the SharePoint Online 'Global Administrator' role."
Connect-SPOService -Url https://contoso-admin.sharepoint.com -Credential $spoCred

# Bag the storage and resource metrics for the tanant
$tenantMetrics = Get-SPOTenant
$tenantMetricsTable = $tenantMetrics | Select-Object StorageQuota, ResourceQuota, ResourceQuotaAllocated | ConvertTo-Html -Fragment
$totalGBAvailable = [Math]::Round($tenantMetrics.StorageQuota / 1024,2)

# Because the SPO commandlets won't give us a sensible 'how much space is used' metric, we have to add it up ourselves
# Get a list of sites, excluding the public site (as it's unused and deprecated) and OneDrive (as that storage is dealt
# with elsewhere)
$siteList = Get-SPOSite -Limit ALL | Where-Object {$_.Url -ne "http://contoso-public.sharepoint.com/" -and $_.Url -ne "https://contoso-my.sharepoint.com/" -and $_.Url -notlike "https://contoso.sharepoint.com/portals*" -and $_.Url -notlike "http://bot*" -and $_.Template -notlike "GROUP*"}

# Add up the $siteList.StorageUsageCurrent attribute and turn it into a useful number
$totalGBUsed = [Math]::Round(($siteList.StorageUsageCurrent | Measure-Object -Sum).sum /1024,2)

# Convert the contents of the $siteList variable into HTML
$siteMetrics = $siteList | Select-Object Title, Url, StorageUsageCurrent | Sort-Object StorageUsageCurrent -Descending | ConvertTo-Html -Fragment

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
        <p>
            The storage metrics for the SharePoint Online tenant (in MB) are as follows:
        </p>
        <p>
			$tenantMetricsTable
        </p>
        <p style="font-size:20px">
            In total, $totalGBUsed GB are used of $totalGBAvailable GB.
        </p>
        <p>
            The storage metrics for individual site collections (in MB) are as follows:
        </p>
        <p>
			$siteMetrics
        </p>
    </body>
</html>
"@

# Send an email to the site owner with a list of external users
Send-MailMessage -BodyAsHtml -From noreply@contoso.com -To someone@contoso.com -Subject "SharePoint Online Storage and Resource Usage" -Body $messageBody -SmtpServer 127.0.0.1 -Port 1025