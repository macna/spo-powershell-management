# Define credentials and connect to the tenant
$spoCred = Get-AutomationPSCredential -Name 'spoAdminCred'
Connect-SPOService -Url https://contoso-admin.sharepoint.com -Credential $spoCred

# Bag the storage and resource metrics for the tanant
$tenantMetrics = Get-SPOTenant
$tenantMetricsTable = $tenantMetrics | Select-Object StorageQuota, ResourceQuota, ResourceQuotaAllocated | ConvertTo-Html -Fragment
$totalGBAvailable = [Math]::Round($tenantMetrics.StorageQuota / 1024,2)

# Because the SPO commandlets won't give us a sensible 'how much space is used' metric, we have to add it up ourselves
# Get a list of sites, excluding the public site (as it's unused and deprecated) and OneDrive (as that storage is dealt
# with elsewhere)
$siteList = Get-SPOSite | Where-Object {$_.Url -ne "http://contoso-public.sharepoint.com/" -and $_.Url -ne "https://contoso-my.sharepoint.com/"}

# Using $siteList, get the amount of storge used by each site
$storageUsed = ForEach ($site In $siteList) {
    $siteUrl = $site.Url
	Get-SPOSite -Identity $siteUrl -Detailed | Select-Object Title, Url, StorageUsageCurrent, ResourceUsageCurrent
}

# Add up the $storageUsed.StorageUsageCurrent attribute and turn it into a useful number
$totalGBUsed = [Math]::Round(($storageUsed.StorageUsageCurrent | Measure-Object -Sum).sum /1024,2)

# Convert the contents of the $storageUsed variable into HTML
$siteMetrics = $storageUsed | Sort-Object StorageUsageCurrent -Descending | ConvertTo-Html -Fragment

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

# Define credentials for the SMTP relay
$smtpCred = Get-AutomationPSCredential -Name 'smtpCred'

# Send an email to the site owner with a list of external users
Send-MailMessage -UseSsl -BodyAsHtml -To someone@contoso.com -From noreply@contoso.com -SmtpServer smtp.contoso.com -Credential $smtpCred -Subject "SharePoint Online Storage and Resource Usage" -Body $messageBody