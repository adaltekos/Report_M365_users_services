# Install required modules
	#Install-Module SharePointPnPPowerShellOnline
	#Install-Module Microsoft.Graph
	#Install-Module ImportExcel

# Set variables
$filename 		= "" #Complete with filename (ex. Raport_M365_users_services.xlsx)
$localPath 		= "" #Complete with local path (ex. C:\Raporty\)
$siteUrl		= "" #Complete with Url site (ex. https://company.sharepoint.com/sites/it-dep)
$onlinePath		= "" #Complete with path where file is on sharepoint (ex. Shared Documents/Global/)
$tenant			= "" #Complete with tenant name (ex. company.onmicrosoft.com)
$appId			= "" #Complete with ClientId (which is ID of application registered in Azure AD)
$thumbprint		= "" #Complete with Thumbprint (which is certificate thumbprint)

# Connect to SharePoint Online
$pnpConnectParams  = @{
    Url				=  $siteUrl
    Tenant			=  $tenant
    ClientId		=  $appId
    Thumbprint		=  $thumbprint
}
Connect-PnPOnline @pnpConnectParams

# Download file from SharePoint Online to local path
$getPnPFileParams = @{
    Url				= ($onlinePath + $filename)
    Path			= $localPath
    Filename		= $filename
    AsFile			= $true
    Force			= $true
}
Get-PnPFile @getPnPFileParams

Start-Sleep -s 3

# Connect to Microsoft Graph API
$graphParams  = @{
    Tenant					= $tenant
    AppId					= $appId
    CertificateThumbprint	= $thumbprint
}
Connect-Graph @graphParams

# List all users and their services in assigned licenses
$users = Get-MgUser -All
$excel = Open-ExcelPackage -Path ($localPath + $filename)
$excel.raport.Cells["A3:AP300"].Clear()
$zmienna = 3
$miejsce = @('D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP')

foreach ($usr in $users) {
    $licenses = Get-MgUserLicenseDetail -UserId $usr.id
    foreach ($lic in $licenses) {
        $excel.raport.Cells["A$zmienna"].Value = $usr.DisplayName
        $excel.raport.Cells["B$zmienna"].Value = $usr.Mail
        $excel.raport.Cells["C$zmienna"].Value = $lic.SkuPartNumber
        $serviceplan = $lic.ServicePlans
        foreach ($msc in $miejsce) {
            $statusWpisania = $false
            foreach ($plans in $serviceplan) {
                if ($plans.AppliesTo -eq "User") {
                    if ($excel.raport.Cells["$msc" + "1"].Value -eq $plans.ServicePlanName) {
                        $excel.raport.Cells["$msc$zmienna"].Value = $plans.ProvisioningStatus
                        $statusWpisania = $true
                    }
                }
            }
            if (!$statusWpisania) {
                $excel.raport.Cells["$msc$zmienna"].Value = "-"
                $statusWpisania = $false
            }
        }
        $zmienna++
    }
}
Close-ExcelPackage -ExcelPackage $excel

# Upload file from local path to SharePoint Online
$addPnPFileParams = @{
    Folder = $onlinePath
    Path   = ($localPath + $filename)
}
Add-PnPFile @addPnPFileParams
