# Report_M365_users_services

## Description
This PowerShell script downloads a file from SharePoint Online to a local path, connects to Microsoft Graph API, lists all users and their services in assigned licenses, saves this information in an Excel file, and uploads the file to SharePoint Online.

## Prerequisites
- PowerShell version 5.1 or later
- Installed SharePointPnPPowerShellOnline module
- Installed Microsoft.Graph module
- Installed ImportExcel module

## Configuration
Before running the script, the following variables need to be configured:

- `$filename`: Complete with filename (e.g., Raport_M365_users_services.xlsx)
- `$localPath`: Complete with local path where the file will be downloaded (e.g., C:\Raporty\)
- `$siteUrl`: Complete with URL site where SharePoint Online is hosted (e.g., https://company.sharepoint.com/sites/it-dep)
- `$onlinePath`: Complete with path where the file will be uploaded on SharePoint Online (e.g., Shared Documents/Global/)
- `$tenant`: Complete with the tenant name (e.g., company.onmicrosoft.com)
- `$appId`: Complete with the Client ID (which is ID of the application registered in Azure AD)
- `$thumbprint`: Complete with the Thumbprint (which is the certificate thumbprint)
