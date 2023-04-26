$tenant = "MYTENANTNAME"
Import-Module Microsoft.Online.SharePoint.PowerShell
Connect-SPOService -url "https://$tenant-admin.sharepoint.com"
Set-SPOTenant -IsLoopEnabled $true
Disconnect-SPOService