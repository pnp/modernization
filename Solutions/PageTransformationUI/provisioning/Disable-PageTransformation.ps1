$site = Get-PnPSite -Includes ServerRelativeUrl
Write-Host "Disabling page transformation for $($site.ServerRelativeUrl)" -ForegroundColor White

Remove-PnPCustomAction -Identity "CA_PnP_Modernize_SitePages_RIBBON" -Scope Site -Force
Remove-PnPCustomAction -Identity "CA_PnP_Modernize_SitePages_ECB" -Scope Site -Force 
Remove-PnPCustomAction -Identity "CA_PnP_Modernize_WikiPage_RIBBON" -Scope Site -Force
Remove-PnPCustomAction -Identity "CA_PnP_Modernize_WebPartPage_RIBBON" -Scope Site -Force

Write-Host "Done" -ForegroundColor Green