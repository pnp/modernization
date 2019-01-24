<#
.SYNOPSIS
Upgrades your current version to the beta release of the Page Transformation UI solution

.EXAMPLE
PS C:\> .\Upgrade-ToBeta.ps1 -ModernizationCenterUrl https://contoso.sharepoint.com/sites/modernizationcenter -AssetsFolder "..\assets"
#>

param(
    [Parameter(Mandatory=$True)]
    [string] $ModernizationCenterUrl,
    [Parameter(Mandatory=$True)]
    [string] $AssetsFolder
 )

Write-Host ""
Write-Host "Connecting to " $ModernizationCenterUrl -ForegroundColor Yellow
Connect-PnPOnline -Url $ModernizationCenterUrl

Write-Host "Upgrading the Page Transformation Apps in the tenant app catalog"  -ForegroundColor Yellow
# Upload and publish the latest versions of the SPFX packages
Add-PnPApp -Path ($AssetsFolder + "\sharepointpnp-pagetransformation-central.sppkg") -Scope Tenant -Publish -Overwrite
Add-PnPApp -Path ($AssetsFolder + "\sharepointpnp-pagetransformation-client.sppkg") -Scope Tenant -Publish -Overwrite -SkipFeatureDeployment

Write-Host "Upgrading the page transformation central app in the modernization center site collection"  -ForegroundColor Yellow
# Upgrade the version installed in the modernization center
$appToUpgrade = Get-PnPApp | Where-Object {$_.Title -eq "sharepointpnp-pagetransformation-central-solution"}
Update-PnPApp -Identity $appToUpgrade.Id

Write-Host "Upgrading SiteAssets"  -ForegroundColor Yellow
#Copy in the latest version of the classic banner script
$updatedItem = Add-PnPFile -Path ($AssetsFolder + "\pnppagetransformationclassicbanner.js") -Folder SiteAssets

Write-Host ""
Write-Host "**************************************************************************************************" -ForegroundColor Yellow
Write-Host "If you want to have a banner on the classic pages indicating there is a modern version then please re-enable the Page Transformation UI for your site collections."
Write-Host "See https://github.com/SharePoint/sp-dev-modernization/blob/master/Solutions/PageTransformationUI/docs/deploymentguide.md#step-3-enable-the-page-transformation-ui-for-your-site-collections for more details."
Write-Host ""
Write-Host "**************************************************************************************************" -ForegroundColor Yellow
Write-Host ""
Write-Host "Upgrade to the beta release of Page Transformation UI done!" -ForegroundColor Green