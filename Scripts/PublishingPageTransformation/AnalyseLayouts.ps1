<#
.SYNOPSIS
Analyses a set of pages individually by publishing page

IMPORTANT: this requires the PnP PowerShell version 3.8.1904.0 (April 2019) or higher to work!
Version: 1.0

.EXAMPLE
PS C:\> .\AnalyseLayouts.ps1
#>

# Connect to the web holding the pages to modernize
Connect-PnPOnline -Url https://contoso.sharepoint.com/sites/modernizationme

$ctx = Get-PnPContext

$modules = (Get-Module -List "SharePointPnPPowerShellOnline")
$modulePath = ""
if($modules.Count -gt 0){

    $latest = $modules[0]
    $modulePath = $latest.ModuleBase

}else{
    throw "Cannot find PnP Modules"
}

Write-Host "Found module: $($modulePath)"

if($modulePath){

    Add-Type -Path "$($modulePath)\SharePointPnP.Modernization.Framework.dll"

    Write-Host "Analysing Layouts" -ForegroundColor Cyan

    $pages = Get-PnPListItem -List "Pages"
    $analyser = New-Object -TypeName SharePointPnP.Modernization.Framework.Publishing.PageLayoutAnalyser -ArgumentList $ctx

    $pages | ForEach-Object{

        Write-Host "Analysing layout"
        $analyser.AnalysePageLayoutFromPublishingPage($_)

    }

    $analyser.GenerateMappingFile((Get-Location), "PageLayoutMapping.xml")

    Write-Host "Done :-)" -ForegroundColor Green

}else{
    Write-Error "Cannot find the transformation modules"
}