<#
.SYNOPSIS
Transforms a set of publishing pages with the framework directly

IMPORTANT: this requires the PnP PowerShell version 3.8.1904.0 (April 2019) or higher to work!
Version: 1.0

This script depends on two files:
    Web Part Mapping file - to generate run: Export-PnPClientSidePageMapping -BuiltInWebPartMapping -Folder (Get-Location)
    Page Layout mapping file - to generate run: AnalyseLayouts.ps1 (update the Url in the file first)

.EXAMPLE
PS C:\> .\AnalyseLayouts.ps1
#>

# Connect to the web holding the pages to modernize
$creds = Get-Credential
$sourceConnection = Connect-PnPOnline -Url https://contoso.sharepoint.com/sites/sourcemodernizationme -Credentials $creds -ReturnConnection
$targetConnection = Connect-PnPOnline -Url https://contoso.sharepoint.com/sites/targetsite -Credentials $creds -ReturnConnection


$layoutMappingFile = "$(Get-Location)\PageLayoutMapping.xml" # Run AnalyseLayouts.ps1 to get Mapping File
$webPartMappingFile = "$(Get-Location)\webpartmapping.xml"

# Get PnP Modules to find the modernisation assemblies
$modules = (Get-Module -List "SharePointPnPPowerShellOnline")
$modulePath = ""
if($modules.Count -gt 0){
    $latest = $modules[0]
    $modulePath = $latest.ModuleBase
    Write-Host "Found module: $($modulePath)"
}else{
    throw "Cannot find PnP Modules"
}

if($modulePath){

    Add-Type -Path "$($modulePath)\SharePointPnP.Modernization.Framework.dll"

    Write-Host "Transforming pages" -ForegroundColor Cyan

    $pageTransformator = New-Object -TypeName SharePointPnP.Modernization.Framework.Publishing.PublishingPageTransformator `
                            -ArgumentList $sourceConnection.Context, `
                                          $targetConnection.Context, `
                                          $webPartMappingFile, `
                                          $layoutMappingFile
    
    # Generates a markdown report
    $markdownObserver = `
        New-Object -TypeName SharePointPnP.Modernization.Framework.Telemetry.Observers.MarkdownObserver -ArgumentList folder:"$(Get-Location)"                                       
    $pageTransformator.RegisterObserver($markdownObserver);    
 
    # Get Items from Pages Library
    $pages = Get-PnPListItem -List "Pages" -Connection $sourceConnection
    $pages | ForEach-Object{

        # Options for the transformation engine
        $pti = New-Object -TypeName SharePointPnP.Modernization.Framework.Publishing.PublishingPageTransformationInformation -ArgumentList $_
        $pti.Overwrite = $true

        Write-Host "Transforming page...$($_.FieldValues["FileLeafRef"])" -ForegroundColor Cyan 
        $result = $pageTransformator.Transform($pti)
        $pageTransformator.FlushObservers()

        Write-Host "Transformed file $($result)" -ForegroundColor Green
    }

    $pageTransformator.FlushObservers()
    
    Write-Host "Done :-)" -ForegroundColor Green

}else{
    Write-Error "Cannot find the transformation modules"
}