<# 

.Synopsis

    Converts all classic wiki and web part pages in a site. 
    You need to install PnP PowerShell version 3.16.1912.* (December 2019) or higher to use this script.

    Sample includes:
        - Conversion of wiki and web part pages
        - Connecting to MFA or supplying credentials
        - Includes Logging to File, log flushing into single log file        

.Example

    Convert-WikiAndWebPartPages.ps1 -SourceUrl "https://contoso.sharepoint.com/sites/classicteamsite" -TakeSourcePageName:$true

.Notes
    
    Useful references:
        - https://aka.ms/sppnp-pagetransformation
        - https://aka.ms/sppnp-powershell
#>

[CmdletBinding()]
param (

    [Parameter(Mandatory = $true, HelpMessage = "Url of the site containing the pages to modernize")]
    [string]$SourceUrl,

    [Parameter(Mandatory = $false, HelpMessage = "Modern page takes source page name")]
    [bool]$TakeSourcePageName = $false,    

    [Parameter(Mandatory = $false, HelpMessage = "Supply credentials for multiple runs/sites")]
    [PSCredential]$Credentials,

    [Parameter(Mandatory = $false, HelpMessage = "Specify log file location, defaults to the same folder as the script is in")]
    [string]$LogOutputFolder = $(Get-Location)
)
begin
{
    Write-Host "Connecting to " $SourceUrl
    
    if($Credentials)
    {
        Connect-PnPOnline -Url $SourceUrl -Credentials $Credentials
        Start-Sleep -s 3
    }
    else
    {
        Connect-PnPOnline -Url $sourceUrl -SPOManagementShell -ClearTokenCache
        Start-Sleep -s 3
    }
}
process 
{    
    Write-Host "Ensure the modern page feature is enabled..." -ForegroundColor Green
    Enable-PnPFeature -Identity "B6917CB1-93A0-4B97-A84D-7CF49975D4EC" -Scope Web -Force

    Write-Host "Modernizing wiki and web part pages..." -ForegroundColor Gree
    # Get all the pages in the site pages library. 
    # Use paging (-PageSize parameter) to ensure the query works when there are more than 5000 items in the list
    $pages = Get-PnPListItem -List sitepages -PageSize 500

    Write-Host "Pages are fetched, let's start the modernization..." -ForegroundColor Green
    Foreach($page in $pages)
    { 
        $pageName = $page.FieldValues["FileLeafRef"]
        
        if ($page.FieldValues["ClientSideApplicationId"] -eq "b6917cb1-93a0-4b97-a84d-7cf49975d4ec" ) 
        { 
            Write-Host "Page " $page.FieldValues["FileLeafRef"] " is modern, no need to modernize it again" -ForegroundColor Yellow
        } 
        else 
        {
            Write-Host "Processing page $($pageName)" -ForegroundColor Cyan
            
            # -TakeSourcePageName:
            # The old pages will be renamed to Previous_<pagename>.aspx. If you want to 
            # keep the old page names as is then set the TakeSourcePageName to $false. 
            # You then will see the new modern page be named Migrated_<pagename>.aspx

            # -Overwrite:
            # Overwrites the target page (needed if you run the modernization multiple times)
            
            # -LogVerbose:
            # Add this switch to enable verbose logging if you want more details logged

            # KeepPageCreationModificationInformation:
            # Give the newly created page the same page author/editor/created/modified information 
            # as the original page. Remove this switch if you don't like that

            # -CopyPageMetadata:
            # Copies metadata of the original page to the created modern page. Remove this
            # switch if you don't want to copy the page metadata

            ConvertTo-PnPClientSidePage -Identity $page.FieldValues["ID"] `
                                        -Overwrite `
                                        -TakeSourcePageName:$TakeSourcePageName `
                                        -LogType File `
                                        -LogFolder $LogOutputFolder `
                                        -LogSkipFlush `
                                        -KeepPageCreationModificationInformation `
                                        -CopyPageMetadata
        }
    }

    # Write the logs to the folder
    Write-Host "Writing the conversion log file..." -ForegroundColor Green
    Save-PnPClientSidePageConversionLog

    Write-Host "Wiki and web part page modernization complete! :)" -ForegroundColor Green
}
end
{
    Disconnect-PnPOnline
}
