<# 

.Synopsis

    Converts all Delve blog pages in a site

    Sample includes:
        - Conversion of Delve blog pages
        - Connecting to MFA or supplying credentials
        - Includes Logging to File, log flushing into single log file        

.Example

    Convert-DelveBlogPages.ps1 -SourceUrl "https://contoso.sharepoint.com/portals/personal/joedoe" -TargetUrl "https://contoso.sharepoint.com/sites/modernblog" -Credentials Get-Credential

.Notes
    
    Useful references:
        - https://aka.ms/sppnp-pagetransformation
#>

[CmdletBinding()]
param (

    [Parameter(Mandatory = $true, HelpMessage = "Delve blog site url")]
    [string]$SourceUrl,

    [Parameter(Mandatory = $true, HelpMessage = "Target modern communication site url")]
    [string]$TargetUrl,

    [Parameter(Mandatory = $false, HelpMessage = "Supply credentials for multiple runs/sites")]
    [PSCredential]$Credentials,

    [Parameter(Mandatory = $false, HelpMessage = "Specify log file location")]
    [string]$LogOutputFolder = "c:\temp"
)
begin
{
    Write-Host "Connecting to " $SourceUrl
    
    if($Credentials)
    {
        Connect-PnPOnline -Url $SourceUrl -Credentials $Credentials -Verbose
    }
    else
    {
        # Sometimes this fails and a retry is needed
        Connect-PnPOnline -Url $sourceUrl -SPOManagementShell -ClearTokenCache -Verbose
        Start-Sleep -s 3
    }
}
process 
{    

    Write-Host "Modernizing Delve blog pages..." -ForegroundColor Cyan

    $posts = Get-PnPListItem -List "pPg"

    Write-Host "pages fetched"

    Foreach($post in $posts)
    {
        $postTitle = $post.FieldValues["Title"]

        Write-Host " Processing Delve blog post $($postTitle)"

        ConvertTo-PnPClientSidePage -Identity $postTitle `
                                    -DelveBlogPage `
                                    -Overwrite `
                                    -TargetWebUrl $TargetUrl `
                                    -LogType File `
                                    -LogSkipFlush `
                                    -LogFolder $LogOutputFolder `
                                    -KeepPageCreationModificationInformation `
                                    -PostAsNews `
                                    -SetAuthorInPageHeader `
                                    -DelveKeepSubTitle

    }

    # Write the logs to the folder
    Save-PnPClientSidePageConversionLog

    Write-Host "Delve blog site modernization complete! :)" -ForegroundColor Green

    Disconnect-PnPOnline
}