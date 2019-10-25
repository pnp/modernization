<# 

.Synopsis

    Converts all blog pages in a site

    Sample includes:
        - Conversion of blog pages
        - Connecting to MFA or supplying credentials
        - Includes Logging to File, log flushing into single log file        

.Example

    Convert-BlogPages.ps1 -SourceUrl "https://contoso.sharepoint.com/sites/classicblog" -TargetUrl "https://contoso.sharepoint.com/sites/modernblog"

.Notes
    
    Useful references:
        - https://aka.ms/sppnp-pagetransformation
#>

[CmdletBinding()]
param (

    [Parameter(Mandatory = $true, HelpMessage = "Classic blog site url")]
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
        Start-Sleep -s 3
    }
    else
    {
        Connect-PnPOnline -Url $sourceUrl -SPOManagementShell -ClearTokenCache -Verbose
        Start-Sleep -s 3
    }
}
process 
{    

    Write-Host "Modernizing blog pages..." -ForegroundColor Cyan

    $posts = Get-PnPListItem -List "Posts"

    Write-Host "pages fetched"

    Foreach($post in $posts)
    {
        $postTitle = $post.FieldValues["Title"]

        Write-Host " Processing blog post $($postTitle)"

        ConvertTo-PnPClientSidePage -Identity $postTitle `
                                    -BlogPage `
                                    -Overwrite `
                                    -TargetWebUrl $TargetUrl `
                                    -LogType File `
                                    -LogVerbose `
                                    -LogSkipFlush `
                                    -LogFolder $LogOutputFolder `
                                    -KeepPageCreationModificationInformation `
                                    -PostAsNews `
                                    -SetAuthorInPageHeader `
                                    -CopyPageMetadata

    }

    # Write the logs to the folder
    Save-PnPClientSidePageConversionLog

    Write-Host "Blog site modernization complete! :)" -ForegroundColor Green
}