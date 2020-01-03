<# 

Created:      Paul Bullock
Date:         26/06/2019
License:      MIT License (MIT)
Disclaimer:   
Version:      1.0

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.


.Synopsis

    Converts all publishing pages in a site

    Sample includes:
        - Conversion of publishing pages
        - Renaming default/welcome subsite pages
        - Connecting to MFA or supplying credentials
        - Includes Logging to File, log flushing into single log file
        - Post Processing of file after transformation

    To generate mapping files, see Export-PnPClientSidePageMapping cmdlet:
        e.g Export-PnPClientSidePageMapping -CustomPageLayoutMapping -BuiltInWebPartMapping -Folder (Get-Location)

.Example

    $creds = Get-Credential
    Convert-PublishingPages.ps1 -Credentials $creds -SourceSitePartUrl "Intranet-Archive" -TargetSitePartUrl "Intranet"

.Notes
    
    Useful references:
        - https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/convertto-pnpclientsidepage?view=sharepoint-ps
        - https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/?view=sharepoint-ps

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "Source e.g. Intranet-Archive")]
    [string]$SourceSitePartUrl,

    [Parameter(Mandatory = $true, HelpMessage = "Target e.g. Intranet")]
    [string]$TargetSitePartUrl,

    [Parameter(Mandatory = $false, HelpMessage = "Organisation Url Fragment e.g. contoso ")]
    [string]$PartTenant = "contoso",

    [Parameter(Mandatory = $false, HelpMessage = "Supply Credentials for multiple runs/sites")]
    [PSCredential]$Credentials,

    [Parameter(Mandatory = $false, HelpMessage = "Specify Mapping File")]
    [string]$WebPartMappingFile = "webpartmapping.xml",
    
    [Parameter(Mandatory = $false, HelpMessage = "Specify Page Layout File")]
    [string]$PageLayoutMappingFile = "pagelayoutmapping.xml",
    
    [Parameter(Mandatory = $false, HelpMessage = "Specify log file location")]
    [string]$LogOutputFolder = "c:\temp"
)
begin{

    $baseUrl = "https://$($PartTenant).sharepoint.com"
    $sourceSiteUrl = "$($baseUrl)/sites/$($SourceSitePartUrl)"
    $targetSiteUrl = "$($baseUrl)/sites/$($TargetSitePartUrl)"

    Write-Host "Connecting to " $sourceSiteUrl
    
    if($Credentials){

        $sourceConnection = Connect-PnPOnline -Url $sourceSiteUrl -ReturnConnection -Credentials $Credentials
      
        # For post transform processing
        $targetConnection = Connect-PnPOnline -Url $targetSiteUrl -ReturnConnection -Credentials $Credentials
    }
    else
    {
        # For MFA Tenants - UseWebLogin opens a browser window
        $sourceConnection = Connect-PnPOnline -Url $sourceSiteUrl  -ReturnConnection -UseWebLogin
        
        # For post transform processing
        $targetConnection = Connect-PnPOnline -Url $targetSiteUrl  -ReturnConnection -UseWebLogin
    }

    $location = Get-Location
}
process {

    Write-Host "Converting site..." -ForegroundColor Cyan

    $web = Get-PnPWeb -Connection $sourceConnection
    # Use paging (-PageSize parameter) to ensure the query works when there are more than 5000 items in the list
    $pages = Get-PnPListItem -List "Pages" -Connection $sourceConnection -PageSize 500
        
    Foreach($page in $pages){

        $targetFileName = $page.FieldValues["FileLeafRef"]

        Write-Host " Processing $($targetFileName)"

        # If Welcome Page, then Rename, 
        # typical for flattening multiple sites that contain standard page(s) e.g. Welcome.aspx or Default.aspx
        if($targetFileName -eq "Welcome.aspx"){

            $targetFileName  = "Welcome-$($web.Title.Replace(" ", "-")).aspx"
            Write-Host " - Updating Welcome.aspx page to $($targetFileName)" -ForegroundColor Yellow

        }

        if($targetFileName -eq "Default.aspx"){

            $targetFileName  = "Default-$($web.Title.Replace(" ", "-")).aspx"
            Write-Host " - Updating Default.aspx page to $($targetFileName)" -ForegroundColor Yellow

        }

        Write-Host " Modernizing $($targetFileName)..."

        # Use the PageID value instead of the page name in the Identity parameter as that is more performant + it works when there are more than 5000 items in the list
        $result = ConvertTo-PnPClientSidePage -Identity $page.FieldValues["ID"] `
                    -PublishingPage `
                    -TargetWebUrl $targetSiteUrl `
                    -PublishingTargetPageName $targetFileName `
                    -WebPartMappingFile "$($location)\$($WebPartMappingFile)" `
                    -PageLayoutMapping "$($location)\$($PageLayoutMappingFile)" `
                    -Connection $sourceConnection `
                    -DontPublish `
                    -Overwrite `
                    -SkipItemLevelPermissionCopyToClientSidePage `
                    -LogSkipFlush `
                    -LogType File `
                    -CopyPageMetadata `
                    -LogFolder $LogOutputFolder
    
        if($result){

            # Post Processing actions on file
            $transformedItem = Get-PnPFile -Url $result -AsListItem -Connection $targetConnection
            if($transformedItem){
                Write-Host " - Post Processing $($targetFileName)..."
                # Peform changes...
            }

        }

        Write-Host " Modernized $($targetFileName)!"
        break;
    }

    # Write the logs to the folder
    Save-PnPClientSidePageConversionLog

    Write-Host "Script Complete! :)" -ForegroundColor Green
}
