<# 

Created:      Paul Bullock
Date:         07/08/2019
License:      MIT License (MIT)
Disclaimer:   

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.


.Synopsis

    Converts all publishing pages from an on-premises server

    You need to install PnP PowerShell: https://pnp.github.io/powershell/

    Sample includes:
        - Conversion of publishing pages from on-premises
        - Renaming default/welcome subsite pages
        - Includes Logging to File, log flushing into single log file
        - Post Processing of file after transformation

    To generate mapping files, see Export-PnPPageMapping cmdlet:
        e.g Export-PnPPageMapping -CustomPageLayoutMapping -BuiltInWebPartMapping -Folder (Get-Location)

.Example

    Collect Credentials
    
        $sourceCredentials = Get-Credential
        $targetCredentials = Get-Credential
    
    Generate mapping and store in "Mapping-2013" folder
    
        $sourceConn = Connect-PnPOnline http://portal2013 -Credentials $sourceCredentials -TransformationOnPrem -ReturnConnection

    To generate mapping files, see Export-PnPClientSidePageMapping cmdlet
    Get mapping for Single Page:

        Export-PnPPageMapping -CustomPageLayoutMapping -Connection $sourceConn -Folder "$(Get-Location)\Mapping-2010"
    
    Get mapping for all page layouts
        
        Export-PnPPageMapping -PublishingPage "Article-2013-Custom.aspx" -CustomPageLayoutMapping -Connection $sourceConn -Folder "$(Get-Location)\Mapping-2013"
    
    Run this transform script example
    
        .\Convert-OnPremisesPublishingPages.ps1 -SourceCredentials $sourceCredentials -SourceUrl "http://portal2013" -TargetSitePartUrl "PnPKatchup" -TargetCredentials $targetCredentials

.Notes
    
    Useful references:
        - https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/convertto-pnpclientsidepage?view=sharepoint-ps
        - https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/?view=sharepoint-ps
        
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "Source e.g. Intranet-Archive")]
    [string]$SourceUrl,

    [Parameter(Mandatory = $true, HelpMessage = "Target e.g. Intranet")]
    [string]$TargetSitePartUrl,

    [Parameter(Mandatory = $false, HelpMessage = "Organisation Url Fragment e.g. contoso ")]
    [string]$PartTenant = "contoso",

    [Parameter(Mandatory = $false, HelpMessage = "Supply Credentials for multiple runs/sites")]
    [PSCredential]$SourceCredentials,

    [Parameter(Mandatory = $false, HelpMessage = "Supply Credentials for multiple runs/sites")]
    [PSCredential]$TargetCredentials,

    [Parameter(Mandatory = $false, HelpMessage = "Specify Mapping File")]
    [string]$WebPartMappingFile = "webpartmapping.xml",
    
    [Parameter(Mandatory = $false, HelpMessage = "Specify Page Layout File")]
    [string]$PageLayoutMappingFile = "pagelayoutmapping.xml",
    
    [Parameter(Mandatory = $false, HelpMessage = "Specify log file location")]
    [string]$LogOutputFolder = "c:\temp"
)
begin{

    $baseUrl = "https://$($PartTenant).sharepoint.com"
    $sourceSiteUrl = $SourceUrl
    $targetSiteUrl = "$($baseUrl)/sites/$($TargetSitePartUrl)"

    Write-Host "Connecting to " $sourceSiteUrl " and " $targetSiteUrl
    
    # To transform to On-Premises servers you need to create connections to both source and target
    $sourceConnection = Connect-PnPOnline -Url $sourceSiteUrl -TransformationOnPrem -ReturnConnection -Credentials $SourceCredentials
    Write-Host "Connected to " $sourceSiteUrl

    # This connection should target SharePoint Online
    $targetConnection = Connect-PnPOnline -Url $targetSiteUrl -ReturnConnection -Credentials $TargetCredentials
    Write-Host "Connected to " $targetSiteUrl

    $location = Get-Location
}
process {

    Write-Host "Converting site..." -ForegroundColor Cyan

    $pages = Get-PnPListItem -List "Pages" -Connection $sourceConnection
        
    Foreach($page in $pages){

        $targetFileName = $page.FieldValues["FileLeafRef"]

        Write-Host " Processing $($targetFileName)" -ForegroundColor Cyan

        # If Welcome Page, then Rename, 
        # typical for flattening multiple sites that contain standard page(s) e.g. Welcome.aspx or Default.aspx
        if($targetFileName -eq "Welcome.aspx"){

            #$targetFileName  = "Welcome-2013.aspx"
            #Write-Host " - Updating Welcome.aspx page to $($targetFileName)" -ForegroundColor Yellow
            Write-Host "  Skipping welcome.aspx" -ForegroundColor Yellow
            continue
        }

        if($targetFileName -eq "Default.aspx"){

            #$targetFileName  = "Default-2013.aspx"
            #Write-Host " - Updating Default.aspx page to $($targetFileName)" -ForegroundColor Yellow
            Write-Host "  Skipping default.aspx" -ForegroundColor Yellow
            continue
        }

        Write-Host " Modernizing $($targetFileName)..." -ForegroundColor Cyan

        # Use -Connection parameter to pass the source 201X SharePoint connection
        # Use -TargetConnection to pass in the target connection to SharePoint Online modern site,
        #   no need to use -TargetUrl in this case
        $result = ConvertTo-PnPPage -Identity $page.FieldValues["FileLeafRef"] `
                    -PublishingPage `
                    -TargetConnection $targetConnection `
                    -PublishingTargetPageName $targetFileName `
                    -WebPartMappingFile "$($location)\$($WebPartMappingFile)" `
                    -PageLayoutMapping "$($location)\$($PageLayoutMappingFile)" `
                    -Connection $sourceConnection `
                    -Overwrite `
                    -SkipItemLevelPermissionCopyToClientSidePage `
                    -LogType File `
                    -CopyPageMetadata `
                    -LogSkipFlush `
                    -LogFolder $LogOutputFolder
    
        if($result){

            Write-Host "  Transformed page: " $result -ForegroundColor Green

            # Running processing tasks on target connection only
            # Post Processing actions on file
            $transformedItem = Get-PnPFile -Url $result -AsListItem -Connection $targetConnection
            if($transformedItem){
                Write-Host " - Post Processing $($targetFileName)..."
                # Peform changes...
            }
        }else{
            Write-Host "  Page not transformed, check the logs for issues" -ForegroundColor Red
        }
    }

    # Write the logs to the folder
    Save-PnPPageConversionLog

    Write-Host "Script Complete! :)" -ForegroundColor Green
}
