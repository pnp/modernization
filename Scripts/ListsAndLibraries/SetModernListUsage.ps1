<#
.SYNOPSIS
Enables or disables the modern list and library experience at site collection level. The script can handle a single site collection or a list of site collections provided via a CSV file. 

To get the CSV file you can run the Modernization Scanner, version 2.3 or higher, and use the "Lists and Library" mode (see https://aka.ms/sppnp-modernizationscanner) or alternatively 
create the file yourselves:

"https://contoso.sharepoint.com/sites/siteA"
"https://contoso.sharepoint.com/sites/siteB"
"https://contoso.sharepoint.com/sites/siteC"

.EXAMPLE
PS C:\> .\SetModernListUsage.ps1
#>

#region Logging and generic functions
function LogWrite
{
    param([string] $log , [string] $ForegroundColor)

    $global:strmWrtLog.writeLine($log)
    if([string]::IsNullOrEmpty($ForegroundColor))
    {
        Write-Host $log
    }
    else
    {    
        Write-Host $log -ForegroundColor $ForegroundColor
    }
}

function LogError
{
    param([string] $log)
    
    $global:strmWrtError.writeLine($log)
}

function UsageLog
{
    try 
    {
        $cc = Get-PnPContext
        $cc.Load($cc.Web)
        $cc.ClientTag = "SPDev:ModernListUsage"
        $cc.ExecuteQuery()
    }
    catch [Exception] { }
}
#endregion

function SiteCollectionUsesModernLists
{
    param([string] $siteCollectionUrl, 
          [Boolean] $useModern,
          $credentials,
          $tenantContext,
          [string] $adminUPN)
    
    
    #region Ensure access to the site collection, if needed promote the calling account to site collection admin
    # Check if we can access the site...if not let's 'promote' ourselves as site admin
    $adminClaim = "i:0#.f|membership|$adminUPN"    
    $adminWasAdded = $false
    $siteContext = $null    
    $siteCollectionUrl = $siteCollectionUrl.TrimEnd("/");

    Try
    {
        LogWrite "User running script: $adminUPN"
        LogWrite "Connecting to site $siteCollectionUrl"
        $siteContext = Connect-PnPOnline -Url $siteCollectionUrl -Credentials $credentials -Verbose -ReturnConnection
    }
    Catch [Exception]
    {
        # If Access Denied then use tenant API to add current tenant admin user as site collection admin to the current site
        if ($_.Exception.Response.StatusCode -eq "Unauthorized")
        {
            LogWrite "Temporarily adding user $adminUPN as site collection admin"
            Set-PnPTenantSite -Url $siteCollectionUrl -Owners @($adminUPN) -Connection $tenantContext
            $adminWasAdded = $true
            LogWrite "Second attempt to connect to site $siteCollectionUrl"
            $siteContext = Connect-PnPOnline -Url $siteCollectionUrl -Credentials $credentials -Verbose -ReturnConnection
        }
        else 
        {
            $ErrorMessage = $_.Exception.Message
            LogWrite "Error for site $siteCollectionUrl : $ErrorMessage" Red
            LogError $ErrorMessage
            return              
        }
    }
    #endregion

    Try
    {
        #region Adding admin
        # Check if current tenant admin is part of the site collection admins, if not add the account        
        $siteAdmins = $null
        if ($adminWasAdded -eq $false)
        {
            try 
            {
                # Eat exceptions here...resulting $siteAdmins variable will be empty which will trigger the needed actions                
                $siteAdmins = Get-PnPSiteCollectionAdmin -Connection $siteContext -ErrorAction Ignore
            }
            catch [Exception] { }
            
            $adminNeedToBeAdded = $true
            foreach($admin in $siteAdmins)
            {
                if ($admin.LoginName -eq $adminClaim)
                {
                    $adminNeedToBeAdded = $false
                    break
                }
            }

            if ($adminNeedToBeAdded)
            {
                LogWrite "Temporarily adding user $adminUPN as site collection admin"
                Set-PnPTenantSite -Url $siteCollectionUrl -Owners @($adminUPN) -Connection $tenantContext
                $adminWasAdded = $true
            }
        }

        UsageLog
        #endregion
        
        #region Enable/disable the modern list experience at site collection level
        if ($useModern)
        {
            LogWrite "Disabling the modern list blocking feature"
            Disable-PnPFeature -Identity "E3540C7D-6BEA-403C-A224-1A12EAFEE4C4" -Scope Site -Force -Connection $siteContext
        }
        else
        {
            LogWrite "Enabling the modern list blocking feature"
            Enable-PnPFeature -Identity "E3540C7D-6BEA-403C-A224-1A12EAFEE4C4" -Scope Site -Force -Connection $siteContext
        }
        #endregion

        #region Cleanup updated permissions
        LogWrite "Configuration is done, let's cleanup the configured permissions"
    
        # Remove the added site collection admin - obviously this needs to be the final step in the script :-)
        if ($adminWasAdded)
        {
            LogWrite "Remove $adminUPN from site collection administrators"            
            Remove-PnPSiteCollectionAdmin -Owners @($adminUPN) -Connection $siteContext
        }
        #endregion

        LogWrite "Configuration done for site collection $siteCollectionUrl" Green
        
        # Disconnect PnP Powershell from site
        Disconnect-PnPOnline
    }
    Catch [Exception]
    {
        $ErrorMessage = $_.Exception.Message
        LogWrite "Error: $ErrorMessage" Red
        LogError $ErrorMessage

        #region Cleanup updated permissions on error
        # Configuration did not complete...remove the added tenant admin to restore site permissions as final step in the cleanup
        if ($adminWasAdded)
        {
            try 
            {
                # Final step, remove the added site collection admin
                Remove-PnPSiteCollectionAdmin -Owners @($adminUPN) -Connection $siteContext
            }
            catch [Exception] { }
        }
        #endregion

        LogWrite "Configuration failed for site collection $siteCollectionUrl" Red
    } 

}

#######################################################
# MAIN section                                        #
#######################################################

# OVERRIDES
# If you want to automate the run and make the script ask less questions, feel free to hardcode these 2 values below. Otherwise they'll be asked from the user or parsed from the values they input

# Tenant admin url
$tenantAdminUrl = "" # e.g. "https://contoso-admin.sharepoint.com"
# If you use credential manager then specify the used credential manager entry, if left blank you'll be asked for a user/pwd
$credentialManagerCredentialToUse = ""

#region Setup Logging
$date = Get-Date
$logfile = ((Get-Item -Path ".\" -Verbose).FullName + "\ModernListUsage_log_" + $date.ToFileTime() + ".txt")
$global:strmWrtLog=[System.IO.StreamWriter]$logfile
$global:Errorfile = ((Get-Item -Path ".\" -Verbose).FullName + "\ModernListUsage_error_" + $date.ToFileTime() + ".txt")
$global:strmWrtError=[System.IO.StreamWriter]$Errorfile
#endregion

#region Load needed PowerShell modules
# Ensure PnP PowerShell is loaded

$minimumVersion = New-Object System.Version("3.4.1812.2")
if (-not (Get-InstalledModule -Name SharePointPnPPowerShellOnline -MinimumVersion $minimumVersion -ErrorAction Ignore)) 
{
    Install-Module SharePointPnPPowerShellOnline -MinimumVersion $minimumVersion -Scope CurrentUser
}
Import-Module SharePointPnPPowerShellOnline -DisableNameChecking -MinimumVersion $minimumVersion
#endregion

#region Gather set modern list usage run input
# Url of the site collection to remediate
$siteCollectionUrlToConfigure = ""
$enableModern = $false

# Get the input information
$siteURLFile = Read-Host -Prompt "Input either single site collection URL (e.g. https://contoso.sharepoint.com/sites/teamsite1) or name of .CSV file (e.g. SitesWithCustomizations.csv) ?"
if (-not $siteURLFile.EndsWith(".csv"))
{
    $siteCollectionUrlToConfigure = $siteURLFile
}
# If we are using a CSV, we'll need to get the tenant admin url from the user or use the hardcoded one
else 
{
    if ($tenantAdminUrl -eq $null -or $tenantAdminUrl.Length -le 0) 
    {
        $tenantAdminUrl = Read-Host -Prompt "Input the tenant admin site URL (like https://contoso-admin.sharepoint.com)"
    }
}

$enableModernString = Read-Host -Prompt "Do you want to enable modern lists and libraries for this site collection? Enter True for yes, False otherwise"
try 
{
    $enableModern = [System.Convert]::ToBoolean($enableModernString) 
} 
catch [FormatException]
{
    $enableModern = $false
}

# We'll parse the tenantAdminUrl from site url (unless it's set already!)
if ($tenantAdminUrl -eq $null -or $tenantAdminUrl.Length -le 0) 
{
    if ($siteURLFile.IndexOf("/teams") -gt 0) 
    {
        $tenantAdminUrl = $siteURLFile.Substring(0, $siteURLFile.IndexOf("/teams")).Replace(".sharepoint.", "-admin.sharepoint.")
    }
    else 
    {
        $tenantAdminUrl = $siteURLFile.Substring(0, $siteURLFile.IndexOf("/sites")).Replace(".sharepoint.", "-admin.sharepoint.")
    }
}

# Get the tenant admin credentials.
$credentials = $null
$adminUPN = $null
if(![String]::IsNullOrEmpty($credentialManagerCredentialToUse) -and (Get-PnPStoredCredential -Name $credentialManagerCredentialToUse) -ne $null)
{
    $adminUPN = (Get-PnPStoredCredential -Name $credentialManagerCredentialToUse).UserName
    $credentials = $credentialManagerCredentialToUse
}
else
{
    # Prompts for credentials, if not found in the Windows Credential Manager.
    $adminUPN = Read-Host -Prompt "Please enter admin UPN (e.g. admin@contoso.onmicrosoft.com)"
    $pass = Read-host -AsSecureString "Please enter admin password"
    $credentials = new-object management.automation.pscredential $adminUPN,$pass
}

if($credentials -eq $null) 
{
    Write-Host "Error: No credentials supplied." -ForegroundColor Red
    exit 1
}
#endregion

#region Connect to SharePoint
# Get a tenant admin connection, will be reused in the remainder of the script
LogWrite "Connect to tenant admin site $tenantAdminUrl"
$tenantContext = Connect-PnPOnline -Url $tenantAdminUrl -Credentials $credentials -Verbose -ReturnConnection
#endregion

#region Configure the site(s)
if (-not $siteURLFile.EndsWith(".csv"))
{
    # Remediate the given site collection
    SiteCollectionUsesModernLists $siteCollectionUrlToConfigure $enableModern $credentials $tenantContext $adminUPN
}
else 
{
    $csvRows = Import-Csv $siteURLFile -Header SiteCollectionUrl
    
    foreach($row in $csvRows)
    {
        if($row.SiteCollectionUrl -ne "")
        {
            $siteUrl = $row.SiteCollectionUrl
            SiteCollectionUsesModernLists $siteUrl $enableModern $credentials $tenantContext $adminUPN
        }
    }
}
#endregion

#region Close log files
if ($global:strmWrtLog -ne $NULL)
{
    $global:strmWrtLog.Close()
    $global:strmWrtLog.Dispose()
}

if ($global:strmWrtError -ne $NULL)
{
    $global:strmWrtError.Close()
    $global:strmWrtError.Dispose()
}
#endregion