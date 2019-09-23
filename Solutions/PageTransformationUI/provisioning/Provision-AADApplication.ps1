<#
.SYNOPSIS
Creates the Azure AD application needed for the Page Transformation UI solution. This script will be called from Provision-AADApplication.ps1, no need to run it manually
#>

Param (
    [Parameter(Mandatory=$true, Position=0)]
    [String] $AppName,
    [Parameter(Mandatory=$true, Position=1)]
    [String] $HomePageUrl,
    [Parameter(Mandatory=$true, Position=2)]
    [String] $ReplyUrl,
    [Parameter(Mandatory=$true, Position=3)]
    [String] $AppTitle    
)

function GeneratePasswordCredential() {
    $keyId = (New-Guid).ToString();
    $startDate = Get-Date
    $endDate = $startDate.AddYears(10)
    $random = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    [byte[]]$bytes = New-Object byte[] 32
    $random.GetBytes($bytes)
    $appSecret = [System.Convert]::ToBase64String($bytes)
    $passwordCredentials = New-Object Microsoft.Open.AzureAD.Model.PasswordCredential($null, $endDate, $keyId, $startDate, $appSecret)
    return $passwordCredentials
}

function GenerateRequiredResourceAccess() {
    $requiredResourcesAccess = New-Object System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.RequiredResourceAccess]
    
    # Site Collections Full Control Delegated permission request
    $siteCollectionsFullControlDelegated = New-Object Microsoft.Open.AzureAD.Model.RequiredResourceAccess
    $siteCollectionsFullControlDelegated.ResourceAppId = "00000003-0000-0ff1-ce00-000000000000"
    $siteCollectionsFullControlDelegated.ResourceAccess = New-Object System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.ResourceAccess]
    $siteCollectionsFullControlDelegatedPermission = New-Object Microsoft.Open.AzureAD.Model.ResourceAccess
    $siteCollectionsFullControlDelegatedPermission.Id = "56680e0d-d2a3-4ae1-80d8-3c4f2100e3d0"
    $siteCollectionsFullControlDelegatedPermission.Type = "Scope"
    $siteCollectionsFullControlDelegated.ResourceAccess.Add($siteCollectionsFullControlDelegatedPermission)
        
    # Sign In and Read User Profile Delegated permission request
    $signInReadProfileDelegated = New-Object Microsoft.Open.AzureAD.Model.RequiredResourceAccess
    $signInReadProfileDelegated.ResourceAppId = "00000002-0000-0000-c000-000000000000"
    $signInReadProfileDelegated.ResourceAccess = New-Object System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.ResourceAccess]
    $signInReadProfileDelegatedPermission = New-Object Microsoft.Open.AzureAD.Model.ResourceAccess
    $signInReadProfileDelegatedPermission.Id = "311a71cc-e848-46a1-bdf8-97ff7156d8e6"
    $signInReadProfileDelegatedPermission.Type = "Scope"
    $signInReadProfileDelegated.ResourceAccess.Add($signInReadProfileDelegatedPermission)
    
    $requiredResourcesAccess.Add($siteCollectionsFullControlDelegated)
    $requiredResourcesAccess.Add($signInReadProfileDelegated)

    return $requiredResourcesAccess
}

#region Load needed PowerShell modules
# Ensure Azure PowerShell is loaded
$loadAzurePreview = $false # false to use 2.x stable, true to use the preview versions of cmdlets
if (-not (Get-Module -ListAvailable -Name AzureAD))
{
    # Maybe the preview AzureAD PowerShell is installed?
    if (-not (Get-Module -ListAvailable -Name AzureADPreview))
    {
        install-module azuread
    }
    else 
    {
        $loadAzurePreview = $true
    }
}

if ($loadAzurePreview)
{
    Import-Module AzureADPreview
}
else 
{
    Import-Module AzureAD   
}
#endregion

Write-Host "Please provide the credential to access the AAD tenant used by your Office 365 target tenant" -ForegroundColor Yellow

# Connect to Azure AD using modern auth
Connect-AzureAd

# Infrastructural variables
$identifierURI = "https://sharepointpnp.com/Modernization/Framework"
$passwordCredentials = GeneratePasswordCredential
$requiredResourceAccess = GenerateRequiredResourceAccess

# Check if the application already exists
$application = Get-AzureADApplication -Filter "identifierUris/any(uri:uri eq '$identifierURI')"
if ($application) {
    Write-Warning 'Application already registered in the target tenant! You should delete it first.'
    return $null
}        
else {
    $application = New-AzureADApplication `
        -DisplayName $AppTitle `
        -AvailableToOtherTenants $false `
        -Homepage $HomePageUrl `
        -ReplyUrls $ReplyUrl `
        -IdentifierUris $identifierURI `
        -RequiredResourceAccess $requiredResourceAccess `
        -PasswordCredentials $passwordCredentials

    # Creating the Service Principal for the application
    $servicePrincipal = New-AzureADServicePrincipal -AppId $application.AppId

    # Grant permissions
    # Permission grant will need to be done manually

    return @{ Application = $application; ClientId = $application.AppId; ClientSecret = $passwordCredentials.Value; }
}
