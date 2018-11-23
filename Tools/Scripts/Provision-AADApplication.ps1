Param (
    [Parameter(Mandatory=$true, Position=0)]
    [String] $AppName,
    [Parameter(Mandatory=$true, Position=1)]
    [String] $AppTitle    
)

Install-Module AzureADPreview -Force
Import-Module AzureADPreview

function GeneratePasswordCredential() {
    $keyId = (New-Guid).ToString();
    $startDate = Get-Date
    $endDate = $startDate.AddYears(2)
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
    $signInReadProfileDelegated.ResourceAccess.Add($siteCollectionsFullControlDelegatedPermission)
    
    $requiredResourcesAccess.Add($siteCollectionsFullControlDelegated)
    $requiredResourcesAccess.Add($signInReadProfileDelegated)

    return $requiredResourcesAccess
}

Write-Information "Please provide the credential to access the AAD tenant under the cover of your Office 365 target tenant"

# Connect to Azure AD
$credential = Get-Credential
Connect-AzureAd -Credential $credential

# Infrastructural variables
$homePage = "https://$AppName.azurewebsites.net/"
$identifierURI = "https://sharepointpnp.com/Modernization/Framework"
$logoutURI = "http://portal.office.com"
$passwordCredentials = GeneratePasswordCredential
$requiredResourceAccess = GenerateRequiredResourceAccess

# Check if the application already exists
$application = Get-AzureADApplication -Filter "identifierUris/any(uri:uri eq '$identifierURI')"
if ($application) {
    Write-Warning 'Application already registered in the target tenant! You can delete it and create it again.'
}        
else {
    $application = New-AzureADApplication `
        -DisplayName $AppTitle `
        -AvailableToOtherTenants $false `
        -Homepage $homePage `
        -ReplyUrls $homePage `
        -IdentifierUris $identifierURI `
        -LogoutUrl $logoutURI `
        -RequiredResourceAccess $requiredResourceAccess `
        -PasswordCredentials $passwordCredentials
}

return @{ Application = $application; ClientId = $application.AppId; ClientSecret = $passwordCredentials.Value }
