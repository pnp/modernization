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

# Install AzureADPreview if it is missing
$aadPreviwModule = Import-Module AzureADPreview -ErrorAction SilentlyContinue -PassThru
if(!$aadPreviwModule)
{
    Write-Output "Installing AzureADPreview module"
    Install-Module AzureADPreview -Force
}

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
    $signInReadProfileDelegated.ResourceAccess.Add($siteCollectionsFullControlDelegatedPermission)
    
    $requiredResourcesAccess.Add($siteCollectionsFullControlDelegated)
    $requiredResourcesAccess.Add($signInReadProfileDelegated)

    return $requiredResourcesAccess
}

Function GrantOAuthPermission(
    $clientId,
    $clientSecret,
    $permissionResourceId,
    $permissionScope) {

    $resource = "https://graph.windows.net/"
    $authority = "https://login.microsoftonline.com/common"
    $tokenEndpointUri = "$authority/oauth2/token"
    $client_secret = [System.Web.HttpUtility]::UrlEncode($clientSecret)
    $content = "grant_type=client_credentials&client_id=$clientId&client_secret=$client_Secret&resource=$resource"
    
    Write-Host $tokenEndpointUri
    Write-Host $content

    $Stoploop = $false
    [int]$Retrycount = "0"
    
    do {
        try {
            $response = Invoke-RestMethod -Uri $tokenEndpointUri -Body $content -Method Post -UseBasicParsing
            Write-Host "Retrieved Access Token for Azure AD Graph API" -ForegroundColor Yellow

            # Assign access token
            $accessToken = $response.access_token
    
            $headers = @{
                Authorization = "Bearer $accessToken"
            }
    
            $consentBody = @{
                clientId    = $clientId
                consentType = "AllPrincipals"
                startTime   = ((get-date).AddDays(-1)).ToString("yyyy-MM-dd")
                principalId = $null
                resourceId  = $permissionResourceId
                scope       = $permissionScope
                expiryTime  = ((get-date).AddYears(99)).ToString("yyyy-MM-dd")
            }
    
            $consentBody = $consentBody | ConvertTo-Json
    
            Write-Host "Granting permission scope $permissionScope for resource $permissionResourceId" -ForegroundColor White
            $body = Invoke-RestMethod -Uri "https://graph.windows.net/myorganization/oauth2PermissionGrants?api-version=1.6" -Body $consentBody -Method POST -Headers $headers -ContentType "application/json"
            Write-Host "Granted permission scope $permissionScope for resource $permissionResourceId" -ForegroundColor Green
    
            $Stoploop = $true
        }
        catch {
            if ($Retrycount -gt 5) {
                Write-Host "Could not get create OAuth2PermissionGrants after 6 retries." -ForegroundColor Red
                $Stoploop = $true
            }
            else {
                Write-Host "Could not create OAuth2PermissionGrants yet. Retrying in 5 seconds..." -ForegroundColor DarkYellow
                Start-Sleep -Seconds 5
                $Retrycount ++
            }
        }
    }
    While ($Stoploop -eq $false)
}

Write-Host "Please provide the credential to access the AAD tenant under the cover of your Office 365 target tenant" -ForegroundColor White

# Connect to Azure AD
$credential = Get-Credential
Connect-AzureAd -Credential $credential

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

    # Wait 10 seconds
    # Start-Sleep -Seconds 10

    # Grant permissions
    # foreach ($permission in $requiredResourceAccess)
    # {
    #     GrantOAuthPermission `
    #         -clientId $application.AppId `
    #         -clientSecret $passwordCredentials.Value `
    #         -permissionResourceId $permission.ResourceAccess[0].Id `
    #         -permissionScope $permission.ResourceAccess[0].Type        
    # }

    return @{ Application = $application; ClientId = $application.AppId; ClientSecret = $passwordCredentials.Value }
}
