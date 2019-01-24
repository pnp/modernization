<#
.SYNOPSIS
Creates the Azure AD application and Azure Function app which are needed for the Page Transformation UI solution.

.EXAMPLE
PS C:\> .\Provision-ModernizationFramework.ps1 -SubscriptionName "MySubscription" -ResourceGroupName "pnpmodernizationtest1" -ResourceGroupLocation "West Europe" -StorageAccountName "pnpmodernizationtest1" -FunctionAppName "pnpmodernizationtest1" -AppName "SharePointPnP.Modernization" -AppTitle "SharePointPnP.Modernization"
#>

param(
 [Parameter(Mandatory=$True)]
 [string]
 $SubscriptionName,
 [Parameter(Mandatory=$True)]
 [string]
 $ResourceGroupName,
 [Parameter(Mandatory=$True)]
 [string]
 $ResourceGroupLocation,
 [Parameter(Mandatory=$True)]
 [string]
 $StorageAccountName,
 [Parameter(Mandatory=$True)]
 [string]
 $FunctionAppName,
 [Parameter(Mandatory=$True)]
 [string]
 $AppName,
 [Parameter(Mandatory=$True)]
 [string]
 $AppTitle,
 [Parameter(Mandatory=$false)]
 [string]
 $AllowedTenants
)

$aadAppHomePageUrl = "https://$FunctionAppName.azurewebsites.net"
$aadAppReplyUrl = "$aadAppHomePageUrl/.auth/login/aad/callback"

Write-Host "Validating the Azure storage account and Function App availability" -ForegroundColor White

# Install AzureRM command lets, if they are missing
$azureRMModule = Import-Module AzureRM -ErrorAction SilentlyContinue -PassThru
if(!$azureRMModule)
{
    Write-Output "Installing AzureRM module"
    Install-Module AzureRM -Force
}

# Fixing invalid characters in provided names
$StorageAccountName = $StorageAccountName.ToLower().Replace(" ", "").Replace("_", "").Replace("'","").Replace("-","").Replace("'","")
if ($StorageAccountName.Length -gt 24)
{
    $StorageAccountName = $storageAccountName.Substring(0,24)
}
Write-Host ("Storage account that will be used: " + $StorageAccountName) -ForegroundColor White

$FunctionAppName = $FunctionAppName.ToLower().Replace(" ", "").Replace("_", "").Replace("'","").Replace("-","").Replace("'","")
if ($FunctionAppName.Length -gt 60)
{
    $FunctionAppName = $FunctionAppName.Substring(0,60)
}
Write-Host ("Function app name that will be used: " + $FunctionAppName) -ForegroundColor White


# Login to AzureRM
Write-Host "Please provide the credential to access the Azure tenant where the Azure Function app needs to be created" -ForegroundColor Yellow
Login-AzureRmAccount 

# Check if the storage account name was taken or not
if ((Get-AzureRmStorageAccountNameAvailability -Name $StorageAccountName).NameAvailable -eq $false)
{
    Write-Host ("Storage account " + $StorageAccountName + " is already used across Azure. Please pick a unique name for storage account") -ForegroundColor Red
    return 1
}

# Check if the function app name was taken or not via a public DNS lookup
$dnsName = $FunctionAppName + ".azurewebsites.net"
$resolve = Resolve-DnsName $dnsName -ErrorAction Ignore

if ($null -ne $resolve)
{
    Write-Host ("Function app name " + $FunctionAppName + " is already used across Azure. Please pick a unique name for the function app name") -ForegroundColor Red
    return 1
}

Write-Host "Creating the AAD application in the target Office 365 Tenant" -ForegroundColor White

# Register the Azure AD Application and get back ClientId and ClientSecret
$aadApp = .\Provision-AADApplication.ps1 -AppName $AppName -HomePageUrl $aadAppHomePageUrl -ReplyUrl $aadAppReplyUrl -AppTitle $AppTitle

if (($null -eq $aadApp.ClientId) -or ($aadApp.ClientId -eq ""))
{
    Write-Host "Failed to create the AAD application! Stopping the further execution of this script." -ForegroundColor Red
}
else
{
    $appIDToConsent = $aadApp.ClientId
    Write-Host "Created the AAD application in the target Office 365 Tenant" -ForegroundColor Green
    Write-Host ("Client ID: " + $aadApp.ClientId) -ForegroundColor White
    Write-Host ("Client Secret: " + $aadApp.ClientSecret) -ForegroundColor White

    # Select the target Subscription
    $subs = Get-AzureRmSubscription
    $sub = $subs | where { $_.Name -eq $SubscriptionName }
    Select-AzureRmSubscription -TenantId $sub.TenantId -SubscriptionId $sub.SubscriptionId

    # Create the Resource Group if it does not exist
    $resourceGroup = Get-AzureRmResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue
    if(!$resourceGroup)
    {
        Write-Host "Resource group '$ResourceGroupName' does not exist." -ForegroundColor DarkYellow
        if(!$ResourceGroupLocation) {
            Write-Host "To create a new resource group, please enter a location." -ForegroundColor DarkYellow
            $ResourceGroupLocation = Read-Host "ResourceGroupLocation"
        }
        Write-Host "Creating resource group '$ResourceGroupName' in location '$ResourceGroupLocation'" -ForegroundColor White
        New-AzureRmResourceGroup -Name $ResourceGroupName -Location $ResourceGroupLocation
        Write-Host "Created resource group '$ResourceGroupName' in location '$ResourceGroupLocation'" -ForegroundColor Green
    }
    else{
        Write-Host "Using existing resource group '$ResourceGroupName'" -ForegroundColor White
    }

    # Create the Storage Account if it does not exist
    $storageAccount = Get-AzureRmStorageAccount -Name $StorageAccountName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
    if(!$storageAccount)
    {
        Write-Host "Storage account '$StorageAccountName' does not exist." -ForegroundColor DarkYellow
        if(!$ResourceGroupLocation) {
            Write-Host "To create a new storage account, please enter a location." -ForegroundColor DarkYellow
            $ResourceGroupLocation = Read-Host "ResourceGroupLocation"
        }
        Write-Host "Creating storage account '$StorageAccountName' in location '$ResourceGroupLocation'" -ForegroundColor White
        New-AzureRmStorageAccount -ResourceGroupName $ResourceGroupName -Location $ResourceGroupLocation -Name $StorageAccountName -SkuName Standard_LRS
        Write-Host "Created storage account '$StorageAccountName' in location '$ResourceGroupLocation'" -ForegroundColor Green
    }
    else{
        Write-Host "Using existing storage account '$StorageAccountName'" -ForegroundColor White
    }

    # Define the Azure Storage connection string
    $storageKey = (Get-AzureRmStorageAccountKey -ResourceGroupName $ResourceGroupName -Name $StorageAccountName)[0].Value
    $storageConnectionString = "DefaultEndpointsProtocol=https;AccountName=$StorageAccountName;AccountKey=$storageKey;EndpointSuffix=core.windows.net"

    Write-Host "Storage account connection string '$storageConnectionString'" -ForegroundColor White

    # Create the Function App if it does not exist
    $functionApp = Get-AzureRmResource | Where-Object { $_.ResourceName -eq $FunctionAppName -And $_.ResourceType -eq 'Microsoft.Web/Sites' }
    if(!$functionApp)
    {
        Write-Host "FunctionApp '$FunctionAppName' does not exist." -ForegroundColor White
        if (!$ResourceGroupLocation) {
            Write-Host "To create a new function app, please enter a location." -ForegroundColor DarkYellow
            $ResourceGroupLocation = Read-Host "ResourceGroupLocation"
        }
        Write-Host "Creating function app '$FunctionAppName' in location '$ResourceGroupLocation'" -ForegroundColor White
        $functionApp = New-AzureRmResource -ResourceType 'Microsoft.Web/Sites' -ResourceName $FunctionAppName -Kind 'functionapp' -Location $ResourceGroupLocation -ResourceGroupName $ResourceGroupName -Properties @{} -Force
        Write-Host "Created function app '$FunctionAppName' in location '$ResourceGroupLocation'" -ForegroundColor Green
    }

    # Wait 10s for the app to be ready
    Write-Host "Waiting for the function app '$FunctionAppName' to be ready ..." -ForegroundColor White
    Start-Sleep -Seconds 15

    # Configure the Function App Settings
    $appSettings = @{
        'AllowedTenants' = $AllowedTenants;
        'AzureWebJobsDashboard' = $storageConnectionString;
        'AzureWebJobsStorage' = $storageConnectionString;
        'CLIENT_ID' = $aadApp.ClientId;
        'CLIENT_SECRET' = $aadApp.ClientSecret;
        'FUNCTIONS_EXTENSION_VERSION' = "~1";
        'FUNCTIONS_WORKER_RUNTIME' = "dotnet";
        'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING' = $storageConnectionString;
        'WEBSITE_CONTENTSHARE' = $FunctionAppName + "0001";
        'WEBSITE_NODE_DEFAULT_VERSION' = "8.11.1";
        'WEBSITE_RUN_FROM_PACKAGE' = "1";
    }

    # Configure the appSettings
    Write-Host "Configuring appSettings for the function app" -ForegroundColor White
    Set-AzureRmWebApp -Name $FunctionAppName -ResourceGroupName $ResourceGroupName -AppSettings $appSettings
    Write-Host "Configured appSettings for the function app" -ForegroundColor Green

    Write-Host "Waiting for the function app '$FunctionAppName' to be ready ..." -ForegroundColor White
    Start-Sleep -Seconds 10

    # Upload the ZIP file of the function and trigger deployment
    Write-Host "Uploading the source package to the function app" -ForegroundColor White
    $publishingProfile = Invoke-AzureRmResourceAction -ResourceGroupName $ResourceGroupName -ResourceType 'Microsoft.Web/Sites/config' `
        -ResourceName "$FunctionAppName/publishingcredentials" -Action list -ApiVersion 2015-08-01 -Force
    $kuduAuthorizationHeader = ("Basic {0}" -f [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $publishingProfile.properties.publishingUserName, $publishingProfile.properties.publishingPassword))))
    $kuduZipDeployUrl = "https://$FunctionAppName.scm.azurewebsites.net/api/zipdeploy"
    $userAgent = "PnP-Modernization/1.0"
    Invoke-RestMethod -Uri $kuduZipDeployUrl -Headers @{Authorization=$kuduAuthorizationHeader} `
        -UserAgent $userAgent -Method POST  `
        -InFile .\sharepointpnpmodernizationeurope.zip `
        -ContentType "multipart/form-data"
    Write-Host "Uploaded the source package to the function app" -ForegroundColor Green

    # Configure Authentication/Authorization for the Function App
    Write-Host "Configuring Authentication settings for the function app" -ForegroundColor White
    $authResourceName = $FunctionAppName + "/authsettings"
    $auth = Invoke-AzureRmResourceAction -ResourceGroupName $ResourceGroupName -ResourceType 'Microsoft.Web/sites/config' -ResourceName $authResourceName -Action list -ApiVersion 2016-08-01 -Force
    $auth.properties.enabled = "True"
    $auth.properties.unauthenticatedClientAction = "RedirectToLoginPage"
    $auth.properties.tokenStoreEnabled = "True"
    $auth.properties.defaultProvider = "AzureActiveDirectory"
    $auth.properties.isAadAutoProvisioned = "False"
    $auth.properties.clientId = $aadApp.ClientId
    $auth.properties.clientSecret = $aadApp.ClientSecret

    New-AzureRmResource -PropertyObject $auth.properties -ResourceGroupName $ResourceGroupName `
        -ResourceType 'Microsoft.Web/sites/config' -ResourceName $authResourceName `
        -ApiVersion 2016-08-01 -Force

    Write-Host "Configured Authentication settings for the function app" -ForegroundColor Green

    # Configure CORS
    Write-Host "Configuring CORS for the function app" -ForegroundColor White
    $allowedOrigins = @()
    $allowedOrigins += "*"
    $functionAppPropertiesObject = @{cors = @{allowedOrigins= $allowedOrigins}}
    Set-AzureRmResource -PropertyObject $functionAppPropertiesObject -ResourceGroupName $ResourceGroupName `
        -ResourceType 'Microsoft.Web/sites/config' -ResourceName "$FunctionAppName/web" `
        -ApiVersion 2015-08-01 -Force
    Write-Host "Configured CORS for the function app" -ForegroundColor Green

    Write-Host "!***********************************************************************" -ForegroundColor White
    Write-Host "! Final manual step is admin consenting the created Azure AD application" -ForegroundColor Yellow
    Write-Host "! Open a browser session to https://login.microsoftonline.com/common/oauth2/authorize?client_id=$($appIDToConsent)&response_type=code&prompt=admin_consent" -ForegroundColor Yellow
    Write-Host "!***********************************************************************" -ForegroundColor White
    Write-Host ""
    Write-Host "Process completed!" -ForegroundColor Green
    Write-Host "The parameters to continue with the SharePoint installation part are the following" -ForegroundColor White
    Write-Host """AzureAppID""=""$($appIDToConsent)"";""AzureFunction""=""https://$($FunctionAppName).azurewebsites.net""" -ForegroundColor White
}
