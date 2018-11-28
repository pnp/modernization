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
 [Parameter(Mandatory=$True)]
 [string]
 $AllowedTenants
)

# Register the Azure AD Application and get back ClientId and ClientSecret
$aadApp = .\Provision-AADApplication.ps1 -AppName $AppName -AppTitle $AppTitle

# Install AzureRM command lets, if they are missing
$azureRMModule = Import-Module AzureRM -ErrorAction SilentlyContinue -PassThru
if(!$azureRMModule)
{
    Write-Output "Installing AzureADPreview module"
    Install-Module AzureRM -Force
}

# Login to AzureRM
Login-AzureRmAccount 

# Select the target Subscription
$subs = Get-AzureRmSubscription
$sub = $subs | where { $_.Name -eq $SubscriptionName }
Select-AzureRmSubscription -TenantId $sub.TenantId -SubscriptionId $sub.SubscriptionId

# Create the Resource Group if it does not exist
$resourceGroup = Get-AzureRmResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue
if(!$resourceGroup)
{
    Write-Host "Resource group '$ResourceGroupName' does not exist."
    if(!$ResourceGroupLocation) {
        Write-Host "To create a new resource group, please enter a location.";
        $ResourceGroupLocation = Read-Host "ResourceGroupLocation";
    }
    Write-Host "Creating resource group '$ResourceGroupName' in location '$ResourceGroupLocation'";
    New-AzureRmResourceGroup -Name $ResourceGroupName -Location $ResourceGroupLocation
}
else{
    Write-Host "Using existing resource group '$ResourceGroupName'";
}

# Create the Storage Account if it does not exist
$storageAccount = Get-AzureRmStorageAccount -Name $StorageAccountName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
if(!$storageAccount)
{
    Write-Host "Storage account '$StorageAccountName' does not exist."
    if(!$ResourceGroupLocation) {
        Write-Host "To create a new storage account, please enter a location.";
        $ResourceGroupLocation = Read-Host "ResourceGroupLocation";
    }
    Write-Host "Creating storage account '$StorageAccountName' in location '$ResourceGroupLocation'";
    New-AzureRmStorageAccount -ResourceGroupName $ResourceGroupName -Location $ResourceGroupLocation -Name $StorageAccountName -SkuName Standard_LRS
}
else{
    Write-Host "Using existing storage account '$StorageAccountName'";
}

# Define the Azure Storage connection string
$storageKey = (Get-AzureRmStorageAccountKey -ResourceGroupName $ResourceGroupName -Name $StorageAccountName)[0].Value
$storageConnectionString = "DefaultEndpointsProtocol=https;AccountName=$StorageAccountName;AccountKey=$storageKey;EndpointSuffix=core.windows.net"

# Create the Function App if it does not exist
$functionApp = Get-AzureRmResource | Where-Object { $_.ResourceName -eq $FunctionAppName -And $_.ResourceType -eq 'Microsoft.Web/Sites' }
if(!$functionApp)
{
    Write-Host "FunctionApp '$FunctionAppName' does not exist."
    if (!$ResourceGroupLocation) {
        Write-Host "To create a new function app, please enter a location.";
        $ResourceGroupLocation = Read-Host "ResourceGroupLocation";
    }
    $functionApp = New-AzureRmResource -ResourceType 'Microsoft.Web/Sites' -ResourceName $FunctionAppName -Kind 'functionapp' -Location $ResourceGroupLocation -ResourceGroupName $ResourceGroupName -Properties @{} -Force
}

# Wait 10s for the app to be ready
Start-Sleep -Seconds 10

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

# Force NewtonSoft.Json installation
# $module = import-module newtonsoft.json -ErrorAction SilentlyContinue -PassThru
# if(!$module)
# {
#     Write-Output "installing newtonsoft.json"
#     Install-Package Newtonsoft.Json -Force
#     import-module newtonsoft.json
# }
Set-AzureRmWebApp -Name $FunctionAppName -ResourceGroupName $ResourceGroupName -AppSettings $appSettings

# Upload the ZIP file of the function and trigger deployment
$publishingProfile = Invoke-AzureRmResourceAction -ResourceGroupName $ResourceGroupName -ResourceType 'Microsoft.Web/Sites/config' `
    -ResourceName "$FunctionAppName/publishingcredentials" -Action list -ApiVersion 2015-08-01 -Force
$kuduAuthorizationHeader = ("Basic {0}" -f [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $publishingProfile.properties.publishingUserName, $publishingProfile.properties.publishingPassword))))
$kuduZipDeployUrl = "https://$FunctionAppName.scm.azurewebsites.net/api/zipdeploy"
$userAgent = "PnP-Modernization/1.0"
Invoke-RestMethod -Uri $kuduZipDeployUrl -Headers @{Authorization=$kuduAuthorizationHeader} `
    -UserAgent $userAgent -Method POST  `
    -InFile .\sharepointpnpmodernizationeurope.zip `
    -ContentType "multipart/form-data"

# Configure Authentication/Authorization for the Function App
$authResourceName = $FunctionAppName + "/authsettings"
$auth = Invoke-AzureRmResourceAction -ResourceGroupName $ResourceGroupName -ResourceType 'Microsoft.Web/sites/config' -ResourceName $authResourceName -Action list -ApiVersion 2016-08-01 -Force
$auth.properties.enabled = "True"
$auth.properties.unauthenticatedClientAction = "RedirectToLoginPage"
$auth.properties.tokenStoreEnabled = "True"
$auth.properties.defaultProvider = "AzureActiveDirectory"
$auth.properties.isAadAutoProvisioned = "False"
$auth.properties.clientId = $aadApp.ClientId
$auth.properties.clientSecret = $aadApp.ClientSecret
$auth.properties.issuer = "https://login.microsoftonline.com/common/"

New-AzureRmResource -PropertyObject $auth.properties -ResourceGroupName $ResourceGroupName `
    -ResourceType 'Microsoft.Web/sites/config' -ResourceName $authResourceName `
    -ApiVersion 2016-08-01 -Force

# Configure CORS
$allowedOrigins = @()
$allowedOrigins += "*"
$functionAppPropertiesObject = @{cors = @{allowedOrigins= $allowedOrigins}}
Set-AzureRmResource -PropertyObject $functionAppPropertiesObject -ResourceGroupName $ResourceGroupName `
    -ResourceType 'Microsoft.Web/sites/config' -ResourceName "$FunctionAppName/web" `
    -ApiVersion 2015-08-01 -Force
