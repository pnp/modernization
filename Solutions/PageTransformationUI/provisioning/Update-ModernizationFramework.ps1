<#
.SYNOPSIS
Deploys the latest version of the modernization function to the Azure function app

.EXAMPLE
PS C:\> .\Update-ModernizationFramework.ps1 -SubscriptionName "subscription" -ResourceGroupName "group" -FunctionAppName "functionname"
#>

param(
    [Parameter(Mandatory=$True)]
    [string] $SubscriptionName,
    [Parameter(Mandatory=$True)]
    [string] $ResourceGroupName,
    [Parameter(Mandatory=$True)]
    [string] $FunctionAppName
 )

 $FunctionAppName = $FunctionAppName.ToLower().Replace(" ", "").Replace("_", "").Replace("'","").Replace("-","").Replace("'","")
 if ($FunctionAppName.Length -gt 60)
 {
     $FunctionAppName = $FunctionAppName.Substring(0,60)
 }
 Write-Host ("Function app name that will be used: " + $FunctionAppName) -ForegroundColor White

$azureRMModule = Import-Module AzureRM -ErrorAction SilentlyContinue -PassThru
if(!$azureRMModule)
{
    Write-Output "Installing AzureRM module"
    Install-Module AzureRM -Force
}

# Login to AzureRM
Write-Host "Please provide the credential to access the Azure tenant where the Azure Function app was created" -ForegroundColor Yellow
Login-AzureRmAccount

# Select the target Subscription
$subs = Get-AzureRmSubscription
$sub = $subs | where { $_.Name -eq $SubscriptionName }
Select-AzureRmSubscription -TenantId $sub.TenantId -SubscriptionId $sub.SubscriptionId

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

