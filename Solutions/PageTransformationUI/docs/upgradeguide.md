# Upgrade guide

This guide shows how you can upgrade your Page Transformation UI deployment to the latest bits.

## Preparation

- Ensure you are using the developer build or December 2018 or later version of [PnP PowerShell](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps)
- Ensure you've pulled down the latest version from https://github.com/SharePoint/sp-dev-modernization
- You have to be a SharePoint Tenant Admin for this operation

## Step 1: Upgrade the Azure function app

Deploy the latest version of **sharepointpnpmodernizationeurope.zip** (available in the **provisioning**) folder to the Azure Function App.  Below snippet shows the PowerShell script does automate the deployment.

```Powershell
# First navigate to the provisioning folder and then call below script
.\Update-ModernizationFramework.ps1 -SubscriptionName "subscription" -ResourceGroupName "group" -FunctionAppName "functionname"
```

## Step 2: Upgrade the SharePoint side

Upgrade the SharePoint Framework solutions to the latest version by using below script:

```Powershell
# First navigate to the provisioning folder and then call below script

# Connect to the deployed modernization center site
# Update your tenant name
connect-pnponline -Url https://yourtenant.sharepoint.com/sites/ModernizationCenter

# Upload and publish the latest versions of the SPFX packages
Add-PnPApp -Path ..\assets\sharepointpnp-pagetransformation-central.sppkg -Scope Tenant -Publish -Overwrite
Add-PnPApp -Path ..\assets\sharepointpnp-pagetransformation-client.sppkg -Scope Tenant -Publish -Overwrite

# Upgrade the SPFX package used in the modernization center site:
Update-PnPApp -Identity 45cba470-b308-48a9-9e1d-9afde19a3f13
```
