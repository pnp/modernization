# Upgrade guide

This guide shows how you can upgrade your Page Transformation UI deployment to the latest bits (= Beta drop).

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

Upgrade the SharePoint Framework solutions and assets to the latest version by using below script:

```Powershell
# First navigate to the provisioning folder and then call below script
.\Upgrade-ToBeta.ps1 -ModernizationCenterUrl https://contoso.sharepoint.com/sites/modernizationcenter -AssetsFolder "..\assets"
```

## Step 3: Update the site collections making use of the Page Transformation UI solution

One of the new features in the beta version is showing a banner on the classic page, letting users know there's a modern version of this page available. This is done via a user custom action which needs to be added to those site collections.

### Option A: Use the existing tools to disable and re-enable page transformation for a site collection

Use the existing tools and again enable page transformation for the site collections that need it as explained in the [Deployment Guide](deploymentguide.md#step-3-enable-the-page-transformation-ui-for-your-site-collections).

### Option B: Add the new user custom action to the site collections

```Powershell
Connect-PnPOnline -Url https://contoso.sharepoint.com/sites/modernizeme
Add-PnPJavaScriptLink -Scope Site -Key "CA_PnP_Modernize_ClassicBanner" -Sequence 1000 -Url "/sites/modernizationcenter/SiteAssets/pnppagetransformationclassicbanner.js?rev=beta.1"
```
