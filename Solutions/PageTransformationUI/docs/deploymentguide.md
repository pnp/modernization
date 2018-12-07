# Deployment guide

## Preparation

- Ensure you are using the developer build or December 2018 or later version of [PnP PowerShell](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps)

## Step 1: Setup the Azure side

You need one Azure AD application and one Azure AD function setup, which you can do by following the [Azure Setup Guide](/Solutions/PageTransformationUI/docs/manual%20setup/modernization-ui-manual-azure.md).

## Step 2: Deploy the SharePoint side

You need to create and configure the Modernization center site collection:

- Navigate to the `provisioning` folder or [download the modernization .pnp file](https://github.com/SharePoint/sp-dev-modernization/blob/master/Solutions/PageTransformationUI/provisioning/modernization.pnp?raw=true)
- Run below PnP PowerShell. Update the parameters before running
  
```PowerShell
# Connect to any given site in your tenant
Connect-PnPOnline -Url https://contoso.sharepoint.com

# Update AzureAppID and AzureFunction before running this
Apply-PnPTenantTemplate -Path .\modernization.pnp -Parameters @{"AzureAppID"="79ad0500-1230-4f7a-a5bb-5e83ce9174f4";"AzureFunction"="https://contosomodernization.azurewebsites.net"}
```

## Step 3: Enable the page transformation UI for your site collections

There's two ways to do this. The easiest is going to your modernization center home page, enter the URL of your site collection and click on **Enable**:

![page transformator setup web part](/Solutions/PageTransformationUI/docs/images/enablepagetransformationwebpart.png)

Alternative approach is applying a PnP provisioning template to the site that you want to enable. This can be done like shown below:

- Load up [PnP PowerShell](http://aka.ms/sppnp-powershell) and navigate to the `/Solutions/PageTransformationUI/assets` folder
- Connect to the site collection that wants to use page transformation via `Connect-PnPOnline -Url https://contoso.sharepoint.com/sites/sitethatwantspagetransformation`
- Apply a template to this site: `Apply-PnPProvisioningTemplate -Path .\clienttemplate.xml`