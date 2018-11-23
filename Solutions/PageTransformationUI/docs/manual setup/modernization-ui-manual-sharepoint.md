# Manual SharePoint setup

Page transformation depends on a SharePoint modernization center site being present and configured, a number of web parts being deployed and finally some user custom actions being installed in the site collections that want to use page transformation.

## Step 1: Deploy the SharePoint Framework solutions

>Important:
>You'll need the December 2018 or more recent version of PnP PowerShell

- Clone this repository to your local machine (git clone https://github.com/SharePoint/sp-dev-modernization.git) and checkout the dev branch (git checkout dev). Alternatively download the repository as zip file and unzip in local folder
- Load up [PnP PowerShell](http://aka.ms/sppnp-powershell) and navigate to the `/Solutions/PageTransformationUI/provisioning` folder
- Connect to your tenant root site collection via `Connect-PnPOnline -Url https://contoso.sharepoint.com` using a tenant admin account
- Apply the Modernization Center tenant template.  Doing this will create the needed communication site and will configure it. **Please provide your Azure App ID and Azure Function URL** so that the solution will use your Azure Function. How to setup the Azure part was discussed in the [Azure setup manual](modernization-ui-manual-azure.md).

```PowerShell
# Update AzureAppID and AzureFunction before running this
Apply-PnPTenantTemplate -Path .\modernization.xml -Parameters @{"AzureAppID"="79ad0500-1230-4f7a-a5bb-5e83ce9174f4", "AzureFunction"="https://contosomodernization.azurewebsites.net"}
```

## Step 2: Making page transformation available for a site collection

Final step is ensuring that the Page Transformation UI elements appear. This is done per site collection by installing a template.

- Load up [PnP PowerShell](http://aka.ms/sppnp-powershell) and navigate to the `/Solutions/PageTransformationUI/assets` folder you downloaded in step 1
- Connect to the site collection that wants to use page transformation via `Connect-PnPOnline -Url https://contoso.sharepoint.com/sites/sitethatwantspagetransformation`
- Apply a template to this site: `Apply-PnPProvisioningTemplate -Path .\clienttemplate.xml`
