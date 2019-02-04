# Deployment guide

## Preparation

- Ensure you are using the developer build or December 2018 or later version of [PnP PowerShell](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps)
- Ensure you're using an Office 365 global admin account

## Step 1: Setup the Azure side

You need one Azure AD application and one Azure AD function setup, which you can do by following the manual steps in the [Azure Setup Guide](/Solutions/PageTransformationUI/docs/manual%20setup/modernization-ui-manual-azure.md) or alternatively you can use a scripted approach to create the needed Azure AD application and Azure Function App as shown below:

- Navigate to the `provisioning` folder
- Open a PowerShell session and run below PowerShell

>**Important:**
> - Update the **SubscriptionName**, **ResourceGroupName**, **ResourceGroupLocation**, **StorageAccountName** and **FunctionAppName** parameters before running.
> - Note that the **FunctionAppName** must not have been used by others: check this by doing an `nslookup <functionappname>.azurewebsites.net`, if the DNS name is found then the function name is already in use. Also the function app name must be lowercase and cannot contain spaces or underscores. See https://docs.microsoft.com/en-us/azure/architecture/best-practices/naming-conventions for details on the naming
> - Note that the storage account name must be between 3 and 24 characters in length, and can include numbers and lowercase letters only. Also the storage account name cannot be taken already: use `Get-AzureRmStorageAccountNameAvailability -Name "mystorageaccount"` to verify if the account name is free
> - **AppName** and **AppTitle** must be equal to **SharePointPnP.Modernization**

```PowerShell
.\Provision-ModernizationFramework.ps1 -SubscriptionName "MySubscription" `
                                       -ResourceGroupName "pnpmodernizationtest1" `
                                       -ResourceGroupLocation "West Europe" `
                                       -StorageAccountName "pnpmodernizationtest1" `
                                       -FunctionAppName "pnpmodernizationtest1" `
                                       -AppName "SharePointPnP.Modernization" `
                                       -AppTitle "SharePointPnP.Modernization"
```

Once the script finishes you see the following type of output:

```Text
Final manual step is admin consenting the created Azure AD application
Open a browser session to https://login.microsoftonline.com/common/oauth2/authorize?client_id=f0e040f0-21e3-4640-ba50-7b56be765b26&response_type=code&prompt=admin_consent
Process completed!
The parameters to continue with the SharePoint installation part are the following
"AzureAppID"="f0e040f0-21e3-4640-ba50-7b56be765b26";"AzureFunction"="https://pnpmodernizationtest1.azurewebsites.net"
```

>**Important:**
>You'll need to perform the admin consenting of the created Azure AD app via the provided URL. Doing so will prompt you to accept the apps permissions for all it's users. When the consent is done you're redirected to the app's redirect url which might show a `Bad Request` message. This message can be safely ignored.

Also note that the last line in the output contains the needed parameter definition to launch step 2, which is described in the next section.

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

>**Note:**
>If you want to host the modernization center site collection under a different URL then `/sites/modernizationcenter` then you can do this by specifying an extra parameter as show below.

```PowerShell
# Update CenterUrl, AzureAppID and AzureFunction before running this
Apply-PnPTenantTemplate -Path .\modernization.pnp -Parameters @{"CenterUrl"="/teams/modernizationcenter";"AzureAppID"="79ad0500-1230-4f7a-a5bb-5e83ce9174f4";"AzureFunction"="https://contosomodernization.azurewebsites.net"}
```

>**Note:**
>The above setup steps are also explained in [a video on the PnP YouTube channel](https://www.youtube.com/watch?v=DK8YMRRyPgw).

## Step 3: Enable the page transformation UI for your site collections

There are two ways to do this.

### Option A: Use the admin web part from the Modernization center site:

The easiest is going to your modernization center home page, enter the URL of your site collection and click on **Enable**:

![page transformator setup web part](/Solutions/PageTransformationUI/docs/images/enablepagetransformationwebpart.png)

### Option B: Use a PowerShell script to configure multiple sites

The script approach enables you to configure either a single site collection or a list of site collections. Load up [PnP PowerShell](http://aka.ms/sppnp-powershell) and navigate to the `/Solutions/PageTransformationUI/provisioning` folder. To configure a single site collection simply connect to the site and call the needed script:

```PowerShell
# Enable page transformation
Connect-PnPOnline -Url https://contoso.sharepoint.com/sites/sitetoconfigure
.\Enable-PageTransformation.ps1
```

>**Note:**
>If you want to host the modernization center site collection under a different URL then `/sites/modernizationcenter` then you can do this by specifying an extra parameter as show below.

```PowerShell
# Enable page transformation
Connect-PnPOnline -Url https://contoso.sharepoint.com/sites/sitetoconfigure
.\Enable-PageTransformation.ps1 -ModernizationCenterUrl "/teams/modernization"
```

```PowerShell
# Disable page transformation
Connect-PnPOnline -Url https://contoso.sharepoint.com/sites/sitetoconfigure
.\Enable-PageTransformation.ps1
```

If you want to configure multiple site collections at once you can specify a CSV file when running `ConfigurePageTransformation.ps1`. To build that CSV file you can use `BuildSiteCSV.ps1`.

>**Note:**
>Both approaches are interchangeable, meaning you can for example use the script approach to enable the page transformation UI integration and then use the admin web part to disable it for a site.

## Help, it's not working

Please consult the [trouble shooting guide](troubleshootingguide.md) to get unblocked.