
# Frequently Asked Questions

## Help, the Page Transformation UI solution is not working in my tenant

Please consult the [trouble shooting guide](troubleshootingguide.md) to get unblocked.

## Can I change the URL in the page banner web part that points to more information?

Yes, this is possible. The used URL can be configured via a tenant storage entity. Start a PnP PowerShell session and use below code to check the current value and set to something else:

```PowerShell
Get-PnPStorageEntity -Key "Modernization_LearnMoreUrl"

# Set a custom URL for the "Learn more" link
Set-PnPStorageEntity -Key "Modernization_LearnMoreUrl" -Value "https://aka.ms/sppnp-modernize" -Description "Url shown in the learn more link"  
```

## Can I configure the page banner to not request feedback on when a page is discarded?

Yes, this is possible by removing the **Modernization_FeedbackList** storage entity or by setting it to an empty string.

```PowerShell
Get-PnPStorageEntity -Key "Modernization_FeedbackList"

# Setting the FeedbackList to blank will make the feedback dialog being skipped
Set-PnPStorageEntity -Key "Modernization_FeedbackList" -Value "" -Description "Name of the created feedback list"
```

## Can I configure additional feedback categories?

No, in the current version this is not possible.

## Can I make the transformation go faster

Creating a modern version of a page is complex process that involves doing a detailed analyzes of the classic + crafting of a new page, meaning it will need some time. Following factors however do speed up things:

- Ensure you're using the latest version of the transformation service. Check the [upgrade guide](upgradeguide.md) to learn how to do this.
- Setup your Azure Function to use a regular App Service plan instead of a consumption based plan: the regular plan keeps your service "alive" and does not incur the cost of activating the service when there's a demand
- Setup the service in the same Azure location as where your tenant is

## My script editor and content editor web parts are not transformed

As there's no 1st party modern web part that can replace the classic script editor and content editor web parts these web parts are not transformed. There however is an open source modern script editor web part that can be used, the steps below described how to update your environment and `webpartmapping.xml` file to make this work:

### Prepare your environment

Install the open source script editor web part (https://github.com/SharePoint/sp-dev-fx-webparts/tree/master/samples/react-script-editor) to your tenant.

### Instructions for creating an updated webpartmapping.xml file

- Grab a default version from https://github.com/SharePoint/sp-dev-modernization/blob/master/Tools/SharePoint.Modernization/SharePointPnP.Modernization.Framework/Nuget/webpartmapping.xml
- Then uncomment the mapping for script editor, content editor and html form web part (lines 193, 204, 521). The default `webpartmapping.xml` already contains the needed mapping for using the custom script editor, but since that web part is not installed by default the mapping is commented out

### Instructions for updating the webpartmapping.xml file for Transformation service usage

The easiest way to deploy the Azure Function app binaries is by using Kudu:

- Go Azure provisioning Function App
- Click on Platform features and select Advanced tools (Kudu) from the Development Tools section
- Drag the updated webpartmapping.xml file inside the `wwwroot` folder pane

## After enabling Page Transformation for a site collection the ribbon buttons are grayed out

The SharePoint ribbon uses a caching in SharePoint Online which causes this problem. You can either wait for the cache to expire or clean the browser cache and re-login again.

## My SharePoint environment uses vanity url's, do I need to do more?

By default only the first vanity domain is added the the "SharePoint Online Client Extensibility Web Application Principal" Azure AD application resulting in following URL's:

 - https://domain1.mycompany.com/_forms/spfxsinglesignon.aspx?redirect
 - https://domain1.mycompany.com/_forms/spfxsinglesignon.aspx
 - https://domain1.mycompany.com/_forms/singlesignon.aspx?redirect
 - https://domain1.mycompany.com/_forms/singlesignon.aspx
 - https://domain1.mycompany.com/

If you've multiple vanity URL's you'll need to add additional redirect URL's for those, simply copy the above 5 URLS's and update the domain1.mycompany.com to the other vanity URL.
