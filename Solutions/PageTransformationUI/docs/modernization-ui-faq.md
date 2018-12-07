
# Frequently Asked Questions

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

No, in the current preview version this is not possible.

## Can I make the transformation go faster

Creating a modern version of a page is complex process that involves doing a detailed analyzes of the classic + crafting of a new page, meaning it will need some time. Following factors however do speed up things:

- Setup your Azure Function to use a regular App Service plan instead of a consumption based plan: the regular plan keeps your service "alive" and does not incur the cost of activating the service when there's a demand
- Setup the service in the same Azure location as where your tenant is
