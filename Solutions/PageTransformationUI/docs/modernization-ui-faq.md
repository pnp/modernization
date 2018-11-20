
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