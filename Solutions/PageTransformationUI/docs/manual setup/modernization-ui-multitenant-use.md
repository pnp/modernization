
# Using a multi-tenant Azure Function app for the Page Transformation UI

If you've setup a multi-tenant Azure Function App or were allowed to access an existing one then follow these steps.

## Azure AD setup

- Ensure you've received the application ID of the Azure AD application that's used to protect the multi-tenant Azure Function app
- Craft a URL to admin consent this Azure AD application:

    - Use this URL, but replace the client_id guid with the application ID you've received: https://login.microsoftonline.com/common/oauth2/authorize?client_id=99ad0500-1230-4f7a-a5bb-5e83ce9374f4&response_type=code&prompt=admin_consent 
    - Login with your Azure AD admin credentials and consent the application

## SharePoint setup

- Use [PnP PowerShell](http://aka.ms/sppnp-powershell) to point your setup to the multi-tenant service

```PowerShell
Set-PnPStorageEntity -Key "Modernization_AzureADApp" -Value "99ad0500-1230-4f7a-a5bb-5e83ce9374f4" -Description "ID of the multi-tenant Azure AD app is used for page transformation"
Set-PnPStorageEntity -Key "Modernization_FunctionHost" -Value "https://multitenantmodernization.azurewebsites.net" -Description "Host of the multi-tenant SharePoint PnP Modernization service"
```