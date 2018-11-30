# Using a .pnp file

When you use the .xml template you'll also need the assets folder present as the .xml version of the template refers to assets. It's often easier to use the provided .pnp template or package your own version of the .pnp file via:

```Powershell
$t = Read-PnPTenantTemplate modernization.xml
Save-PnPTenantTemplate -Template $t -Out modernization.pnp
```

Applying the SharePoint PnP template then works like this:

```Powershell
# Deploy SharePoint component
Apply-PnPTenantTemplate -Path .\modernization.pnp -Parameters @{"AzureAppID"="79ad0500-1230-4f7a-a5bb-5e83ce9174f4";"AzureFunction"="https://contosomodernization.azurewebsites.net"}
```