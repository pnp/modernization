# Manual setup of the Page Transformation UI

If you want to manually configure the Page Transformation UI, if you want to get deep understanding of the solution is built then reading the manual setup guide will help you. This setup guide consists out of a mandatory part and an optional part for multi-tenant service usage.

## Mandatory steps in setting up the Page Transformation UI

Since the Page Transformation UI solution depends on Azure and SharePoint the setup is split accordingly. Follow below 2 guides to get up and running:

1. [Configure Azure (Azure AD application and Azure Function app)](modernization-ui-manual-azure.md)
2. [Configure SharePoint (solution deployment, modernization center site and enabling page transformation for your sites)](modernization-ui-manual-sharepoint.md)

## Optional steps

You can also configure the Azure Function app created in previous chapter as a multi-tenant function which will than allow you to share this Azure function app with multiple tenants needing page transformation.

- [Configure your Azure Function app to be used from multiple tenants](modernization-ui-multitenant-setup.md)
- [Configure your SharePoint tenant to consume a multi-tenant page transformation service](modernization-ui-multitenant-use.md)
