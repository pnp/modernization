Multi-tenant admin consent url: https://login.microsoftonline.com/common/oauth2/authorize?client_id=79ad0500-1230-4f7a-a5bb-5e83ce9174f4&response_type=code&prompt=admin_consent
Sample URLs to test this demo function: 
https://sharepointpnpmodernizationdemo.azurewebsites.net/api/Demo?SiteUrl=https%3A%2F%2Fbertonline.sharepoint.com%2Fsites%2Fpermissive
https://sharepointpnpmodernizationdemo.azurewebsites.net/api/Demo?SiteUrl=https%3A%2F%2Fbertonline.sharepoint.com%2Fsites%2Fpermissive&PageUrl=https%3A%2F%2Fbertonline.sharepoint.com%2Fsites%2Fpermissive%2FSitePages%2FClassicPage.aspx
https://sharepointpnpmodernizationdemo.azurewebsites.net/api/Demo?SiteUrl=https%3A%2F%2Fofficedevpnp.sharepoint.com%2Fsites%2Fmodernizeme5&PageUrl=https%3A%2F%2Fofficedevpnp.sharepoint.com%2Fsites%2Fmodernizeme5%2FSitePages%2Fgiro2018.aspx

Background:
Multi-tenant Azure Function: https://stackoverflow.com/questions/44720343/connect-azure-function-to-microsoft-graph-as-a-multi-tenant-application-using-az
App Service authentication overview: https://docs.microsoft.com/en-us/azure/app-service/app-service-authentication-overview


Host.json details: https://github.com/Azure/azure-functions-host/wiki/host.json

