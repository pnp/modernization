# The SharePoint PnP Page Transformation UI solution

The Page Transformation UI solution makes it possible for end users to request a modern version of a wiki or web part page. The generated modern page will have a page banner web part on top of the page which will allow the user to keep the generated page or discard it. When the user discards the page the solution will show a feedback dialog asking for a reason why the page was not good.

Below diagram shows the high level architecture of the solution:

1. From any of the UI elements the users triggers the creation of a modern version of the selected wiki or web part page. This will be done by calling a "central" proxy page which is hosted in the modernization center site collection
2. The "central" proxy page contains an SPFX web part that makes a call to an Azure AD secured Azure Function
3. The Azure Function uses the [SharePoint Modernization Framework](https://www.nuget.org/packages/SharePointPnPModernizationOnline) to create a modern version of the page. This created modern version does contain a banner web part which provides the end user with the option to keep or discard the created page. Important to understand is that this modern page is a **new** page with name like migrated_oldpagename.aspx
4. If the page is discard a feedback dialog is shown asking the user for a reason why the page was not good. This information is then stored in a central list in the modernization center site collection. If the users keeps the page then the modern page gets the name of the original page and the original page is renamed with an old_ prefix

![page transformator web part](docs/images/PageTransformationUIarchitecture.png)

<img src="https://telemetry.sharepointpnp.com/sp-dev-modernization/solutions/PageTransformationUI" />