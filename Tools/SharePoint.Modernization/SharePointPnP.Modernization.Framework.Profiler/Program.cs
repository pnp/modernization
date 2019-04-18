using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Publishing;
using SharePointPnP.Modernization.Framework.Telemetry.Observers;
using SharePointPnP.Modernization.Framework.Transform;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Profiler
{
    class Program
    {
        static void Main(string[] args)
        {
            SharePointOnlineCredentials creds = new SharePointOnlineCredentials(AppSetting("Username"), ConvertToSecureString(AppSetting("Password")));

            using (var targetClientContext = new ClientContext(AppSetting("SPOTargetSiteUrl")))
            {
                targetClientContext.Credentials = creds;

                using (var sourceClientContext = new ClientContext(AppSetting("SPODevSiteUrl")))
                {
                    sourceClientContext.Credentials = creds;

                    //"C:\github\sp-dev-modernization\Tools\SharePoint.Modernization\SharePointPnP.Modernization.Framework.Tests\Transform\Publishing\custompagelayoutmapping.xml"
                    //"C:\temp\mappingtest.xml"
                    //var pageTransformator = new PublishingPageTransformator(sourceClientContext, targetClientContext , @"C:\github\sp-dev-modernization\Tools\SharePoint.Modernization\SharePointPnP.Modernization.Framework.Tests\Transform\Publishing\custompagelayoutmapping.xml");
                    var pageTransformator = new PublishingPageTransformator(sourceClientContext, targetClientContext, @"C:\temp\mappingtest.xml");
                    pageTransformator.RegisterObserver(new MarkdownObserver(folder: "c:\\temp"));
                    
                    var pages = sourceClientContext.Web.GetPagesFromList("Pages", pageNameStartsWith: "article");
                    //var pages = sourceClientContext.Web.GetPagesFromList("Pages", folder:"News");

                    foreach (var page in pages)
                    {
                        PublishingPageTransformationInformation pti = new PublishingPageTransformationInformation(page)
                        {
                            // If target page exists, then overwrite it
                            Overwrite = true,

                            // Don't log test runs
                            SkipTelemetry = true,

                            //RemoveEmptySectionsAndColumns = false,

                            // Configure the page header, empty value means ClientSidePageHeaderType.None
                            //PageHeader = new ClientSidePageHeader(cc, ClientSidePageHeaderType.None, null),

                            // Replace embedded images and iframes with a placeholder and add respective images and video web parts at the bottom of the page
                            // HandleWikiImagesAndVideos = false,

                            // Callout to your custom code to allow for title overriding
                            //PageTitleOverride = titleOverride,

                            // Callout to your custom layout handler
                            //LayoutTransformatorOverride = layoutOverride,

                            // Callout to your custom content transformator...in case you fully want replace the model
                            //ContentTransformatorOverride = contentOverride,
                        };

                        pti.MappingProperties["SummaryLinksToQuickLinks"] = "true";
                        pti.MappingProperties["UseCommunityScriptEditor"] = "true";

                        var result = pageTransformator.Transform(pti);
                        pageTransformator.FlushObservers();
                    }

                    pageTransformator.FlushObservers();
                }
            }

            Console.WriteLine("App Complete, press any key to end!");
            Console.ReadKey();
        }

        private static SecureString ConvertToSecureString(string password)
        {
            if (password == null)
                throw new ArgumentNullException("password");

            var securePassword = new SecureString();

            foreach (char c in password)
                securePassword.AppendChar(c);

            securePassword.MakeReadOnly();
            return securePassword;
        }

        public static string AppSetting(string key)
        {
#if !NETSTANDARD2_0
            return ConfigurationManager.AppSettings[key];
#else
            try
            {
                return configuration.AppSettings.Settings[key].Value;
            }
            catch
            {
                return null;
            }
#endif
        }
    }
}
