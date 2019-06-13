using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SharePointPnP.Modernization.Framework.Publishing;
using SharePointPnP.Modernization.Framework.Telemetry.Observers;
using System;

namespace SharePointPnP.Modernization.Framework.Tests.Transform.Publishing
{
    [TestClass]
    public class OnPremisesPublishingPageTests
    {
        [TestMethod]
        public void BasicOnPremPublishingPageTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateOnPremisesClientContext())
                {
                    //"C:\github\sp-dev-modernization\Tools\SharePoint.Modernization\SharePointPnP.Modernization.Framework.Tests\Transform\Publishing\custompagelayoutmapping.xml"
                    //"C:\temp\onprem-mapping-all-test.xml.xml"
                    var pageTransformator = new PublishingPageTransformator(sourceClientContext, targetClientContext , @"C:\temp\onprem-mapping-all-test.xml");
                    pageTransformator.RegisterObserver(new MarkdownObserver(folder: "c:\\temp", includeVerbose:true));
                    pageTransformator.RegisterObserver(new UnitTestLogObserver());

                    //Quick test if SP2010 has this API property.
                    sourceClientContext.ExecuteQuery();
                    Console.WriteLine(sourceClientContext.ServerLibraryVersion);

                    var pages = sourceClientContext.Web.GetPagesFromList("Pages", "Article-2010");
                    //var pages = sourceClientContext.Web.GetPagesFromList("Pages", folder:"News");

                    foreach (var page in pages)
                    {
                        PublishingPageTransformationInformation pti = new PublishingPageTransformationInformation(page)
                        {
                            // If target page exists, then overwrite it
                            Overwrite = true,

                            // Don't log test runs
                            SkipTelemetry = true,      

                            //Permissions are unlikely to work given cross domain
                            KeepPageSpecificPermissions = false,
                            
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
                    }

                    pageTransformator.FlushObservers();
                }
            }
        }

        [TestMethod]
        public void OnPremPageLayout_AnalyzeByPages_Test()
        {
            using (var context = TestCommon.CreateOnPremisesClientContext())
            {
                var pages = context.Web.GetPagesFromList("Pages");
                var analyzer = new PageLayoutAnalyser(context);
                int errorCount = 0;
                foreach (var page in pages)
                {
                    try
                    {
                        analyzer.AnalysePageLayoutFromPublishingPage(page);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error {0} {1}", ex.Message, ex.StackTrace);
                        errorCount++;
                    }
                }

                Console.WriteLine("Error Count {0}", errorCount);
                Assert.IsTrue((errorCount == 0));
                analyzer.GenerateMappingFile("c:\\temp", "onprem-mapping-test.xml");
            }
        }

        [TestMethod]
        public void OnPremPageLayout_AnalyseAll_Test()
        {
            using (var context = TestCommon.CreateOnPremisesClientContext())
            {
                
                var analyzer = new PageLayoutAnalyser(context);
                analyzer.AnalyseAll();                

                analyzer.GenerateMappingFile("c:\\temp", "onprem-mapping-all-test.xml");
            }
        }


    }
}
