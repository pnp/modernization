using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SharePointPnP.Modernization.Framework.Cache;
using SharePointPnP.Modernization.Framework.Publishing;
using SharePointPnP.Modernization.Framework.Telemetry.Observers;

namespace SharePointPnP.Modernization.Framework.Tests.Transform.Replay
{
    [TestClass]
    public class ReplayPublishingPageTests
    {

        [TestMethod]
        public void InitialPublishingPage()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevSiteUrl")))
                {
                    var pageTransformator = new PublishingPageTransformator(sourceClientContext, targetClientContext);
                    //pageTransformator.RegisterObserver(new MarkdownObserver(folder: "c:\\temp", includeVerbose: true));
                    pageTransformator.RegisterObserver(new UnitTestLogObserver());

                    var pages = sourceClientContext.Web.GetPagesFromList("Pages", "Hot-Off-The-Press-New-Chilling-Truth-About-Sauce", "News");

                    foreach (var page in pages)
                    {
                        PublishingPageTransformationInformation pti = new PublishingPageTransformationInformation(page)
                        {
                            // If target page exists, then overwrite it
                            Overwrite = true,

                            // Don't log test runs
                            SkipTelemetry = true,

                            KeepPageCreationModificationInformation = false,

                            // Callout to your custom layout handler
                            //LayoutTransformatorOverride = layoutOverride,

                            // Callout to your custom content transformator...in case you fully want replace the model
                            //ContentTransformatorOverride = contentOverride,
                            
                            // Initially cache the page out webpart layout
                            IsReplayCapture = true

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
        public void ReplayPublishingPageTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevSiteUrl")))
                {
                    var pageTransformator = new PublishingPageTransformator(sourceClientContext, targetClientContext);
                    //pageTransformator.RegisterObserver(new MarkdownObserver(folder: "c:\\temp", includeVerbose:true));
                    pageTransformator.RegisterObserver(new UnitTestLogObserver());
                    
                    var pages = sourceClientContext.Web.GetPagesFromList("Pages", "Hot-Off-The-Press-New-Chilling-Truth-About-Sauce", "News");

                    foreach (var page in pages)
                    {
                        PublishingPageTransformationInformation pti = new PublishingPageTransformationInformation(page)
                        {
                            // If target page exists, then overwrite it
                            Overwrite = true,

                            // Don't log test runs
                            SkipTelemetry = true,  
                            
                            KeepPageCreationModificationInformation = false,
                            
                            TargetPageName = "New-Layout-Hot-Off-The-Press.aspx",

                            // Callout to your custom layout handler
                            //LayoutTransformatorOverride = layoutOverride,

                            // Callout to your custom content transformator...in case you fully want replace the model
                            //ContentTransformatorOverride = contentOverride,

                            ReplayLayoutChangeBasedOn = "Hot-Off-The-Press-New-Chilling-Truth-About-Sauce.aspx"

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
        public void ReplayPageLayoutAnalyzeByPage()
        {
            using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevSiteUrl")))
            {
                var pages = sourceClientContext.Web.GetPagesFromList("Pages", "Hot-Off-The-Press-New-Chilling-Truth-About-Sauce", "News");
                var analyzer = new PageLayoutAnalyser(sourceClientContext);

                foreach (var page in pages)
                {
                    analyzer.AnalysePageLayoutFromPublishingPage(page);
                }

                analyzer.GenerateMappingFile("c:\\temp", "replay-mapping-example.xml");
            }
        }



    }
}
