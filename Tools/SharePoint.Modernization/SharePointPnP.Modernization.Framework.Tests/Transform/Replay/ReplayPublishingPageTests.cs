using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SharePointPnP.Modernization.Framework.Cache;
using SharePointPnP.Modernization.Framework.Entities;
using SharePointPnP.Modernization.Framework.Publishing;
using SharePointPnP.Modernization.Framework.Telemetry;
using SharePointPnP.Modernization.Framework.Telemetry.Observers;
using SharePointPnP.Modernization.Framework.Transform;
using System.Collections;
using System.Collections.Generic;

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

                    // TODO This test may require so tweaking of the cache to fake a read from the source page


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

                            IsReplayLayout = true

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


        [TestMethod]
        public void ScanTargetPageWebPartLayoutForChangesTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
               
                PublishingPageTransformationInformation pti = new PublishingPageTransformationInformation()
                {
                    // If target page exists, then overwrite it
                    Overwrite = true,

                    // Don't log test runs
                    SkipTelemetry = true,

                    KeepPageCreationModificationInformation = false,
                                        
                    // Initially cache the page out webpart layout
                    IsReplayCapture = true,

                    TargetPageFolder = "News",

                    TargetPageName = "Hot-Off-The-Press-New-Chilling-Truth-About-Sauce.aspx"

                };

                // Fake Cache Data
                var captureData = new ReplayPageCaptureData()
                {
                    PageLayoutName = "NewsPageLayout",
                    PageName = "Hot-Off-The-Press-New-Chilling-Truth-About-Sauce.aspx",
                    PageUrl = "News/Hot-Off-The-Press-New-Chilling-Truth-About-Sauce.aspx",
                    ReplayWebPartLocations =
                    {
                        new ReplayWebPartLocation() { Row = 0, Column = 0, Order = 0,
                            SourceWebPartId = System.Guid.Empty,
                            SourceWebPartType = "SharePointPnP.Modernization.WikiTextPart",
                            TargetWebPartInstanceId = System.Guid.Parse("{dd6cd4ff-1b62-4d3c-9d3e-5348a3fa5403}"),
                            TargetWebPartTypeId = "Text"
                        },
                        new ReplayWebPartLocation() {  Column = 0, Order= 1, Row= 0,
                            SourceWebPartId = System.Guid.Parse("{d4dfc251-980c-4ddf-9ca4-64838ffed864}"),
                            SourceWebPartType = "Microsoft.SharePoint.Publishing.WebControls.SummaryLinkWebPart, Microsoft.SharePoint.Publishing, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c",
                            TargetWebPartInstanceId = System.Guid.Parse("{234bf1bc-8f62-4586-9e75-24225007767e}"),
                            TargetWebPartTypeId = "c70391ea-0b10-4ee9-b2b4-006d3fcad0cd"
                        },
                    }
                };
                //CacheManager.Instance.SetReplayCaptureData(captureData);
                               
                // Observers
                var observers = new List<ILogObserver>() { new UnitTestLogObserver() };
                ReplayPageLayout rpl = new ReplayPageLayout(pti, targetClientContext, "NewsPageLayout", observers);
                var result = rpl.ScanTargetPageWebPartLayoutForChanges(captureData);

                Assert.IsNotNull(result);
                Assert.IsTrue(result != default);
            }
        }



    }
}
