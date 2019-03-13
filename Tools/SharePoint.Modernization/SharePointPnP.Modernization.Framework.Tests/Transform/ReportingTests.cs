using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SharePointPnP.Modernization.Framework.Telemetry.Observers;
using SharePointPnP.Modernization.Framework.Transform;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Tests.Transform
{
    [TestClass]
    public class ReportingTests
    {

        [TestMethod]
        public void Reporting_SameSite_WebPartPageTest()
        {
            using (var sourceClientContext = TestCommon.CreateClientContext())
            {
                var pageTransformator = new PageTransformator(sourceClientContext);
                pageTransformator.RegisterObserver(new UnitTestLogObserver()); //Registers the unit test observer to log details for testing

                var pages = sourceClientContext.Web.GetPages("wpp");

                foreach (var page in pages)
                {
                    PageTransformationInformation pti = new PageTransformationInformation(page)
                    {
                        // If target page exists, then overwrite it
                        Overwrite = true,

                        // Don't log test runs
                        SkipTelemetry = true,

                        // ModernizationCenter options
                        ModernizationCenterInformation = new ModernizationCenterInformation()
                        {
                            AddPageAcceptBanner = true
                        },

                        // Give the migrated page a specific prefix, default is Migrated_
                        TargetPagePrefix = "Converted_",
    
                        // Replace embedded images and iframes with a placeholder and add respective images and video web parts at the bottom of the page
                        HandleWikiImagesAndVideos = false,

                    };

                    pageTransformator.Transform(pti);
                    pageTransformator.FlushObservers();
                }

            }

            Assert.Inconclusive(TestCommon.InconclusiveNoAutomatedChecksMessage);

        }

        [TestMethod]
        public void Reporting_SameSite_WebPartPage_SingleObserverTest()
        {
            using (var sourceClientContext = TestCommon.CreateClientContext())
            {
                var pageTransformator = new PageTransformator(sourceClientContext);
                // This test will use only the default observer

                var pages = sourceClientContext.Web.GetPages("wpp").Take(1);

                foreach (var page in pages)
                {
                    PageTransformationInformation pti = new PageTransformationInformation(page)
                    {
                        // If target page exists, then overwrite it
                        Overwrite = true,

                        // Don't log test runs
                        SkipTelemetry = true,

                        // ModernizationCenter options
                        ModernizationCenterInformation = new ModernizationCenterInformation()
                        {
                            AddPageAcceptBanner = true
                        },

                        // Give the migrated page a specific prefix, default is Migrated_
                        TargetPagePrefix = "Converted_",

                        // Replace embedded images and iframes with a placeholder and add respective images and video web parts at the bottom of the page
                        HandleWikiImagesAndVideos = false,

                    };

                    pageTransformator.Transform(pti);
                    pageTransformator.FlushObservers();
                }

            }

            Assert.Inconclusive(TestCommon.InconclusiveNoAutomatedChecksMessage);

        }

        [TestMethod]
        public void Reporting_CrossSite_WebPartPage_SingleObserverTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext())
                {
                    var pageTransformator = new PageTransformator(sourceClientContext, targetClientContext);
                    pageTransformator.RegisterObserver(new MarkdownObserver());

                    var pages = sourceClientContext.Web.GetPages("wpp").Take(1);

                    foreach (var page in pages)
                    {
                        PageTransformationInformation pti = new PageTransformationInformation(page)
                        {
                            // If target page exists, then overwrite it
                            Overwrite = true,

                            // Don't log test runs
                            SkipTelemetry = true,

                            // Replace embedded images and iframes with a placeholder and add respective images and video web parts at the bottom of the page
                            HandleWikiImagesAndVideos = false,

                            //Include the assets within the file transfer
                            CopyReferencedAssets = true

                        };

                        pageTransformator.Transform(pti);
                        pageTransformator.FlushSpecificObserver<MarkdownObserver>();
                    }

                }
            }

            Assert.Inconclusive(TestCommon.InconclusiveNoAutomatedChecksMessage);

        }
    }
}
