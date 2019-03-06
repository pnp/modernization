using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Transform;
using OfficeDevPnP.Core.Pages;
using SharePointPnP.Modernization.Framework.Pages;
using SharePointPnP.Modernization.Framework.Entities;
using System.Linq;

namespace SharePointPnP.Modernization.Framework.Tests.Transform
{
    [TestClass]
    public class LoggingTests
    {

        [TestMethod]
        public void Error_LoggingTest()
        {

            // Deliberate Error
            var pageTransformator = new PageTransformator(null);
            pageTransformator.RegisterObserver(new UnitTestLogObserver()); // Example of registering an observer, this can be anything really.

            PageTransformationInformation pti = new PageTransformationInformation(null);

            // Should capture a argument exception
            Assert.ThrowsException<ArgumentNullException>(() =>
            {
                pageTransformator.Transform(pti);
            });

        }

        [TestMethod]
        public void NormalOperation_LoggingTest()
        {

            using (var sourceClientContext = TestCommon.CreateClientContext())
            {
                var pageTransformator = new PageTransformator(sourceClientContext);
                pageTransformator.RegisterObserver(new UnitTestLogObserver()); // Example of registering an observer, this can be anything really.

                var pages = sourceClientContext.Web.GetPages("wk").Take(1);

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
                }
            }
        }
    }
}
