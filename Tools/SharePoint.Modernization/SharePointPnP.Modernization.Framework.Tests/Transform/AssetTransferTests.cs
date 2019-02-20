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
    public class AssetTransferTests
    {
        
        /// <summary>
        /// This test validates with SharePoint the entire operation
        /// </summary>
        [TestMethod]
        public void AssetTransfer_TransformTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext())
                {
                    var pageTransformator = new PageTransformator(sourceClientContext, targetClientContext);

                    var pages = sourceClientContext.Web.GetPages("p");

                    foreach (var page in pages)
                    {
                        PageTransformationInformation pti = new PageTransformationInformation(page)
                        {
                            // If target page exists, then overwrite it
                            Overwrite = true,

                            // Don't log test runs
                            SkipTelemetry = true,

                            IncludeReferencedAssets = true
                        };

                        pageTransformator.Transform(pti);
                    }
                }
            }
        }
    }
}
