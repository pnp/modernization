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
        public void AssetTransfer_CopyAssetToTargetLocation_SmallFileTest()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext())
                {
                    AssetTransfer assetTransfer = new AssetTransfer(sourceClientContext, targetClientContext);

                    // Very crude test - ensure the site is setup for this ahead of the test
                    var sourceFileServerRelativeUrl = "/sites/PnPTransformationSource/SiteCollectionImages/extra8_500x500.jpg";
                    var targetLocation = "/sites/PnPTransformationTarget/Shared%20Documents"; //Shared Documents for example, Site Assets may not exist on vanilla sites

                    assetTransfer.CopyAssetToTargetLocation(sourceFileServerRelativeUrl, targetLocation);
                }
            }
        }

        /// <summary>
        /// This test validates with SharePoint the entire operation
        /// </summary>
        [TestMethod]
        public void AssetTransfer_CopyAssetToTargetLocation_LargeFileTest()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext())
                {
                    AssetTransfer assetTransfer = new AssetTransfer(sourceClientContext, targetClientContext);

                    // Very crude test - ensure the site is setup for this ahead of the test
                    // Note this file is not included in this project assets due to its licensing. Pls find a > 3MB file to use as a test.
                    var sourceFileServerRelativeUrl = "/sites/PnPTransformationSource/SiteCollectionImages/bigstock-Html-Web-Code-57446159.jpg";
                    var targetLocation = "/sites/PnPTransformationTarget/Shared%20Documents"; //Shared Documents for example, Site Assets may not exist on vanilla sites

                    assetTransfer.CopyAssetToTargetLocation(sourceFileServerRelativeUrl, targetLocation);
                }
            }
        }
    }
}
