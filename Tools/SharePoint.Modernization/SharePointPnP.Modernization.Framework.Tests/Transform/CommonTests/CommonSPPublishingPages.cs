using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SharePointPnP.Modernization.Framework.Publishing;
using SharePointPnP.Modernization.Framework.Telemetry.Observers;
using Microsoft.SharePoint.Client;
using static SharePointPnP.Modernization.Framework.Tests.TestCommon;

namespace SharePointPnP.Modernization.Framework.Tests.Transform.CommonTests
{
    /// <summary>
    /// Summary description for CommonSP_PublishingPages
    /// </summary>
    [TestClass]
    public class CommonSPPublishingPages
    {
        public CommonSPPublishingPages()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        [TestCategory(TestCategories.SP2010)]
        [TestMethod]
        public void AllCommonPages_SP2010()
        {
            TransformPage(SPPlatformVersion.SP2010);
        }

        [TestCategory(TestCategories.SP2013)]
        [TestMethod]
        public void AllCommonPages_SP2013()
        {
            TransformPage(SPPlatformVersion.SP2013);
        }

        [TestCategory(TestCategories.SP2016)]
        [TestMethod]
        public void AllCommonPages_SP2016()
        {
            TransformPage(SPPlatformVersion.SP2016);
        }

        [TestCategory(TestCategories.SP2019)]
        [TestMethod]
        public void AllCommonPages_SP2019()
        {
            TransformPage(SPPlatformVersion.SP2019);
        }

        // Common Tests
        private void TransformPage(SPPlatformVersion version, string fullPageLayoutMapping = "", string pageNameStartsWith = "Common")
        {

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateSPPlatformClientContext(version, TransformType.PublishingPage))
                {

                    var  pageTransformator = new PublishingPageTransformator(sourceClientContext, targetClientContext, fullPageLayoutMapping);
                    pageTransformator.RegisterObserver(new MarkdownObserver(folder: TestContext.ResultsDirectory, includeVerbose: true));
                    pageTransformator.RegisterObserver(new UnitTestLogObserver());

                    var pages = sourceClientContext.Web.GetPagesFromList("Pages", pageNameStartsWith);
                    pages.FailTestIfZero();

                    foreach (var page in pages)
                    {
                        // Not great on efficiency but need the name
                        var pageName = page.EnsureProperty(o => o.File.Name);
                                               
                        PublishingPageTransformationInformation pti = new PublishingPageTransformationInformation(page)
                        {
                            // If target page exists, then overwrite it
                            Overwrite = true,

                            // Don't log test runs
                            SkipTelemetry = true,

                            //Permissions are unlikely to work given cross domain
                            KeepPageSpecificPermissions = false,

                            //Update target to include SP version
                            TargetPageName = TestCommon.UpdatePageToIncludeVersion(version, pageName)
                        };

                        Console.WriteLine("SharePoint Version: {0}", pti.SourceVersion);

                        pti.MappingProperties["SummaryLinksToQuickLinks"] = "true";
                        pti.MappingProperties["UseCommunityScriptEditor"] = "true";

                        var result = pageTransformator.Transform(pti);
                    }

                    pageTransformator.FlushObservers();

                }
            }
        }

    }
}
