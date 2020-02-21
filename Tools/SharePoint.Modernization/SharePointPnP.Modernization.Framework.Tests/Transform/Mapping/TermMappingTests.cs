using System;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SharePointPnP.Modernization.Framework.Publishing;
using SharePointPnP.Modernization.Framework.Telemetry.Observers;
using SharePointPnP.Modernization.Framework.Transform;
using SharePointPnP.Modernization.Framework.Utilities;

namespace SharePointPnP.Modernization.Framework.Tests.Transform.Mapping
{
    [TestClass]
    public class TermMappingTests
    {
        [TestMethod]
        public void TermMappingFileLoadTest()
        {
            FileManager fm = new FileManager();
            var mapping = fm.LoadTermMappingFile(@"..\..\Transform\Mapping\term_mapping_sample.csv");

            Assert.IsTrue(mapping.Count > 0);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void TermMappingFileNotFoundTest()
        {
            FileManager fm = new FileManager();
            var mapping = fm.LoadTermMappingFile(@"..\..\Transform\Mapping\idontexist_sample.csv");
        }

        [TestMethod]
        public void TermMappingTransformatorTest_PassThrough()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateOnPremisesClientContext())
                {
                    PublishingPageTransformationInformation pti = new PublishingPageTransformationInformation()
                    {
                        // If target page exists, then overwrite it
                        Overwrite = true,

                        // Don't log test runs
                        SkipTelemetry = true,

                        //Permissions are should work given cross domain with mapping
                        KeepPageSpecificPermissions = true,

                        // Term store mapping
                        TermMappingFile = string.Empty,
                        SkipTermStoreMapping = false
                    };

                    TermTransformator termTransformator = new TermTransformator(pti, sourceClientContext, targetClientContext, null);

                    var inputLabel = "pass-through-test";
                    var inputGuid = Guid.NewGuid();
                    var result = termTransformator.Transform(new Entities.TermData() { TermGuid = inputGuid, TermLabel = inputLabel });
                    Console.WriteLine(inputLabel + " and " + inputGuid);

                    Assert.AreEqual(inputLabel, result.TermLabel);
                    Assert.AreEqual(inputGuid, result.TermGuid);
                }
            }
        }

        [TestMethod]
        public void BasicOnlineWikiPage_TermMappingTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevTeamSiteUrl")))
                {
                    var pageTransformator = new PageTransformator(sourceClientContext, targetClientContext);
                    //pageTransformator.RegisterObserver(new MarkdownObserver(folder: "c:\\temp", includeVerbose: true));
                    pageTransformator.RegisterObserver(new UnitTestLogObserver());

                    var pages = sourceClientContext.Web.GetPagesFromList("Site Pages", pageNameStartsWith: "Common-WikiPageTest");

                    pages.FailTestIfZero();

                    foreach (var page in pages)
                    {
                        PageTransformationInformation pti = new PageTransformationInformation(page)
                        {
                            // If target page exists, then overwrite it
                            Overwrite = true,

                            // Don't log test runs
                            SkipTelemetry = true,

                            //Permissions are unlikely to work given cross domain
                            KeepPageSpecificPermissions = false,

                            // Term store mapping
                            TermMappingFile = @"..\..\Transform\Mapping\term_mapping_sample.csv",

                            //Should process default mapping
                            SkipTermStoreMapping = false,

                            CopyPageMetadata = true

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



        [TestMethod]
        public void BasicOnlinePublishingPage_TermDefaultTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevSiteUrl")))
                {
                    var pageTransformator = new PublishingPageTransformator(sourceClientContext, targetClientContext, @"C:\temp\spo-mapping-all-test.xml");
                    //pageTransformator.RegisterObserver(new MarkdownObserver(folder: "c:\\temp", includeVerbose: true));
                    pageTransformator.RegisterObserver(new UnitTestLogObserver());

                    var pages = sourceClientContext.Web.GetPagesFromList("Pages", pageNameStartsWith: "Article-PnP-Example");

                    pages.FailTestIfZero();

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

                            // Term store mapping
                            TermMappingFile = string.Empty,

                            //Should process default mapping
                            SkipTermStoreMapping = false

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

        [TestMethod]
        public void BasicOnlinePublishingPage_TermMappingTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevSiteUrl")))
                {
                    var pageTransformator = new PublishingPageTransformator(sourceClientContext, targetClientContext, @"C:\temp\onprem-mapping-all-test.xml");
                    //pageTransformator.RegisterObserver(new MarkdownObserver(folder: "c:\\temp", includeVerbose: true));
                    pageTransformator.RegisterObserver(new UnitTestLogObserver());


                    var pages = sourceClientContext.Web.GetPagesFromList("Pages", folder: "News", pageNameStartsWith: "Kitchen");

                    pages.FailTestIfZero();

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

                            // Term store mapping
                            TermMappingFile = @"..\..\Transform\Mapping\term_mapping_sample.csv",

                            SkipTermStoreMapping = true

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

        [TestMethod]
        public void BasicOnPremPublishingPage_TermMappingTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateOnPremisesClientContext())
                {
                    var pageTransformator = new PublishingPageTransformator(sourceClientContext, targetClientContext, @"C:\temp\onprem-mapping-all-test.xml");
                    //pageTransformator.RegisterObserver(new MarkdownObserver(folder: "c:\\temp", includeVerbose: true));
                    pageTransformator.RegisterObserver(new UnitTestLogObserver());


                    var pages = sourceClientContext.Web.GetPagesFromList("Pages", folder: "News", pageNameStartsWith: "Our-new-IT-suite-is-mint");

                    pages.FailTestIfZero();

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

                            // Term store mapping
                            TermMappingFile = @"..\..\Transform\Mapping\term_mapping_sample.csv",

                            SkipTermStoreMapping = true

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

        [TestMethod]
        public void BasicOnlineSitePage_TermTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOPublishingSite")))
                {
                    var pageTransformator = new PublishingPageTransformator(sourceClientContext, targetClientContext, @"C:\temp\onprem-mapping-all-test.xml");
                    //pageTransformator.RegisterObserver(new MarkdownObserver(folder: "c:\\temp", includeVerbose: true));
                    pageTransformator.RegisterObserver(new UnitTestLogObserver());


                    var pages = sourceClientContext.Web.GetPagesFromList("Pages", folder: "News", pageNameStartsWith: "Kitchen");

                    pages.FailTestIfZero();

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

                            // Term store mapping
                            TermMappingFile = @"..\..\Transform\Mapping\term_mapping_sample.csv",

                            SkipTermStoreMapping = false

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

        [TestMethod]
        public void CacheTermStoreSiteCollectionByIdTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevSiteUrl")))
                {
                    TermTransformator termTransformator = new TermTransformator(null, sourceClientContext, targetClientContext);

                    // Target - Katchup - dc2cca62-cf9a-4a61-b672-4aa7d0b6a179
                    // Source - Categories - e757dcf5-f443-42e9-98c6-5842861099cb (site collection term set)
                    termTransformator.CacheTermsFromTermStore(
                        new Guid("e757dcf5-f443-42e9-98c6-5842861099cb"),
                        new Guid("dc2cca62-cf9a-4a61-b672-4aa7d0b6a179"));

                    // Need to have the term store populated values
                    var result = Cache.CacheManager.Instance.GetTransformTermCacheTermById(sourceClientContext, new Guid("ac625b0a-0459-4d23-bc96-0970abd1029d"));
                    var expectedLabel = "Announcements";

                    Assert.AreEqual(expectedLabel, result.TermLabel);
                }
            }
        }

        [TestMethod]
        public void CacheTermStoreSiteCollectionByNameTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevSiteUrl")))
                {
                    TermTransformator termTransformator = new TermTransformator(null, sourceClientContext, targetClientContext);

                    // Target - Katchup - dc2cca62-cf9a-4a61-b672-4aa7d0b6a179
                    // Source - Categories - e757dcf5-f443-42e9-98c6-5842861099cb (site collection term set)
                    termTransformator.CacheTermsFromTermStore(
                        new Guid("e757dcf5-f443-42e9-98c6-5842861099cb"),
                        new Guid("dc2cca62-cf9a-4a61-b672-4aa7d0b6a179"));

                    // Need to have the term store populated values
                    var expectedLabel = "Announcements";
                    var result = Cache.CacheManager.Instance.GetTransformTermCacheTermByName(sourceClientContext, expectedLabel);

                    Assert.IsTrue(result.Count > 0);

                    result.ForEach(o => Console.WriteLine("Cached Term: {0} {1} ", o.TermSetId, o.TermPath, o.TermLabel));
                                       
                }
            }
        }

        [TestMethod]
        public void CacheTermStoreTenantStoreTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevSiteUrl")))
                {
                    TermTransformator termTransformator = new TermTransformator(null, sourceClientContext, targetClientContext);

                    // Target - Katchup - dc2cca62-cf9a-4a61-b672-4aa7d0b6a179
                    // Source - Categories - e757dcf5-f443-42e9-98c6-5842861099cb (site collection term set)
                    termTransformator.CacheTermsFromTermStore(
                        new Guid("e757dcf5-f443-42e9-98c6-5842861099cb"),
                        new Guid("dc2cca62-cf9a-4a61-b672-4aa7d0b6a179"));

                    // Need to have the term store populated values
                    var result = Cache.CacheManager.Instance.GetTransformTermCacheTermById(sourceClientContext, new Guid("c9cbb11b-77ed-4890-ae24-ee103002c46b"));
                    var expectedLabel = "PnPTransform";

                    Assert.AreEqual(expectedLabel, result.TermLabel);
                }
            }
        }

        [TestMethod]
        public void GetTermSetPathsTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevSiteUrl")))
                {
                    TermTransformator termTransformator = new TermTransformator(null, sourceClientContext, targetClientContext);

                    // Target - Katchup - dc2cca62-cf9a-4a61-b672-4aa7d0b6a179
                    // Source - Categories - e757dcf5-f443-42e9-98c6-5842861099cb (site collection term set)
                    termTransformator.CacheTermsFromTermStore(
                        new Guid("e757dcf5-f443-42e9-98c6-5842861099cb"),
                        new Guid("dc2cca62-cf9a-4a61-b672-4aa7d0b6a179"));

                    var results = TermTransformator.GetAllTermsFromTermSet(new Guid("e757dcf5-f443-42e9-98c6-5842861099cb"), sourceClientContext);
                    foreach(var result in results)
                    {
                        Console.WriteLine($"ID: {result.Key} {result.Value.TermPath}");
                    }

                    Assert.IsTrue(results.Count > 0); //Super simple
                }
            }
        }


        [TestMethod]
        public void ValidateTermById_PositiveTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevSiteUrl")))
                {
                    TermTransformator termTransformator = new TermTransformator(null, sourceClientContext, targetClientContext);

                    // Target - Katchup - dc2cca62-cf9a-4a61-b672-4aa7d0b6a179
                    // Source - Categories - e757dcf5-f443-42e9-98c6-5842861099cb (site collection term set)
                    termTransformator.CacheTermsFromTermStore(
                        new Guid("e757dcf5-f443-42e9-98c6-5842861099cb"),
                        new Guid("dc2cca62-cf9a-4a61-b672-4aa7d0b6a179"));

                    // Need to have the term store populated values
                    var result = termTransformator.ResolveTermInCache(sourceClientContext, new Guid("ac625b0a-0459-4d23-bc96-0970abd1029d"));
                    var expectedLabel = "Announcements";

                    Assert.AreEqual(expectedLabel, result.TermLabel);
                }
            }
        }

        [TestMethod]
        public void ValidateTermById_NegativeTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevSiteUrl")))
                {
                    TermTransformator termTransformator = new TermTransformator(null, sourceClientContext, targetClientContext);

                    // Target - Katchup - dc2cca62-cf9a-4a61-b672-4aa7d0b6a179
                    // Source - Categories - e757dcf5-f443-42e9-98c6-5842861099cb (site collection term set)
                    termTransformator.CacheTermsFromTermStore(
                        new Guid("e757dcf5-f443-42e9-98c6-5842861099cb"),
                        new Guid("dc2cca62-cf9a-4a61-b672-4aa7d0b6a179"));

                    // Need to have the term store populated values
                    // Announcements
                    var result = termTransformator.ResolveTermInCache(sourceClientContext, new Guid("11111111-2222-3333-4444-0970abd1029d"));
                    
                    Assert.IsTrue(result == default);
                }
            }
        }
    }
}
