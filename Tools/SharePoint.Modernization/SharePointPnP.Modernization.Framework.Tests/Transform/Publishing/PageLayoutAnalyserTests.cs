using System;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SharePointPnP.Modernization.Framework.Publishing;

namespace SharePointPnP.Modernization.Framework.Tests.Transform.Publishing
{
    [TestClass]
    public class PageLayoutAnalyserTests
    {
        [TestMethod]
        public void PageLayoutAnalyse_SimpleReadOutput()
        {
            using (var sourceClientContext = TestCommon.CreateClientContext())
            {
                var pageLayoutAnalyser = new PageLayoutAnalyser(sourceClientContext);
                pageLayoutAnalyser.RegisterObserver(new UnitTestLogObserver());
                                
                //This will need option for target output location
                var result = pageLayoutAnalyser.GenerateMappingFile();
                Console.WriteLine("Mapping file: {0}", result);

                Assert.IsNotNull(result);

            }
        }

        [TestMethod]
        public void PageLayoutAnalyse_GetPageLayouts()
        {
            using (var sourceClientContext = TestCommon.CreateClientContext())
            {
                var pageLayoutAnalyser = new PageLayoutAnalyser(sourceClientContext);
                pageLayoutAnalyser.RegisterObserver(new UnitTestLogObserver());

                var result = pageLayoutAnalyser.GetAllPageLayouts();


                //This will need option for target output location
                Assert.IsNotNull(result);
                Assert.IsTrue(result.Count > 0);

            }
        }

        [TestMethod]
        public void PageLayoutAnalyse_AnalyseAllWithOutput()
        {
            using (var sourceClientContext = TestCommon.CreateClientContext())
            {
                var pageLayoutAnalyser = new PageLayoutAnalyser(sourceClientContext);
                pageLayoutAnalyser.RegisterObserver(new UnitTestLogObserver());

                pageLayoutAnalyser.AnalyseAll();
                var result = pageLayoutAnalyser.GenerateMappingFile();

                //This will need option for target output location
                Assert.IsNotNull(result);
                
            }
        }

        [TestMethod]
        public void PageLayoutAnalyse_AnalyseSingleWithOutput()
        {
            using (var sourceClientContext = TestCommon.CreateClientContext())
            {
                // Source Context could be a site collection
                ClientContext contextToUse;
                if (sourceClientContext.Web.IsSubSite())
                {
                    string siteCollectionUrl = sourceClientContext.Site.EnsureProperty(o => o.Url);
                    contextToUse = sourceClientContext.Clone(siteCollectionUrl);
                }
                else
                {
                    contextToUse = sourceClientContext;
                }

                var pageLayoutAnalyser = new PageLayoutAnalyser(sourceClientContext);
                pageLayoutAnalyser.RegisterObserver(new UnitTestLogObserver());

                var layout = contextToUse.Web.GetFileByServerRelativeUrl("/sites/PnPTransformationSource/_catalogs/masterpage/ArticleCustom.aspx");

                var result = string.Empty;
                if(layout!= null){
                    ListItem item = layout.EnsureProperty(o=> o.ListItemAllFields);
                    
                    pageLayoutAnalyser.AnalysePageLayout(item);
                    result = pageLayoutAnalyser.GenerateMappingFile();
                }



                Assert.IsTrue(result != string.Empty);

            }
        }

        [TestMethod]
        public void PageLayoutAnalyse_GetListOfWebPartsAssemblyReference()
        {
            var result = WebParts.GetListOfWebParts();

            result.ForEach(o =>
            {
                Console.WriteLine(o);
            });

            Assert.IsNotNull(result);
            Assert.IsTrue(result.Count > 0);
        }

    }
}
