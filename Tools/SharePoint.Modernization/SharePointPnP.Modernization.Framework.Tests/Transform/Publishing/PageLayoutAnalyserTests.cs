using System;
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

                var result = pageLayoutAnalyser.GetPageLayouts();


                //This will need option for target output location
                Assert.IsNotNull(result);
                Assert.IsTrue(result.Count > 0);

            }
        }

        [TestMethod]
        public void PageLayoutAnalyse_AnalyseWithOutput()
        {
            using (var sourceClientContext = TestCommon.CreateClientContext())
            {
                var pageLayoutAnalyser = new PageLayoutAnalyser(sourceClientContext);
                pageLayoutAnalyser.RegisterObserver(new UnitTestLogObserver());

                pageLayoutAnalyser.Analyse();
                var result = pageLayoutAnalyser.GenerateMappingFile();

                //This will need option for target output location
                Assert.IsNotNull(result);
                
            }
        }

    }
}
