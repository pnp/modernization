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

                Assert.IsNotNull(result);

            }
        }
    }
}
