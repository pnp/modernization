using Microsoft.VisualStudio.TestTools.UnitTesting;
using SharePointPnP.Modernization.Framework.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Tests.Transform
{
    [TestClass]
    public class MappingTests
    {
        [TestMethod]
        public void UrlMappingFileLoadTest()
        {
            FileManager fm = new FileManager();
            var mapping = fm.LoadUrlMappingFile(@"..\..\Transform\Mapping\urlmapping_sample.csv");

            Assert.IsTrue(mapping.Count > 0);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void UrlMappingFileNotFoundTest()
        {
            FileManager fm = new FileManager();
            var mapping = fm.LoadUrlMappingFile(@"..\..\Transform\Mapping\idontexist_sample.csv");
        }

    }
}
