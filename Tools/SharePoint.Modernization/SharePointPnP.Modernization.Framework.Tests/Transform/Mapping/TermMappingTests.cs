using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
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

    }
}
