using Microsoft.VisualStudio.TestTools.UnitTesting;
using SharePointPnP.Modernization.Framework.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Tests.Transform.Mapping
{
    [TestClass]
    public class UserMappingTests
    {

        [TestMethod]
        public void UrlMappingFileLoadTest()
        {
            FileManager fm = new FileManager();
            var mapping = fm.LoadUserMappingFile(@"..\..\Transform\Mapping\urlmapping_sample.csv");

            Assert.IsTrue(mapping.Count > 0);
        }

    }
}
