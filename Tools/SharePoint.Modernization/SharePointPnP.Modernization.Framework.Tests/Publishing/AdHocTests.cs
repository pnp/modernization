using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SharePointPnP.Modernization.Framework.Publishing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Tests.Publishing
{
    [TestClass]
    public class AdHocTests
    {

        [TestMethod]
        public void TestMethod1()
        {
            using (ClientContext cc = TestCommon.CreateClientContext())
            {
                PageLayoutManager m = new PageLayoutManager(cc, null);
                var result = m.LoadPageLayoutMappingFile(@"..\..\..\SharePointPnP.Modernization.Framework\Publishing\pagelayoutmapping_sample.xml");
            }
        }

    }
}
