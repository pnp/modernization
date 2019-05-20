using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SharePointPnP.Modernization.Framework.Publishing;

namespace SharePointPnP.Modernization.Framework.Tests.Transform.Publishing
{
    [TestClass]
    public class AdHocTests
    {

        [TestMethod]
        public void TestMethod1()
        {
            using (ClientContext cc = TestCommon.CreateClientContext())
            {
                PageLayoutManager m = new PageLayoutManager(null);
                var result = m.LoadPageLayoutMappingFile(@"..\..\..\SharePointPnP.Modernization.Framework\Publishing\pagelayoutmapping_sample.xml");
            }
        }

    }
}
