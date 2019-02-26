using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Transform;
using OfficeDevPnP.Core.Pages;
using SharePointPnP.Modernization.Framework.Pages;
using SharePointPnP.Modernization.Framework.Entities;
using System.Linq;

namespace SharePointPnP.Modernization.Framework.Tests.Transform
{
    [TestClass]
    public class LoggingTests
    {

        [TestMethod]
        public void Error_LoggingTest()
        {

            // Deliberate Error
            var pageTransformator = new PageTransformator(null);
            pageTransformator.RegisterObserver(new UnitTestLogObserver());

            PageTransformationInformation pti = new PageTransformationInformation(null);

            // Should capture a argument exception
            Assert.ThrowsException<ArgumentNullException>(() =>
            {
                pageTransformator.Transform(pti);
            });

        }





    }
}
