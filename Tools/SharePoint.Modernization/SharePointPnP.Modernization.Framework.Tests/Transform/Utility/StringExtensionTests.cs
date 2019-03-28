using Microsoft.VisualStudio.TestTools.UnitTesting;
using SharePointPnP.Modernization.Framework.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Tests.Transform.Utility
{
    [TestClass]
    public class StringExtensionTests
    {
        [TestMethod]
        public void StringExtension_ContainsTest()
        {
            var original = "ThisIsAFolderInAPath/Testing";
            var partialText = "thisisafolderinapath/testing";

            var result = original.ContainsIgnoringCasing(partialText);

            Assert.IsTrue(result);
        }

        [TestMethod]
        public void StringExtension_ContainsPartialTest()
        {
            var original = "ThisIsAFolderInAPath/Testing";
            var partialText = "thisisafolderinapath";

            var result = original.ContainsIgnoringCasing(partialText);

            Assert.IsTrue(result);
        }

        [TestMethod]
        public void StringExtension_NotContainsTest()
        {
            var original = "ThisIsAFolderInAPath/Testing";
            var partialText = "somethingcompletelydifferent";

            var result = original.ContainsIgnoringCasing(partialText);

            Assert.IsFalse(result);
        }

        [TestMethod]
        public void StringExtension_EmptyCheckContainsTest()
        {
            var original = "ThisIsAFolderInAPath/Testing";
            var partialText = "";

            var result = original.ContainsIgnoringCasing(partialText);

            Assert.IsTrue(result);
        }

        [TestMethod]
        public void StringExtension_Emptyheck2ContainsTest()
        {
            var original = "";
            var partialText = "somethingcompletelydifferent";

            var result = original.ContainsIgnoringCasing(partialText);

            Assert.IsFalse(result);
        }
    }
}
