using Microsoft.VisualStudio.TestTools.UnitTesting;
using SharePointPnP.Modernization.Framework.Publishing;
using SharePointPnP.Modernization.Framework.Transform;
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
        public void UserMappingFileLoadTest()
        {
            FileManager fm = new FileManager();
            var mapping = fm.LoadUserMappingFile(@"..\..\Transform\Mapping\usermapping_sample.csv");

            Assert.IsTrue(mapping.Count > 0);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void UserMappingFileNotFoundTest()
        {
            FileManager fm = new FileManager();
            var mapping = fm.LoadUrlMappingFile(@"..\..\Transform\Mapping\idontexist_sample.csv");
        }

        [TestMethod]
        public void GetUPNFromAccountTest()
        {


            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateOnPremisesClientContext())
                {
                    PublishingPageTransformationInformation pti = new PublishingPageTransformationInformation()
                    {
                        // If target page exists, then overwrite it
                        Overwrite = true,

                        // Don't log test runs
                        SkipTelemetry = true,

                        //Permissions are should work given cross domain with mapping
                        KeepPageSpecificPermissions = true,

                        // Replace User Mapping
                        UserMappingFile = @"..\..\Transform\Mapping\usermapping_sample.csv"
                    };

                    UserTransformator userTransformator = new UserTransformator(pti, sourceClientContext, targetClientContext, null, false);

                    var result = userTransformator.SearchSourceDomainForUPN("user", "test.user3");
                    Console.WriteLine(result);

                    Assert.IsTrue(!string.IsNullOrEmpty(result));

                }
            }
        }

        [TestMethod]
        public void ResolveDomainFriendlyNameTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateOnPremisesClientContext())
                {
                    PublishingPageTransformationInformation pti = new PublishingPageTransformationInformation()
                    {
                        // If target page exists, then overwrite it
                        Overwrite = true,

                        // Don't log test runs
                        SkipTelemetry = true,

                        //Permissions are should work given cross domain with mapping
                        KeepPageSpecificPermissions = true,

                        // Replace User Mapping
                        UserMappingFile = @"..\..\Transform\Mapping\usermapping_sample.csv"
                    };

                    UserTransformator userTransformator = new UserTransformator(pti, sourceClientContext, targetClientContext, null, false);

                    var result = userTransformator.ResolveFriendlyDomainToLdapDomain("ALPHADELTA");
                    Console.WriteLine(result);

                    Assert.IsTrue(!string.IsNullOrEmpty(result));

                }
            }
        }

        [TestMethod]
        public void GetComputerDomainTest()
        {

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateOnPremisesClientContext())
                {

                    UserTransformator userTransformator = new UserTransformator(null, sourceClientContext, targetClientContext, null, false);

                    var result = userTransformator.GetFriendlyComputerDomain();
                    Console.WriteLine(result);

                    Assert.IsTrue(!string.IsNullOrEmpty(result));

                }
            }


        }

        [TestMethod]
        public void GetLDAPConnectingStringTest()
        {
            UserTransformator userTransformator = new UserTransformator(null, null, null, null, false);

            var result = userTransformator.GetLDAPConnectionString();
            Console.WriteLine(result);

            Assert.IsTrue(!string.IsNullOrEmpty(result));
        }

    }
}
