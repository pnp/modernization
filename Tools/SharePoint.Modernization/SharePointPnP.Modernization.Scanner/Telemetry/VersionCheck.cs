using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SharePoint.Modernization.Scanner.Telemetry
{
    public class VersionCheck
    {
        public const string versionFileUrl = "https://raw.githubusercontent.com/SharePoint/sp-dev-modernization/dev/Tools/SharePoint.Modernization/Releases/version.txt";
        public const string newVersionDownloadUrl = "https://aka.ms/sppnp-modernizationscanner";

        public static Tuple<string, string> LatestVersion()
        {
            string latestVersion = "";
            string currentVersion = "";

            try
            {
                var coreAssembly = Assembly.GetExecutingAssembly();
                currentVersion = ((AssemblyFileVersionAttribute)coreAssembly.GetCustomAttribute(typeof(AssemblyFileVersionAttribute))).Version;

                using (var wc = new System.Net.WebClient())
                {
                    Random random = new Random();
                    latestVersion = wc.DownloadString(versionFileUrl + "?random=" + random.Next().ToString());
                }

                if (!string.IsNullOrEmpty(latestVersion))
                {
                    latestVersion = latestVersion.Replace("\\r", "").Replace("\\t", "");

                    var versionOld = new Version(currentVersion);
                    if (Version.TryParse(latestVersion, out Version versionNew))
                    {
                        if (versionOld.CompareTo(versionNew) >= 0)
                        {
                            // version is not newer
                            latestVersion = null;
                        }
                    }
                    else
                    {
                        // We could not get the version file
                        latestVersion = null;
                    }
                }

            }
            catch(Exception ex)
            {
                // Something went wrong
                latestVersion = null;
            }

            return new Tuple<string, string>(currentVersion, latestVersion);
        }

    }
}
