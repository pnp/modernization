using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Extensions
{
    /// <summary>
    /// Extension methods for the client context class
    /// </summary>
    public static class ClientContextExtensions
    {
        /// <summary>
        /// Determine the minimum library version
        /// </summary>
        /// <param name="clientContext">client context</param>
        /// <param name="minimallyRequiredVersion">Required Version</param>
        /// <returns></returns>
        public static bool HasMinimalServerLibraryVersion(this ClientRuntimeContext clientContext, Version minimallyRequiredVersion)
        {
            bool hasMinimalVersion = false;

            try
            {
                clientContext.ExecuteQueryRetry();
                hasMinimalVersion = clientContext.ServerLibraryVersion.CompareTo(minimallyRequiredVersion) >= 0;
            }
            catch (PropertyOrFieldNotInitializedException)
            {
                // swallow the exception.
            }
            try
            {
                Uri urlUri = new Uri(clientContext.Url);
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create($"{urlUri.Scheme}://{urlUri.DnsSafeHost}:{urlUri.Port}/_vti_pvt/service.cnf");
                request.UseDefaultCredentials = true;

                var response = request.GetResponse();

                using (var dataStream = response.GetResponseStream())
                {
                    // Open the stream using a StreamReader for easy access.
                    using (System.IO.StreamReader reader = new System.IO.StreamReader(dataStream))
                    {
                        // Read the content.Will be in this format
                        // vti_encoding:SR|utf8-nl
                        // vti_extenderversion: SR | 15.0.0.4505

                        string version = reader.ReadToEnd().Split('|')[2].Trim();

                        // Only compare the first three digits
                        var compareToVersion = new Version(minimallyRequiredVersion.Major, minimallyRequiredVersion.Minor, minimallyRequiredVersion.Build, 0);
                        hasMinimalVersion = new Version(version.Split('.')[0].ToInt32(), 0, version.Split('.')[3].ToInt32(), 0).CompareTo(compareToVersion) >= 0;
                    }
                }
            }
            catch (WebException ex)
            {
                //TODO Add logging here - swallow for now
            }

            return hasMinimalVersion;
        }

        /// <summary>
        /// Is the connected server SharePoint 2010
        /// </summary>
        /// <param name="clientContext"></param>
        /// <returns></returns>
        public static bool IsSharePoint2010(this ClientRuntimeContext clientContext)
        {
            //Get 2013 version. - TEMP

            try
            {
                // This works in 2010
                clientContext.ExecuteQueryRetry();
                return clientContext.ServerLibraryVersion.CompareTo(Constants.MinimumRequiredVersion_SP2013) >= 0;
            }
            catch (PropertyOrFieldNotInitializedException)
            {
                // swallow the exception.
            }

            return false;
        }
    }
}
