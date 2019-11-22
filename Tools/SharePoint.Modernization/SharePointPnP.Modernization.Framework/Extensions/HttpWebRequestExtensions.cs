using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;

namespace System.Net
{
    public static class HttpWebRequestExtensions
    {

        public static void AddAuthenticationData(this HttpWebRequest httpWebRequest, ClientContext cc)
        {

            if (cc.Credentials != null)
            {
                httpWebRequest.Credentials = cc.Credentials;
            }
            else
            {
                httpWebRequest.CookieContainer = new CookieManager().GetCookies(cc);
            }            

        }

        private static EventHandler<WebRequestEventArgs> CollectCookiesHandler(CookieContainer authCookies)
        {
            return (s, e) =>
            {
                if (authCookies == null || (authCookies != null && authCookies.Count == 0))
                {
                    authCookies = CopyContainer(e.WebRequestExecutor.WebRequest.CookieContainer);
                }
            };
        }

        private static CookieContainer CopyContainer(CookieContainer container)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                BinaryFormatter formatter = new BinaryFormatter();
                formatter.Serialize(stream, container);
                stream.Seek(0, SeekOrigin.Begin);
                return (CookieContainer)formatter.Deserialize(stream);
            }
        }
    }
}
