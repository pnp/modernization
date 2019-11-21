using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace System.Net
{
    public static class HttpWebRequestExtensions
    {

        internal static void AddAuthenticationData(this HttpWebRequest httpWebRequest, ClientContext cc)
        {

            if (cc.Credentials != null)
            {
                httpWebRequest.Credentials = cc.Credentials;
            }
            else
            {
                CookieContainer authCookiesContainer = null;
                EventHandler<WebRequestEventArgs> cookieInterceptorHandler = CollectCookiesHandler(authCookiesContainer);
                try
                {
                    // Hookup a custom handler, assumes the original handler placing the cookies is ran first
                    cc.ExecutingWebRequest += cookieInterceptorHandler;
                    // Trigger the handler to fire by loading something
                    cc.Load(cc.Web, p => p.Url);
                    cc.ExecuteQuery();
                }
                catch(Exception ex)
                {
                    // Eating the exception
                }
                finally
                {
                    // Disconnect the handler as we don't need it anymore
                    cc.ExecutingWebRequest -= cookieInterceptorHandler;
                }

                if (authCookiesContainer != null && authCookiesContainer.Count > 0)
                {
                    httpWebRequest.CookieContainer = authCookiesContainer;
                }
            }            

        }

        private static EventHandler<WebRequestEventArgs> CollectCookiesHandler(CookieContainer authCookies)
        {
            return (s, e) =>
            {
                authCookies = e.WebRequestExecutor.WebRequest.CookieContainer;
            };
        }


    }
}
