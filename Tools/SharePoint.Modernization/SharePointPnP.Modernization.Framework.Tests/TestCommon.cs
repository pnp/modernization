using Microsoft.SharePoint.Client;
using System;
using System.Configuration;
using System.Net;
using System.Security;

namespace SharePointPnP.Modernization.Framework.Tests
{
    static class TestCommon
    {

        #region Defaults

        /// <summary>
        /// Common warning that the test is used to perform the process and not yet automated in checks/validation of results.
        /// </summary>
        public static string InconclusiveNoAutomatedChecksMessage { get { return "Does not yet have automated checks, please manually check the results of the test"; } }

        #endregion
                     
        #region methods
        
        public static ClientContext CreateClientContext()
        {
            return InternalCreateContext(AppSetting("SPODevSiteUrl"));
        }

        public static ClientContext CreateClientContext(string url)
        {
            return InternalCreateContext(url, SourceContextMode.SPO);
        }
        
        public static ClientContext CreateOnPremisesClientContext()
        {
            return InternalCreateContext(AppSetting("SPOnPremDevSiteUrl"), SourceContextMode.OnPremises);
        }

        public static ClientContext CreateOnPremisesClientContext(string url)
        {
            return InternalCreateContext(url, SourceContextMode.OnPremises);
        }

        /// <summary>
        /// SharePoint Online Admin Context
        /// </summary>
        /// <returns></returns>
        public static ClientContext CreateTenantClientContext()
        {
            return InternalCreateContext(AppSetting("SPOTenantUrl"), SourceContextMode.SPO);
        }

        private static ClientContext InternalCreateContext(string contextUrl, SourceContextMode sourceContextMode = SourceContextMode.SPO)
        {
            string siteUrl;
            
            // Read configuration data
            // Trim trailing slashes
            siteUrl = contextUrl.TrimEnd(new[] { '/' });
            
            if (string.IsNullOrEmpty(siteUrl))
            {
                throw new ConfigurationErrorsException("Site Url in App.config are not set up.");
            }

            ClientContext context = new ClientContext(contextUrl);
            context.RequestTimeout = 1000 * 60 * 15;

            if (sourceContextMode == SourceContextMode.SPO)
            {

                if (!string.IsNullOrEmpty(AppSetting("SPOCredentialManagerLabel")))
                {
                    var tempCred = OfficeDevPnP.Core.Utilities.CredentialManager.GetCredential(AppSetting("SPOCredentialManagerLabel"));
                    context.Credentials = new SharePointOnlineCredentials(tempCred.UserName, tempCred.SecurePassword);
                }
                else
                {
                    if (!String.IsNullOrEmpty(AppSetting("SPOUserName")) &&
                        !String.IsNullOrEmpty(AppSetting("SPOPassword")))
                    {
                        context.Credentials = new SharePointOnlineCredentials(AppSetting("SPOUserName"), 
                                GetSecureString(AppSetting("SPOPassword")));

                    }
                    else if (!String.IsNullOrEmpty(AppSetting("AppId")) &&
                             !String.IsNullOrEmpty(AppSetting("AppSecret")))
                    {
                        OfficeDevPnP.Core.AuthenticationManager am = new OfficeDevPnP.Core.AuthenticationManager();
                        context = am.GetAppOnlyAuthenticatedContext(contextUrl, AppSetting("AppId"), AppSetting("AppSecret"));
                    }
                    else
                    {
                        throw new ConfigurationErrorsException("Credentials in App.config are not set up.");
                    }
                }

            }


            if(sourceContextMode == SourceContextMode.OnPremises)
            {

                if (!string.IsNullOrEmpty(AppSetting("SPOnPremCredentialManagerLabel")))
                {
                    var tempCred = OfficeDevPnP.Core.Utilities.CredentialManager.GetCredential(AppSetting("SPOnPremCredentialManagerLabel"));

                    // username in format domain\user means we're testing in on-premises
                    if (tempCred.UserName.IndexOf("\\") > 0)
                    {
                        string[] userParts = tempCred.UserName.Split('\\');
                        context.Credentials = new NetworkCredential(userParts[1], tempCred.SecurePassword, userParts[0]);
                    }
                    else
                    {
                        throw new ConfigurationErrorsException("Credentials in App.config are not set up for on-premises.");
                    }
                }
                else
                {
                    if (!String.IsNullOrEmpty(AppSetting("SPOnPremUserName")) &&
                        !String.IsNullOrEmpty(AppSetting("SPOnPremPassword")))
                    {
                        string[] userParts = AppSetting("SPOnPremUserName").Split('\\');
                        context.Credentials = new NetworkCredential(userParts[1], GetSecureString(AppSetting("SPOnPremPassword")), userParts[0]);
                    }
                    else if (!String.IsNullOrEmpty(AppSetting("SPOnPremAppId")) &&
                             !String.IsNullOrEmpty(AppSetting("SPOnPremAppSecret")))
                    {
                        OfficeDevPnP.Core.AuthenticationManager am = new OfficeDevPnP.Core.AuthenticationManager();
                        context = am.GetAppOnlyAuthenticatedContext(contextUrl, AppSetting("AppId"), AppSetting("AppSecret"));
                    }
                    else
                    {
                        throw new ConfigurationErrorsException("Tenant credentials in App.config are not set up.");
                    }
                }

            }
          
            return context;
        }
        #endregion


        #region Utility

        /// <summary>
        /// Secure Passwords
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        private static SecureString GetSecureString(string input)
        {
            if (string.IsNullOrEmpty(input))
                throw new ArgumentException("Input string is empty and cannot be made into a SecureString", "input");

            var secureString = new SecureString();
            foreach (char c in input.ToCharArray())
                secureString.AppendChar(c);

            return secureString;
        }
        

        /// <summary>
        /// Get Settings from Config files
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public static string AppSetting(string key)
        {
#if !NETSTANDARD2_0
            return ConfigurationManager.AppSettings[key];
#else
            try
            {
                return configuration.AppSettings.Settings[key].Value;
            }
            catch
            {
                return null;
            }
#endif
        }
               
        #endregion

        public enum SourceContextMode
        {
            SPO,
            OnPremises
        }
    }
}
