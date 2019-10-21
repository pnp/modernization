using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Framework.TimerJobs.Enums;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;

namespace SharePoint.Modernization.Scanner.Utilities
{
    internal class SiteInformation
    {
        internal string SiteUrl { get; set; }
        internal string Title { get; set; }
        internal bool? AllowGuestUserSignIn { get; set; }
        internal bool? ExternalSharing { get; set; }
        internal bool? ShareByEmailEnabled { get; set; }
        internal bool? ShareByLinkEnabled { get; set; }
        internal DateTime LastActivityOn { get; set; }
        internal int PageViews { get; set; }
        internal int PagesVisited { get; set; }
    }

    /// <summary>
    /// Class used to detect Sites.Read.All permissions and deal with the consequences of that
    /// </summary>
    internal class AppOnlyManager
    {
        private static readonly string SitesInformationListUrl = "DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO";
        private static readonly string SitesInformationListAllUrl = "DO_NOT_DELETE_SPLIST_TENANTADMIN_ALL_SITES_AGGREGA";
        private static readonly string SitesListAllQuery = @"<View Scope=""RecursiveAll"">
                                                            <Query>
                                                               <Where>
                                                                  <IsNull>
                                                                     <FieldRef Name='TimeDeleted' />
                                                                  </IsNull>
                                                               </Where>
                                                               <OrderBy>
                                                                  <FieldRef Name='SiteUrl' Ascending='False' />
                                                               </OrderBy>
                                                            </Query>
                                                            <ViewFields>
                                                               <FieldRef Name='SiteUrl' />
                                                               <FieldRef Name='TemplateName' />
                                                            </ViewFields>
                                                            <RowLimit Paged=""TRUE"">1000</RowLimit>
                                                          </View>";
        private static readonly string SitesListQuery = @"<View Scope=""RecursiveAll"">
                                                            <Query>
                                                               <Where>
                                                                  <IsNull>
                                                                     <FieldRef Name='TimeDeleted' />
                                                                  </IsNull>
                                                               </Where>
                                                               <OrderBy>
                                                                  <FieldRef Name='SiteUrl' Ascending='False' />
                                                               </OrderBy>
                                                            </Query>
                                                            <ViewFields>
                                                               <FieldRef Name='SiteUrl' />
                                                               <FieldRef Name='Title' />
                                                               <FieldRef Name='AllowGuestUserSignIn' />
                                                               <FieldRef Name='ExternalSharing' />
                                                               <FieldRef Name='ShareByEmailEnabled' />
                                                               <FieldRef Name='ShareByLinkEnabled' />
                                                               <FieldRef Name='LastActivityOn' />
                                                               <FieldRef Name='PageViews' />
                                                               <FieldRef Name='PagesVisited' />
                                                            </ViewFields>
                                                            <RowLimit Paged=""TRUE"">1000</RowLimit>
                                                          </View>";

        #region Construction
        internal AppOnlyManager()
        {
            this.SiteInformation = new List<SiteInformation>();
        }
        #endregion

        #region Properties
        internal List<SiteInformation> SiteInformation { get; }
        #endregion

        /// <summary>
        /// Fetch list of site collections based upon enumerating lists in tenant admin
        /// </summary>
        /// <param name="tenantAdminClientContext">Tenant admin client context</param>
        /// <param name="addedSites">Provides list of (wildcard) sites to resolve</param>
        /// <param name="excludeOD4B">Skip OD4B sites</param>
        /// <returns>List of resolved site collections</returns>
        internal List<string> ResolveSitesWithoutFullControl(ClientContext tenantAdminClientContext, List<string> addedSites, bool excludeOD4B)
        {
            List<string> enumeratedSites = new List<string>();
            List<string> resolvedSites = new List<string>();

            // Populate general list with site information, note that this list will not contain the personal sites
            this.LoadSites(tenantAdminClientContext);

            foreach (var site in addedSites)
            {
                if (site.Contains("*"))
                {
                    if (enumeratedSites.Count == 0)
                    {
                        this.LoadAllSites(tenantAdminClientContext, enumeratedSites, excludeOD4B);
                    }

                    string searchString = site.Substring(0, site.IndexOf("*")).ToLower();

                    foreach(var enumeratedSite in enumeratedSites)
                    {
                        if (enumeratedSite.Contains(searchString))
                        {
                            if (!resolvedSites.Contains(enumeratedSite))
                            {
                                resolvedSites.Add(enumeratedSite);
                            }
                        }
                    }
                }
                else
                {
                    resolvedSites.Add(site);
                }
            }

            return resolvedSites;
        }

        /// <summary>
        /// Verifies if the Azure AD app-only authentication is configured with Sites.FullControl.All
        /// </summary>
        /// <param name="options">Scanner options</param>
        /// <param name="sites">Resolved list of site collections</param>
        /// <returns>True if we're running under with the Sites.FullControl.All role</returns>
        internal bool AppOnlyTokenHasFullControl(Options options, List<string> sites)
        {
            // Skip if not app-only
            if (options.AuthenticationTypeProvided() == AuthenticationType.Office365 || options.AuthenticationTypeProvided() == AuthenticationType.NetworkCredentials || options.AuthenticationTypeProvided() == AuthenticationType.AppOnly)
            {
                return true;
            }

            // get a valid site url from the sites list
            string url = GetSiteUrl(sites);

            if (string.IsNullOrEmpty(url))
            {
                return true;
            }

            if (options.AuthenticationTypeProvided() == AuthenticationType.AzureADAppOnly)
            {
                string roleToCheck = "Sites.FullControl.All";
                string accessToken = GetAzureADAppOnlyToken(options, url);

                if (!string.IsNullOrEmpty(accessToken))
                {
                    var handler = new JwtSecurityTokenHandler();
                    var token = handler.ReadJwtToken(accessToken);

                    if (token != null)
                    {
                        if (token.Payload.ContainsKey("roles"))
                        {
                            if (token.Payload["roles"].ToString().IndexOf(roleToCheck, StringComparison.InvariantCultureIgnoreCase) > 0)
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            return false;
        }

        #region Helper methods
        private void LoadSites(ClientContext tenantAdminClientContext)
        {
            tenantAdminClientContext.Web.EnsureProperty(p => p.Url);
            var sitesList = tenantAdminClientContext.Web.GetList($"{tenantAdminClientContext.Web.Url}/Lists/{SitesInformationListUrl}");
            tenantAdminClientContext.ExecuteQueryRetry();

            // Query the list to obtain the sites to return
            CamlQuery camlQuery = new CamlQuery
            {
                ViewXml = SitesListQuery
            };

            do
            {
                var sites = sitesList.GetItems(camlQuery);
                sitesList.Context.Load(sites, i => i.IncludeWithDefaultProperties(li => li.FieldValuesAsText), i => i.ListItemCollectionPosition);
                sitesList.Context.ExecuteQueryRetry();
                foreach (var site in sites)
                {
                    if (this.SiteInformation.Where(p => p.SiteUrl.Equals(site["SiteUrl"].ToString())).FirstOrDefault() == null)
                    {
                        DateTime.TryParse(site["LastActivityOn"]?.ToString(), out DateTime lastActivityOn);

                        var siteInfo = new SiteInformation()
                        {
                            SiteUrl = site["SiteUrl"].ToString(),
                            Title = site["Title"] != null ? site["Title"].ToString() : "",
                            PagesVisited = site["PagesVisited"] != null ? int.Parse(site["PagesVisited"].ToString()) : 0,
                            PageViews = site["PageViews"] != null ? int.Parse(site["PageViews"].ToString()) : 0,
                            LastActivityOn = lastActivityOn
                        };

                        if (site["AllowGuestUserSignIn"] != null)
                        {
                            siteInfo.AllowGuestUserSignIn = bool.Parse(site["AllowGuestUserSignIn"].ToString());
                        }

                        if (site["ExternalSharing"] != null)
                        {
                            if (site["ExternalSharing"].ToString().Equals("On", StringComparison.InvariantCultureIgnoreCase))
                            {
                                siteInfo.ExternalSharing = true;
                            }
                            else
                            {
                                siteInfo.ExternalSharing = false;
                            }
                        }

                        if (site["ShareByEmailEnabled"] != null)
                        {
                            siteInfo.ShareByEmailEnabled = bool.Parse(site["ShareByEmailEnabled"].ToString());
                        }

                        if (site["ShareByLinkEnabled"] != null)
                        {
                            siteInfo.ShareByLinkEnabled = bool.Parse(site["ShareByLinkEnabled"].ToString());
                        }

                        this.SiteInformation.Add(siteInfo);
                    }
                }
                camlQuery.ListItemCollectionPosition = sites.ListItemCollectionPosition;

            } while (camlQuery.ListItemCollectionPosition != null);
        }

        private void LoadAllSites(ClientContext tenantAdminClientContext, List<string> foundSites, bool excludeOD4B)
        {
            tenantAdminClientContext.Web.EnsureProperty(p => p.Url);
            var sitesList = tenantAdminClientContext.Web.GetList($"{tenantAdminClientContext.Web.Url}/Lists/{SitesInformationListAllUrl}");
            tenantAdminClientContext.ExecuteQueryRetry();

            // Query the list to obtain the sites to return
            CamlQuery camlQuery = new CamlQuery
            {
                ViewXml = SitesListAllQuery
            };

            do
            {
                var sites = sitesList.GetItems(camlQuery);
                sitesList.Context.Load(sites, i => i.IncludeWithDefaultProperties(li => li.FieldValuesAsText), i => i.ListItemCollectionPosition);
                sitesList.Context.ExecuteQueryRetry();
                foreach (var site in sites)
                {
                    if (!foundSites.Contains(site["SiteUrl"].ToString().ToLower()))
                    {
                        if (excludeOD4B)
                        {
                            if (site["TemplateName"] != null && site["TemplateName"].ToString().StartsWith("SPSPERS#", StringComparison.InvariantCultureIgnoreCase))
                            {
                                continue;
                            }
                        }

                        foundSites.Add(site["SiteUrl"].ToString().ToLower());
                    }
                }
                camlQuery.ListItemCollectionPosition = sites.ListItemCollectionPosition;

            } while (camlQuery.ListItemCollectionPosition != null);
        }

        private string GetAzureADAppOnlyToken(Options options, string siteUrl, AzureEnvironment environment = AzureEnvironment.Production)
        {
            var certfile = System.IO.File.OpenRead(options.CertificatePfx);
            var certificateBytes = new byte[certfile.Length];
            certfile.Read(certificateBytes, 0, (int)certfile.Length);
            var cert = new X509Certificate2(
                certificateBytes,
                options.CertificatePfxPassword,
                X509KeyStorageFlags.Exportable |
                X509KeyStorageFlags.MachineKeySet |
                X509KeyStorageFlags.PersistKeySet);

            var clientContext = new ClientContext(siteUrl);

            string authority = string.Format(CultureInfo.InvariantCulture, "{0}/{1}/", new OfficeDevPnP.Core.AuthenticationManager().GetAzureADLoginEndPoint(environment), options.AzureTenant);

            var authContext = new AuthenticationContext(authority);

            var clientAssertionCertificate = new ClientAssertionCertificate(options.ClientID, cert);

            var host = new Uri(siteUrl);

            string accessToken = null;

            clientContext.ExecutingWebRequest += (sender, args) =>
            {
                var ar = Task.Run(() => authContext
                    .AcquireTokenAsync(host.Scheme + "://" + host.Host + "/", clientAssertionCertificate))
                    .GetAwaiter().GetResult();
                args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + ar.AccessToken;
                accessToken = ar.AccessToken;
            };

            clientContext.Load(clientContext.Web, p => p.Title);
            clientContext.ExecuteQueryRetry();

            return accessToken;
        }

        private string GetSiteUrl(List<string> sites)
        {
            string siteUrl = null;

            if (sites.Count > 0)
            {
                // grab first url and remove wildcard character if needed
                var url = sites[0].Replace("*", "");

                if (Uri.TryCreate(url, UriKind.Absolute, out Uri siteUri))
                {
                    siteUrl = $"{siteUri.Scheme}://{siteUri.DnsSafeHost}";
                }
            }

            return siteUrl;
        }

        #region not used
        //private string GetAzureACSAppOnlyToken(Options options, string siteUrl, AzureEnvironment environment = AzureEnvironment.Production)
        //{
        //    string accessToken = null;

        //    var am = new AuthenticationManager();
        //    using (var clientContext = am.GetAppOnlyAuthenticatedContext(siteUrl, options.ClientID, options.ClientSecret))
        //    {
        //        clientContext.ExecutingWebRequest += (sender, args) =>
        //        {
        //            var authHeader = args.WebRequestExecutor.RequestHeaders["Authorization"];
        //            accessToken = authHeader.Replace("Bearer", "").Trim();
        //        };

        //        clientContext.Load(clientContext.Web, p => p.Title);
        //        clientContext.ExecuteQueryRetry();
        //    }

        //    return accessToken;
        //}
        #endregion

        #endregion
    }
}
