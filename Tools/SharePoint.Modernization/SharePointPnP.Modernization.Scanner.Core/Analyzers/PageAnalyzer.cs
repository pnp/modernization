using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client;
using SharePoint.Modernization.Scanner.Core.Results;
using SharePointPnP.Modernization.Framework.Entities;

namespace SharePoint.Modernization.Scanner.Core.Analyzers
{
    /// <summary>
    /// Analyses a page
    /// </summary>
    public class PageAnalyzer : BaseAnalyzer
    {
        #region Variables
        private static readonly Guid FeatureId_Web_HomePage = new Guid("F478D140-B148-4038-9CB0-84A8F1E4BE09");
        // Feature indicating the site was group connected (groupified)
        private static readonly Guid FeatureId_Web_GroupHomepage = new Guid("E3DC7334-CEC0-4D2C-8B90-E4857698FC4E");
        // Root site homepage with XsltListViewWebPart and GettingStartedWebPart
        private static readonly string DefaultRootHtml = "<divclass=\"<tableid=\"layoutsTable\"style=\"width&#58;100%;\"><tbody><trstyle=\"vertical-align&#58;top;\"><tdstyle=\"width&#58;100%;padding&#58;0px;\"><divclass=\"ms-rte-layoutszone-outer\"style=\"width&#58;100%;\"><divclass=\"ms-rte-layoutszone-inner\"style=\"min-height&#58;60px;word-wrap&#58;break-word;margin&#58;0px;border&#58;0px;\"><divclass=\"ms-rtestate-readms-rte-wpbox\"><divclass=\"ms-rtestate-read\"id=\"div_\"></div><divclass=\"ms-rtestate-read\"id=\"vid_\"style=\"display&#58;none;\"></div></div><divclass=\"ms-rtestate-readms-rte-wpbox\"><divclass=\"ms-rtestate-read\"id=\"div_\"></div><divclass=\"ms-rtestate-read\"id=\"vid_\"style=\"display&#58;none;\"></div></div></div></div></td></tr></tbody></table><spanid=\"layoutsData\"style=\"display&#58;none;\">false,false,1</span></div>";
        // Homepage with SiteFeedWebPart, XsltListViewWebPart and GettingStartedWebPart
        private static readonly string DefaultHtml = "<divclass=\"<tableid=\"layoutsTable\"style=\"width&#58;100%;\"><tbody><trstyle=\"vertical-align&#58;top;\"><tdcolspan=\"2\"><divclass=\"ms-rte-layoutszone-outer\"style=\"width&#58;100%;\"><divclass=\"ms-rte-layoutszone-inner\"style=\"word-wrap&#58;break-word;margin&#58;0px;border&#58;0px;\"><divclass=\"ms-rtestate-readms-rte-wpbox\"><divclass=\"ms-rtestate-read\"id=\"div_\"></div><divclass=\"ms-rtestate-read\"id=\"vid_\"style=\"display&#58;none;\"></div></div></div></div></td></tr><trstyle=\"vertical-align&#58;top;\"><tdstyle=\"width&#58;49.95%;\"><divclass=\"ms-rte-layoutszone-outer\"style=\"width&#58;100%;\"><divclass=\"ms-rte-layoutszone-inner\"style=\"word-wrap&#58;break-word;margin&#58;0px;border&#58;0px;\"><divclass=\"ms-rtestate-readms-rte-wpbox\"><divclass=\"ms-rtestate-read\"id=\"div_\"></div><divclass=\"ms-rtestate-read\"id=\"vid_\"style=\"display&#58;none;\"></div></div></div></div></td><tdclass=\"ms-wiki-columnSpacing\"style=\"width&#58;49.95%;\"><divclass=\"ms-rte-layoutszone-outer\"style=\"width&#58;100%;\"><divclass=\"ms-rte-layoutszone-inner\"style=\"word-wrap&#58;break-word;margin&#58;0px;border&#58;0px;\"><divclass=\"ms-rtestate-readms-rte-wpbox\"><divclass=\"ms-rtestate-read\"id=\"div_\"></div><divclass=\"ms-rtestate-read\"id=\"vid_\"style=\"display&#58;none;\"></div></div></div></div></td></tr></tbody></table><spanid=\"layoutsData\"style=\"display&#58;none;\">true,false,2</span></div>";
        // Home page web part configurations
        private static readonly string TeamSiteDefaultWebParts = "TeamSiteDefaultWebParts";
        private static readonly string TeamSiteRemovedWebParts = "TeamSiteRemovedWebParts";
        private static readonly string TeamSiteCustomListsOnly = "TeamSiteCustomListsOnly";
        // Fields
        private static readonly string Field_WikiField = "WikiField";
        private static readonly string Field_FileRefField = "FileRef";
        private static readonly string Field_FileLeafRef = "FileLeafRef";

        private const string CAMLQueryByExtension = @"
                <View Scope='Recursive'>
                  <Query>
                    <Where>
                      <Contains>
                        <FieldRef Name='File_x0020_Type'/>
                        <Value Type='text'>aspx</Value>
                      </Contains>
                    </Where>
                  </Query>
                </View>";
        
        private List<Dictionary<string, string>> pageSearchResults;
        #endregion

        #region Construction
        /// <summary>
        /// page analyzer construction
        /// </summary>
        /// <param name="url">Url of the web to be analyzed</param>
        /// <param name="siteColUrl">Url of the site collection hosting this web</param>
        public PageAnalyzer(string url, string siteColUrl, ModernizationScanJob scanJob, List<Dictionary<string, string>> pageSearchResults) : base(url, siteColUrl, scanJob)
        {
            this.pageSearchResults = pageSearchResults;
        }
        #endregion

        /// <summary>
        /// Analyses a page
        /// </summary>
        /// <param name="cc">ClientContext instance used to retrieve page data</param>
        /// <returns>Duration of the page analysis</returns>
        public override TimeSpan Analyze(ClientContext cc)
        {
            try
            {
                base.Analyze(cc);
                Web web = cc.Web;
                cc.Web.EnsureProperties(p => p.WebTemplate, p => p.Configuration, p => p.Features);

                var homePageUrl = web.WelcomePage;
                if (string.IsNullOrEmpty(homePageUrl))
                {
                    // Will be case when the site home page is a web part page
                    homePageUrl = "default.aspx";
                }

                var listsToScan = web.GetListsToScan();
                var sitePagesLibraries = listsToScan.Where(p => p.BaseTemplate == (int)ListTemplateType.WebPageLibrary);

                if (sitePagesLibraries.Count() > 0)
                {
                    foreach (var sitePagesLibrary in sitePagesLibraries)
                    {
                        CamlQuery query = new CamlQuery
                        {
                            ViewXml = CAMLQueryByExtension
                        };
                        var pages = sitePagesLibrary.GetItems(query);
                        web.Context.Load(pages);
                        web.Context.ExecuteQueryRetry();

                        if (pages.FirstOrDefault() != null)
                        {
                            DateTime start;
                            bool forceCheckout = sitePagesLibrary.ForceCheckout;
                            foreach (var page in pages)
                            {
                                string pageUrl = null;
                                try
                                {                                    
                                    if (page.FieldValues.ContainsKey(Field_FileRefField) && !String.IsNullOrEmpty(page[Field_FileRefField].ToString()))
                                    {
                                        pageUrl = page[Field_FileRefField].ToString();
                                    }
                                    else
                                    {
                                        //skip page
                                        continue;
                                    }

                                    start = DateTime.Now;
                                    PageScanResult pageResult = new PageScanResult()
                                    {
                                        SiteColUrl = this.SiteCollectionUrl,
                                        SiteURL = this.SiteUrl,
                                        PageUrl = pageUrl,
                                        Library = sitePagesLibrary.RootFolder.ServerRelativeUrl,
                                    };

                                    // Is this page the web's home page?
                                    if (pageUrl.EndsWith(homePageUrl, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        pageResult.HomePage = true;
                                    }

                                    // Get the type of the page
                                    pageResult.PageType = page.PageType();

                                    // Get page web parts
                                    var pageAnalysis = page.WebParts(this.ScanJob.PageTransformation);
                                    if (pageAnalysis != null)
                                    {
                                        pageResult.Layout = pageAnalysis.Item1.ToString().Replace("Wiki_", "").Replace("WebPart_", "");
                                        pageResult.WebParts = pageAnalysis.Item2;
                                    }

                                    // Determine if this site contains a default "uncustomized" home page
                                    bool isUncustomizedHomePage = false;
                                    try
                                    {
                                        string pageName = "";
                                        if (page.FieldValues.ContainsKey(Field_FileLeafRef) && !String.IsNullOrEmpty(page[Field_FileLeafRef].ToString()))
                                        {
                                            pageName = page[Field_FileLeafRef].ToString();
                                        }

                                        if (pageResult.HomePage && web.WebTemplate == "STS" && web.Configuration == 0 && pageName.Equals("home.aspx", StringComparison.InvariantCultureIgnoreCase))
                                        {
                                            bool homePageModernizationOptedOut = web.Features.Where(f => f.DefinitionId == FeatureId_Web_HomePage).Count() > 0;
                                            if (!homePageModernizationOptedOut)
                                            {
                                                bool siteWasGroupified = web.Features.Where(f => f.DefinitionId == FeatureId_Web_GroupHomepage).Count() > 0;
                                                if (!siteWasGroupified)
                                                {
                                                    var wiki = page.FieldValues[Field_WikiField].ToString();
                                                    if (!string.IsNullOrEmpty(wiki))
                                                    {
                                                        var isHtmlUncustomized = IsHtmlUncustomized(wiki);

                                                        if (isHtmlUncustomized)
                                                        {
                                                            string pageType = GetPageWebPartInfo(pageResult.WebParts);

                                                            if (pageType == TeamSiteDefaultWebParts)
                                                            {
                                                                page.ContentType.EnsureProperty(p => p.DisplayFormTemplateName);
                                                                if (page.ContentType.DisplayFormTemplateName == "WikiEditForm")
                                                                {
                                                                    isUncustomizedHomePage = true;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch(Exception ex)
                                    {
                                        // no point in failing the scan if something goes wrong here
                                    }
                                    finally
                                    {
                                        pageResult.UncustomizedHomePage = isUncustomizedHomePage;
                                    }

                                    // Get page change information
                                    pageResult.ModifiedAt = page.LastModifiedDateTime();
                                    pageResult.ModifiedBy = page.LastModifiedBy();

                                    // Grab this page from the search results to connect view information                                
                                    string fullPageUrl = $"https://{new Uri(this.SiteCollectionUrl).DnsSafeHost}{pageUrl}";
                                    if (pageResult.HomePage)
                                    {
                                        fullPageUrl = this.SiteUrl;
                                    }

                                    if (!this.ScanJob.SkipUsageInformation && this.pageSearchResults != null)
                                    {
                                        var searchPage = this.pageSearchResults.Where(x => x.Values.Contains(fullPageUrl)).FirstOrDefault();
                                        if (searchPage != null)
                                        {
                                            // Recent = last 14 days
                                            pageResult.ViewsRecent = searchPage["ViewsRecent"].ToInt32();
                                            pageResult.ViewsRecentUniqueUsers = searchPage["ViewsRecentUniqueUsers"].ToInt32();
                                            pageResult.ViewsLifeTime = searchPage["ViewsLifeTime"].ToInt32();
                                            pageResult.ViewsLifeTimeUniqueUsers = searchPage["ViewsLifeTimeUniqueUsers"].ToInt32();
                                        }
                                    }

                                    if (!this.ScanJob.PageScanResults.TryAdd(pageResult.PageUrl, pageResult))
                                    {
                                        ScanError error = new ScanError()
                                        {
                                            Error = $"Could not add page scan result for {pageResult.PageUrl}",
                                            SiteColUrl = this.SiteCollectionUrl,
                                            SiteURL = this.SiteUrl,
                                            Field1 = "PageAnalyzer",
                                        };
                                        this.ScanJob.ScanErrors.Push(error);
                                    }
                                    var duration = new TimeSpan((DateTime.Now.Subtract(start).Ticks));
                                }
                                catch(Exception ex)
                                {
                                    ScanError error = new ScanError()
                                    {
                                        Error = ex.Message,
                                        SiteColUrl = this.SiteCollectionUrl,
                                        SiteURL = this.SiteUrl,
                                        Field1 = "MainPageAnalyzerLoop",
                                        Field2 = ex.StackTrace,
                                        Field3 = pageUrl
                                    };

                                    // Send error to telemetry to make scanner better
                                    if (this.ScanJob.ScannerTelemetry != null)
                                    {
                                        this.ScanJob.ScannerTelemetry.LogScanError(ex, error);
                                    }

                                    this.ScanJob.ScanErrors.Push(error);
                                }
                            }
                        }
                    }
                }
            }
            finally
            {
                this.StopTime = DateTime.Now;
            }

            // return the duration of this scan
            return new TimeSpan((this.StopTime.Subtract(this.StartTime).Ticks));
        }

        #region Helper methods
        /// <summary>
        /// Determines if wiki page contains custom html content by removing GUIDs and whitespace from wiki field content and comparing result with default html.
        /// </summary>
        /// <param name="wikiHtml">Html string of wiki field.</param>
        /// <returns></returns>
        private static bool IsHtmlUncustomized(string wikiHtml)
        {
            Regex guidRegex = new Regex("([0-9A-Fa-f]{8}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{12})");
            Regex externalClassRegex = new Regex(@"(ExternalClass.{32}"">)");

            string trimmed = wikiHtml.Replace("\r", "").Replace("\n", "").Replace("\r\n", "").Replace(" ", "").Trim();
            trimmed = guidRegex.Replace(trimmed, "");
            trimmed = externalClassRegex.Replace(trimmed, "");

            return trimmed == DefaultHtml || trimmed == DefaultRootHtml;
        }

        /// <summary>
        /// Returns the category of a page based on the number and types of webparts on the page for the homepage modernization campaign.
        /// </summary>
        /// <param name="webparts">Collection of webparts on the page.</param>
        /// <returns></returns>
        private static string GetPageWebPartInfo(List<WebPartEntity> webparts)
        {
            const string gettingStartedWebPart = "Microsoft.SharePoint.WebPartPages.GettingStartedWebPart";
            const string siteFeedWebPart = "Microsoft.SharePoint.Portal.WebControls.SiteFeedWebPart";
            const string xsltListViewWebPart = "Microsoft.SharePoint.WebPartPages.XsltListViewWebPart";
            const string listViewWebPart = "Microsoft.SharePoint.WebPartPages.ListViewWebPart";
            const string other = "Other";
            string pageType = other;

            Dictionary<string, int> webpartCounts = new Dictionary<string, int>();
            List<string> webpartNames = new List<string>();

            webpartCounts.Add(gettingStartedWebPart, 0);
            webpartCounts.Add(siteFeedWebPart, 0);
            webpartCounts.Add(xsltListViewWebPart, 0);
            webpartCounts.Add(listViewWebPart, 0);
            webpartCounts.Add(other, 0);

            if (webparts != null)
            {
                foreach (var webpart in webparts)
                {
                    string webpartName = webpart.TypeShort();
                    if (!string.IsNullOrEmpty(webpartName))
                    {
                        webpartNames.Add(webpartName);

                        if (webpartCounts.ContainsKey(webpartName))
                        {
                            webpartCounts[webpartName]++;
                        }
                        else
                        {
                            webpartCounts[other]++;
                        }
                    }
                }
            }

            if (webpartCounts[other] == 0)
            {
                if (webpartCounts[siteFeedWebPart] <= 1 && webpartCounts[gettingStartedWebPart] <= 1)
                {
                    if (webpartCounts[listViewWebPart] == 0 && webpartCounts[xsltListViewWebPart] == 0)
                    {
                        pageType = TeamSiteRemovedWebParts;
                    }
                    else if (webpartCounts[xsltListViewWebPart] == 1 && webpartCounts[listViewWebPart] == 0)
                    {
                        pageType = TeamSiteDefaultWebParts;
                    }
                    else if (webpartCounts[xsltListViewWebPart] >= 1 || webpartCounts[listViewWebPart] >= 1)
                    {
                        pageType = TeamSiteCustomListsOnly;
                    }
                }
            }

            return pageType;
        }
        #endregion

    }
}
