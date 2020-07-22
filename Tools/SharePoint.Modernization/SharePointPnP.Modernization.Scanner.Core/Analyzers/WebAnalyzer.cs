using System;
using System.Linq;
using Microsoft.SharePoint.Client;
using SharePoint.Modernization.Scanner.Core.Results;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace SharePoint.Modernization.Scanner.Core.Analyzers
{
    /// <summary>
    /// Web object analyzer
    /// </summary>
    public class WebAnalyzer : BaseAnalyzer
    {
        // Modern page experience - Site feature that needs to be enabled to support modern page creation
        public static readonly Guid FeatureId_Web_ModernPage = new Guid("B6917CB1-93A0-4B97-A84D-7CF49975D4EC");
        // Modern list experience - Web block feature that can be enabled to prevent modern library experience in the web
        public static readonly Guid FeatureId_Web_ModernList = new Guid("52E14B6F-B1BB-4969-B89B-C4FAA56745EF");
        // PublishingWeb SharePoint Server Publishing - Web. Publishing feature will prevent modern pages
        public static readonly Guid FeatureId_Web_Publishing = new Guid("94C94CA6-B32F-4DA9-A9E3-1F3D343D7ECB");
        // Site Page content type
        public static readonly string SitePageContentTypeId = "0x0101009D1CB255DA76424F860D91F20E6C4118";
        // Column indicating the clientside application id feature
        public static readonly string ClientSideApplicationId = "ClientSideApplicationId";
        // Stores page customization information
        public Dictionary<string, CustomizedPageStatus> MasterPageGalleryCustomization = null;
        // OOB master pages
        private static List<string> excludeMasterPage = new List<string>
                                                        {
                                                            "v4.master",
                                                            "minimal.master",
                                                            "seattle.master",
                                                            "oslo.master",
                                                            "default.master",
                                                            "app.master",
                                                            "mwsdefault.master",
                                                            "mwsdefaultv4.master",
                                                            "mwsdefaultv15.master",
                                                            "mysite15.master", // mysite host
                                                            "boston.master" // modern group sites
                                                        };
        // Stores the page search results for all pages in the site collection
        private List<Dictionary<string, string>> pageSearchResults;

        #region Construction
        /// <summary>
        /// Web analyzer construction
        /// </summary>
        /// <param name="url">Url of the web to be analyzed</param>
        /// <param name="siteColUrl">Url of the site collection hosting this web</param>
        /// <param name="scanJob">Job that launched this analyzer</param>
        /// <param name="pageSearchResults">Site collection level search results that were already loaded</param>
        public WebAnalyzer(string url, string siteColUrl, ModernizationScanJob scanJob, List<Dictionary<string, string>> pageSearchResults) : base(url, siteColUrl, scanJob)
        {
            this.pageSearchResults = pageSearchResults;
        }
        #endregion

        /// <summary>
        /// Analyze the web
        /// </summary>
        /// <param name="cc">ClientContext of the web to be analyzed</param>
        /// <returns>Duration of the analysis</returns>
        public override TimeSpan Analyze(ClientContext cc)
        {
            try
            {
                base.Analyze(cc);

                // Ensure needed data is loaded
                Web web = cc.Web;
                web.EnsureProperties(p => p.UserCustomActions, p => p.AlternateCssUrl, p => p.CustomMasterUrl, p => p.MasterUrl, p => p.Features, p => p.WebTemplate, p => p.Configuration, p => p.HasUniqueRoleAssignments, p => p.AllProperties);

                if (cc.Web.WebTemplate.Equals("TEAMCHANNEL", StringComparison.InvariantCultureIgnoreCase))
                {
                    return new TimeSpan((this.StopTime.Subtract(this.StartTime).Ticks));
                }

                // Log in Site scan data that the scanned web is a sub site
                if (web.IsSubSite())
                {
                    SiteScanResult siteScanData;
                    if (this.ScanJob.SiteScanResults.TryGetValue(this.SiteCollectionUrl, out siteScanData))
                    {
                        if (!siteScanData.SubSites)
                        {
                            var clonedSiteScandata = siteScanData.Clone();
                            clonedSiteScandata.SubSites = true;

                            if (!this.ScanJob.SiteScanResults.TryUpdate(this.SiteCollectionUrl, clonedSiteScandata, siteScanData))
                            {
                                ScanError error = new ScanError()
                                {
                                    Error = $"Could not add update site scan result for {this.SiteCollectionUrl} from web scan of {this.SiteUrl}",
                                    SiteColUrl = this.SiteCollectionUrl,
                                    SiteURL = this.SiteUrl,
                                    Field1 = "WebAnalyzer",
                                };
                                this.ScanJob.ScanErrors.Push(error);
                            }
                        }
                    }

                    // Check if we've broken permission inheritance in this site collection
                    if (web.HasUniqueRoleAssignments)
                    {
                        SiteScanResult siteScanData2;
                        if (this.ScanJob.SiteScanResults.TryGetValue(this.SiteCollectionUrl, out siteScanData2))
                        {
                            if (!siteScanData2.SubSitesWithBrokenPermissionInheritance)
                            {
                                var clonedSiteScandata = siteScanData2.Clone();
                                clonedSiteScandata.SubSitesWithBrokenPermissionInheritance = true;

                                if (!this.ScanJob.SiteScanResults.TryUpdate(this.SiteCollectionUrl, clonedSiteScandata, siteScanData2))
                                {
                                    ScanError error = new ScanError()
                                    {
                                        Error = $"Could not add update site scan result for {this.SiteCollectionUrl} from web scan of {this.SiteUrl}",
                                        SiteColUrl = this.SiteCollectionUrl,
                                        SiteURL = this.SiteUrl,
                                        Field1 = "WebAnalyzer",
                                    };
                                    this.ScanJob.ScanErrors.Push(error);
                                }
                            }
                        }
                    }                    
                }

                // Perform specific analysis work
                WebScanResult scanResult = new WebScanResult()
                {
                    SiteColUrl = this.SiteCollectionUrl,
                    SiteURL = this.SiteUrl,
                };

                // Log used web template
                if (web.WebTemplate != null)
                {
                    scanResult.WebTemplate = $"{web.WebTemplate}#{web.Configuration}";
                }

                // Page feature check: users can disable this to prevent modern page creation
                scanResult.ModernPageWebFeatureDisabled = web.Features.Where(f => f.DefinitionId == FeatureId_Web_ModernPage).Count() == 0;
                // List feature check: users can enabled this to prevent modern lists from working
                scanResult.ModernListWebBlockingFeatureEnabled = web.Features.Where(f => f.DefinitionId == FeatureId_Web_ModernList).Count() > 0;
                // Publishing web feature enabled
                scanResult.WebPublishingFeatureEnabled = web.Features.Where(f => f.DefinitionId == FeatureId_Web_Publishing).Count() > 0;

                // Site is using the publishing "Pages" library?
                if (scanResult.WebPublishingFeatureEnabled)
                {
                    if (!(scanResult.WebTemplate.Equals("BICENTERSITE#0", StringComparison.InvariantCultureIgnoreCase) ||
                          scanResult.WebTemplate.Equals("BLANKINTERNET#0", StringComparison.InvariantCulture) ||
                          scanResult.WebTemplate.Equals("SRCHCEN#0", StringComparison.InvariantCulture) ||
                          scanResult.WebTemplate.Equals("CMSPUBLISHING#0", StringComparison.InvariantCulture) ||
                          scanResult.WebTemplate.Equals("SRCHCENTERLITE#0", StringComparison.InvariantCulture) ||
                          scanResult.WebTemplate.Equals("POINTPUBLISHINGHUB#0", StringComparison.InvariantCulture) ||
                          scanResult.WebTemplate.Equals("POINTPUBLISHINGTOPIC#0", StringComparison.InvariantCulture) ||
                          scanResult.WebTemplate.Equals("ENTERWIKI#0", StringComparison.InvariantCulture)))
                    {
                        var pagesLibrary = web.GetListsToScan().Where(p => p.BaseTemplate == 850).FirstOrDefault();
                        if (pagesLibrary != null)
                        {
                            if (pagesLibrary.ItemCount > 0)
                            {
                                scanResult.PublishingPagesLibraryContainsPages = true;
                            }
                        }
                    }
                }

                // Log permission inheritance details
                if (web.IsSubSite() && web.HasUniqueRoleAssignments)
                {
                    scanResult.BrokenPermissionInheritance = web.HasUniqueRoleAssignments;

                    if (!this.ScanJob.SkipUserInformation)
                    {
                        scanResult.Owners = web.GetOwners();
                        scanResult.Members = web.GetMembers();
                        scanResult.Visitors = web.GetVisitors();
                        scanResult.EveryoneClaimsGranted = web.ClaimsHaveRoleAssignment(this.ScanJob.EveryoneClaim, this.ScanJob.EveryoneExceptExternalUsersClaim);
                    }
                }

                // First check for a search site override
                if (web.AllProperties.FieldValues.ContainsKey("SRCH_ENH_FTR_URL_SITE"))
                {
                    scanResult.SearchCenterUrl = web.AllProperties.FieldValues["SRCH_ENH_FTR_URL_SITE"] as string;
                }
                else if (web.AllProperties.FieldValues.ContainsKey("SRCH_ENH_FTR_URL_WEB"))
                {
                    scanResult.SearchCenterUrl = web.AllProperties.FieldValues["SRCH_ENH_FTR_URL_WEB"] as string;
                }

                // If no search site override check for a search results page override
                if (string.IsNullOrEmpty(scanResult.SearchCenterUrl))
                {
                    CheckForCustomSearchExperience(web, scanResult, "SRCH_SB_SET_SITE");
                    if (string.IsNullOrEmpty(scanResult.SearchCenterUrl))
                    {
                        CheckForCustomSearchExperience(web, scanResult, "SRCH_SB_SET_WEB");
                    }
                }

                // If the web template is STS#0, GROUP#0 or SITEPAGEPUBLISHING#0 then the feature was activated by SPO, other templates never got it
                scanResult.ModernPageFeatureWasEnabledBySPO = false;
                if (scanResult.WebTemplate.Equals("STS#0", StringComparison.InvariantCultureIgnoreCase) ||
                    scanResult.WebTemplate.Equals("STS#3", StringComparison.InvariantCultureIgnoreCase) ||
                    scanResult.WebTemplate.Equals("GROUP#0", StringComparison.InvariantCulture) ||
                    scanResult.WebTemplate.Equals("SITEPAGEPUBLISHING#0", StringComparison.InvariantCulture))
                {
                    // Since we did not enable the feature for all STS#0 sites (sites with publishing active did not get it, nor sites having a large number of webpart/wiki pages) we
                    // check if it should have been turned on by checking for the "site page" content type being added to the site pages library                
                    var listsToScan = web.GetListsToScan();
                    var sitePagesLibrary = listsToScan.Where(p => p.BaseTemplate == (int)ListTemplateType.WebPageLibrary).FirstOrDefault();

                    if (sitePagesLibrary != null)
                    {
                        cc.Load(sitePagesLibrary, p => p.ContentTypes.Include(c => c.Id));
                        cc.ExecuteQueryRetry();

                        var sitePageContentTypeFound = sitePagesLibrary.ContentTypes.Where(c => c.Id.StringValue.StartsWith(SitePageContentTypeId, StringComparison.InvariantCultureIgnoreCase)).OrderBy(c => c.Id.StringValue.Length).FirstOrDefault();
                        if (sitePageContentTypeFound != null)
                        {
                            scanResult.ModernPageFeatureWasEnabledBySPO = true;
                        }
                    }
                }

                // Get information about the master pages used
                if (!string.IsNullOrEmpty(web.MasterUrl) && !excludeMasterPage.Contains(web.MasterUrl.Substring(web.MasterUrl.LastIndexOf("/") + 1).ToLower()))
                {
                    scanResult.MasterPage = web.MasterUrl;
                }
                if (!string.IsNullOrEmpty(web.CustomMasterUrl) && !excludeMasterPage.Contains(web.CustomMasterUrl.Substring(web.CustomMasterUrl.LastIndexOf("/") + 1).ToLower()))
                {
                    scanResult.CustomMasterPage = web.CustomMasterUrl;
                }

                if (!string.IsNullOrEmpty(web.AlternateCssUrl))
                {
                    scanResult.AlternateCSS = web.AlternateCssUrl;
                }

                // Get the user custom actions
                scanResult.WebUserCustomActions = web.UserCustomActions.Analyze(this.SiteCollectionUrl, this.SiteUrl);

                // Get home page from the web and check whether it's a modern page or not
                scanResult.ModernHomePage = false;

                var homePageUrl = web.WelcomePage;
                if (string.IsNullOrEmpty(homePageUrl))
                {
                    // Will be case when the site home page is a web part page
                    homePageUrl = "default.aspx";
                }
                var homepageName = System.IO.Path.GetFileName(homePageUrl);

                var sitePagesLibraryForWeb = web.GetListsToScan().Where(p => p.BaseTemplate == (int)ListTemplateType.WebPageLibrary).FirstOrDefault();
                if (sitePagesLibraryForWeb != null && homePageUrl.StartsWith("SitePages", StringComparison.InvariantCultureIgnoreCase))
                {
                    var homePageFile = web.GetFileByServerRelativeUrl($"{(web.ServerRelativeUrl.Length == 1 ? "" : web.ServerRelativeUrl)}/{homePageUrl}");
                    cc.Load(homePageFile, f => f.ListItemAllFields, f => f.Exists);
                    cc.ExecuteQueryRetry();
                    if (homePageFile.Exists)
                    {
                        var item = homePageFile.ListItemAllFields;
                        if (item.FieldValues.ContainsKey(ClientSideApplicationId) && item[ClientSideApplicationId] != null && item[ClientSideApplicationId].ToString().Equals(FeatureId_Web_ModernPage.ToString(), StringComparison.InvariantCultureIgnoreCase))
                        {
                            scanResult.ModernHomePage = true;
                        }
                    }
                }

                // Push information from root web to respective SiteScanResult object
                if (!cc.Web.IsSubSite())
                {
                    SiteScanResult siteScanData;
                    if (this.ScanJob.SiteScanResults.TryGetValue(this.SiteCollectionUrl, out siteScanData))
                    {
                        var clonedSiteScandata = siteScanData.Clone();
                        clonedSiteScandata.WebPublishingFeatureEnabled = scanResult.WebPublishingFeatureEnabled;
                        clonedSiteScandata.ModernPageWebFeatureDisabled = scanResult.ModernPageWebFeatureDisabled;
                        clonedSiteScandata.ModernPageFeatureWasEnabledBySPO = scanResult.ModernPageFeatureWasEnabledBySPO;
                        clonedSiteScandata.ModernListWebBlockingFeatureEnabled = scanResult.ModernListWebBlockingFeatureEnabled;
                        clonedSiteScandata.ModernHomePage = scanResult.ModernHomePage;
                        clonedSiteScandata.MasterPage = (!String.IsNullOrEmpty(scanResult.MasterPage) || !String.IsNullOrEmpty(scanResult.CustomMasterPage));
                        clonedSiteScandata.AlternateCSS = !String.IsNullOrEmpty(scanResult.AlternateCSS);
                        clonedSiteScandata.WebUserCustomActions = scanResult.WebUserCustomActions;

                        if (scanResult.PublishingPagesLibraryContainsPages && clonedSiteScandata.PublishingPagesUsed == false)
                        {
                            clonedSiteScandata.PublishingPagesUsed = true;
                        }

                        if (!this.ScanJob.SiteScanResults.TryUpdate(this.SiteCollectionUrl, clonedSiteScandata, siteScanData))
                        {
                            ScanError error = new ScanError()
                            {
                                Error = $"Could not add update site scan result for {this.SiteCollectionUrl} from web scan of {this.SiteUrl}",
                                SiteColUrl = this.SiteCollectionUrl,
                                SiteURL = this.SiteUrl,
                                Field1 = "WebAnalyzer",
                            };
                            this.ScanJob.ScanErrors.Push(error);
                        }
                    }
                }
                else
                {
                    if (scanResult.PublishingPagesLibraryContainsPages)
                    {
                        if (this.ScanJob.SiteScanResults.TryGetValue(this.SiteCollectionUrl, out SiteScanResult siteScanData))
                        {
                            var clonedSiteScandata = siteScanData.Clone();

                            if (clonedSiteScandata.PublishingPagesUsed == false)
                            {
                                clonedSiteScandata.PublishingPagesUsed = true;
                            }

                            if (!this.ScanJob.SiteScanResults.TryUpdate(this.SiteCollectionUrl, clonedSiteScandata, siteScanData))
                            {
                                ScanError error = new ScanError()
                                {
                                    Error = $"Could not add update site scan result for {this.SiteCollectionUrl} from web scan of {this.SiteUrl}",
                                    SiteColUrl = this.SiteCollectionUrl,
                                    SiteURL = this.SiteUrl,
                                    Field1 = "WebAnalyzer",
                                };
                                this.ScanJob.ScanErrors.Push(error);
                            }
                        }
                    }
                }

                // Persist web results
                if (!this.ScanJob.WebScanResults.TryAdd(this.SiteUrl, scanResult))
                {
                    ScanError error = new ScanError()
                    {
                        Error = $"Could not add web scan result for {this.SiteUrl}",
                        SiteColUrl = this.SiteCollectionUrl,
                        SiteURL = this.SiteUrl,
                        Field1 = "WebAnalyzer",
                    };
                    this.ScanJob.ScanErrors.Push(error);
                }

                if (Options.IncludeLists(this.ScanJob.Mode))
                {
                    // Kickoff the list analyzer
                    var listAnalyzer = new ListAnalyzer(this.SiteUrl, this.SiteCollectionUrl, this.ScanJob);
                    listAnalyzer.Analyze(cc);
                }

                if (Options.IncludePage(this.ScanJob.Mode))
                {
                    // Kickoff the page analysing
                    var pageAnalyzer = new PageAnalyzer(this.SiteUrl, this.SiteCollectionUrl, this.ScanJob, this.pageSearchResults);
                    pageAnalyzer.Analyze(cc);
                }

                if (Options.IncludePublishing(this.ScanJob.Mode))
                {
                    // Kickoff publishing analysis
                    PublishingAnalyzer publishingAnalyzer = null;
                    if (this.ScanJob.SiteScanResults.TryGetValue(this.SiteCollectionUrl, out SiteScanResult siteScanData))
                    {
                        publishingAnalyzer = new PublishingAnalyzer(this.SiteUrl, this.SiteCollectionUrl, this.ScanJob, scanResult, siteScanData);
                    }
                    else
                    {
                        publishingAnalyzer = new PublishingAnalyzer(this.SiteUrl, this.SiteCollectionUrl, this.ScanJob, scanResult, null);
                    }

                    // Assign the current masterpage gallery results
                    publishingAnalyzer.MasterPageGalleryCustomization = this.MasterPageGalleryCustomization;
                    // Run the publishing analyzer
                    publishingAnalyzer.Analyze(cc);

                    // Store the masterpage gallery results
                    this.MasterPageGalleryCustomization = publishingAnalyzer.MasterPageGalleryCustomization;
                }

                if (Options.IncludeInfoPath(this.ScanJob.Mode))
                {
                    // Kick off InfoPath analysis
                    var infoPathAnalyzer = new InfoPathAnalyzer(this.SiteUrl, this.SiteCollectionUrl, this.ScanJob);
                    infoPathAnalyzer.Analyze(cc);
                }

                if (Options.IncludeBlog(this.ScanJob.Mode))
                {
                    // Kick off Blog analysis
                    var blogAnalyzer = new BlogAnalyzer(this.SiteUrl, this.SiteCollectionUrl, this.ScanJob);
                    blogAnalyzer.Analyze(cc);
                }

                // Place workflow as last scan as it's reloading the web.Lists with different properties. The GetListsToScan method will not reload and hence cause missing properties otherwise
                if (Options.IncludeWorkflow(this.ScanJob.Mode) || Options.IncludeWorkflowWithDetails(this.ScanJob.Mode))
                {
                    // Kick off workflow analysis
                    var workflowAnalyzer = new WorkflowAnalyzer(this.SiteUrl, this.SiteCollectionUrl, this.ScanJob);
                    workflowAnalyzer.Analyze(cc);
                }
            }
            finally
            {
                this.StopTime = DateTime.Now;
            }

            // return the duration of this scan
            return new TimeSpan((this.StopTime.Subtract(this.StartTime).Ticks));
        }

        private static void CheckForCustomSearchExperience(Web web, WebScanResult scanResult, string searchCenterUrlProperty)
        {
            if (web.AllProperties.FieldValues.ContainsKey(searchCenterUrlProperty))
            {
                var searchSettingsValue = web.AllProperties.FieldValues[searchCenterUrlProperty] as string;
                if (!string.IsNullOrEmpty(searchSettingsValue))
                {
                    // Convert the settings into a typed object
                    var searchSettings = JsonConvert.DeserializeAnonymousType(searchSettingsValue, new
                    {
                        Inherit = false,
                        ResultsPageAddress = String.Empty,
                        ShowNavigation = false,
                    });

                    if (searchSettings != null && !searchSettings.Inherit)
                    {
                        // Return the search results page URL of the current web
                        scanResult.SearchCenterUrl = searchSettings?.ResultsPageAddress;
                    }
                }
            }
        }

    }
}
